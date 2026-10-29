# utils/reader.py
#
# Single place to read everything the tool needs:
#   - mapping sheets from the open Excel workbook (xlwings)
#   - data files stored in <workbook folder>\data\ (one entry per source)
#
# To add a new data source, add one entry to SOURCES. No new reader file.

from dataclasses import dataclass
from pathlib import PureWindowsPath
from typing import Any, Callable

import pandas as pd
import xlwings as xw
from openpyxl import load_workbook


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def dataframe_from_values(values):

    values = [
        row for row in values
        if row[0] is not None
    ]

    if not values:
        return pd.DataFrame()

    return pd.DataFrame(
        values[1:],
        columns=values[0]
    )


# ------------------------------------------------------------------
# Data sources (files inside the "data" folder)
# ------------------------------------------------------------------

@dataclass(frozen=True)
class DataSource:
    filename: str
    loader: Callable[[str], Any]


def _load_discover(path):

    return pd.read_excel(
        path,
        sheet_name=0,
        header=8,
        engine="calamine"
    )


def _load_eda(path):

    return load_workbook(
        path,
        data_only=True
    )


SOURCES = {
    "discover": DataSource("discover.xlsx", _load_discover),
    "eda": DataSource("eda.xlsx", _load_eda),
    "model": DataSource("model.xlsx", _load_eda),
}


# ------------------------------------------------------------------
# Mapping sheets: key -> (sheet name, range, fallback sheet or None)
# ------------------------------------------------------------------

MAPPINGS = {
    "periods": (
        "Periods Mapping (Template)",
        "E1:E200",
        "Periods Mapping",
    ),
    "variables": (
        "Variables Mapping",
        "A1:D50",
        None,
    ),
    "brands": (
        "Brands Mapping",
        "A1:D100",
        None,
    ),
}


# ------------------------------------------------------------------
# Reader
# ------------------------------------------------------------------

class WorkbookReader:

    def __init__(self, workbook_path):

        path = PureWindowsPath(workbook_path)

        self.workbook_path = workbook_path
        self.workbook_name = path.name
        self.data_dir = path.parent / "data"
        self._wb = None

    @property
    def wb(self):

        if self._wb is None:
            self._wb = xw.books[self.workbook_name]

        return self._wb

    # ---- mapping sheets ------------------------------------------

    def mapping(self, key):

        sheet, cell_range, fallback = MAPPINGS[key]

        try:
            ws = self.wb.sheets[sheet]
        except Exception:
            if fallback is None:
                raise
            ws = self.wb.sheets[fallback]
        if sheet == "Periods Mapping (Template)":
            values = ws.range(cell_range).value
            values = [v for v in values if v is not None]
            return pd.DataFrame(
                values[1:],
                columns=[values[0]],
            )

        return dataframe_from_values(
            ws.range(cell_range).value
        )

    def periods(self):
        return self.mapping("periods")

    def variables(self):
        return self.mapping("variables")

    def brands(self):
        return self.mapping("brands")

    def mappings(self):
        """Returns (periods_df, variables_df, brands_df)."""
        return (
            self.periods(),
            self.variables(),
            self.brands(),
        )

    # ---- data files ----------------------------------------------

    def source(self, name):

        if name not in SOURCES:
            raise KeyError(
                f"Unknown source '{name}'. "
                f"Available: {list(SOURCES)}"
            )

        src = SOURCES[name]

        return src.loader(
            str(self.data_dir / src.filename)
        )
