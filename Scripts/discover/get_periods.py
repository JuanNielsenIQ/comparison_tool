# discover/get_periods.py

import sys
from datetime import datetime, timedelta

import time
import pandas as pd
import win32com.client


def get_periods(discover_path):

    df = pd.read_excel(
        discover_path,
        sheet_name=0,
        header=8,
        engine="calamine",
    )

    periods_col = None

    for col in df.columns:

        if str(col).strip().lower() == "periods":

            periods_col = col
            break

    if periods_col is None:

        raise ValueError(
            "Periods column not found."
        )

    periods = (
        df[periods_col]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    )

    return periods


def harmonize_period(period_text):

    try:

        date_text = str(period_text)[-8:]

        return datetime.strptime(
            date_text,
            "%d/%m/%y",
        )

    except Exception:

        return None


def build_week_year(date_value):

    if date_value is None:
        return None

    week = date_value.isocalendar().week

    return f"{week}-{date_value.year}"


def get_target_workbook(
    excel,
    workbook_path,
):

    for wb in excel.Workbooks:

        if wb.FullName.lower() == workbook_path.lower():

            return wb

    raise Exception(
        f"Workbook not found: {workbook_path}"
    )




def get_sheet(workbook, sheet_name):
    try:
        return workbook.Worksheets(sheet_name)

    except Exception:
        raise Exception(
            f"Worksheet '{sheet_name}' not found."
        )



def write_periods(
    workbook_path,
    periods,
):

    excel = win32com.client.GetObject(
        Class="Excel.Application"
    )

    wb = get_target_workbook(
        excel,
        workbook_path,
    )

    ws = get_sheet(
        wb,
        "Periods Mapping",
    )

    # Clear generated columns only
    ws.Range("A:D").ClearContents()

    # Prevent Excel from converting
    # 1-2024 -> Jan-24
    ws.Columns("C:C").NumberFormat = "@"

    # Headers
    ws.Cells(1, 1).Value = "discover_weeks"
    ws.Cells(1, 2).Value = "periods_harmonized"
    ws.Cells(1, 3).Value = "week_year"
    ws.Cells(1, 4).Value = "end_period"
    ws.Cells(1, 5).Value = "custom_period"

    row = 2

    for period in periods:

        harmonized = harmonize_period(
            period
        )

        week_year = build_week_year(
            harmonized
        )

        ws.Cells(row, 1).Value = period

        if harmonized is not None:

            ws.Cells(row, 2).Value = harmonized
            ws.Cells(row, 2).NumberFormat = "yyyy-mm-dd"

            end_period = harmonized + timedelta(days=1)

            ws.Cells(row, 4).Value = end_period
            ws.Cells(row, 4).NumberFormat = "mm/dd/yyyy"

        if week_year is not None:

            ws.Cells(row, 3).Value = str(week_year)

        # Column E intentionally left blank for custom_period.

        row += 1

    ws.Columns("A:E").AutoFit()


def main(workbook_path):

    discover_path = (
        workbook_path.rsplit("\\", 1)[0]
        + "\\data\\discover.xlsx"
    )

    periods = get_periods(
        discover_path
    )

    write_periods(
        workbook_path,
        periods,
    )

if __name__ == "__main__":

    import traceback
    from pathlib import Path
    import sys

    try:
        main(sys.argv[1])

    except Exception:
        error_text = traceback.format_exc()

        log_file = Path(sys.argv[1]).parent / "python_error.log"

        with open(log_file, "w", encoding="utf-8") as f:
            f.write(error_text)

        raise