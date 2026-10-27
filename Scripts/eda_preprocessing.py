# eda_preprocessing.py

import sys

from eda.parser import build_results
from eda.preprocessing import build_eda_results
from utils.reader import WorkbookReader
from utils.table_output import write_report


def main():

    if len(sys.argv) < 2:
        raise ValueError(
            "Workbook path not provided."
        )

    workbook_path = sys.argv[1]

    reader = WorkbookReader(workbook_path)

    # Mapping tables
    periods_df, variables_mapping_df, brands_mapping_df = (
        reader.mappings()
    )

    # Raw EDA extraction
    raw_results = build_results(
        reader.source("eda")
    )

    # Transform to generic structure
    results = build_eda_results(
        raw_results,
        variables_mapping_df,
        brands_mapping_df,
        periods_df,
    )

    # Output
    write_report(
        workbook_path=workbook_path,
        results=results,
        periods_df=periods_df,
        worksheet_name="Executive Summary",
        section_name="EDA",
    )


if __name__ == "__main__":
    main()
