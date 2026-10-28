# discover_preprocessing.py

import sys

from discover.preprocessing import build_discover_results
from utils.reader import WorkbookReader
from utils.table_output import write_report


def main():

    if len(sys.argv) < 2:
        raise ValueError(
            "Workbook path not provided."
        )

    workbook_path = sys.argv[1]

    reader = WorkbookReader(workbook_path)

    periods_df, variables_mapping_df, brands_mapping_df = (
        reader.mappings()
    )

    discover_df = reader.source("discover")

    results = build_discover_results(
        discover_df=discover_df,
        discover_mapping_df=variables_mapping_df,
        brands_mapping_df=brands_mapping_df,
        workbook_path=workbook_path
    )

    write_report(
        workbook_path=workbook_path,
        results=results,
        periods_df=periods_df,
        worksheet_name="Executive Summary",
        section_name="DISCOVER",
        variables_mapping_df=variables_mapping_df,
        brands_mapping_df=brands_mapping_df,
        is_discover=True
    )


if __name__ == "__main__":
    main()
