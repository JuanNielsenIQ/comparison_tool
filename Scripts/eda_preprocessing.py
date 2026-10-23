# eda_preprocessing.py

import sys

from eda.reader import build_results

from eda.preprocessing import (
    build_eda_results,
)

from discover.reader import (
    get_periods_mapping,
    get_brands_mapping,
    get_variables_mapping,
)

from utils.table_output import (
    write_report,
)


def main():

    if len(sys.argv) < 2:

        raise ValueError(
            "Workbook path not provided."
        )

    workbook_path = sys.argv[1]

    # ---------------------------------
    # Mapping tables
    # ---------------------------------

    periods_df = get_periods_mapping(
        workbook_path
    )

    brands_mapping_df = (
        get_brands_mapping(
            workbook_path
        )
    )

    variables_mapping_df = (
        get_variables_mapping(
            workbook_path
        )
    )

    # ---------------------------------
    # Raw EDA extraction
    # ---------------------------------

    raw_results = build_results(
        workbook_path
    )

    # ---------------------------------
    # Transform to generic structure
    # ---------------------------------
    results = build_eda_results(
        raw_results,
        variables_mapping_df,
        brands_mapping_df,
        periods_df,
    )

    # ---------------------------------
    # Output
    # ---------------------------------

    write_report(
        workbook_path=workbook_path,
        results=results,
        periods_df=periods_df,
        worksheet_name="Executive Summary",
        section_name="EDA",
    )

if __name__ == "__main__":
    main()