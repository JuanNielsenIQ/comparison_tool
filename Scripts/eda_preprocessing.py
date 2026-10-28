# eda_preprocessing.py

import sys

from eda.parser import build_results
from eda.postprocessing import (
    build_formula_variables,
    add_total_metric,
    add_no_promo_metric,
)
from utils.reader import WorkbookReader
from utils.table_output import write_report


def main():

    if len(sys.argv) < 2:

        raise ValueError(
            "Workbook path not provided."
        )

    workbook_path = sys.argv[1]

    reader = WorkbookReader(
        workbook_path
    )

    #
    # Mapping tables
    #

    periods_df, variables_df, brands_df = (
        reader.mappings()
    )

    #
    # Source workbook
    #

    eda_wb = reader.source(
        "eda"
    )

    #
    # Parse EDA results
    #

    results = build_results(
        eda_wb=eda_wb,
        variables_df=variables_df,
        periods_df=periods_df,
    )

    #
    # Post-process EDA results
    #

    results_formulas = build_formula_variables(
        results=results,
        variables_df=variables_df,
    )

    results_totals = add_total_metric(
        eda_wb=eda_wb,
        results=results_formulas,
        period_labels=periods_df["custom_period"].dropna().astype(str).str.strip().drop_duplicates().tolist()
    )

    results_no_promo = add_no_promo_metric(
        results=results_totals,
        total_variable="TOTAL",
        promo_variable="ANY PROMO",
        output_variable="NO PROMO"
    )


    #
    # Write Executive Summary
    #

    write_report(
        workbook_path=workbook_path,
        results=results_no_promo,
        periods_df=periods_df,
        worksheet_name="Executive Summary",
        section_name="EDA",
        variables_mapping_df=None,
        brands_mapping_df=None,
    )


if __name__ == "__main__":
    main()