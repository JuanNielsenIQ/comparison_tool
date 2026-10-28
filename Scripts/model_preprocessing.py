# eda_preprocessing.py

import sys

from model.duetos_parset import build_duetos_results
from model.parser import build_results
from eda.postprocessing import (
    build_formula_variables,
    add_no_promo_metric,
)
from model.totals import add_total_metric_model
from utils.reader import WorkbookReader
from utils.table_output import write_brand_metric, write_report


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

    target_brand = brands_df.iloc[0, 0]
    competitor_brands = brands_df.iloc[0, 1:].dropna().astype(str).str.strip().tolist()

    #
    # Source workbook
    #

    eda_wb = reader.source(
        "model"
    )

    #
    # Parse EDA results
    #
    results = build_results(
        model_wb=eda_wb,
        variables_df=variables_df,
        periods_df=periods_df,
        target_brand=brands_df['Model'].iloc[0],
        competitor_brands=brands_df['Model'].iloc[1:].dropna().astype(str).tolist()
    )

    results_formulas = build_formula_variables(
        results=results,
        variables_df=variables_df,
    )

    results_totals = add_total_metric_model(
        model_wb=eda_wb,
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
        section_name="Model",
        variables_mapping_df=None,
        brands_mapping_df=None,
    )


    duetos_results = build_duetos_results(
        model_wb=eda_wb,
        target_brand=brands_df['Model'].iloc[0],
        competitor_brands=brands_df['Model'].iloc[1:].dropna().astype(str).tolist(),
        variables_df=variables_df
    )


    write_brand_metric(
        workbook_path=workbook_path,
        results=duetos_results,
        section_name="MODEL",
        metric_name="DUE TOS",
    )


if __name__ == "__main__":
    main()

