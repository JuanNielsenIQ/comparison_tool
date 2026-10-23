# discover_processing.py

import sys

from discover.reader import (
    load_all_inputs,
)
from discover.preprocessing import (
    build_discover_results
)

workbook_path = sys.argv[1]

(
    periods_df,
    discover_mapping_df,
    discover_df,
    brands_mapping
) = load_all_inputs(

    workbook_path
)

print(periods_df.head())
print(discover_mapping_df.head())
print(discover_df.head())


results = build_discover_results(
    discover_df,
    periods_df,
    discover_mapping_df,
    brands_mapping
)

from discover.reader import load_all_inputs
from discover.preprocessing import build_discover_results
from utils.table_output import write_report

(
    periods_df,
    discover_mapping_df,
    discover_df,
    brands_mapping
) = load_all_inputs(
    workbook_path
)




results = build_discover_results(
    discover_df,
    periods_df,
    discover_mapping_df,
    brands_mapping
)


write_report(
    workbook_path=workbook_path,
    results=results,
    periods_df=periods_df,
    worksheet_name="Executive Summary",
    section_name="DISCOVER",
)
