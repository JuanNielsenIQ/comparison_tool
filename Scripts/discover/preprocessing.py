# discover/preprocessing.py
import win32com.client

from utils.table_output import get_workbook

def normalize(value):

    if value is None:
        return ""

    return (
        str(value)
        .strip()
        .lower()
        .replace("-", "_")
        .replace(" ", "_")
    )


def get_brand_column(discover_df):

    candidates = [
        "BRAND",
        "SC BRAND",
    ]

    for col in candidates:

        if col in discover_df.columns:
            return col

    raise ValueError(
        "Neither BRAND nor SC BRAND found."
    )

def build_period_lookup(
    workbook_path,
):

    excel = win32com.client.GetObject(
        Class="Excel.Application"
    )

    wb = get_workbook(
        excel,
        workbook_path,
    )

    ws = wb.Worksheets(
        "Periods Mapping"
    )

    mapping = {}

    row = 2

    while True:

        discover_week = ws.Cells(
            row,
            1,  # A
        ).Value

        custom_period = ws.Cells(
            row,
            5,  # E
        ).Value

        if (
            discover_week is None
            and custom_period is None
        ):
            break

        if (
            discover_week is not None
            and custom_period is not None
        ):

            mapping[
                discover_week
            ] = str(
                custom_period
            ).strip()

        row += 1

    return mapping

def build_brand_mapping(brands_mapping_df):

    mapping = {}

    for _, row in brands_mapping_df.iterrows():

        discover_brand = row["Discover"]
        executive_brand = row["Brand"]

        if (
            discover_brand is not None
            and executive_brand is not None
            and str(discover_brand).strip() != ""
            and str(executive_brand).strip() != ""
        ):

            mapping[
                str(discover_brand).strip()
            ] = str(executive_brand).strip()

    return mapping

def build_discover_results(
    discover_df,
    discover_mapping_df,
    brands_mapping_df,
    workbook_path
):

    discover_df = discover_df.copy()

    period_lookup = build_period_lookup(
        workbook_path=workbook_path
    )

    discover_df["custom_period"] = (
        discover_df["Periods"] # TODO: verify that always it is "Periods" column
        .map(period_lookup)
    )

    discover_df = discover_df[
        discover_df["custom_period"]
        .notna()
    ]

    brand_col = get_brand_column(
        discover_df
    )

    brands = sorted(
        discover_df[brand_col]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    results = {}

    for brand in brands:

        results[brand] = {}

    mapping_rows = (
        discover_mapping_df
        .dropna(subset=["variables"])
    )

    for _, mapping in mapping_rows.iterrows():

        measure = str(
            mapping["variables"]
        ).strip()

        discover_column = str(
            mapping["discover_report"]
        ).strip()

        if discover_column == "":
            continue

        if discover_column not in discover_df.columns:

            print(
                f"WARNING: "
                f"{discover_column} "
                f"not found in Discover."
            )

            continue

        for brand in brands:

            if measure not in results[brand]:

                results[brand][measure] = {}

            brand_df = discover_df[
                discover_df[brand_col]
                == brand
            ]

            custom_periods = (
                brand_df["custom_period"]
                .dropna()
                .unique()
            )

            for period in custom_periods:

                value = (
                    brand_df[
                        brand_df["custom_period"]
                        == period
                    ][discover_column]
                    .fillna(0)
                    .sum()
                )

                results[brand][measure][
                    str(period)
                ] = value

    # ---------------------------------
    # Apply Brand Mapping
    # ---------------------------------


    mapped_results = {}
    brand_mapping = build_brand_mapping(
        brands_mapping_df
    )

    mapped_results = {
        brand_mapping.get(brand, brand): measures
        for brand, measures in results.items()
    }

    return mapped_results

