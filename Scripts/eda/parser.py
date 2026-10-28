
def build_variable_mapping(variables_df):

    variables_df = variables_df.copy()

    variables_df.columns = (
        variables_df.columns
        .astype(str)
        .str.strip()
        .str.lower()
    )

    mapping = {}

    required = [
        "variables",
        "eda_report",
    ]

    variables_df = variables_df.dropna(
        subset=required
    )

    for _, row in variables_df.iterrows():

        summary_variable = (
            str(row["variables"])
            .strip()
        )

        eda_variable = (
            str(row["eda_report"])
            .strip()
            .lower()
        )

        if summary_variable and eda_variable:

            mapping[
                eda_variable
            ] = summary_variable

    return mapping


def normalize_name(value):

    return (
        str(value)
        .strip()
        .lower()
        .replace(" ", "_")
    )


def resolve_metadata_variable(
    predictor_variable,
    brand,
    variable_mapping,
):
    """
    Resolve a predictor variable using the
    metadata mapping.

    Example:

    Mapping:
        brand_tpr

    Brand:
        Coca Cola

    Resolves to:
        coca_cola_tpr
    """

    predictor_variable = normalize_name(
        predictor_variable
    )

    normalized_brand = normalize_name(
        brand
    )

    for (
        eda_variable,
        summary_variable,
    ) in variable_mapping.items():

        resolved_variable = (
            normalize_name(
                eda_variable
            )
            .replace(
                "brand",
                normalized_brand
            )
        )

        if predictor_variable == resolved_variable:

            return summary_variable

    return None


def get_project_info(eda_wb):

    project_ws = eda_wb[
        "project_info"
    ]

    target_brand = str(
        project_ws["I2"].value
    ).strip()

    competitor_brands = []

    row = 2

    while True:

        value = project_ws[
            f"J{row}"
        ].value

        if value in (
            None,
            "",
        ):
            break

        competitor_brands.append(
            str(value).strip()
        )

        row += 1

    return (
        target_brand,
        competitor_brands,
    )


def get_period_labels(periods_df):

    periods = (
        periods_df["custom_period"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .tolist()
    )

    if not periods:

        raise ValueError(
            "At least one period is required."
        )

    return periods

def get_predictor_columns(
    pred_ws,
    period_labels,
):

    variable_col = None
    period_cols = {}

    last_row1 = ""

    expected_periods = {
        str(period).strip().lower()
        for period in period_labels
    }

    for col in range(
        1,
        pred_ws.max_column + 1,
    ):

        row1 = pred_ws.cell(
            1,
            col,
        ).value

        row2 = pred_ws.cell(
            2,
            col,
        ).value

        row1 = (
            ""
            if row1 is None
            else str(row1).strip()
        )

        row2 = (
            ""
            if row2 is None
            else str(row2).strip()
        )

        # Handle merged headers

        if row1:
            last_row1 = row1
        else:
            row1 = last_row1

        if row1.lower() == "variable":

            variable_col = col

        elif row1.lower() == "aggregated":

            period_name = row2.lower()

            if (
                period_name
                in expected_periods
            ):

                period_cols[
                    row2
                ] = col

    if variable_col is None:

        raise ValueError(
            "Variable column not found."
        )

    missing_periods = [
        period
        for period in period_labels
        if period.lower()
        not in {
            key.lower()
            for key in period_cols
        }
    ]

    if missing_periods:

        raise ValueError(
            "Missing Aggregated columns "
            f"for periods: {missing_periods}"
        )

    ordered_period_cols = [
        period_cols[period]
        for period in period_labels
    ]

    return (
        variable_col,
        ordered_period_cols,
    )

def build_results(
    eda_wb,
    variables_df,
    periods_df,
):

    variable_mapping = (
        build_variable_mapping(
            variables_df
        )
    )

    period_labels = (
        get_period_labels(
            periods_df
        )
    )

    target_brand, competitor_brands = (
        get_project_info(
            eda_wb
        )
    )

    pred_ws = eda_wb[
        "Predictors Summary"
    ]

    variable_col, period_cols = (
        get_predictor_columns(
            pred_ws,
            period_labels
        )
    )

    if len(period_cols) != len(
        period_labels
    ):

        raise ValueError(
            f"Found {len(period_cols)} aggregated columns "
            f"but {len(period_labels)} periods were configured."
        )

    all_brands = (
        [target_brand]
        + competitor_brands
    )

    results = {
        brand: {
            variable: {
                period: 0
                for period in period_labels
            }
            for variable in (
                variable_mapping.values()
            )
        }
        for brand in all_brands
    }

    for row in range(
        4,
        pred_ws.max_row + 1,
    ):

        predictor_variable = pred_ws.cell(
            row,
            variable_col,
        ).value

        if predictor_variable is None:
            continue

        values = {}

        for period, col in zip(
            period_labels,
            period_cols,
        ):

            value = pred_ws.cell(
                row,
                col,
            ).value

            values[period] = (
                0
                if value is None
                else value
            )

        for brand in all_brands:

            summary_variable = (
                resolve_metadata_variable(
                    predictor_variable=
                    predictor_variable,
                    brand=brand,
                    variable_mapping=
                    variable_mapping,
                )
            )

            if summary_variable is None:
                continue

            results[brand][
                summary_variable
            ] = values.copy()

    return results
