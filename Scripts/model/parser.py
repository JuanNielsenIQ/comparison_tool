# model/parser.py

from eda.parser import (
    build_variable_mapping,
    normalize_name,
    resolve_metadata_variable,
    get_project_info,
    get_period_labels,
)

def get_model_columns(
    ws,
    period_labels,
):

    variable_col = None
    period_cols = {}

    expected_periods = {
        str(period).strip().lower(): period
        for period in period_labels
    }

    last_row1 = ""

    for col in range(
        1,
        ws.max_column + 1,
    ):

        row1 = ws.cell(
            1,
            col
        ).value

        row2 = ws.cell(
            2,
            col
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

        #
        # Handle merged headers
        #

        if row1:

            last_row1 = row1

        else:

            row1 = last_row1

        #
        # Variable column
        #

        if row1.lower() == "variable":

            variable_col = col

        #
        # Volume section
        #

        elif row1.lower() == "volume":

            if (
                row2.lower()
                in expected_periods
            ):

                period_cols[
                    expected_periods[
                        row2.lower()
                    ]
                ] = col

    if variable_col is None:

        raise ValueError(
            "Variable column not found."
        )

    missing_periods = [
        period
        for period in period_labels
        if period not in period_cols
    ]

    if missing_periods:

        raise ValueError(
            f"Missing volume columns for periods: "
            f"{missing_periods}"
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
    model_wb,
    variables_df,
    periods_df,
    target_brand,
    competitor_brands
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

    ws = model_wb[
        "Model Result"
    ]

    variable_col, period_cols = (
        get_model_columns(
            ws,
            period_labels,
        )
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
        3,
        ws.max_row + 1,
    ):

        predictor_variable = ws.cell(
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

            value = ws.cell(
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