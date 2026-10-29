# model/duetos_parser.py

from eda.parser import (
    normalize_name,
)


from eda.parser import (
    build_variable_mapping,
    resolve_metadata_variable,
)

def get_duetos_columns(ws):

    variable_col = None

    duetos_cols = {}

    inside_duetos = False

    for col in range(
        1,
        ws.max_column + 1,
    ):

        row1 = ws.cell(
            1,
            col,
        ).value

        row2 = ws.cell(
            2,
            col,
        ).value

        #
        # Variable column
        #

        if (
            row1 is not None
            and str(row1)
            .strip()
            .lower()
            == "variable"
        ):
            variable_col = col

        #
        # Start Due Tos section
        #

        if (
            row1 is not None
            and str(row1)
            .strip()
            .lower()
            == "due_tos (%)"
        ):

            inside_duetos = True

        #
        # Stop when next row1 header begins
        #

        elif (
            inside_duetos
            and row1 is not None
        ):

            break

        #
        # Capture period headers
        #

        if (
            inside_duetos
            and row2 is not None
        ):

            duetos_cols[
                str(row2).strip()
            ] = col

    if variable_col is None:

        raise ValueError(
            "Variable column not found."
        )

    if not duetos_cols:

        raise ValueError(
            "Due_tos (%) columns not found."
        )

    return (
        variable_col,
        duetos_cols,
    )


def build_duetos_results(
    model_wb,
    target_brand,
    competitor_brands,
    variables_df,
):

    ws = model_wb[
        "Model Result"
    ]

    variable_col, duetos_cols = (
        get_duetos_columns(ws)
    )

    variable_mapping = (
        build_variable_mapping(
            variables_df
        )
    )

    all_brands = (
        [target_brand]
        + competitor_brands
    )

    results = {
        brand: {}
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

        predictor_variable = (
            normalize_name(
                predictor_variable
            )
        )

        period_values = {}

        for (
            period_name,
            col
        ) in duetos_cols.items():

            period_values[
                period_name
            ] = (
                ws.cell(
                    row,
                    col,
                ).value
                or 0
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

            results[
                brand
            ][
                summary_variable
            ] = period_values.copy()

            break

    return results