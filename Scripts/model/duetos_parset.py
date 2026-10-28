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
    duetos_col = None

    for col in range(
        1,
        ws.max_column + 1,
    ):

        header = ws.cell(
            1,
            col,
        ).value

        if header is None:
            continue

        header = (
            str(header)
            .strip()
            .lower()
        )

        if header == "variable":

            variable_col = col

        elif header == "due_tos (%)":

            duetos_col = col

    if variable_col is None:

        raise ValueError(
            "Variable column not found."
        )

    if duetos_col is None:

        raise ValueError(
            "Due_tos (%) column not found."
        )

    return (
        variable_col,
        duetos_col,
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

    variable_col, duetos_col = (
        get_duetos_columns(
            ws
        )
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

        value = (
            ws.cell(
                row,
                duetos_col,
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
            ] = value

            break

    return results