# eda/parser.py

# Parsing an already loaded EDA workbook (see utils/reader.py)
# Extracting:
# target brand
# competitor brands
# Detecting:
# Variable column
# Aggregated P1
# Aggregated P2
# Building and returning the results dictionary

import re


def empty_metrics():

    return {
        "tpr": {"P1": None, "P2": None},
        "regular_price": {"P1": None, "P2": None},
        "feature": {"P1": 0, "P2": 0},
        "display": {"P1": 0, "P2": 0},
    }


def get_project_info(eda_wb):

    project_ws = eda_wb["project_info"]

    target_brand = str(
        project_ws["I2"].value
    ).strip()

    competitor_brands = []

    row = 2

    while True:

        value = project_ws[f"J{row}"].value

        if value in (None, ""):
            break

        competitor_brands.append(
            str(value).strip()
        )

        row += 1

    return target_brand, competitor_brands


def get_predictor_columns(pred_ws):

    variable_col = None
    p1_col = None
    p2_col = None

    last_row1 = ""

    for col in range(1, pred_ws.max_column + 1):

        row1 = pred_ws.cell(1, col).value
        row2 = pred_ws.cell(2, col).value

        row1 = "" if row1 is None else str(row1).strip()
        row2 = "" if row2 is None else str(row2).strip()

        # Handle merged headers
        if row1 != "":
            last_row1 = row1
        else:
            row1 = last_row1

        if row1.lower() == "variable":
            variable_col = col

        if row1.lower() == "aggregated":

            if row2.lower() in ["year 1", "p1"]:
                p1_col = col

            elif row2.lower() in ["year 2", "p2"]:
                p2_col = col

    if variable_col is None:
        raise ValueError(
            "Variable column not found"
        )

    if p1_col is None:
        raise ValueError(
            "P1 column not found"
        )

    if p2_col is None:
        raise ValueError(
            "P2 column not found"
        )

    return variable_col, p1_col, p2_col


def build_results(eda_wb):

    target_brand, competitor_brands = (
        get_project_info(eda_wb)
    )

    pred_ws = eda_wb[
        "Predictors Summary"
    ]

    variable_col, p1_col, p2_col = (
        get_predictor_columns(pred_ws)
    )

    results = {
        "target_brand": {
            target_brand: empty_metrics()
        },
        "competitor_brands": {
            brand: empty_metrics()
            for brand in competitor_brands
        }
    }

    all_brands = (
        [target_brand]
        + competitor_brands
    )

    for row in range(4, pred_ws.max_row + 1):

        variable = pred_ws.cell(
            row,
            variable_col
        ).value

        if variable is None:
            continue

        variable = str(variable).strip()

        p1 = pred_ws.cell(
            row,
            p1_col
        ).value

        p2 = pred_ws.cell(
            row,
            p2_col
        ).value

        for brand in all_brands:

            container = (
                results["target_brand"]
                if brand == target_brand
                else results[
                    "competitor_brands"
                ]
            )

            # ------------------
            # TPR
            # ------------------

            if variable.lower() == (
                f"{brand}_tpr"
            ).lower():

                container[brand]["tpr"] = {
                    "P1": p1,
                    "P2": p2,
                }

            # ------------------
            # REGULAR PRICE
            # ------------------

            candidates = [
                f"{brand}_regularprice",
                f"{brand}_regular_price"
            ]

            if variable.lower() in [
                c.lower()
                for c in candidates
            ]:

                container[brand][
                    "regular_price"
                ] = {
                    "P1": p1,
                    "P2": p2,
                }

            # ------------------
            # FEATURE
            # ------------------

            feature_pattern = (
                rf"^FEATURE_"
                rf"{re.escape(brand)}_"
            )

            if re.match(
                feature_pattern,
                variable,
                re.IGNORECASE
            ):

                container[brand][
                    "feature"
                ]["P1"] += p1 or 0

                container[brand][
                    "feature"
                ]["P2"] += p2 or 0

            # ------------------
            # DISPLAY
            # ------------------

            display_pattern = (
                rf"^DISPLAY_"
                rf"{re.escape(brand)}_"
            )

            if re.match(
                display_pattern,
                variable,
                re.IGNORECASE
            ):

                container[brand][
                    "display"
                ]["P1"] += p1 or 0

                container[brand][
                    "display"
                ]["P2"] += p2 or 0

    return results