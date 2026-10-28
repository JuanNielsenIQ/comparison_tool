
def build_formula_variables(
    results,
    variables_df,
):
    """
    Build variables whose EDA_Report
    contains formulas separated by '+'.
    """

    variables_df = variables_df.copy()

    variables_df.columns = (
        variables_df.columns
        .astype(str)
        .str.strip()
        .str.lower()
    )

    formula_rows = variables_df[
        variables_df["eda_report"]
        .astype(str)
        .str.contains(
            r"\+",
            regex=True,
            na=False,
        )
    ]

    for _, row in formula_rows.iterrows():

        summary_variable = (
            str(row["variables"])
            .strip()
        )

        components = [
            item.strip()
            for item in str(
                row["eda_report"]
            ).split("+")
        ]

        component_variables = []

        for component in components:

            match = variables_df[
                variables_df[
                    "eda_report"
                ]
                .astype(str)
                .str.strip()
                .str.lower()
                == component.lower()
            ]

            if match.empty:
                continue

            component_variables.append(
                str(
                    match.iloc[0][
                        "variables"
                    ]
                ).strip()
            )

        for brand, measures in results.items():

            periods = {}

            for variable in component_variables:

                if variable not in measures:
                    continue

                for period, value in (
                    measures[
                        variable
                    ].items()
                ):

                    periods.setdefault(
                        period,
                        0
                    )

                    periods[
                        period
                    ] += value

            results[brand][
                summary_variable
            ] = periods

    return results


def add_total_metric(
    eda_wb,
    results,
    period_labels,
):
    """
    Add TOTAL metric from
    Brand Summary sheet.

    Output:

    results[brand]["TOTAL"] = {
        "Year 1": value,
        "Year 2": value,
        ...
    }
    """

    ws = eda_wb[
        "Brand Summary"
    ]

    # Period values start in column E
    first_period_col = 5

    brand_totals = {}

    for row in range(
        4,
        ws.max_row + 1,
    ):

        brand = ws.cell(
            row,
            1
        ).value

        if brand is None:
            continue

        brand = str(
            brand
        ).strip()

        values = {}

        for idx, period in enumerate(
            period_labels
        ):

            col = (
                first_period_col
                + idx
            )

            values[period] = (
                ws.cell(
                    row,
                    col
                ).value
                or 0
            )

        brand_totals[
            brand
        ] = values

    for brand in results:

        if brand in brand_totals:

            results[brand][
                "TOTAL"
            ] = brand_totals[
                brand
            ]

    return results


def add_no_promo_metric(
    results,
    total_variable="TOTAL",
    promo_variable="ANY PROMO",
    output_variable="NO PROMO",
):
    """
    Create:

        NO PROMO = TOTAL - ANY PROMO

    for every brand and period.
    """

    for brand, measures in results.items():

        if total_variable not in measures:
            continue

        if promo_variable not in measures:
            continue

        total_values = measures[
            total_variable
        ]

        promo_values = measures[
            promo_variable
        ]

        no_promo = {}

        for period in total_values:

            total = total_values.get(
                period,
                0
            )

            promo = promo_values.get(
                period,
                0
            )

            no_promo[
                period
            ] = total - promo

        results[brand][
            output_variable
        ] = no_promo

    return results