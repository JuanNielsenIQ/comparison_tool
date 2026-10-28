def add_total_metric_model(
    model_wb,
    results,
    period_labels,
):
    """
    Add TOTAL metric from
    'Model Summary by Type'.

    Logic:

    - Find Volume columns.
    - Handle merged Volume headers.
    - Match configured periods.
    - Ignore Agg.
    - TOTAL(period) =
        Row 4(period) +
        Row 5(period)
    """

    ws = model_wb[
        "Model Summary by Type"
    ]

    period_cols = {}

    expected_periods = {
        str(period).strip().lower(): period
        for period in period_labels
    }

    last_row1 = ""

    #
    # Find Volume columns
    #

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
        # Volume + matching period
        #

        if (
            row1.lower() == "volume"
            and row2.lower()
            in expected_periods
        ):

            period_cols[
                expected_periods[
                    row2.lower()
                ]
            ] = col

    #
    # Validation
    #

    missing_periods = [
        period
        for period in period_labels
        if period not in period_cols
    ]

    if missing_periods:

        raise ValueError(
            "Missing Volume columns "
            f"for periods: {missing_periods}. "
            f"Found periods: {list(period_cols.keys())}"
        )

    #
    # Build TOTAL metric
    #

    total_values = {}

    for period in period_labels:

        col = period_cols[
            period
        ]

        row4_value = (
            ws.cell(
                4,
                col,
            ).value
            or 0
        )

        row5_value = (
            ws.cell(
                5,
                col,
            ).value
            or 0
        )

        total_values[
            period
        ] = (
            row4_value
            + row5_value
        )

    #
    # Apply TOTAL to every brand
    #

    for brand in results:

        results[brand][
            "TOTAL"
        ] = total_values.copy()

    return results