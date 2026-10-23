def get_period_labels(periods_df):

    periods = (
        periods_df["custom_period"]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    )

    if len(periods) < 2:

        raise ValueError(
            "At least two periods are required."
        )

    return periods


def build_brand_mapping(
    brands_mapping_df,
):

    mapping = {}

    for _, row in brands_mapping_df.iterrows():

        summary_brand = row["Brand"]
        eda_brand = row["EDA"]

        if (
            summary_brand is not None
            and eda_brand is not None
            and str(summary_brand).strip() != ""
            and str(eda_brand).strip() != ""
        ):

            mapping[
                str(eda_brand).strip()
            ] = str(summary_brand).strip()

    return mapping


def build_variable_mapping(
    variables_mapping_df,
):

    mapping = {}

    mapping_df = (
        variables_mapping_df
        .dropna(subset=["variables"])
        .dropna(subset=["eda_report"])
    )

    for _, row in mapping_df.iterrows():

        summary_variable = (
            str(row["variables"])
            .strip()
        )

        eda_variable = (
            str(row["eda_report"])
            .strip()
        )

        if (
            summary_variable != ""
            and eda_variable != ""
        ):

            mapping[
                eda_variable
            ] = summary_variable

    return mapping


def build_eda_results(
    raw_results,
    variables_mapping_df,
    brands_mapping_df,
    periods_df,
):

    period_labels = get_period_labels(
        periods_df
    )

    brand_mapping = build_brand_mapping(
        brands_mapping_df
    )

    variable_mapping = build_variable_mapping(
        variables_mapping_df
    )

    target_brand = next(
        iter(raw_results["target_brand"])
    )

    all_brands = {
        target_brand:
            raw_results["target_brand"][
                target_brand
            ]
    }

    all_brands.update(
        raw_results["competitor_brands"]
    )

    output = {}

    for eda_brand, metrics in all_brands.items():

        summary_brand = brand_mapping.get(
            eda_brand,
            eda_brand
        )

        output[summary_brand] = {}

        #
        # Regular mapped metrics
        #

        for eda_metric, summary_variable in (
            variable_mapping.items()
        ):

            if eda_metric not in metrics:
                continue

            metric_values = metrics[
                eda_metric
            ]

            output[
                summary_brand
            ][summary_variable] = {}

            #
            # Dynamically assign
            # every available period
            #

            for idx, period in enumerate(
                period_labels
            ):

                source_key = (
                    "P" + str(idx + 1)
                )

                output[
                    summary_brand
                ][summary_variable][period] = (
                    metric_values.get(
                        source_key
                    )
                )

        #
        # ANY PROMO
        #
        # If not explicitly mapped
        #

        if (
            "ANY PROMO"
            not in output[
                summary_brand
            ]
        ):

            output[
                summary_brand
            ]["ANY PROMO"] = {}

            for period in period_labels:

                feature = (
                    output[
                        summary_brand
                    ]
                    .get(
                        "FEATURE",
                        {}
                    )
                    .get(
                        period,
                        0
                    )
                    or 0
                )

                regular_price = (
                    output[
                        summary_brand
                    ]
                    .get(
                        "REGULAR PRICE",
                        {}
                    )
                    .get(
                        period,
                        0
                    )
                    or 0
                )

                tpr = (
                    output[
                        summary_brand
                    ]
                    .get(
                        "TPR",
                        {}
                    )
                    .get(
                        period,
                        0
                    )
                    or 0
                )

                display = (
                    output[
                        summary_brand
                    ]
                    .get(
                        "DISPLAY",
                        {}
                    )
                    .get(
                        period,
                        0
                    )
                    or 0
                )

                output[
                    summary_brand
                ]["ANY PROMO"][period] = (
                    feature
                    + regular_price
                    + tpr
                    + display
                )

    return output