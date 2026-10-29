import win32com.client


def get_workbook(
    excel,
    workbook_path,
):

    for wb in excel.Workbooks:

        if (
            wb.FullName.lower()
            == workbook_path.lower()
        ):
            return wb

    raise Exception(
        f"Workbook not found: {workbook_path}"
    )


def get_period_labels(periods_df, workbook_path, is_discover=False):

    if not is_discover:
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
                f"Expected at least 2 periods, found {periods}"
            )

        return periods

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


    periods = []
    seen = set()

    row = 2

    while True:

        value = ws.Cells(
            row,
            5,  # Column E
        ).Value

        if value is None:
            break

        value = str(value).strip()

        if (
            value
            and value not in seen
        ):

            periods.append(
                value
            )

            seen.add(
                value
            )

        row += 1

    return periods
    


def initialize_layout(ws):

    if ws.Cells(2, 1).Value is None:
        ws.Cells(2, 1).Value = "BRAND"

    if ws.Cells(2, 2).Value is None:
        ws.Cells(2, 2).Value = "SIGN"

    if ws.Cells(2, 3).Value is None:
        ws.Cells(2, 3).Value = "VARIABLE"


def find_or_create_section(
    ws,
    section_name,
    period_labels,
):
    """
    Returns the first column of the section.

    If the section does not exist:
        - Creates period columns
        - Creates VAR column
        - Appends after the last used header
    """

    #
    # Existing section
    #

    for col in range(4, 201):

        value = ws.Cells(
            1,
            col
        ).Value

        if (
            value is not None
            and str(value).strip().upper()
            == section_name.upper()
        ):

            return col

    #
    # Find first empty header column
    #

    section_col = (
        ws.Cells(
            2,
            ws.Columns.Count
        )
        .End(-4159)  # xlToLeft
        .Column
        + 1
    )

    #
    # Section name
    #

    ws.Cells(
        1,
        section_col
    ).Value = section_name

    current_col = section_col

    #
    # Period columns
    #

    for period in period_labels:

        ws.Cells(
            2,
            current_col
        ).Value = period

        current_col += 1

    #
    # VAR columns
    #

    for idx in range(
        len(period_labels) - 1
    ):

        current_period = (
            period_labels[idx]
        )

        next_period = (
            period_labels[idx + 1]
        )

        ws.Cells(
            2,
            current_col
        ).Value = (
            f"VAR_{current_period}_vs_{next_period}"
        )

        current_col += 1
    
    return section_col



def calculate_var(
    p1,
    p2,
):

    if p1 in (None, 0):
        return None

    if p2 is None:
        return None

    return (p2 / p1) - 1


def format_section(
    ws,
    section_columns,
    last_row,
):

    if last_row < 3:
        return

    for header, col in section_columns.items():

        rng = ws.Range(
            ws.Cells(
                3,
                col,
            ),
            ws.Cells(
                last_row,
                col,
            ),
        )

        if str(header).upper().startswith(
            "VAR_"
        ):

            rng.NumberFormat = "0%"

        else:

            rng.NumberFormat = "0.00"

def sheet_has_structure(ws):

    return (
        ws.Cells(3, 1).Value is not None
        and
        ws.Cells(3, 3).Value is not None
    )


def build_sheet_lookup(ws):

    lookup = {}

    row = 3

    while True:

        brand = ws.Cells(row, 1).Value
        variable = ws.Cells(row, 3).Value

        if (
            brand is None
            and variable is None
        ):
            break

        if (
            brand is not None
            and variable is not None
        ):

            lookup[
                (
                    str(brand).strip().upper(),
                    str(variable).strip().upper(),
                )
            ] = row

        row += 1

    return lookup


def get_or_create_worksheet(
    wb,
    worksheet_name,
):

    try:

        return wb.Worksheets(
            worksheet_name
        )

    except Exception:

        ws = wb.Worksheets.Add()

        ws.Name = (
            worksheet_name
        )

        return ws


def write_report(
    workbook_path,
    results,
    periods_df,
    worksheet_name="Executive Summary",
    section_name="DISCOVER",
    variables_mapping_df=None,
    brands_mapping_df=None,
    is_discover=False,
    brand_mapping_column=None,
):

    excel = win32com.client.GetObject(
        Class="Excel.Application"
    )


    brand_lookup = {}

    if (
        brands_mapping_df is not None
        and brand_mapping_column is not None
    ):
        brand_lookup = dict(
            zip(
                brands_mapping_df[
                    brand_mapping_column
                ]
                .astype(str)
                .str.strip()
                .str.upper(),

                brands_mapping_df.iloc[:, 0]
                .astype(str)
                .str.strip(),
            )
        )

    if brand_lookup:

        mapped_results = {}

        for (
            source_brand,
            measures
        ) in results.items():

            target_brand = (
                brand_lookup.get(
                    str(source_brand)
                    .strip()
                    .upper()
                )
            )

            if target_brand is None:
                continue

            mapped_results[
                target_brand
            ] = measures

        results = mapped_results


    wb = get_workbook(
        excel,
        workbook_path,
    )

    ws = get_or_create_worksheet(
        wb,
        worksheet_name,
    )

    initialize_layout(ws)

    period_labels = get_period_labels(
        periods_df,
        workbook_path,
        is_discover=is_discover
    )

    #
    # Build ordered brands
    #

    if brands_mapping_df is not None:

        brand_order = (
            brands_mapping_df.iloc[:, 0]
            .dropna()
            .astype(str)
            .str.strip()
            .tolist()
        )

    else:

        brand_order = list(
            results.keys()
        )

    #
    # Build ordered variables
    #

    if variables_mapping_df is not None:

        variable_order = (
            variables_mapping_df[
                "variables"
            ]
            .dropna()
            .astype(str)
            .str.strip()
            .tolist()
        )

    else:

        all_vars = set()

        for measures in (
            results.values()
        ):
            all_vars.update(
                measures.keys()
            )

        variable_order = list(
            all_vars
        )

    #
    # Create section
    #

    
    section_col = (
        find_or_create_section(
            ws,
            section_name,
            period_labels
        )
    )

    #
    # Map section columns
    #

    section_columns = {}

    current_col = section_col

    for period in period_labels:

        section_columns[
            period
        ] = current_col

        current_col += 1


    for idx in range(
        len(period_labels) - 1
    ):

        current_period = (
            period_labels[idx]
        )

        next_period = (
            period_labels[idx + 1]
        )

        section_columns[
            f"VAR_{current_period}_vs_{next_period}"
        ] = current_col

        current_col += 1
    #
    # Existing structure
    #

    if sheet_has_structure(ws):

        lookup = (
            build_sheet_lookup(
                ws
            )
        )

        for brand in brand_order:

            if brand not in results:
                continue

            measures = results[
                brand
            ]

            for variable in (
                variable_order
            ):

                if variable not in measures:
                    continue

                key = (
                    brand.upper(),
                    variable.upper(),
                )

                if key not in lookup:
                    continue

                row = lookup[key]

                periods = measures[
                    variable
                ]

                #
                # Write periods
                #

                for period in (
                    period_labels
                ):

                    ws.Cells(
                        row,
                        section_columns[
                            period
                        ]
                    ).Value = (
                        periods.get(
                            period
                        )
                    )

                #
                # VAR uses first/last period
                #
                for idx in range(
                    len(period_labels) - 1
                ):

                    current_period = (
                        period_labels[idx]
                    )

                    next_period = (
                        period_labels[idx + 1]
                    )

                    var_value = calculate_var(
                        periods.get(
                            current_period
                        ),
                        periods.get(
                            next_period
                        ),
                    )

                    ws.Cells(
                        row,
                        section_columns[
                            f"VAR_{current_period}_vs_{next_period}"
                        ]
                    ).Value = var_value

        last_row = max(
            lookup.values(),
            default=3,
        )

    #
    # Empty structure
    #

    else:

        row = 3

        for brand in brand_order:

            if brand not in results:
                continue

            measures = results[
                brand
            ]

            for variable in (
                variable_order
            ):

                if variable not in measures:
                    continue

                periods = measures[
                    variable
                ]

                ws.Cells(
                    row,
                    1
                ).Value = brand

                ws.Cells(
                    row,
                    3
                ).Value = variable

                for period in (
                    period_labels
                ):

                    ws.Cells(
                        row,
                        section_columns[
                            period
                        ]
                    ).Value = (
                        periods.get(
                            period
                        )
                    )

                for idx in range(
                    len(period_labels) - 1
                ):

                    current_period = (
                        period_labels[idx]
                    )

                    next_period = (
                        period_labels[idx + 1]
                    )

                    var_value = calculate_var(
                        periods.get(
                            current_period
                        ),
                        periods.get(
                            next_period
                        ),
                    )

                    ws.Cells(
                        row,
                        section_columns[
                            f"VAR_{current_period}_vs_{next_period}"
                        ]
                    ).Value = var_value

                row += 1

        last_row = row - 1

    format_section(
        ws,
        section_columns,
        last_row,
    )

    excel.Calculate()

    return last_row


def build_brand_variable_lookup(ws):

    lookup = {}

    row = 3

    while True:

        brand = ws.Cells(
            row,
            1
        ).Value

        variable = ws.Cells(
            row,
            3
        ).Value

        if (
            brand is None
            and variable is None
        ):
            break

        if (
            brand is not None
            and variable is not None
        ):

            lookup[
                (
                    str(brand)
                    .strip()
                    .upper(),

                    str(variable)
                    .strip()
                    .upper(),
                )
            ] = row

        row += 1

    return lookup

def write_brand_metric(
    workbook_path,
    results,
    worksheet_name="Executive Summary",
    section_name="MODEL",
    metric_name="DUE TOS (%)",
    brands_mapping_df=None,
    brand_mapping_column=None,
):

    #
    # Brand mapping
    #

    brand_lookup = {}

    if (
        brands_mapping_df is not None
        and brand_mapping_column is not None
    ):
        brand_lookup = dict(
            zip(
                brands_mapping_df[
                    brand_mapping_column
                ]
                .astype(str)
                .str.strip()
                .str.upper(),

                brands_mapping_df.iloc[:, 0]
                .astype(str)
                .str.strip(),
            )
        )

    if brand_lookup:

        mapped_results = {}

        for (
            source_brand,
            measures
        ) in results.items():

            target_brand = (
                brand_lookup.get(
                    str(source_brand)
                    .strip()
                    .upper()
                )
            )

            if target_brand is None:
                continue

            mapped_results[
                target_brand
            ] = measures

        results = mapped_results

    #
    # Workbook / Worksheet
    #

    excel = win32com.client.GetObject(
        Class="Excel.Application"
    )

    wb = get_workbook(
        excel,
        workbook_path,
    )

    ws = get_or_create_worksheet(
        wb,
        worksheet_name,
    )

    initialize_layout(ws)

    #
    # Discover period headers
    #

    period_headers = []

    for brand_values in results.values():

        for variable_values in (
            brand_values.values()
        ):

            period_headers = list(
                variable_values.keys()
            )

            break

        if period_headers:
            break

    if not period_headers:
        return {}

    #
    # Find section
    #

    section_start = None

    for col in range(
        4,
        200,
    ):

        value = ws.Cells(
            1,
            col,
        ).Value

        if (
            value is not None
            and str(value).strip().upper()
            == section_name.upper()
        ):

            section_start = col
            break

    #
    # Create section if needed
    #

    if section_start is None:

        section_start = (
            ws.Cells(
                2,
                ws.Columns.Count,
            )
            .End(-4159)  # xlToLeft
            .Column
            + 1
        )

        ws.Cells(
            1,
            section_start,
        ).Value = section_name

    #
    # Build metric columns
    # Reuse existing headers
    #

    metric_cols = {}

    for period_header in (
        period_headers
    ):

        header_name = (
            f"{metric_name}_{period_header}"
        )

        existing_col = None

        #
        # Look for existing header
        #

        for col in range(
            section_start,
            200,
        ):

            value = ws.Cells(
                2,
                col,
            ).Value

            if (
                value is not None
                and str(value).strip().upper()
                == header_name.upper()
            ):

                existing_col = col
                break

        #
        # Reuse existing column
        #

        if existing_col is not None:

            metric_cols[
                period_header
            ] = existing_col

            continue

        #
        # Create new column
        #

        new_col = section_start

        while (
            ws.Cells(
                2,
                new_col,
            ).Value
            is not None
        ):
            new_col += 1

        ws.Cells(
            2,
            new_col,
        ).Value = header_name

        metric_cols[
            period_header
        ] = new_col

    #
    # Build row lookup
    #

    lookup = (
        build_brand_variable_lookup(
            ws
        )
    )

    last_row = max(
        lookup.values(),
        default=3,
    )

    #
    # Clear existing data
    #

    for col in (
        metric_cols.values()
    ):

        ws.Range(
            ws.Cells(
                3,
                col,
            ),
            ws.Cells(
                last_row,
                col,
            ),
        ).ClearContents()

    #
    # Write values
    #

    for (
        brand,
        variables,
    ) in results.items():

        for (
            variable,
            period_values,
        ) in variables.items():

            key = (
                str(brand)
                .strip()
                .upper(),

                str(variable)
                .strip()
                .upper(),
            )

            if key not in lookup:
                continue

            row = lookup[
                key
            ]

            for (
                period_header,
                value,
            ) in period_values.items():

                col = (
                    metric_cols.get(
                        period_header
                    )
                )

                if col is None:
                    continue

                ws.Cells(
                    row,
                    col,
                ).Value = value

    #
    # Format
    #

    for col in (
        metric_cols.values()
    ):

        ws.Range(
            ws.Cells(
                3,
                col,
            ),
            ws.Cells(
                last_row,
                col,
            ),
        ).NumberFormat = "0%"

    excel.Calculate()

    return metric_cols
