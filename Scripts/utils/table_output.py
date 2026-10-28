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

        if len(periods) != 2:

            raise ValueError(
                f"Expected 2 periods, found {periods}"
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
    # VAR column
    #

    ws.Cells(
        2,
        current_col
    ).Value = "VAR"

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
    section_col,
    last_row,
):

    if last_row < 3:
        return

    ws.Range(
        ws.Cells(3, section_col + 2),
        ws.Cells(last_row, section_col + 2)
    ).NumberFormat = "0%"

    ws.Range(
        ws.Cells(3, section_col),
        ws.Cells(last_row, section_col + 1)
    ).NumberFormat = "0.00"


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
    is_discover=False
):

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

    section_columns[
        "VAR"
    ] = current_col

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

                first_period = (
                    period_labels[0]
                )

                last_period = (
                    period_labels[-1]
                )

                var_value = (
                    calculate_var(
                        periods.get(
                            first_period
                        ),
                        periods.get(
                            last_period
                        ),
                    )
                )

                ws.Cells(
                    row,
                    section_columns[
                        "VAR"
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

                first_period = (
                    period_labels[0]
                )

                last_period = (
                    period_labels[-1]
                )

                var_value = (
                    calculate_var(
                        periods.get(
                            first_period
                        ),
                        periods.get(
                            last_period
                        ),
                    )
                )

                ws.Cells(
                    row,
                    section_columns[
                        "VAR"
                    ]
                ).Value = var_value

                row += 1

        last_row = row - 1

    format_section(
        ws,
        section_col,
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
):

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

    metric_col = None

    #
    # Reuse existing metric column
    # if it already exists
    #

    for col in range(4, 200):

        metric = ws.Cells(
            2,
            col
        ).Value

        if (
            metric is not None
            and str(metric).strip().upper()
            == metric_name.upper()
        ):

            metric_col = col
            break

    #
    # Metric not found:
    # create it under the section
    #

    if metric_col is None:

        section_start = None

        for col in range(4, 200):

            section = ws.Cells(
                1,
                col
            ).Value

            if (
                section is not None
                and str(section).strip().upper()
                == section_name.upper()
            ):

                section_start = col
                break

        #
        # Existing section
        #

        if section_start is not None:

            metric_col = section_start

            while (
                ws.Cells(
                    2,
                    metric_col
                ).Value
                is not None
            ):
                metric_col += 1

        #
        # New section
        #

        else:

            metric_col = 4

            while (
                ws.Cells(
                    2,
                    metric_col
                ).Value
                is not None
            ):
                metric_col += 1

            ws.Cells(
                1,
                metric_col
            ).Value = section_name

        ws.Cells(
            2,
            metric_col
        ).Value = metric_name

    #
    # Build lookup:
    # (BRAND, VARIABLE) -> ROW
    #

    lookup = build_brand_variable_lookup(
        ws
    )

    #
    # Clear existing values
    # before rewriting
    #

    last_row = max(
        lookup.values(),
        default=3,
    )

    ws.Range(
        ws.Cells(
            3,
            metric_col,
        ),
        ws.Cells(
            last_row,
            metric_col,
        ),
    ).ClearContents()

    #
    # Write values
    #

    for brand, variables in results.items():

        for variable, value in variables.items():

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

            row = lookup[key]

            ws.Cells(
                row,
                metric_col
            ).Value = value

    #
    # Format
    #

    ws.Range(
        ws.Cells(
            3,
            metric_col,
        ),
        ws.Cells(
            last_row,
            metric_col,
        ),
    ).NumberFormat = "0.00"

    excel.Calculate()

    return metric_col