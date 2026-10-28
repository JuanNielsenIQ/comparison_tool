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


def get_period_labels(periods_df):

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

    col = 4

    while col <= 200:

        value = ws.Cells(1, col).Value

        if (
            value is not None
            and
            str(value).strip().upper()
            == section_name.upper()
        ):
            return col

        col += 3

    col = 4

    while ws.Cells(1, col).Value is not None:
        col += 3

    ws.Cells(1, col).Value = section_name

    ws.Cells(2, col).Value = period_labels[0]
    ws.Cells(2, col + 1).Value = period_labels[1]
    ws.Cells(2, col + 2).Value = "VAR"

    return col


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


def write_report(
    workbook_path,
    results,
    periods_df,
    worksheet_name="Executive Summary",
    section_name="DISCOVER",
):

    excel = win32com.client.GetObject(
        Class="Excel.Application"
    )

    wb = get_workbook(
        excel,
        workbook_path,
    )

    ws = wb.Worksheets(
        worksheet_name
    )

    initialize_layout(ws)

    period_labels = get_period_labels(
        periods_df
    )

    p1_name = period_labels[0]
    p2_name = period_labels[1]

    section_col = find_or_create_section(
        ws,
        section_name,
        period_labels,
    )

    # =====================================
    # MODE 1
    # Existing structure already exists
    # =====================================

    if sheet_has_structure(ws):

        lookup = build_sheet_lookup(ws)

        for brand, measures in results.items():

            for variable, periods in measures.items():

                key = (
                    str(brand).strip().upper(),
                    str(variable).strip().upper(),
                )

                if key not in lookup:
                    continue

                row = lookup[key]

                p1 = periods.get(p1_name)
                p2 = periods.get(p2_name)

                var_value = calculate_var(
                    p1,
                    p2,
                )

                ws.Cells(row, section_col).Value = p1
                ws.Cells(row, section_col + 1).Value = p2
                ws.Cells(row, section_col + 2).Value = var_value

        last_row = max(
            lookup.values()
        )

    # =====================================
    # MODE 2
    # Empty sheet
    # =====================================

    else:

        row = 3

        for brand, measures in results.items():

            for variable, periods in measures.items():

                p1 = periods.get(p1_name)
                p2 = periods.get(p2_name)

                var_value = calculate_var(
                    p1,
                    p2,
                )

                ws.Cells(row, 1).Value = brand
                ws.Cells(row, 3).Value = variable

                ws.Cells(row, section_col).Value = p1
                ws.Cells(row, section_col + 1).Value = p2
                ws.Cells(row, section_col + 2).Value = var_value

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
    metric_name="DUE TOS",
):

    excel = win32com.client.GetObject(
        Class="Excel.Application"
    )

    wb = get_workbook(
        excel,
        workbook_path,
    )

    ws = wb.Worksheets(
        worksheet_name
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