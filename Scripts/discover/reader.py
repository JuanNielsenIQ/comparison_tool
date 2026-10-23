# discover/reader.py

import pandas as pd
import xlwings as xw


def get_workbook_name(workbook_path):

    return workbook_path.split("\\")[-1]


def dataframe_from_values(values):

    values = [
        row for row in values
        if row[0] is not None
    ]

    if not values:
        return pd.DataFrame()

    headers = values[0]

    return pd.DataFrame(
        values[1:],
        columns=headers
    )


def get_periods_mapping(workbook_path):

    workbook_name = get_workbook_name(
        workbook_path
    )

    wb = xw.books[workbook_name]

    try: 
        ws = wb.sheets["Periods Mapping"]

        values = ws.range(
            "A1:E200"
        ).value

        return dataframe_from_values(
            values
        )

    except:
        ws = wb.sheets["Periods Mapping (Template)"]

        values = ws.range(
            "A1:E200"
        ).value

        return dataframe_from_values(
            values
        )





def get_variables_mapping(workbook_path):

    workbook_name = get_workbook_name(
        workbook_path
    )

    wb = xw.books[workbook_name]

    ws = wb.sheets["Variables Mapping"]

    values = ws.range(
        "A1:D50"
    ).value

    return dataframe_from_values(
        values
    )


def get_brands_mapping(workbook_path):

    workbook_name = get_workbook_name(
        workbook_path
    )

    wb = xw.books[workbook_name]

    ws = wb.sheets["Brands Mapping"]

    values = ws.range(
        "A1:D100"
    ).value

    return dataframe_from_values(
        values
    )


def load_discover_data(workbook_path):

    discover_path = (
        workbook_path.rsplit("\\", 1)[0]
        + "\\data\\discover.xlsx"
    )

    discover_df = pd.read_excel(
        discover_path,
        sheet_name=0,
        header=8,
        engine="calamine"
    )

    return discover_df


def load_all_inputs(workbook_path):

    periods_df = get_periods_mapping(
        workbook_path
    )

    variables_mapping_df = (
        get_variables_mapping(
            workbook_path
        )
    )

    discover_df = (
        load_discover_data(
            workbook_path
        )
    )

    brands_mapping =(
        get_brands_mapping(
            workbook_path
        )
    )

    return (
        periods_df,
        variables_mapping_df,
        discover_df,
        brands_mapping
    )