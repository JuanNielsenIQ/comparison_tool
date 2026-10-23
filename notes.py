import pandas as pd

discover_path = r"C:\Users\grju6002\Documents\macros\data\discover.xlsx"

df = pd.read_excel(
    discover_path,
    sheet_name=0,
    header=8,
    engine="calamine"
)

print(df.head())