import sys
import win32com.client

# Full path passed from VBA
workbook_path = sys.argv[1]

excel = win32com.client.GetObject(Class="Excel.Application")

# Find the matching open workbook
wb = None

for candidate in excel.Workbooks:
    if candidate.FullName.lower() == workbook_path.lower():
        wb = candidate
        break

if wb is None:
    raise Exception(f"Workbook not found: {workbook_path}")

ws = wb.ActiveSheet

ws.Range("A1").Value = "Workbook found!"
ws.Range("B1").Value = workbook_path

from pathlib import Path

Path(r"C:\Users\grju6002\Documents\macros\python_ran.txt").write_text("I ran")