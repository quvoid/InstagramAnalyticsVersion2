"""
Inspect Croma Spreadsheet Sheets and Columns
"""

import sys, openpyxl

sys.stdout.reconfigure(encoding="utf-8")

wb = openpyxl.load_workbook("croma_spreadsheet_export.xlsx", data_only=True)
print(f"Total Sheets in Croma Spreadsheet: {len(wb.sheetnames)}")
print("Sheet Names:")
for i, s in enumerate(wb.sheetnames, 1):
    ws = wb[s]
    print(f"  {i:>2}. {s:<30} (Rows: {ws.max_row:>4}, Cols: {ws.max_column})")

print("\nInspecting first 3 rows of each sheet:")
for s in wb.sheetnames[:6]:
    ws = wb[s]
    print(f"\n--- Sheet: {s} ---")
    for r in range(1, min(4, ws.max_row + 1)):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, min(12, ws.max_column + 1))]
        print(f"  Row {r}: {row_vals}")
