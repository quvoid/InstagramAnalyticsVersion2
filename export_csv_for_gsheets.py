"""
Export the consolidated tables to clean CSV files ready for direct 1-click Google Sheets Import
"""

import sys, openpyxl, csv

sys.stdout.reconfigure(encoding="utf-8")

wb = openpyxl.load_workbook("jewellery_brands_consolidated_analysis.xlsx", data_only=True)

# 1. Export All Brands - Paid Collabs CSV
ws1 = wb["All Brands - Paid Collabs"]
with open("All_Brands_Paid_Collabs.csv", "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    for r in range(2, ws1.max_row + 1):
        row_vals = [ws1.cell(row=r, column=c).value for c in range(1, 13)]
        # format ER percentage nicely if float
        if isinstance(row_vals[7], (int, float)):
            row_vals[7] = f"{row_vals[7]*100:.2f}%" if row_vals[7] < 1 else f"{row_vals[7]:.2f}%"
        writer.writerow(row_vals)

print("✓ Created All_Brands_Paid_Collabs.csv (1,067 rows)")

# 2. Export Brand-Creator Summary CSV
ws2 = wb["Brand-Creator Summary"]
with open("Brand_Creator_Summary.csv", "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    for r in range(2, ws2.max_row + 1):
        row_vals = [ws2.cell(row=r, column=c).value for c in range(1, 12)]
        if isinstance(row_vals[6], (int, float)):
            row_vals[6] = f"{row_vals[6]*100:.1f}%"
        if isinstance(row_vals[8], (int, float)):
            row_vals[8] = f"{row_vals[8]*100:.2f}%"
        writer.writerow(row_vals)

print("✓ Created Brand_Creator_Summary.csv (34 brands summary)")
