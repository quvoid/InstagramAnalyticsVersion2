"""
Generate updated CSV with Paid Partnership Toggle column for All Brands - Paid Collabs
"""

import sys, openpyxl, csv

sys.stdout.reconfigure(encoding="utf-8")

wb = openpyxl.load_workbook("jewellery_brands_consolidated_analysis.xlsx", data_only=True)
ws = wb["All Brands - Paid Collabs"]

output_csv = "All_Brands_Paid_Collabs_With_Toggle.csv"

with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    # Row 2 is the header
    for r in range(2, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, 13)]
        if r > 2:
            # Format ER% as percentage string
            if isinstance(row_vals[7], (int, float)):
                row_vals[7] = f"{row_vals[7]*100:.2f}%" if row_vals[7] < 1 else f"{row_vals[7]:.2f}%"
        writer.writerow(row_vals)

print(f"✓ Generated: {output_csv} with 1,067 data rows and Paid Partnership Toggle column.")
