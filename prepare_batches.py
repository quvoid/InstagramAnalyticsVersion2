"""
Prepare batches of data for Google Sheets update_cells MCP tool
"""

import csv, json

with open("All_Brands_Paid_Collabs_With_Toggle.csv", encoding="utf-8-sig") as f:
    rows = list(csv.reader(f))

# Title row + Header + 1,067 data rows = 1,069 rows total
title_row = ["Consolidated Creator-Owned Paid Partnerships ('Post owned by partner') — 1067 Total Posts (Toggle ON: 4 · Toggle OFF: 1063)"] + [""] * 11
full_data = [title_row] + rows

with open("batch1.json", "w", encoding="utf-8") as f:
    json.dump(full_data[:500], f)

with open("batch2.json", "w", encoding="utf-8") as f:
    json.dump(full_data[500:], f)

print(f"Prepared Batch 1 (Rows 1-500): {len(full_data[:500])} rows")
print(f"Prepared Batch 2 (Rows 501-1069): {len(full_data[500:])} rows")
