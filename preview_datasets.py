"""
Deep inspect columns and sample rows from YouTube and Google Maps datasets
"""

import sys, openpyxl

sys.stdout.reconfigure(encoding="utf-8")

yt_path = r"C:\Users\omkar\Documents\antigravity\keen-bardeen\youtube_mall_master_analysis.xlsx"
maps_path = r"C:\Users\omkar\OneDrive\Desktop\ScrapePlaces\data\all_malls_reviews.xlsx"

print("="*80)
print("GOOGLE MAPS REVIEWS DATASET SCHEMA & PREVIEW")
print("="*80)

wb_maps = openpyxl.load_workbook(maps_path, read_only=True)
ws_maps = wb_maps.active
headers_maps = [cell.value for cell in next(ws_maps.iter_rows(max_row=1))]
print(f"Total Rows: {ws_maps.max_row} | Columns ({len(headers_maps)}): {headers_maps}\n")

print("Sample 3 Reviews:")
for idx, row in enumerate(ws_maps.iter_rows(min_row=2, max_row=4, values_only=True), 1):
    print(f"Row {idx}: {row[:8]}")

print("\n" + "="*80)
print("YOUTUBE MASTER ANALYSIS DATASET SCHEMA & PREVIEW")
print("="*80)

wb_yt = openpyxl.load_workbook(yt_path, read_only=True)
for s in wb_yt.sheetnames:
    ws = wb_yt[s]
    headers_yt = [cell.value for cell in next(ws.iter_rows(max_row=2))]
    print(f"Tab: {s} | Columns: {headers_yt[:6]}...")
