"""
Inspect external YouTube and Google Maps reviews datasets
"""

import os, sys, openpyxl

sys.stdout.reconfigure(encoding="utf-8")

yt_path = r"C:\Users\omkar\Documents\antigravity\keen-bardeen\youtube_mall_master_analysis.xlsx"
maps_path = r"C:\Users\omkar\OneDrive\Desktop\ScrapePlaces\data\all_malls_reviews.xlsx"
maps_dir = r"C:\Users\omkar\OneDrive\Desktop\ScrapePlaces\data"

print("="*80)
print("INSPECTING EXTERNAL DATASETS")
print("="*80)

# Check YouTube Excel
if os.path.exists(yt_path):
    print(f"\n[FOUND] YouTube Master Analysis: {yt_path}")
    wb_yt = openpyxl.load_workbook(yt_path, read_only=True)
    print(f"Total Sheets in YouTube Workbook: {len(wb_yt.sheetnames)}")
    for s in wb_yt.sheetnames:
        ws = wb_yt[s]
        print(f"  • Tab: {s:<35} | Rows: {ws.max_row}")
else:
    print(f"\n[NOT FOUND] YouTube Master Analysis at {yt_path}")

# Check Google Maps Reviews Excel / Directory
if os.path.exists(maps_path):
    print(f"\n[FOUND] Google Maps Reviews: {maps_path}")
    wb_maps = openpyxl.load_workbook(maps_path, read_only=True)
    print(f"Total Sheets in Google Maps Workbook: {len(wb_maps.sheetnames)}")
    for s in wb_maps.sheetnames:
        ws = wb_maps[s]
        print(f"  • Tab: {s:<35} | Rows: {ws.max_row}")
elif os.path.exists(maps_dir):
    print(f"\n[DIRECTORY FOUND] {maps_dir}")
    files = os.listdir(maps_dir)
    print("Files in directory:")
    for f in files[:20]:
        sz = os.path.getsize(os.path.join(maps_dir, f))
        print(f"  • {f} ({sz/1024:.1f} KB)")
else:
    print(f"\n[NOT FOUND] Google Maps Reviews at {maps_path}")
