"""
Extract all 'Post owned by partner (collab)' posts from the 10 electronics brands in croma_spreadsheet_export.xlsx
"""

import sys, openpyxl, json, csv
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

wb = openpyxl.load_workbook("croma_spreadsheet_export.xlsx", data_only=True)

BRAND_SHEETS = [
    "Croma",
    "Vijay Sales",
    "Reliance Digital",
    "Electronics Mart",
    "Bajaj Electronics",
    "Sathya",
    "Pai International",
    "Aditya Vision",
    "Great Eastern",
    "Tata Neu"
]

all_collab_posts = []

print("Extracting 'Post owned by partner (collab)' from all 10 electronics brands:\n")

for b_sheet in BRAND_SHEETS:
    if b_sheet not in wb.sheetnames:
        print(f"Sheet {b_sheet} not found in workbook")
        continue
    ws = wb[b_sheet]
    brand_posts = []
    
    # Read rows
    header = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]
    
    for r in range(3, ws.max_row + 1):
        handle = ws.cell(row=r, column=1).value
        cat = ws.cell(row=r, column=2).value
        followers = ws.cell(row=r, column=3).value or 0
        likes = ws.cell(row=r, column=4).value or 0
        comments = ws.cell(row=r, column=5).value or 0
        er = ws.cell(row=r, column=6).value or 0
        post_date = ws.cell(row=r, column=7).value or "N/A"
        post_url = ws.cell(row=r, column=8).value
        via = ws.cell(row=r, column=9).value
        caption = ws.cell(row=r, column=10).value or ""
        
        if not post_url or not handle:
            continue
            
        # Filter for genuine creator / partner owned posts
        via_str = str(via or "").strip()
        if "Post owned by partner" in via_str or "collab" in via_str.lower() or cat == "Creator/Individual":
            brand_posts.append({
                "brand": b_sheet,
                "handle": str(handle).strip(),
                "followers": int(float(followers)) if followers else 0,
                "likes": int(float(likes)) if likes else 0,
                "comments": int(float(comments)) if comments else 0,
                "er_pct": round(float(er), 2) if er else 0.0,
                "post_date": str(post_date).split(" ")[0],
                "url": str(post_url).strip(),
                "via": via_str or "Post owned by partner (collab)",
                "caption": str(caption)[:250].replace("\n", " ")
            })
            
    print(f"  • {b_sheet:<22}: {len(brand_posts):>4} partner-owned posts ({len(set(p['handle'].lower() for p in brand_posts))} unique creators)")
    all_collab_posts.extend(brand_posts)

# Deduplicate by (brand, url)
seen_keys = set()
unique_posts = []
for p in all_collab_posts:
    k = (p["brand"], p["url"])
    if k not in seen_keys:
        seen_keys.add(k)
        unique_posts.append(p)

print(f"\nTotal Unique Partner-Owned Posts across all 10 Brands: {len(unique_posts)}")

with open("croma_raw_extracted_posts.json", "w", encoding="utf-8") as f:
    json.dump(unique_posts, f, indent=2)

print("✓ Saved croma_raw_extracted_posts.json")
