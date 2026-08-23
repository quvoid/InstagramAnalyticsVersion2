"""
Analyze and filter Kalyan Jewellers sheet data for paid partnership creators
"""

import json, sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

with open(r"C:/Users/omkar/.gemini/antigravity/brain/d196c916-bdf1-40a6-835a-3259489a070c/.system_generated/steps/633/output.txt", encoding="utf-8") as f:
    data = json.load(f)

values = data["valueRanges"][0].get("values", [])
title_row = values[0] if len(values) > 0 else []
header_row = values[1] if len(values) > 1 else []
rows = values[2:] if len(values) > 2 else []

print(f"Title: {title_row}")
print(f"Header: {header_row}")
print(f"Total Rows: {len(rows)}")

# Check detection methods and handles
detected_via = Counter()
handle_counts = Counter()

for r in rows:
    r_padded = r + [""] * (10 - len(r))
    h, fol, l, c, er, date, url, via, cap, stat = r_padded[:10]
    detected_via[via] += 1
    handle_counts[h] += 1

print("\nDetection methods breakdown:")
for k, v in detected_via.items():
    print(f"  {k}: {v}")

print(f"\nTotal unique handles: {len(handle_counts)}")
print("All handles and row counts:")
for h, cnt in handle_counts.most_common():
    print(f"  {h:<32}: {cnt} rows")
