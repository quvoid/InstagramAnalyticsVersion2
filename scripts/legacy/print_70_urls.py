"""
Output all 70 Toggle ON URLs grouped by Brand with Creator Handle and Date
"""

import sys, json

sys.stdout.reconfigure(encoding="utf-8")

with open("api_toggle_ground_truth.json", encoding="utf-8") as f:
    d = json.load(f)

true_on = [v for v in d.values() if v.get("is_paid_partnership")]
true_on.sort(key=lambda x: (x["brand"], x["index"]))

# Also grab dates from All_Brands_Paid_Collabs.csv
import csv
with open("All_Brands_Paid_Collabs.csv", encoding="utf-8-sig") as f_csv:
    csv_rows = {r[9]: r[8] for r in list(csv.reader(f_csv))[1:]}

print(f"Total: {len(true_on)} posts\n")

current_brand = ""
for idx, p in enumerate(true_on, 1):
    b = p["brand"]
    h = p["handle"]
    u = p["url"]
    dt = csv_rows.get(u, "N/A")
    if b != current_brand:
        current_brand = b
        print(f"\n### 🏛️ {current_brand}")
        print("| # | Creator Handle | Post Date | Direct Instagram URL |")
        print("|---|---|---|---|")
    print(f"| {idx} | `{h}` | {dt} | [{u}]({u}) |")
