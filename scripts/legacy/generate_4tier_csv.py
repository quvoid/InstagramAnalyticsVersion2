"""
Generate clean CSV exports for 4-tier hierarchy
"""

import sys, json, csv

sys.stdout.reconfigure(encoding="utf-8")

with open("unified_master_dataset.json", encoding="utf-8") as f:
    master_records = json.load(f)

# Sort master records strictly by Tier ascending, then Views descending
master_records.sort(key=lambda x: (x["tier"], -x["views"]))

# Write All_Brands_4Tier_Master.csv
with open("All_Brands_4Tier_Master.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow([
        "#", "Hierarchy Tier", "Brand Name", "Creator Handle", "Followers",
        "Views / Plays", "Likes", "Comments", "Like-to-View %", "Creator ER%",
        "Post Date", "Direct Instagram URL", "Boost Status & Classification", "Caption Preview"
    ])
    for idx, p in enumerate(master_records, 1):
        w.writerow([
            idx,
            p["tier_name"],
            p["brand"],
            p["handle"],
            p["followers"],
            p["views"],
            p["likes"],
            p["comments"],
            f"{p['like_rate_pct']:.2f}%",
            f"{p['er_pct']:.2f}%",
            p["post_date"],
            p["url"],
            f"{p['boost_status']}: {p['reason']}" if p['reason'] else p['boost_status'],
            p["caption"]
        ])

print("✓ Saved All_Brands_4Tier_Master.csv")
