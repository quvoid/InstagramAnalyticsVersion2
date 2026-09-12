"""
Build the Unified 4-Tier Master Dataset across all 36 Brands (1,179 Posts)
Tiers:
  Tier 1: 🟢 Toggle ON + 🚀 Boosted (Paid Label + Paid Media Spend)
  Tier 2: 🟢 Toggle ON + ⚪ Organic (Paid Label + Organic Reach)
  Tier 3: 🚀 Toggle OFF + 🚀 Boosted (Collab Only + Paid Media Spend)
  Tier 4: ⚪ Toggle OFF + ⚪ Organic (Collab Only + Standard Organic Reach / Noise)
"""

import sys, json, csv, re
sys.stdout.reconfigure(encoding="utf-8")

# 1. Load 70 verified Toggle ON posts & their boost analysis
with open("boost_analysis_70_videos.json", encoding="utf-8") as f:
    boost_70 = json.load(f)

# 2. Load 997 Toggle OFF posts & their boost analysis
with open("boost_analysis_toggle_off.json", encoding="utf-8") as f:
    boost_off_997 = json.load(f)

# 3. Load 112 GIVA & Palmonas posts
with open("giva_palmonas_scraped.json", encoding="utf-8") as f:
    giva_pal_112 = json.load(f)

# Combine into master list
all_master_records = []

# Process 70 Toggle ON posts
for p in boost_70:
    is_boosted = "Boosted" in p.get("boost_status", "")
    all_master_records.append({
        "brand": p["brand"],
        "handle": p["handle"],
        "followers": p["followers"],
        "views": p["views"],
        "likes": p["likes"],
        "comments": p["comments"],
        "like_rate_pct": p["like_rate_pct"],
        "er_pct": round(((p["likes"] + p["comments"]) / p["followers"] * 100), 2) if p["followers"] > 0 else 0.0,
        "post_date": p["post_date"],
        "url": p["url"],
        "is_paid_partnership": True,
        "is_boosted": is_boosted,
        "tier": 1 if is_boosted else 2,
        "tier_name": "Tier 1: Toggle ON + Boosted" if is_boosted else "Tier 2: Toggle ON + Organic",
        "boost_status": p.get("boost_status", ""),
        "reason": p.get("reason", ""),
        "caption": ""
    })

# Process 997 Toggle OFF posts from original 34 brands
for p in boost_off_997:
    is_boosted = p.get("is_boosted", False)
    all_master_records.append({
        "brand": p["brand"],
        "handle": p["handle"],
        "followers": p["followers"],
        "views": p["views"],
        "likes": p["likes"],
        "comments": p["comments"],
        "like_rate_pct": p["like_rate_pct"],
        "er_pct": round(((p["likes"] + p["comments"]) / p["followers"] * 100), 2) if p["followers"] > 0 else 0.0,
        "post_date": p["post_date"],
        "url": p["url"],
        "is_paid_partnership": False,
        "is_boosted": is_boosted,
        "tier": 3 if is_boosted else 4,
        "tier_name": "Tier 3: Toggle OFF + Boosted" if is_boosted else "Tier 4: Toggle OFF + Organic (Noise)",
        "boost_status": p.get("verdict", ""),
        "reason": p.get("reason", ""),
        "caption": ""
    })

# Process 112 GIVA & Palmonas posts
for p in giva_pal_112:
    is_boosted = p.get("is_boosted", False)
    # As verified, GIVA and Palmonas are executed via Co-Author (Toggle OFF)
    all_master_records.append({
        "brand": p["brand"],
        "handle": p["handle"],
        "followers": p["followers"],
        "views": p["views"],
        "likes": p["likes"],
        "comments": p["comments"],
        "like_rate_pct": p["like_rate"],
        "er_pct": p["er"],
        "post_date": p["date"],
        "url": p["url"],
        "is_paid_partnership": False,
        "is_boosted": is_boosted,
        "tier": 3 if is_boosted else 4,
        "tier_name": "Tier 3: Toggle OFF + Boosted" if is_boosted else "Tier 4: Toggle OFF + Organic (Noise)",
        "boost_status": p.get("boost_status", ""),
        "reason": "Paid media views detected" if is_boosted else "Standard organic reach",
        "caption": p.get("caption", "")
    })

# Also populate captions from All_Brands_Paid_Collabs.csv
with open("All_Brands_Paid_Collabs.csv", encoding="utf-8-sig") as f_csv:
    csv_caps = {r[9]: r[11] for r in list(csv.reader(f_csv))[1:]}

for r in all_master_records:
    if not r["caption"] and r["url"] in csv_caps:
        r["caption"] = csv_caps[r["url"]]

# Deduplicate by URL just in case
seen_urls = set()
unique_master_records = []
for r in all_master_records:
    if r["url"] not in seen_urls:
        seen_urls.add(r["url"])
        unique_master_records.append(r)

print(f"Total Unique Posts in Unified Dataset: {len(unique_master_records)}")

# Tier breakdown
from collections import Counter
tier_counts = Counter(r["tier"] for r in unique_master_records)
print(f"• Tier 1 (🟢 Toggle ON + 🚀 Boosted): {tier_counts[1]}")
print(f"• Tier 2 (🟢 Toggle ON + ⚪ Organic): {tier_counts[2]}")
print(f"• Tier 3 (🚀 Toggle OFF + 🚀 Boosted): {tier_counts[3]}")
print(f"• Tier 4 (⚪ Toggle OFF + ⚪ Organic / Noise): {tier_counts[4]}")

with open("unified_master_dataset.json", "w", encoding="utf-8") as f:
    json.dump(unique_master_records, f, indent=2)

print("\n✓ Saved unified_master_dataset.json")
