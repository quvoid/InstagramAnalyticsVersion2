"""
Format Boosted Toggle ON and Boosted Toggle OFF lists for GIVA and Palmonas
"""

import sys, json

sys.stdout.reconfigure(encoding="utf-8")

with open("giva_palmonas_scraped.json", encoding="utf-8") as f:
    posts = json.load(f)

boosted_on = [p for p in posts if p.get("is_boosted") and p.get("is_paid_partnership")]
boosted_off = [p for p in posts if p.get("is_boosted") and not p.get("is_paid_partnership")]

# Sort by Views descending
boosted_on.sort(key=lambda x: x["views"], reverse=True)
boosted_off.sort(key=lambda x: x["views"], reverse=True)

print("### 🚀 Group 1: Boosted + 🟢 Toggle ON (21 Posts)\n")
print("| # | Brand | Creator Handle | Views | Likes | Like-to-View % | Direct Instagram URL |")
print("|---|---|---|---|---|---|---|")
for idx, p in enumerate(boosted_on, 1):
    b = p["brand"]
    h = p["handle"]
    v = f"{p['views']:,}"
    l = f"{p['likes']:,}"
    r = f"{p['like_rate']:.2f}%"
    u = p["url"]
    print(f"| {idx} | {b} | `{h}` | {v} | {l} | {r} | [{u}]({u}) |")

print("\n\n### 🚀 Group 2: Boosted + ⚪ Toggle OFF (27 Posts)\n")
print("| # | Brand | Creator Handle | Views | Likes | Like-to-View % | Direct Instagram URL |")
print("|---|---|---|---|---|---|---|")
for idx, p in enumerate(boosted_off, 1):
    b = p["brand"]
    h = p["handle"]
    v = f"{p['views']:,}"
    l = f"{p['likes']:,}"
    r = f"{p['like_rate']:.2f}%"
    u = p["url"]
    print(f"| {idx} | {b} | `{h}` | {v} | {l} | {r} | [{u}]({u}) |")
