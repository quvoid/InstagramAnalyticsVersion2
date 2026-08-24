"""
Inspect 1-Year Footwear Scrape Results
"""

import sys, json, openpyxl

sys.stdout.reconfigure(encoding="utf-8")

with open("footwear_1year_4tier_dataset.json", encoding="utf-8") as f:
    posts = json.load(f)

print(f"Total 1-Year Collab Posts: {len(posts)}")

wb = openpyxl.load_workbook("footwear_sneaker_brands_master_analysis.xlsx", data_only=True)
print(f"Total Sheets in Master Excel: {len(wb.sheetnames)}")

ws_sum = wb["Executive Summary"]
print("\nExecutive Summary Table:")
for r in range(2, 6):
    row_vals = [ws_sum.cell(row=r, column=c).value for c in range(1, 16)]
    print(f"  Row {r}: {row_vals}")

print("\nPer-Brand Breakdown:")
for b in ["Skechers India", "Gully Labs", "Comet"]:
    b_posts = [p for p in posts if p["brand"] == b]
    u_c = len(set(p["handle"].lower() for p in b_posts))
    t1 = [p for p in b_posts if p["tier"] == 1]
    t2 = [p for p in b_posts if p["tier"] == 2]
    t3 = [p for p in b_posts if p["tier"] == 3]
    t4 = [p for p in b_posts if p["tier"] == 4]
    
    dates = [p["date"] for p in b_posts if p["date"] != "N/A"]
    date_range = f"{min(dates)} to {max(dates)}" if dates else "N/A"
    
    print(f"\n👟 {b} (Date Range: {date_range}):")
    print(f"   • Total Posts: {len(b_posts)} | Unique Creators: {u_c}")
    print(f"   • 🟢 Tier 1 (Toggle ON + Boosted): {len(t1)}")
    print(f"   • 🟢 Tier 2 (Toggle ON + Organic): {len(t2)}")
    print(f"   • 🚀 Tier 3 (Toggle OFF + Boosted): {len(t3)}")
    print(f"   • ⚪ Tier 4 (Noise / Unboosted):    {len(t4)}")
    print(f"   • 💎 High-Intent Paid Total:        {len(t1)+len(t2)+len(t3)} ({(len(t1)+len(t2)+len(t3))/len(b_posts)*100:.1f}%)")
    
    print("   • Top 5 Collaborations (by Views / Reach):")
    b_posts_sorted = sorted(b_posts, key=lambda x: x["views"], reverse=True)
    for p in b_posts_sorted[:5]:
        print(f"     - {p['tier_name'][:6]} | {p['handle']:<20} | Views: {p['views']:>10,} | Likes: {p['likes']:>8,} | Date: {p['date']} | URL: {p['url']}")
