"""
Verify and print comprehensive audit summary across all 6 spiritual brands
"""

import sys, json
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

with open("spiritual_2year_4tier_dataset.json", encoding="utf-8") as f:
    posts = json.load(f)

with open("spiritual_creators_profile_metrics.json", encoding="utf-8") as f:
    profiles = json.load(f)

print("="*75)
print(f"SPIRITUAL & RUDRAKSHA BRANDS 2-YEAR MASTER AUDIT (Aug 2024 – Aug 2026)")
print(f"Total Posts: {len(posts):,}  |  Unique Creators: {len(profiles):,}")
print("="*75)

print("\n1. 4-TIER COLLABORATION BREAKDOWN:")
tier_counts = Counter(p["tier_name"] for p in posts)
for t, cnt in sorted(tier_counts.items()):
    print(f"  • {t:<36}: {cnt:>4} posts ({cnt/len(posts)*100:>4.1f}%)")

print("\n2. CREATOR SIZE / AUDIENCE TIER BREAKDOWN (617 Creators):")
prof_tiers = Counter(p["creator_tier"] for p in profiles)
for pt, cnt in prof_tiers.most_common():
    print(f"  • {pt:<36}: {cnt:>4} creators ({cnt/len(profiles)*100:>4.1f}%)")

print("\n3. VIDEO CONTENT GENRE BREAKDOWN (1,010 Posts):")
genre_counts = Counter(p["video_genre"] for p in posts)
for g, cnt in genre_counts.most_common():
    print(f"  • {g:<50}: {cnt:>4} posts ({cnt/len(posts)*100:>4.1f}%)")

print("\n4. BRAND-BY-BRAND BREAKDOWN:")
brands = ["Divine Hindu", "Isha Life India", "Japam", "House of Rudra", "Rudralife", "Nepa Rudraksha"]
for b in brands:
    b_posts = [p for p in posts if p["brand"] == b]
    b_creators = len(set(p["handle"].lower() for p in b_posts))
    t1 = sum(1 for p in b_posts if p["tier"] == 1)
    t2 = sum(1 for p in b_posts if p["tier"] == 2)
    t3 = sum(1 for p in b_posts if p["tier"] == 3)
    t4 = sum(1 for p in b_posts if p["tier"] == 4)
    hi = t1 + t2 + t3
    print(f"  [{b:<18}] Total: {len(b_posts):>3} posts | Creators: {b_creators:>3} | High-Intent Paid: {hi:>3} (T1:{t1} T2:{t2} T3:{t3}) | Noise T4:{t4:>3}")
print("="*75)
