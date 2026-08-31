"""
Print detailed statistics for Palmonas 1-year audit
"""

import sys, json
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

with open("palmonas_1year_4tier_dataset.json", encoding="utf-8") as f:
    posts = json.load(f)

with open("palmonas_creators_profile_metrics.json", encoding="utf-8") as f:
    profiles = json.load(f)

print("="*80)
print(f"PALMONAS (@palmonas_official) 1-YEAR DEEP AUDIT (Aug 2025 – Aug 2026)")
print(f"Total Collab Posts: {len(posts)}  |  Unique Creators: {len(profiles)}")
print("="*80)

print("\n1. 4-TIER PARTNERSHIP HIERARCHY:")
tc = Counter(p["tier_name"] for p in posts)
for t in ["Tier 1: Toggle ON + Boosted", "Tier 2: Toggle ON + Organic", "Tier 3: Toggle OFF + Boosted", "Tier 4: Toggle OFF + Organic (Noise)"]:
    print(f"  • {t:<36}: {tc.get(t, 0):>3} posts")

t123 = sum(tc.get(t, 0) for t in ["Tier 1: Toggle ON + Boosted", "Tier 2: Toggle ON + Organic", "Tier 3: Toggle OFF + Boosted"])
print(f"  💎 High-Intent Paid Media Adoption: {t123}/{len(posts)} ({t123/len(posts)*100:.1f}%)")

print("\n2. CREATOR SIZING TIERS (122 CREATORS):")
tier_c = Counter(p["creator_tier"] for p in profiles)
for t, cnt in tier_c.most_common():
    print(f"  • {t:<36}: {cnt:>3} creators")

print("\n3. TOP 10 CREATORS BY AUDIENCE REACH:")
for idx, p in enumerate(profiles[:10], 1):
    print(f"  [{idx:>2}] {p['handle']:<24} | Followers: {p['followers']:>10,} | Tier: {p['creator_tier']:<32} | Name: {p['full_name']}")

print("\n4. TOP 10 COLLABORATION POSTS BY VIEWS:")
sorted_by_v = sorted(posts, key=lambda x: x["views"], reverse=True)
for idx, p in enumerate(sorted_by_v[:10], 1):
    print(f"  [{idx:>2}] {p['tier_name']:<32} | {p['handle']:<22} | Views: {p['views']:>10,} | Likes: {p['likes']:>7,} | ER: {p['er_pct']:>5.2f}% | Genre: {p['video_genre']}")
    print(f"       URL: {p['url']}")
print("="*80)
