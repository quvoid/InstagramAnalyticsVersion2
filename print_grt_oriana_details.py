"""
Print full audit details for GRT Oriana
"""

import sys, json
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

with open("grt_oriana_2year_4tier_dataset.json", encoding="utf-8") as f:
    posts = json.load(f)

with open("grt_oriana_creators_profile_metrics.json", encoding="utf-8") as f:
    profiles = json.load(f)

print("="*75)
print(f"GRT ORIANA (@grtoriana) 2-YEAR MASTER AUDIT (Aug 2024 – Aug 2026)")
print(f"Total Collab Posts: {len(posts)}  |  Unique Creators: {len(profiles)}")
print("="*75)

print("\n1. CREATORS PROFILE METRICS:")
for p in profiles:
    print(f"  • {p['handle']:<22} | Followers: {p['followers']:>10,} | Tier: {p['creator_tier']:<32} | Name: {p['full_name']}")

print("\n2. 4-TIER BREAKDOWN:")
tc = Counter(p["tier_name"] for p in posts)
for t, cnt in tc.items():
    print(f"  • {t:<35}: {cnt} posts")

print("\n3. ALL COLLABORATION POSTS:")
for idx, p in enumerate(posts, 1):
    print(f"  [{idx}] {p['tier_name']:<32} | {p['handle']:<22} | Date: {p['date']} | Views: {p['views']:>8,} | Likes: {p['likes']:>6,} | ER: {p['er_pct']:>5.2f}% | Genre: {p['video_genre']}")
    print(f"      URL: {p['url']}")
    print(f"      Caption: {p['caption'][:100]}...")
print("="*75)
