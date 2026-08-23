"""
Inspect Croma dataset and check likes, followers, ER, and boost signals
"""

import sys, json
sys.stdout.reconfigure(encoding="utf-8")

with open("croma_raw_extracted_posts.json", encoding="utf-8") as f:
    posts = json.load(f)

print(f"Total posts: {len(posts)}")

# Sort by likes descending to see top performing creator posts
posts.sort(key=lambda x: x["likes"], reverse=True)

print("\nTop 15 Creator Collabs by Likes across all 10 Electronics Brands:")
for i, p in enumerate(posts[:15], 1):
    print(f"[{i:>2}] {p['brand']:<18} | {p['handle']:<22} | Fols: {p['followers']:>10,} | Likes: {p['likes']:>8,} | Date: {p['post_date']}")
    print(f"     URL: {p['url']}")
