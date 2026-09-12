"""
Extract and clean Facebook Page posts for GRT Jewellers with views and metrics
"""

import json, sys, re

sys.stdout.reconfigure(encoding="utf-8")

with open("grt_facebook_all_posts.json", encoding="utf-8") as f:
    raw_posts = json.load(f)

# Filter out JS bundles / non-post artifacts
clean_posts = []
seen_texts = set()

# Combine video nodes and text permalinks
for p in raw_posts:
    url = p.get("url", "")
    cap = p.get("caption", "").strip()
    views = p.get("video_views", 0)
    reacts = p.get("reactions", 0)
    dt = p.get("date", "N/A")
    pid = p.get("post_id", "")
    
    if "xx.fbcdn.net" in url or url.endswith("/about") or url.endswith("/photos") or url.endswith("/reels_tab"):
        continue
    
    if pid == "152403158122022" and not cap:
        continue
        
    clean_posts.append(p)

print(f"Total clean Facebook posts: {len(clean_posts)}")
for i, p in enumerate(clean_posts, 1):
    print(f"[{i:>2}] Date: {p['date']} | Views: {p['video_views']:>8,} | Reacts: {p['reactions']:>5,} | URL: {p['url']}")
    if p['caption']:
        print(f"     Text: {p['caption'][:120]}...")
