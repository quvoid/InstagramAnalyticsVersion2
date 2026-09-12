"""
Extract clean Facebook post URLs, dates, and campaign text for GRT Jewellers
"""

import json, sys

with open("grt_facebook_all_posts.json", encoding="utf-8") as f:
    posts = json.load(f)

clean = []
seen = set()

for p in posts:
    url = p.get("url", "")
    dt = p.get("date", "N/A")
    vw = p.get("video_views", 0)
    rc = p.get("reactions", 0)
    cap = p.get("caption", "").strip().replace("\n", " ")
    pid = p.get("post_id", "")
    
    if not url or "facebook.com" not in url or "xx.fbcdn.net" in url:
        continue
    if url.endswith("/about") or url.endswith("/photos") or url.endswith("/reels_tab"):
        continue
    if pid == "152403158122022" and not cap:
        continue
        
    # Clean up URL format
    if "posts/Uzpf" in url:
        # permalink with base64 ID
        pass
        
    if url not in seen:
        seen.add(url)
        clean.append({
            "date": dt,
            "url": url,
            "views": vw,
            "reactions": rc,
            "caption": cap
        })

print(f"Total clean Facebook URLs: {len(clean)}")
for i, p in enumerate(clean, 1):
    print(f"\n[{i}] Date: {p['date']} | Views: {p['views']:,} | Reacts: {p['reactions']:,}")
    print(f"    URL: {p['url']}")
    if p['caption']:
        print(f"    Text: {p['caption'][:150]}...")
