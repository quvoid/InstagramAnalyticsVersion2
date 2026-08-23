"""
Inspect 3 specific Instagram URLs for paid partnership status, owner, coauthors, caption, sponsor tags
"""

import sys, json
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")

urls = [
    "https://www.instagram.com/p/DHTLc4rSsK9/",
    "https://www.instagram.com/p/DEhnY2TvBZa/",
    "https://www.instagram.com/p/DYuTVFSSnc8/",
]

session = make_session()
mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

for i, u in enumerate(urls, 1):
    sc = extract_shortcode(u)
    mid = shortcode_to_id(sc)
    
    print(f"\n{'='*70}")
    print(f"[{i}] URL: {u} (Shortcode: {sc}, Media ID: {mid})")
    print(f"{'='*70}")
    
    r = session.get(f"https://i.instagram.com/api/v1/media/{mid}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=15)
    if r.status_code == 200:
        items = r.json().get("items", [])
        if items:
            item = items[0]
            owner = item.get("user", {}).get("username")
            full_name = item.get("user", {}).get("full_name")
            coauthors = [c.get("username") for c in item.get("coauthor_producers", [])]
            is_paid = item.get("is_paid_partnership", False)
            sponsors = item.get("sponsor_tags", [])
            cap = (item.get("caption") or {}).get("text", "")
            taken_at = item.get("taken_at")
            likes = item.get("like_count", 0) or 0
            comments = item.get("comment_count", 0) or 0
            
            print(f"Owner / Publisher: @{owner} ({full_name})")
            print(f"Co-authors: {coauthors}")
            print(f"Meta 'is_paid_partnership' flag: {is_paid}")
            print(f"Sponsor tags: {sponsors}")
            print(f"Likes: {likes:,} | Comments: {comments:,}")
            print(f"Caption:\n{cap}\n")
    else:
        print(f"HTTP Error {r.status_code}: {r.text[:200]}")
