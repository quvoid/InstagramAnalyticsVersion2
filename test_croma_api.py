"""
Test live API scan on 20 top posts from Electronics spreadsheet
"""

import sys, json, time
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

with open("croma_raw_extracted_posts.json", encoding="utf-8") as f:
    posts = json.load(f)

# Sort by likes
posts.sort(key=lambda x: x["likes"], reverse=True)

print("Testing top 20 electronics posts for live API Toggle and Views:\n")

for idx, p in enumerate(posts[:20], 1):
    u = p["url"]
    sc = extract_shortcode(u)
    mid = shortcode_to_id(sc)
    
    try:
        r = session.get(f"https://i.instagram.com/api/v1/media/{mid}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=8)
        if r.status_code == 200:
            it = r.json().get("items", [])[0]
            is_paid = bool(it.get("is_paid_partnership", False))
            play_cnt = it.get("play_count") or it.get("view_count") or 0
            likes = it.get("like_count") or p["likes"]
            print(f"[{idx:>2}] {p['brand']:<18} | {p['handle']:<20} | Toggle: {'🟢 ON' if is_paid else '⚪ OFF'} | Views: {play_cnt:>10,} | Likes: {likes:>8,}")
        else:
            print(f"[{idx:>2}] {p['brand']:<18} | {p['handle']:<20} | HTTP {r.status_code}")
    except Exception as e:
        print(f"[{idx:>2}] Error: {e}")
    time.sleep(0.3)
