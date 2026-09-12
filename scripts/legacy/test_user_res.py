"""
Test user ID resolution and feed extraction for giva.co and palmonas_official
"""

import sys, json, re
from scrape_bulk import make_session, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

for u in ["giva.co", "palmonas_official"]:
    r = session.get(f"https://i.instagram.com/api/v1/users/search/?q={u}", headers=mob_hdrs, cookies=COOKIES)
    print(f"Search {u} -> Status: {r.status_code}")
    if r.status_code == 200:
        users = r.json().get("users", [])
        for user in users[:3]:
            username = user.get("username")
            pk = user.get("pk")
            fn = user.get("full_name")
            print(f"  • @{username} (ID: {pk}, Name: {fn})")
            if username == u:
                # Test fetching user info endpoint
                r_info = session.get(f"https://i.instagram.com/api/v1/users/{pk}/info/", headers=mob_hdrs, cookies=COOKIES)
                if r_info.status_code == 200:
                    u_obj = r_info.json().get("user", {})
                    fols = u_obj.get("follower_count")
                    posts = u_obj.get("media_count")
                    print(f"    Followers: {fols:,} | Posts: {posts:,}")
                    
                # Test fetching feed
                r_feed = session.get(f"https://i.instagram.com/api/v1/feed/user/{pk}/", headers=mob_hdrs, cookies=COOKIES)
                if r_feed.status_code == 200:
                    items = r_feed.json().get("items", [])
                    print(f"    Feed items returned: {len(items)}")
                    for it in items[:2]:
                        code = it.get("code")
                        co = [c.get("username") for c in it.get("coauthor_producers", [])]
                        print(f"      - {code} | Co-authors: {co}")
