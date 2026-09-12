"""
Inspect user object from mobile API for giva.co and palmonas_official
"""

import sys, json
from scrape_bulk import make_session, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

for u in ["giva.co", "palmonas_official"]:
    r = session.get(f"https://i.instagram.com/api/v1/users/search/?q={u}", headers=mob_hdrs, cookies=COOKIES)
    if r.status_code == 200:
        users = r.json().get("users", [])
        for user in users:
            if user.get("username") == u:
                pk = user.get("pk")
                print(f"\n[+] Found {u} -> User ID: {pk}")
                r_info = session.get(f"https://i.instagram.com/api/v1/users/{pk}/info/", headers=mob_hdrs, cookies=COOKIES)
                if r_info.status_code == 200:
                    u_obj = r_info.json().get("user", {})
                    print(f"    Full Name: {u_obj.get('full_name')}")
                    print(f"    Followers: {u_obj.get('follower_count') or u_obj.get('follower_count_raw')}")
                    print(f"    Media Count: {u_obj.get('media_count')}")
                    print(f"    Biography: {u_obj.get('biography')[:100]}...")
                
                # Fetch feed
                r_feed = session.get(f"https://i.instagram.com/api/v1/feed/user/{pk}/", headers=mob_hdrs, cookies=COOKIES)
                if r_feed.status_code == 200:
                    items = r_feed.json().get("items", [])
                    print(f"    Feed items: {len(items)}")
                    for item in items[:5]:
                        code = item.get("code")
                        coauthors = [c.get("username") for c in item.get("coauthor_producers", [])]
                        is_paid = item.get("is_paid_partnership", False)
                        print(f"      - {code} | Coauthors: {coauthors} | is_paid: {is_paid}")
