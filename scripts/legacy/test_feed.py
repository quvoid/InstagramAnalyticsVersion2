"""
Inspect user info keys
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
        for user in r.json().get("users", []):
            if user.get("username") == u:
                pk = user.get("pk")
                print(f"User: {u} (ID: {pk})")
                r_feed = session.get(f"https://i.instagram.com/api/v1/feed/user/{pk}/", headers=mob_hdrs, cookies=COOKIES)
                print(f"Feed status: {r_feed.status_code}")
                if r_feed.status_code == 200:
                    d = r_feed.json()
                    print(f"Items count: {len(d.get('items', []))}")
                    print(f"Next max ID: {d.get('next_max_id')}")
                    for it in d.get("items", [])[:3]:
                        print("  - Item code:", it.get("code"), "User:", it.get("user", {}).get("username"), "Coauthors:", [c.get("username") for c in it.get("coauthor_producers", [])])
