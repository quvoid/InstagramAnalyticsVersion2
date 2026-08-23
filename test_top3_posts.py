"""
Inspect top 3 posts in detail
"""

import sys, json
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

urls = [
    "https://www.instagram.com/p/DYoxNMDvmy_/", # Kriti Sanon
    "https://www.instagram.com/p/DYXDLF-oani/", # Aditya Roy Kapur
    "https://www.instagram.com/p/Db8S4vusSSn/", # Devishi Madaan
]

for u in urls:
    sc = extract_shortcode(u)
    mid = shortcode_to_id(sc)
    r = session.get(f"https://i.instagram.com/api/v1/media/{mid}/info/", headers=mob_hdrs, cookies=COOKIES)
    print(f"URL: {u} (ID: {mid}) -> HTTP {r.status_code}", flush=True)
    if r.status_code == 200:
        items = r.json().get("items", [])
        if items:
            it = items[0]
            print(f"  Owner: @{it.get('user', {}).get('username')}", flush=True)
            print(f"  Coauthors: {[c.get('username') for c in it.get('coauthor_producers', [])]}", flush=True)
            print(f"  is_paid_partnership: {it.get('is_paid_partnership')}", flush=True)
            print(f"  sponsor_tags: {it.get('sponsor_tags')}", flush=True)
            print(f"  caption: {it.get('caption', {}).get('text', '')[:100]}...", flush=True)
        else:
            print("  No items in JSON", flush=True)
