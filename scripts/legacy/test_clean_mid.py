"""
Test clean numeric media ID
"""

import sys, json
from scrape_bulk import make_session, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

with open("giva_palmonas_scraped.json", encoding="utf-8") as f:
    posts = json.load(f)

for p in posts[:8]:
    raw_mid = str(p["media_id"])
    clean_mid = raw_mid.split("_")[0]
    r = session.get(f"https://i.instagram.com/api/v1/media/{clean_mid}/info/", headers=mob_hdrs, cookies=COOKIES)
    print(f"Clean MID: {clean_mid} ({p['url']}) -> HTTP {r.status_code}", flush=True)
    if r.status_code == 200:
        it = r.json().get("items", [])[0]
        print(f"  Brand: {p['brand']} | Creator: {p['handle']}", flush=True)
        print(f"  is_paid_partnership: {it.get('is_paid_partnership')}", flush=True)
        print(f"  sponsor_tags: {it.get('sponsor_tags')}", flush=True)
        print(f"  coauthors: {[c.get('username') for c in it.get('coauthor_producers', [])]}", flush=True)
