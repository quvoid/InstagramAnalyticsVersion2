"""
Test profile scraping methods on 5 creator handles
"""

import sys, json, time
from scrape_bulk import make_session, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

web_hdrs = {
    "accept": "*/*",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "x-ig-app-id": "936619743392459",
}

handles = ["surya_14kumar", "kartikaaryan", "thevishnukaushal", "winonakicks", "itskivitime"]

for h in handles:
    print(f"Testing @{h}...")
    try:
        # Web profile info
        r = session.get(f"https://www.instagram.com/api/v1/users/web_profile_info/?username={h}", headers={**web_hdrs, "referer": f"https://www.instagram.com/{h}/"}, cookies=COOKIES, timeout=10)
        if r.status_code == 200:
            u = r.json().get("data", {}).get("user", {})
            fols = u.get("edge_followed_by", {}).get("count", 0)
            fn = u.get("full_name")
            ver = u.get("is_verified")
            posts = u.get("edge_owner_to_timeline_media", {}).get("count", 0)
            print(f"  ✓ [Web API] @{h}: Fols={fols:,} | Name='{fn}' | Verified={ver} | Posts={posts}")
        else:
            print(f"  [-] Web API returned {r.status_code}")
    except Exception as e:
        print(f"  [-] Error: {e}")
    time.sleep(0.5)
