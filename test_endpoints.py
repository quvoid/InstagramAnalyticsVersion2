"""
Test Instagram media endpoints
"""

import sys, json
from scrape_bulk import make_session, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()

# Let's test standard web GraphQL or mobile endpoints
urls_to_test = [
    ("https://i.instagram.com/api/v1/media/3969362841668769216/info/", {
        "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
        "x-ig-app-id": "1217981644879628",
    }),
    ("https://www.instagram.com/p/DcWAnlSTVXA/?__a=1&__d=dis", {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "x-ig-app-id": "936619743392459",
        "accept": "*/*",
    }),
    ("https://i.instagram.com/api/v1/clips/user/", {
        "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
        "x-ig-app-id": "936619743392459",
    })
]

for url, hdrs in urls_to_test[:2]:
    r = session.get(url, headers=hdrs, cookies=COOKIES)
    print(f"URL: {url[:60]}... -> HTTP {r.status_code} | Text: {r.text[:150]}")
