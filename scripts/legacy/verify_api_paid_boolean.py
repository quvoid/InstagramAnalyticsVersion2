"""
Check exact is_paid_partnership boolean from Instagram API for candidate posts
"""

import sys, json
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()
mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

test_urls = [
    ("Prithi Narayanan", "https://www.instagram.com/p/DY4dm7XTFkk/"),
    ("R. Ashwin", "https://www.instagram.com/p/DZPtsIdC9jP/"),
    ("Aarya Lakshmi", "https://www.instagram.com/p/DMHbNBkTbYN/"),
    ("Chaitra Vasudevan", "https://www.instagram.com/p/DHTLc4rSsK9/"),
]

print("--- Live Instagram API is_paid_partnership Verification ---\n")

for name, u in test_urls:
    sc = extract_shortcode(u)
    mid = shortcode_to_id(sc)
    r = session.get(f"https://i.instagram.com/api/v1/media/{mid}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=12)
    if r.status_code == 200:
        item = r.json().get("items", [])[0]
        is_paid = item.get("is_paid_partnership", False)
        coauthors = [c.get("username") for c in item.get("coauthor_producers", [])]
        print(f"Name: {name:<20} | URL: {u}")
        print(f"  • is_paid_partnership boolean: {is_paid}")
        print(f"  • Co-authors: {coauthors}")
        print()
    else:
        print(f"Error {r.status_code} for {name}")
