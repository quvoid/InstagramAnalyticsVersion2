"""
Fetch full caption and is_paid_partnership flag from Instagram API for sample posts
"""

import sys, json, time
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()
mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

with open("All_Brands_Paid_Collabs_With_Toggle.csv", encoding="utf-8-sig") as f:
    import csv
    rows = list(csv.reader(f))[1:]

print("Testing 15 sample posts for full captions and Meta is_paid_partnership status:\n")

for i, r in enumerate(rows[:15], 1):
    brand = r[1]
    handle = r[2]
    url = r[9]
    sc = extract_shortcode(url)
    mid = shortcode_to_id(sc)
    
    r_api = session.get(f"https://i.instagram.com/api/v1/media/{mid}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=12)
    if r_api.status_code == 200:
        item = r_api.json().get("items", [])[0]
        is_paid = item.get("is_paid_partnership", False)
        coauthors = [c.get("username") for c in item.get("coauthor_producers", [])]
        cap = (item.get("caption") or {}).get("text", "")
        
        # Check for disclosures anywhere in the FULL text
        cap_l = cap.lower()
        has_disclosure = is_paid or any(t in cap_l for t in ["#ad", "#collab", "#sponsored", "#paidpartnership", "in collaboration with", "collab with"])
        
        print(f"[{i:>2}] {brand:<25} | {handle:<22} | is_paid_partnership: {str(is_paid):<5} | Disclosure in full cap: {has_disclosure}")
        # print hashtags found
        tags = [t for t in cap.split() if t.startswith("#")]
        print(f"     Hashtags ({len(tags)}): {' '.join(tags[:8])}")
    else:
        print(f"[{i:>2}] {brand:<25} | {handle} | HTTP Error {r_api.status_code}")
    time.sleep(0.8)
