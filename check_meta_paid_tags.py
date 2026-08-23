"""
Check sample creator posts for formal 'Paid Partnership' label and sponsored hashtags
"""

import json, sys
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

with open(r"C:/Users/omkar/.gemini/antigravity/brain/d196c916-bdf1-40a6-835a-3259489a070c/.system_generated/steps/471/output.txt", encoding="utf-8") as f:
    data = json.load(f)

rows = data["valueRanges"][0].get("values", [])[2:]

sub_brands = {"@grt.diamonds", "@grt.silverjewellery", "@grt.silverarticles", "@grtoriana"}
industry_media = {"@platinumevara", "@platinumdaysoflove", "@menofplatinum", "@eventart_india", "@chennaitimestoi"}

creator_posts = [r for r in rows if r[0] not in sub_brands and r[0] not in industry_media and len(r) > 7 and "owned by partner" in r[7]]

print(f"Total 'Post owned by partner' creator rows: {len(creator_posts)}")
print(f"Sample testing first 10 posts for Meta Paid Partnership handshake / tags:")

session = make_session()
mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

for i, r in enumerate(creator_posts[:10], 1):
    handle = r[0]
    url = r[6]
    sc = extract_shortcode(url)
    mid = shortcode_to_id(sc)
    
    r_api = session.get(f"https://i.instagram.com/api/v1/media/{mid}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=12)
    if r_api.status_code == 200:
        items = r_api.json().get("items", [])
        if items:
            item = items[0]
            is_paid = item.get("is_paid_partnership", False)
            sponsors = item.get("sponsor_tags", [])
            coauthors = [c.get("username") for c in item.get("coauthor_producers", [])]
            cap = (item.get("caption") or {}).get("text", "")
            has_ad_tag = any(t in cap.lower() for t in ["#ad", "#collab", "#paidpartnership", "#sponsored", "grt jewellers"])
            print(f"[{i}] {handle:<24} | is_paid_partnership: {is_paid} | sponsors: {sponsors} | coauthors: {coauthors} | #ad/#collab in cap: {has_ad_tag}")
    else:
        print(f"[{i}] {handle} API status: {r_api.status_code}")
