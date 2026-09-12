"""
Direct Web JSON / HTML inspection of the top 3 posts to check exact Paid Partnership label & HTML tags
"""

import sys, json, re
from scrape_bulk import make_session, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()

hdrs = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "x-ig-app-id": "936619743392459",
    "accept": "*/*",
}

urls = [
    "https://www.instagram.com/p/DYoxNMDvmy_/", # Kriti
    "https://www.instagram.com/p/DYXDLF-oani/", # Aditya
    "https://www.instagram.com/p/Db8S4vusSSn/", # Devishi
    "https://www.instagram.com/p/Db8HO0PSDnH/", # Palak
    "https://www.instagram.com/p/DbLFfPHPHJ6/", # Avantika
]

for u in urls:
    r = session.get(u + "?__a=1&__d=dis", headers=hdrs, cookies=COOKIES)
    print(f"\nURL: {u} -> Status: {r.status_code}")
    if r.status_code == 200:
        try:
            d = r.json()
            # Try to find items / graphql
            items = d.get("items") or [d.get("graphql", {}).get("shortcode_media")]
            if items and items[0]:
                it = items[0]
                is_paid = it.get("is_paid_partnership")
                sponsor = it.get("sponsor_tags")
                owner = it.get("user", {}).get("username") or it.get("owner", {}).get("username")
                coauthors = [c.get("username") for c in it.get("coauthor_producers", [])]
                print(f"  Owner: @{owner}")
                print(f"  Co-authors: {coauthors}")
                print(f"  is_paid_partnership: {is_paid}")
                print(f"  sponsor_tags: {sponsor}")
        except Exception as e:
            print(f"  JSON parse error: {e}")
    else:
        # Check HTML page text
        r_html = session.get(u, headers=hdrs, cookies=COOKIES)
        print(f"  HTML Status: {r_html.status_code}")
        has_paid_label = "Paid partnership" in r_html.text or "paid_partnership" in r_html.text or "sponsor_tags" in r_html.text
        print(f"  'Paid partnership' in HTML: {has_paid_label}")
        # Search for is_paid_partnership in html
        matches = re.findall(r'"is_paid_partnership":\s*(true|false)', r_html.text, re.IGNORECASE)
        print(f"  is_paid_partnership matches in HTML: {matches}")
