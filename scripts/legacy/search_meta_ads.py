"""
Search Meta Ad Library for GRT Jewellers and find boosted creator partnership ads
"""

import sys, json, re, html, time
from curl_cffi import requests

sys.stdout.reconfigure(encoding="utf-8")

session = requests.Session(impersonate="chrome120")

# 1. First search for Page ID of GRT Jewellers
headers = {
    "accept": "*/*",
    "accept-language": "en-US,en;q=0.9",
    "origin": "https://www.facebook.com",
    "referer": "https://www.facebook.com/ads/library/?active_status=all&ad_type=all&country=IN&q=GRT%20Jewellers",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "x-fb-lsd": "AVq4rT3G",
}

# Fetch search page
page_url = "https://www.facebook.com/ads/library/?active_status=all&ad_type=all&country=IN&q=GRT%20Jewellers&sort_data[direction]=desc&sort_data[mode]=relevancy_monthly_grouped&search_type=keyword_unordered&media_type=all"

print("Fetching Ad Library initial HTML...")
r = session.get(page_url, timeout=20)
print(f"Status: {r.status_code}, Length: {len(r.text)}")

# Extract LSD token and DTSG if present
lsd_m = re.search(r'name="lsd"\s+value="([^"]+)"', r.text) or re.search(r'"LSD",\[\],{"token":"([^"]+)"}', r.text)
lsd_token = lsd_m.group(1) if lsd_m else "AVq4rT3G"
print("LSD Token:", lsd_token)

# Extract collated ad cards or json payloads from SSR
ad_data = []
for script_m in re.finditer(r'<script[^>]*>(.*?)</script>', r.text, re.DOTALL):
    s_txt = script_m.group(1)
    if "adArchiveID" in s_txt or "collationGroupID" in s_txt:
        # Search for JSON blocks
        for m in re.finditer(r'\{"adArchiveID":.*?"\}', s_txt):
            try:
                ad_data.append(m.group(0)[:200])
            except Exception:
                pass

print(f"Ad data snippets found in SSR: {len(ad_data)}")

# 2. Query async search endpoint
async_url = "https://www.facebook.com/ads/library/async/search_ads/"
payload = {
    "active_status": "all",
    "ad_type": "all",
    "country": "IN",
    "q": "GRT Jewellers",
    "sort_data[direction]": "desc",
    "sort_data[mode]": "relevancy_monthly_grouped",
    "search_type": "keyword_unordered",
    "media_type": "all",
    "lsd": lsd_token,
}

r2 = session.post(async_url, data=payload, headers=headers, timeout=20)
print(f"Async search status: {r2.status_code}, Length: {len(r2.text)}")

if r2.status_code == 200:
    clean_txt = r2.text.replace("for (;;);", "").strip()
    try:
        data = json.loads(clean_txt)
        print("Async Response keys:", list(data.keys()))
        payload_data = data.get("payload", {})
        results = payload_data.get("results", [])
        print(f"Total Ad Results retrieved: {len(results)}")
        
        # Analyze ads for boosted / creator partnerships
        creator_ads = []
        for ad_group in results:
            for ad in ad_group:
                # check ad properties
                page_name = ad.get("pageName", "")
                page_id = ad.get("pageID", "")
                bylines = ad.get("bylines", "")
                publisher_platforms = ad.get("publisherPlatform", [])
                snapshot = ad.get("snapshot", {})
                body = snapshot.get("body", {}).get("text", "")
                title = snapshot.get("title", "")
                caption = snapshot.get("caption", "")
                is_partnership = bool(bylines) or "with" in page_name.lower() or "with" in title.lower() or "@" in body
                
                ad_info = {
                    "ad_id": ad.get("adArchiveID"),
                    "page_name": page_name,
                    "page_id": page_id,
                    "bylines": bylines,
                    "platforms": publisher_platforms,
                    "start_date": ad.get("startDateFormatted"),
                    "is_active": ad.get("isActive"),
                    "body_preview": body[:120],
                    "is_partnership": is_partnership
                }
                if is_partnership:
                    creator_ads.append(ad_info)
                print(f"Ad ID: {ad_info['ad_id']} | Page: {page_name} | Platforms: {publisher_platforms} | Start: {ad_info['start_date']} | Active: {ad_info['is_active']} | Partnership: {is_partnership}")
        
        print(f"\nTotal Detected Partnership / Creator Boosted Ads: {len(creator_ads)}")
    except Exception as e:
        print(f"JSON Parse Error: {e}")
        print("Sample raw text:", clean_txt[:300])
