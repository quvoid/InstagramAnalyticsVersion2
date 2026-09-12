"""
Fetch OpenGraph description for the 20 handles
"""

import sys, json, time, re
from curl_cffi import requests

sys.stdout.reconfigure(encoding="utf-8")

handles = [
    "capsulindia", "amazonfashionin", "bombaysweetshop", "casabacardiin",
    "cmf.tech", "districtupdates", "evaxfried", "indiansneakerfestival",
    "kanchhiiii", "kommunedelhincr", "medusaindia", "kauraverse",
    "leada.in", "niviasports", "myntra", "rahasyafragrances",
    "parvaazmusic", "skecherscricket", "thethirdspacedelhi", "yuzenmatcha"
]

s = requests.Session(impersonate="chrome120")
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1"
}

results = {}

for h in handles:
    fols = 0
    fn = h
    try:
        r = s.get(f"https://www.instagram.com/{h}/", headers=headers, timeout=10)
        # Search for: <meta property="og:description" content="12K Followers, 450 Following, 120 Posts ...
        m = re.search(r'([0-9.,KMBkmb]+)\s+Followers,\s*([0-9.,KMBkmb]+)\s+Following,\s*([0-9.,KMBkmb]+)\s+Posts', r.text)
        if m:
            raw = m.group(1).upper().replace(",", "")
            if "M" in raw: fols = int(float(raw.replace("M", "")) * 1000000)
            elif "K" in raw: fols = int(float(raw.replace("K", "")) * 1000)
            else: fols = int(float(raw))
        
        # Search for full name in title
        m2 = re.search(r'<title>([^(<]+)\s*\(@', r.text)
        if m2:
            fn = m2.group(1).strip()
    except Exception as e:
        print(f"Error {h}: {e}")
        
    print(f"@{h:<24} -> Followers: {fols:>10,} | Name: {fn}")
    results[h] = {"followers": fols, "full_name": fn}
    time.sleep(0.3)

with open("og_20_results.json", "w", encoding="utf-8") as f:
    json.dump(results, f, indent=2)
