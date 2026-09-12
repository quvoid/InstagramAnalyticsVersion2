"""
Fast fetch follower counts for the 20 handles using public web profiles & mobile API
"""

import sys, json, time, re
import requests

sys.stdout.reconfigure(encoding="utf-8")

handles = [
    "capsulindia", "amazonfashionin", "bombaysweetshop", "casabacardiin",
    "cmf.tech", "districtupdates", "evaxfried", "indiansneakerfestival",
    "kanchhiiii", "kommunedelhincr", "medusaindia", "kauraverse",
    "leada.in", "niviasports", "myntra", "rahasyafragrances",
    "parvaazmusic", "skecherscricket", "thethirdspacedelhi", "yuzenmatcha"
]

results = {}

s = requests.Session()
s.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
})

for idx, h in enumerate(handles, 1):
    fols = 0
    fn = h
    ver = False
    
    try:
        r = s.get(f"https://www.instagram.com/{h}/", timeout=8)
        if r.status_code == 200:
            # Look for meta description: "<meta content="53K Followers, 120 Following, 432 Posts..." name="description" />"
            m = re.search(r'content="([0-9.,KMBkmb]+)\s+Followers,\s*([0-9.,KMBkmb]+)\s+Following,\s*([0-9.,KMBkmb]+)\s+Posts', r.text)
            if m:
                raw_fols = m.group(1).upper().replace(",", "")
                if "M" in raw_fols:
                    fols = int(float(raw_fols.replace("M", "")) * 1000000)
                elif "K" in raw_fols:
                    fols = int(float(raw_fols.replace("K", "")) * 1000)
                else:
                    fols = int(float(raw_fols))
            
            # Look for title or full name
            m_title = re.search(r'<title>([^(<]+)\s*\(@', r.text)
            if m_title:
                fn = m_title.group(1).strip()
    except Exception as e:
        print(f"Error for @{h}: {e}")
        
    print(f"[{idx:>2}/{len(handles)}] @{h:<24} -> Followers: {fols:>10,} | Name: {fn}")
    results[h] = {"followers": fols, "full_name": fn, "verified": ver}
    time.sleep(0.5)

with open("resolved_20_handles.json", "w", encoding="utf-8") as f:
    json.dump(results, f, indent=2)
