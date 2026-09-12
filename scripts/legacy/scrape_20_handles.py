"""
Scrape exact live follower counts for the 20 handles
"""

import sys, json, time, re
from scrape_bulk import make_session, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()

handles = [
    "capsulindia", "amazonfashionin", "bombaysweetshop", "casabacardiin",
    "cmf.tech", "districtupdates", "evaxfried", "indiansneakerfestival",
    "kanchhiiii", "kommunedelhincr", "medusaindia", "kauraverse",
    "leada.in", "niviasports", "myntra", "rahasyafragrances",
    "parvaazmusic", "skecherscricket", "thethirdspacedelhi", "yuzenmatcha"
]

web_hdrs = {
    "accept": "*/*",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "x-ig-app-id": "936619743392459",
}

results = {}

for idx, h in enumerate(handles, 1):
    fols = 0
    fn = h
    ver = False
    
    # 1. Try web profile info
    try:
        r = session.get(f"https://www.instagram.com/api/v1/users/web_profile_info/?username={h}", headers={**web_hdrs, "referer": f"https://www.instagram.com/{h}/"}, cookies=COOKIES, timeout=10)
        if r.status_code == 200:
            u = r.json().get("data", {}).get("user", {})
            fols = u.get("edge_followed_by", {}).get("count", 0)
            fn = u.get("full_name") or h
            ver = bool(u.get("is_verified", False))
    except Exception:
        pass
        
    # 2. Try mobile search/info if 0
    if not fols:
        try:
            r = session.get(f"https://i.instagram.com/api/v1/users/search/?q={h}", headers={"User-Agent": "Instagram 269.0.0.18.75 Android", "x-ig-app-id": "936619743392459"}, cookies=COOKIES, timeout=8)
            if r.status_code == 200:
                for u in r.json().get("users", []):
                    if u.get("username", "").lower() == h.lower():
                        pk = u.get("pk")
                        r2 = session.get(f"https://i.instagram.com/api/v1/users/{pk}/info/", headers={"User-Agent": "Instagram 269.0.0.18.75 Android", "x-ig-app-id": "936619743392459"}, cookies=COOKIES, timeout=8)
                        if r2.status_code == 200:
                            ud = r2.json().get("user", {})
                            fols = ud.get("follower_count") or u.get("follower_count", 0)
                            fn = ud.get("full_name") or u.get("full_name", h)
                            ver = bool(ud.get("is_verified", False))
                        break
        except Exception:
            pass
            
    # 3. Try public HTML scrape if still 0
    if not fols:
        try:
            r = session.get(f"https://www.instagram.com/{h}/", headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}, cookies=COOKIES, timeout=10)
            if r.status_code == 200:
                m = re.search(r'([0-9,.]+[KMB]?)\s+Followers', r.text, re.IGNORECASE)
                if m:
                    raw_str = m.group(1).replace(",", "")
                    if "M" in raw_str:
                        fols = int(float(raw_str.replace("M", "")) * 1000000)
                    elif "K" in raw_str:
                        fols = int(float(raw_str.replace("K", "")) * 1000)
                    else:
                        fols = int(float(raw_str))
        except Exception:
            pass

    results[h] = {"followers": fols, "full_name": fn, "verified": ver}
    print(f"[{idx:>2}/{len(handles)}] @{h:<24} -> Followers: {fols:>10,} | Name: {fn}")
    time.sleep(0.3)

with open("resolved_20_handles.json", "w", encoding="utf-8") as f:
    json.dump(results, f, indent=2)
