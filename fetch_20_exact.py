"""
Test mobile search on 20 handles
"""

import sys, json, time
from scrape_bulk import make_session, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

handles = [
    "capsulindia", "amazonfashionin", "bombaysweetshop", "casabacardiin",
    "cmf.tech", "districtupdates", "evaxfried", "indiansneakerfestival",
    "kanchhiiii", "kommunedelhincr", "medusaindia", "kauraverse",
    "leada.in", "niviasports", "myntra", "rahasyafragrances",
    "parvaazmusic", "skecherscricket", "thethirdspacedelhi", "yuzenmatcha"
]

results = {}

for h in handles:
    r = session.get(f"https://i.instagram.com/api/v1/users/search/?q={h}", headers=mob_hdrs, cookies=COOKIES, timeout=10)
    if r.status_code == 200:
        users = r.json().get("users", [])
        matched = False
        for u in users:
            if u.get("username", "").lower() == h.lower():
                matched = True
                pk = u.get("pk")
                # get full info
                r2 = session.get(f"https://i.instagram.com/api/v1/users/{pk}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=10)
                if r2.status_code == 200:
                    ud = r2.json().get("user", {})
                    fols = ud.get("follower_count") or u.get("follower_count", 0)
                    fn = ud.get("full_name") or u.get("full_name", h)
                    ver = bool(ud.get("is_verified", False))
                    posts = ud.get("media_count", 0)
                    fing = ud.get("following_count", 0)
                    results[h] = {
                        "followers": fols,
                        "full_name": fn,
                        "verified": ver,
                        "total_posts": posts,
                        "following": fing
                    }
                    print(f"✓ @{h:<24} -> Followers: {fols:>10,} | Name: {fn}")
                    break
        if not matched and users:
            u = users[0]
            pk = u.get("pk")
            r2 = session.get(f"https://i.instagram.com/api/v1/users/{pk}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=10)
            if r2.status_code == 200:
                ud = r2.json().get("user", {})
                fols = ud.get("follower_count") or u.get("follower_count", 0)
                fn = ud.get("full_name") or u.get("full_name", h)
                ver = bool(ud.get("is_verified", False))
                posts = ud.get("media_count", 0)
                fing = ud.get("following_count", 0)
                results[h] = {
                    "followers": fols,
                    "full_name": fn,
                    "verified": ver,
                    "total_posts": posts,
                    "following": fing
                }
                print(f"✓ @{h:<24} (top match @{u.get('username')}) -> Followers: {fols:>10,} | Name: {fn}")
    else:
        print(f"[-] @{h} HTTP {r.status_code}")
    time.sleep(0.4)

with open("resolved_20_exact_metrics.json", "w", encoding="utf-8") as f:
    json.dump(results, f, indent=2)

print("\n✓ Finished fetching exact metrics for all 20 handles!")
