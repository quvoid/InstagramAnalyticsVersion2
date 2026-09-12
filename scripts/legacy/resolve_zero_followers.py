"""
Check creators with 0 followers and scrape their follower counts
"""

import sys, json, time
from scrape_bulk import make_session, COOKIES

sys.stdout.reconfigure(encoding="utf-8")

with open("footwear_creators_profile_metrics.json", encoding="utf-8") as f:
    creators = json.load(f)

zero_fol_creators = [c for c in creators if c["followers"] == 0]
print(f"Total creators with 0 followers: {len(zero_fol_creators)}")

session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

for idx, c in enumerate(zero_fol_creators, 1):
    h = c["raw_handle"]
    try:
        # Search API to get pk
        r = session.get(f"https://i.instagram.com/api/v1/users/search/?q={h}", headers=mob_hdrs, cookies=COOKIES, timeout=8)
        if r.status_code == 200:
            for u in r.json().get("users", []):
                if u.get("username", "").lower() == h.lower():
                    pk = u.get("pk")
                    fn = u.get("full_name") or h
                    ver = bool(u.get("is_verified", False))
                    fols = u.get("follower_count", 0)
                    
                    # Fetch /info/ for exact count
                    r2 = session.get(f"https://i.instagram.com/api/v1/users/{pk}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=8)
                    if r2.status_code == 200:
                        ud = r2.json().get("user", {})
                        fols = ud.get("follower_count") or fols
                        fn = ud.get("full_name") or fn
                        ver = bool(ud.get("is_verified", ver))
                        
                    c["followers"] = fols
                    c["full_name"] = fn
                    c["verified"] = ver
                    c["scrape_status"] = "Scraped OK (Resolved)"
                    print(f"[{idx:>2}/{len(zero_fol_creators)}] ✓ @{h:<22} -> Followers: {fols:>9,} | Name: {fn}")
                    break
            else:
                print(f"[{idx:>2}/{len(zero_fol_creators)}] ⚠ @{h:<22} -> Not matched in search")
        else:
            print(f"[{idx:>2}/{len(zero_fol_creators)}] ⚠ @{h:<22} -> HTTP {r.status_code}")
    except Exception as e:
        print(f"[{idx:>2}/{len(zero_fol_creators)}] ⚠ @{h:<22} -> Error: {e}")
    time.sleep(0.3)

with open("footwear_creators_profile_metrics.json", "w", encoding="utf-8") as f:
    json.dump(creators, f, indent=2)

print("\n✓ Finished resolving follower counts!")
