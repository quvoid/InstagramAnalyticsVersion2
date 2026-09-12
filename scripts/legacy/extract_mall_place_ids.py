"""
Extract exact Instagram Location IDs for KOPA Pune and Lulu Mall Y Junction
"""

import sys, json, re
from curl_cffi import requests as cffi_requests
from api_wrapper.client import DEFAULT_IG_COOKIES, DEFAULT_IG_HEADERS

sys.stdout.reconfigure(encoding="utf-8")
s = cffi_requests.Session(impersonate="chrome120")

handles_to_check = [
    {"handle": "kopapune", "mall": "KOPA Mall, Pune (Koregaon Park)", "city": "Pune"},
    {"handle": "lulumallhyderabad", "mall": "Lulu Mall Hyderabad (Kukatpally Y Junction)", "city": "Hyderabad"},
    {"handle": "inorbitcyberabad", "mall": "Inorbit Mall Cyberabad", "city": "Hyderabad"},
    {"handle": "phoenixavenueofstars", "mall": "Phoenix Avenue of Stars / Marketcity Pune", "city": "Pune"},
    {"handle": "phoenix_millennium", "mall": "Phoenix Mall of the Millennium Wakad", "city": "Pune"},
    {"handle": "pavillionpune", "mall": "The Pavillion Pune (SB Road)", "city": "Pune"},
    {"handle": "seasons_mall", "mall": "Seasons Mall Pune (Magarpatta)", "city": "Pune"},
    {"handle": "amanoramall", "mall": "Amanora Mall Pune (Hadapsar)", "city": "Pune"},
    {"handle": "nexus_hyderabad", "mall": "Nexus Hyderabad Mall (Kukatpally)", "city": "Hyderabad"},
    {"handle": "sarathcitycapital.hyd", "mall": "Sarath City Capital Mall (Kondapur)", "city": "Hyderabad"},
    {"handle": "gvkone", "mall": "GVK One Mall (Banjara Hills)", "city": "Hyderabad"},
]

print("="*80)
print("EXTRACTING VERIFIED LOCATION IDS VIA FEED CLIPS")
print("="*80)

master_places = []

for h in handles_to_check:
    uname = h["handle"]
    try:
        r_prof = s.get(f"https://www.instagram.com/{uname}/", timeout=10)
        pk_m = re.search(r'"profilePage_(\d+)"', r_prof.text) or re.search(r'"props":{"id":"(\d+)"', r_prof.text)
        if not pk_m:
            print(f"Could not find PK for @{uname}")
            continue
        pk = int(pk_m.group(1))
        
        # Fetch first 12 feed posts to get embedded location tags
        f_url = f"https://i.instagram.com/api/v1/feed/user/{pk}/?count=12"
        r = s.get(f_url, headers=DEFAULT_IG_HEADERS, cookies=DEFAULT_IG_COOKIES, timeout=10)
        
        found_loc = None
        if r.status_code == 200:
            items = r.json().get("items", [])
            for it in items:
                loc = it.get("location")
                if loc and loc.get("pk"):
                    found_loc = loc
                    break
                    
        # If no post location, search place api directly
        if not found_loc:
            p_url = f"https://i.instagram.com/api/v1/fbsearch/places/?query={h['mall']}"
            r_p = s.get(p_url, headers=DEFAULT_IG_HEADERS, cookies=DEFAULT_IG_COOKIES, timeout=10)
            if r_p.status_code == 200:
                p_items = r_p.json().get("items", [])
                if p_items:
                    found_loc = p_items[0].get("location")

        if found_loc:
            loc_id = found_loc.get("pk") or found_loc.get("facebook_places_id")
            loc_name = found_loc.get("name")
            addr = found_loc.get("address", "N/A")
            city = found_loc.get("city", h["city"])
            lat = found_loc.get("lat")
            lng = found_loc.get("lng")
            
            print(f"✓ {h['mall']:<45} | Loc ID: {str(loc_id):<18} | Place Name: {loc_name} | Lat, Lng: ({lat}, {lng})")
            master_places.append({
                "mall_name": h["mall"],
                "instagram_handle": f"@{uname}",
                "city": h["city"],
                "instagram_place_id": str(loc_id),
                "place_name": loc_name,
                "address": addr,
                "latitude": lat,
                "longitude": lng,
                "explore_url": f"https://www.instagram.com/explore/locations/{loc_id}/"
            })
        else:
            print(f"✗ No location found for {h['mall']}")
    except Exception as e:
        print(f"✗ Error on @{uname}: {e}")

with open("master_mall_place_ids.json", "w", encoding="utf-8") as f:
    json.dump(master_places, f, indent=2)

print("\nSaved master_mall_place_ids.json")
