"""
Test and resolve PKs for 10 Pune & Hyderabad Malls
"""

import sys, json, re
from curl_cffi import requests as cffi_requests
from api_wrapper.client import DEFAULT_IG_COOKIES, DEFAULT_IG_HEADERS

sys.stdout.reconfigure(encoding="utf-8")

malls = [
    {"name": "Phoenix Avenue of Stars Pune", "handle": "phoenixavenueofstars", "city": "Pune"},
    {"name": "Seasons Mall Pune", "handle": "seasons_mall", "city": "Pune"},
    {"name": "The Pavillion Pune", "handle": "pavillionpune", "city": "Pune"},
    {"name": "Phoenix Mall of the Millennium Wakad", "handle": "phoenix_millennium", "city": "Pune"},
    {"name": "Amanora Mall Pune", "handle": "amanoramall", "city": "Pune"},
    {"name": "Nexus Hyderabad Mall", "handle": "nexus_hyderabad", "city": "Hyderabad"},
    {"name": "Sarath City Capital Mall Hyderabad", "handle": "sarathcitycapital.hyd", "city": "Hyderabad"},
    {"name": "Inorbit Mall Cyberabad", "handle": "inorbitcyberabad", "city": "Hyderabad"},
    {"name": "Lulu Mall Hyderabad", "handle": "lulumallhyderabad", "city": "Hyderabad"},
    {"name": "GVK One Mall Hyderabad", "handle": "gvkone", "city": "Hyderabad"},
]

s = cffi_requests.Session(impersonate="chrome120")

print("="*70)
print("RESOLVING PKS & PROFILE INFO FOR 10 MALLS")
print("="*70)

resolved = []
for m in malls:
    h = m["handle"].lower().strip()
    try:
        r = s.get(f"https://www.instagram.com/{h}/", timeout=12)
        pk_m = re.search(r'"profilePage_(\d+)"', r.text) or re.search(r'"props":{"id":"(\d+)"', r.text) or re.search(r'"user_id":"(\d+)"', r.text)
        fols_m = re.search(r'([0-9.,KMBkmb]+)\s+Followers', r.text)
        
        pk = int(pk_m.group(1)) if pk_m else None
        fols = fols_m.group(1) if fols_m else "N/A"
        
        print(f"✓ {m['name']:<40} | @{h:<24} | PK: {str(pk):<12} | Followers: {fols}")
        m["pk"] = pk
        m["followers"] = fols
        resolved.append(m)
    except Exception as e:
        print(f"✗ Error on {h}: {e}")

with open("malls_resolved_meta.json", "w", encoding="utf-8") as f:
    json.dump(resolved, f, indent=2)

print("\nSaved malls_resolved_meta.json")
