"""
Verify all 20 handles and pure tier breakdown
"""

import sys, json
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

with open("footwear_creators_profile_metrics.json", encoding="utf-8") as f:
    profiles = json.load(f)

print(f"Total Profiles in Database: {len(profiles)}")
tc = Counter(p["creator_tier"] for p in profiles)
print("\nPure Creator Tier Breakdown (All 91 Profiles):")
for k, v in tc.most_common():
    print(f"  • {k:<35}: {v:>2} creators")

handles_20 = [
    "capsulindia", "amazonfashionin", "bombaysweetshop", "casabacardiin",
    "cmf.tech", "districtupdates", "evaxfried", "indiansneakerfestival",
    "kanchhiiii", "kommunedelhincr", "medusaindia", "kauraverse",
    "leada.in", "niviasports", "myntra", "rahasyafragrances",
    "parvaazmusic", "skecherscricket", "thethirdspacedelhi", "yuzenmatcha"
]

print("\nExact Followers & Pure Tiers for the 20 requested handles:")
for p in profiles:
    if p["raw_handle"] in handles_20:
        print(f"  @{p['raw_handle']:<24} -> Followers: {p['followers']:>10,} | Tier: {p['creator_tier']}")
