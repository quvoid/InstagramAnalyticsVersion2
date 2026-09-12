"""
Inspect detection methods and check sample posts for formal Meta Paid Partnership tags
"""

import json, sys

with open(r"C:/Users/omkar/.gemini/antigravity/brain/d196c916-bdf1-40a6-835a-3259489a070c/.system_generated/steps/471/output.txt", encoding="utf-8") as f:
    data = json.load(f)

rows = data["valueRanges"][0].get("values", [])[2:]

sub_brands = {"@grt.diamonds", "@grt.silverjewellery", "@grt.silverarticles", "@grtoriana"}
industry_media = {"@platinumevara", "@platinumdaysoflove", "@menofplatinum", "@eventart_india", "@chennaitimestoi"}

by_method = {}
for r in rows:
    r_padded = r + [""] * (10 - len(r))
    h, fol, l, c, er, date, url, via, cap, stat = r_padded[:10]
    m = via
    by_method.setdefault(m, []).append({"handle": h, "url": url, "date": date, "caption": cap})

for m, rlist in by_method.items():
    creators_only = [r for r in rlist if r["handle"] not in sub_brands and r["handle"] not in industry_media]
    u_urls = len(set(r["url"] for r in creators_only))
    u_handles = len(set(r["handle"] for r in creators_only))
    print(f"Detection Method: {m}")
    print(f"  Total rows: {len(rlist)}")
    print(f"  Creator rows (excl. sub-brands/media): {len(creators_only)}")
    print(f"  Unique creator URLs: {u_urls}")
    print(f"  Unique creator handles: {u_handles}")
    print()
