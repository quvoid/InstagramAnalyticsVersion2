"""
List all partner usernames for GRT Jewellers
"""

import json, sys

with open(r"C:/Users/omkar/.gemini/antigravity/brain/d196c916-bdf1-40a6-835a-3259489a070c/.system_generated/steps/471/output.txt", encoding="utf-8") as f:
    data = json.load(f)

rows = data["valueRanges"][0].get("values", [])[2:]

sub_brands = {"@grt.diamonds", "@grt.silverjewellery", "@grt.silverarticles", "@grtoriana"}
industry_media = {"@platinumevara", "@platinumdaysoflove", "@menofplatinum", "@eventart_india", "@chennaitimestoi"}

collab_creators = {}
all_creators = {}

for r in rows:
    r_padded = r + [""] * (10 - len(r))
    h, fol, l, c, er, date, url, via, cap, stat = r_padded[:10]
    if h in sub_brands or h in industry_media:
        continue
    
    entry = {"handle": h, "followers": fol, "likes": l, "comments": c, "er": er, "via": via, "last_date": date, "url": url}
    
    if h not in all_creators:
        all_creators[h] = entry
    
    if "owned by partner" in via:
        if h not in collab_creators:
            collab_creators[h] = entry

print(f"Total Creator-Owned Paid Partners: {len(collab_creators)}")
print(f"Total Overall Creator Partners: {len(all_creators)}")

print("\n--- 71 Filtered Creator-Owned Sponsored Partners ---")
for i, (h, d) in enumerate(collab_creators.items(), 1):
    print(f"{i:>2}. {h:<30} | Followers: {d['followers']:>10} | ER%: {d['er']:>6} | Date: {d['last_date']}")
