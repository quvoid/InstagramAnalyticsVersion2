"""
Export all creator partnership posts with Creator, Date, and URL
"""

import json, sys

with open(r"C:/Users/omkar/.gemini/antigravity/brain/d196c916-bdf1-40a6-835a-3259489a070c/.system_generated/steps/471/output.txt", encoding="utf-8") as f:
    data = json.load(f)

rows = data["valueRanges"][0].get("values", [])[2:]

sub_brands = {"@grt.diamonds", "@grt.silverjewellery", "@grt.silverarticles", "@grtoriana"}
industry_media = {"@platinumevara", "@platinumdaysoflove", "@menofplatinum", "@eventart_india", "@chennaitimestoi"}

creator_posts = []
for r in rows:
    r_padded = r + [""] * (10 - len(r))
    h, fol, l, c, er, date, url, via, cap, stat = r_padded[:10]
    if h in sub_brands or h in industry_media:
        continue
    if "owned by partner" in via:
        creator_posts.append({
            "handle": h,
            "followers": fol,
            "date": date,
            "url": url,
            "caption": cap[:90].replace("\n", " ") if cap else "No caption"
        })

# Sort by date descending
creator_posts.sort(key=lambda x: x["date"], reverse=True)

print(f"Total Creator-Owned Paid Posts: {len(creator_posts)}")
for i, p in enumerate(creator_posts, 1):
    print(f"{i:>2}. {p['handle']:<28} | Date: {p['date']} | {p['url']}")
