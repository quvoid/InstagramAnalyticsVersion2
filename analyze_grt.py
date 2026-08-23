"""
Analyze GRT Jewellers sheet data from Google Spreadsheet export
"""

import json

with open(r"C:/Users/omkar/.gemini/antigravity/brain/d196c916-bdf1-40a6-835a-3259489a070c/.system_generated/steps/471/output.txt", encoding="utf-8") as f:
    data = json.load(f)

values = data["valueRanges"][0].get("values", [])
title_row = values[0]
header_row = values[1]
rows = values[2:]

sub_brands = {"@grt.diamonds", "@grt.silverjewellery", "@grt.silverarticles", "@grtoriana"}
industry_media = {"@platinumevara", "@platinumdaysoflove", "@menofplatinum", "@eventart_india", "@chennaitimestoi"}

creator_rows = []
sub_brand_rows = []
industry_media_rows = []

for r in rows:
    r_padded = r + [""] * (10 - len(r))
    h, fol, l, c, er, date, url, via, cap, stat = r_padded[:10]
    entry = {
        "handle": h,
        "followers": fol,
        "likes": l,
        "comments": c,
        "er": er,
        "date": date,
        "url": url,
        "via": via,
        "caption": cap,
        "status": stat
    }
    if h in sub_brands:
        sub_brand_rows.append(entry)
    elif h in industry_media:
        industry_media_rows.append(entry)
    else:
        creator_rows.append(entry)

u_creators = len(set(r["handle"] for r in creator_rows))
all_urls = set(r["url"] for r in creator_rows + sub_brand_rows + industry_media_rows if r["url"])
creator_urls = set(r["url"] for r in creator_rows if r["url"])
sub_brand_urls = set(r["url"] for r in sub_brand_rows if r["url"])
industry_urls = set(r["url"] for r in industry_media_rows if r["url"])

collab_owned = sum(1 for r in creator_rows if "owned by partner" in r["via"])
coauthor = sum(1 for r in creator_rows if "Coauthor" in r["via"])

print(f"Total Rows in Sheet: {len(rows)}")
print(f"Total Unique Posts Logged: {len(all_urls)}")
print(f"\n1. Individual Creator / Influencer / Celebrity Collabs:")
print(f"   • Total Rows: {len(creator_rows)}")
print(f"   • Unique Creators: {u_creators}")
print(f"   • Unique Posts featuring Creators: {len(creator_urls)}")
print(f"   • Post owned by Creator (Partnership Collab): {collab_owned}")
print(f"   • Co-authored Tag Collabs: {coauthor}")

print(f"\n2. Sub-Brand Self-Referential Cross-Promotions (@grt.diamonds, @grt.silverjewellery, etc.):")
print(f"   • Total Rows: {len(sub_brand_rows)}")
print(f"   • Unique Sub-Brand Handles: {len(sub_brands)}")
print(f"   • Unique Posts: {len(sub_brand_urls)}")
print(f"   • Posts featuring ONLY Sub-Brands (no external creator): {len(sub_brand_urls - creator_urls)}")

print(f"\n3. Industry & Media Partnerships (PGI Platinum, Times of India, Event Art):")
print(f"   • Total Rows: {len(industry_media_rows)}")
print(f"   • Unique Handles: {len(industry_media)}")
print(f"   • Unique Posts: {len(industry_urls)}")

# Top creators by appearances / followers
from collections import Counter
creator_counts = Counter(r["handle"] for r in creator_rows)
print(f"\nTop Creator Handles by Appearances:")
for h, cnt in creator_counts.most_common(15):
    fol = next(r["followers"] for r in creator_rows if r["handle"] == h)
    er = next(r["er"] for r in creator_rows if r["handle"] == h)
    print(f"   {h:<30} | {cnt} posts | Followers: {fol:>10} | ER%: {er}")
