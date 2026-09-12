"""
Deep scan of all 1,067 paid collab captions for paid partnership tags, ad hashtags, and disclosures
"""

import csv, json, re, sys

sys.stdout.reconfigure(encoding="utf-8")

with open("All_Brands_Paid_Collabs_With_Toggle.csv", encoding="utf-8-sig") as f:
    rows = list(csv.reader(f))

header = rows[0]
data_rows = rows[1:]

print(f"Total rows to analyze: {len(data_rows)}")

# Let's check various keyword patterns in captions
keywords = {
    "#ad": re.compile(r"(?i)#ad\b|#ad\||#ad_"),
    "#collab / #collaboration": re.compile(r"(?i)#collab\b|#collaboration\b|#incollaborationwith\b"),
    "#sponsored": re.compile(r"(?i)#sponsored\b|#spon\b"),
    "#paidpartnership": re.compile(r"(?i)#paidpartnership\b|#paidcollab\b|paid partnership"),
    "#brandambassador / #ambassador": re.compile(r"(?i)#brandambassador\b|#ambassador\b"),
    "#partner": re.compile(r"(?i)#partner\b|#brandpartner\b"),
}

keyword_matches = {k: 0 for k in keywords}
on_rows = []

for idx, r in enumerate(data_rows, 1):
    brand = r[1]
    handle = r[2]
    url = r[9]
    cap = r[11]
    
    matched_tags = []
    for tag_name, pattern in keywords.items():
        if pattern.search(cap):
            matched_tags.append(tag_name)
            keyword_matches[tag_name] += 1
            
    if matched_tags:
        on_rows.append({
            "index": idx,
            "brand": brand,
            "handle": handle,
            "tags": matched_tags,
            "url": url,
            "caption": cap[:120].replace("\n", " ")
        })

print("\n--- Keyword Detection Breakdown across 1,067 Posts ---")
for k, v in keyword_matches.items():
    print(f"  • {k:<32}: {v:>3} posts")

print(f"\nTotal Posts with Paid / Collab Disclosures in Caption: {len(on_rows)} ({len(on_rows)/len(data_rows)*100:.1f}%)")

print("\n--- Sample Posts with Paid / Ad Disclosures ---")
for p in on_rows[:20]:
    print(f"Row {p['index']:>4} | {p['brand']:<25} | {p['handle']:<24} | Tags: {p['tags']}")
    print(f"         URL: {p['url']}")
    print(f"         Text: {p['caption']}...")
    print()
