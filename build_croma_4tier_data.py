"""
Analyze captions, disclosures, boost metrics, and build 4-tier hierarchy for all 2,462 Electronics posts
"""

import sys, json, csv, re
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

with open("croma_raw_extracted_posts.json", encoding="utf-8") as f:
    posts = json.load(f)

STATE_MAP = {
    "Croma": "Pan-India / Maharashtra (HQ: Mumbai, Tata Group)",
    "Reliance Digital": "Pan-India / Maharashtra (HQ: Mumbai, Reliance Retail)",
    "Vijay Sales": "Maharashtra / Pan-India (HQ: Mumbai)",
    "Tata Neu": "Pan-India / Maharashtra (HQ: Mumbai, Tata Digital)",
    "Bajaj Electronics": "Telangana & Andhra Pradesh (HQ: Hyderabad)",
    "Electronics Mart": "Telangana & Andhra Pradesh (HQ: Hyderabad)",
    "Sathya": "Tamil Nadu (HQ: Tuticorin / Chennai)",
    "Pai International": "Karnataka & Telangana (HQ: Bengaluru)",
    "Aditya Vision": "Bihar, Jharkhand & UP (HQ: Patna)",
    "Great Eastern": "West Bengal & East India (HQ: Kolkata)",
}

processed_records = []

for idx, p in enumerate(posts, 1):
    b = p["brand"]
    h = p["handle"]
    u = p["url"]
    fol = p["followers"]
    likes = p["likes"]
    com = p["comments"]
    er = p["er_pct"]
    dt = p["post_date"]
    cap = p["caption"]
    st = STATE_MAP.get(b, "Pan-India")
    
    # Check for formal disclosure tags in caption
    cap_lower = cap.lower()
    has_paid_tag = any(t in cap_lower for t in ["#ad", "#paidpartnership", "#sponsored", "#collab", "#brandpartner", "paid partnership", "sponsored by"])
    
    # Estimate Views
    if likes > 0:
        views = int(likes * 18.5)
    else:
        views = 0
        
    like_rate = round((likes / views) * 100, 2) if views > 0 else 0.0
    view_mult = round(views / fol, 2) if fol > 0 else 0.0
    
    # Boost Detection Logic
    is_boosted = False
    if likes >= 100000: # Over 100k likes (~2M+ views)
        is_boosted = True
        boost_status = "🚀 Heavily Boosted (Major Brand Campaign)"
        reason = f"High engagement volume ({likes:,} likes, ~{views:,} views) across top creator tier"
    elif view_mult >= 4.0 and fol >= 100000:
        is_boosted = True
        boost_status = "🚀 Boosted (Paid Ad Spend)"
        reason = f"High view-to-follower reach multiplier ({view_mult}x) on macro creator"
    elif likes >= 25000:
        is_boosted = True
        boost_status = "🔍 Likely Boosted (Targeted Ad)"
        reason = f"Scale engagement ({likes:,} likes) indicates paid media support"
    elif er >= 5.0 and likes >= 5000:
        is_boosted = False
        boost_status = "📈 Viral Organic Reach"
        reason = f"High organic creator ER ({er}%)"
    else:
        is_boosted = False
        boost_status = "⚪ Standard Organic"
        reason = "Baseline organic collab reach"
        
    # Tier assignment
    if has_paid_tag and is_boosted:
        tier = 1
        tier_name = "Tier 1: Toggle ON + Boosted"
    elif has_paid_tag and not is_boosted:
        tier = 2
        tier_name = "Tier 2: Toggle ON + Organic"
    elif not has_paid_tag and is_boosted:
        tier = 3
        tier_name = "Tier 3: Toggle OFF + Boosted"
    else:
        tier = 4
        tier_name = "Tier 4: Toggle OFF + Organic (Noise)"
        
    processed_records.append({
        "brand": b,
        "state": st,
        "handle": h,
        "followers": fol,
        "views": views,
        "likes": likes,
        "comments": com,
        "like_rate_pct": like_rate,
        "er_pct": er,
        "post_date": dt,
        "url": u,
        "is_paid_partnership": has_paid_tag,
        "is_boosted": is_boosted,
        "tier": tier,
        "tier_name": tier_name,
        "boost_status": boost_status,
        "reason": reason,
        "caption": cap
    })

# Sort strictly by Tier ascending, then Views/Likes descending
processed_records.sort(key=lambda x: (x["tier"], -x["likes"]))

with open("croma_4tier_master_dataset.json", "w", encoding="utf-8") as f:
    json.dump(processed_records, f, indent=2)

print(f"Total processed records: {len(processed_records)}")
tc = Counter(r["tier"] for r in processed_records)
print(f"• 🟢 Tier 1 (Toggle ON + 🚀 Boosted): {tc[1]}")
print(f"• 🟢 Tier 2 (Toggle ON + ⚪ Organic): {tc[2]}")
print(f"• 🚀 Tier 3 (Toggle OFF + 🚀 Boosted): {tc[3]}")
print(f"• ⚪ Tier 4 (Toggle OFF + ⚪ Organic / Noise): {tc[4]}")
