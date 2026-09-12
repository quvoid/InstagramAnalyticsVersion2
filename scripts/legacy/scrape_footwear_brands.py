"""
Scrape and Analyze All Creator Collaborations for Footwear & Sneaker Brands:
1. Skechers India (@skechersindia)
2. Gully Labs (@gullylabs)
3. Comet (@thecometuniverse)

Extracts all collaborator posts across Tiers 1, 2, 3, and 4 (including unboosted noise),
evaluates Meta disclosures, view metrics, ER%, boost status, and builds master Excel & CSV.
"""

import sys, os, json, time, re, csv
from datetime import datetime, timezone
from collections import defaultdict, Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")

session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

BRANDS = [
    {"name": "Skechers India", "username": "skechersindia", "state": "Pan-India / Maharashtra (HQ: Mumbai)"},
    {"name": "Gully Labs", "username": "gullylabs", "state": "Pan-India / Delhi NCR (HQ: New Delhi, D2C)"},
    {"name": "Comet", "username": "thecometuniverse", "state": "Pan-India / Karnataka (HQ: Bengaluru, D2C)"},
]

# Internal non-creator accounts to ignore
INTERNAL_ACCOUNTS = {
    "skechersindia", "skechers", "skechersperformance", "skechers_south_asia",
    "gullylabs", "gully_labs",
    "thecometuniverse", "comet_universe", "wearcomet",
}

def resolve_user_id(username):
    url = f"https://i.instagram.com/api/v1/users/search/?q={username}"
    try:
        r = session.get(url, headers=mob_hdrs, cookies=COOKIES, timeout=10)
        if r.status_code == 200:
            for u in r.json().get("users", []):
                if u.get("username", "").lower() == username.lower():
                    return {
                        "pk": u.get("pk"),
                        "username": u.get("username"),
                        "full_name": u.get("full_name"),
                        "followers": u.get("follower_count", 0),
                    }
    except Exception as e:
        print(f"Error resolving @{username}: {e}")
    return {"pk": None, "username": username, "full_name": username, "followers": 0}

def fetch_feed(user_id, max_pages=15):
    items = []
    max_id = ""
    for p in range(1, max_pages + 1):
        url = f"https://i.instagram.com/api/v1/feed/user/{user_id}/"
        if max_id:
            url += f"?max_id={max_id}"
        try:
            r = session.get(url, headers=mob_hdrs, cookies=COOKIES, timeout=12)
            if r.status_code == 200:
                data = r.json()
                f_items = data.get("items", [])
                items.extend(f_items)
                max_id = data.get("next_max_id")
                print(f"    [Feed] Page {p:>2}: {len(f_items)} items fetched (Total: {len(items)})", flush=True)
                if not max_id or len(f_items) == 0:
                    break
                time.sleep(0.4)
            else:
                break
        except Exception:
            break
    return items

def fetch_clips(user_id, max_pages=15):
    items = []
    max_id = ""
    for p in range(1, max_pages + 1):
        url = "https://i.instagram.com/api/v1/clips/user/"
        payload = {"target_user_id": str(user_id), "page_size": 30}
        if max_id:
            payload["max_id"] = str(max_id)
        try:
            r = session.post(url, headers=mob_hdrs, data=payload, cookies=COOKIES, timeout=12)
            if r.status_code == 200:
                data = r.json()
                clips = [it.get("media") for it in data.get("items", []) if it.get("media")]
                items.extend(clips)
                paging = data.get("paging_info", {})
                max_id = paging.get("max_id")
                print(f"    [Clips] Page {p:>2}: {len(clips)} reels fetched (Total: {len(items)})", flush=True)
                if not paging.get("more_available") or not max_id:
                    break
                time.sleep(0.4)
            else:
                break
        except Exception:
            break
    return items

def get_creator_followers(handle):
    try:
        sess = make_session()
        r = sess.get(f"https://i.instagram.com/api/v1/users/search/?q={handle}", headers=mob_hdrs, cookies=COOKIES, timeout=8)
        if r.status_code == 200:
            for u in r.json().get("users", []):
                if u.get("username", "").lower() == handle.lower():
                    pk = u.get("pk")
                    r2 = sess.get(f"https://i.instagram.com/api/v1/users/{pk}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=8)
                    if r2.status_code == 200:
                        ud = r2.json().get("user", {})
                        return ud.get("follower_count") or u.get("follower_count") or 0
                    return u.get("follower_count") or 0
    except Exception:
        pass
    return 0

print("="*75, flush=True)
print("SCRAPING SKECHERS INDIA, GULLY LABS & COMET CREATOR COLLABORATIONS", flush=True)
print("="*75, flush=True)

all_collab_posts = []

for b in BRANDS:
    b_name = b["name"]
    b_user = b["username"]
    b_state = b["state"]
    
    print(f"\n[+] Resolving user profile for {b_name} (@{b_user})...", flush=True)
    u_info = resolve_user_id(b_user)
    b_pk = u_info["pk"]
    
    if not b_pk:
        print(f"[-] Could not resolve user ID for {b_user}", flush=True)
        continue
        
    print(f"    User ID: {b_pk} | Followers: {u_info['followers']:,} | Full Name: {u_info['full_name']}", flush=True)
    print("    Fetching Feed & Clips...", flush=True)
    
    f_items = fetch_feed(b_pk, max_pages=15)
    c_items = fetch_clips(b_pk, max_pages=15)
    
    seen_ids = set()
    unique_items = []
    for it in f_items + c_items:
        pk = str(it.get("pk") or it.get("id"))
        if pk and pk not in seen_ids:
            seen_ids.add(pk)
            unique_items.append(it)
            
    print(f"    -> Total unique posts fetched: {len(unique_items)}", flush=True)
    
    brand_posts = []
    for it in unique_items:
        owner = it.get("user", {}).get("username", "").lower()
        coauthors = [c.get("username", "").lower() for c in it.get("coauthor_producers", [])]
        is_paid = bool(it.get("is_paid_partnership", False))
        code = it.get("code") or ""
        post_url = f"https://www.instagram.com/p/{code}/" if code else ""
        
        cap_obj = it.get("caption") or {}
        cap_text = cap_obj.get("text", "") if isinstance(cap_obj, dict) else str(cap_obj)
        taken_at = it.get("taken_at")
        date_str = datetime.fromtimestamp(taken_at, tz=timezone.utc).strftime("%Y-%m-%d") if taken_at else "N/A"
        
        play_count = it.get("play_count") or it.get("view_count") or 0
        like_count = it.get("like_count") or 0
        comment_count = it.get("comment_count") or 0
        
        if not play_count and like_count:
            play_count = int(like_count * 18.5)
            
        creator_handle = ""
        # Check partner ownership or coauthor
        if owner != b_user.lower():
            if owner not in INTERNAL_ACCOUNTS:
                creator_handle = f"@{owner}"
        elif coauthors:
            ext = [c for c in coauthors if c not in INTERNAL_ACCOUNTS and c != b_user.lower()]
            if ext:
                creator_handle = f"@{ext[0]}"
                
        if not creator_handle:
            continue
            
        brand_posts.append({
            "brand": b_name,
            "state": b_state,
            "handle": creator_handle,
            "raw_handle": creator_handle.replace("@", ""),
            "url": post_url,
            "shortcode": code,
            "media_id": str(it.get("pk") or ""),
            "date": date_str,
            "is_paid_partnership": is_paid,
            "views": play_count,
            "likes": like_count,
            "comments": comment_count,
            "via": "Post owned by partner (collab)",
            "caption": cap_text[:250].replace("\n", " ").replace("\r", " ")
        })
        
    print(f"    -> Extracted {len(brand_posts)} creator collaboration posts for {b_name}!", flush=True)
    all_collab_posts.extend(brand_posts)

# Resolve creator followers concurrently
unique_handles = list(set(p["raw_handle"] for p in all_collab_posts))
print(f"\n[+] Resolving follower counts for {len(unique_handles)} unique creators...", flush=True)

followers_cache = {}
with ThreadPoolExecutor(max_workers=15) as executor:
    fut_to_h = {executor.submit(get_creator_followers, h): h for h in unique_handles}
    for fut in as_completed(fut_to_h):
        h = fut_to_h[fut]
        followers_cache[h] = fut.result()

# Evaluate Boost and 4-Tier Assignment
for p in all_collab_posts:
    raw_h = p["raw_handle"]
    fols = followers_cache.get(raw_h, 0)
    p["followers"] = fols
    
    views = p["views"]
    likes = p["likes"]
    comments = p["comments"]
    cap_lower = p["caption"].lower()
    
    # Check disclosure tags in caption
    has_disclosure = any(t in cap_lower for t in ["#ad", "#paidpartnership", "#sponsored", "#collab", "#brandpartner", "paid partnership"])
    is_formal_paid = p["is_paid_partnership"] or has_disclosure
    
    like_rate = round((likes / views) * 100, 2) if views > 0 else 0.0
    view_mult = round(views / fols, 2) if fols > 0 else 0.0
    er = round(((likes + comments) / fols) * 100, 2) if fols > 0 else 0.0
    
    p["like_rate_pct"] = like_rate
    p["view_multiplier"] = view_mult
    p["er_pct"] = er
    
    # Boost Detection Logic
    is_boosted = False
    if views >= 500000 and like_rate < 0.35:
        is_boosted = True
        boost_status = "🚀 Heavily Boosted (Paid Ad Spend)"
        reason = f"High view count ({views:,}) with sub-0.35% like rate ({like_rate}%) indicates paid video ads campaign"
    elif view_mult >= 5.0 and like_rate < 0.70:
        is_boosted = True
        boost_status = "🚀 Boosted (Paid Ad Spend)"
        reason = f"High view multiplier ({view_mult}x followers) combined with low engagement rate ({like_rate}%)"
    elif view_mult >= 3.0 and like_rate < 1.00 and views >= 80000:
        is_boosted = True
        boost_status = "🔍 Likely Boosted (Targeted Ad)"
        reason = f"Disproportionate views ({views:,}) relative to likes ({likes:,})"
    elif er >= 4.0 and like_rate >= 2.00:
        is_boosted = False
        boost_status = "📈 Viral Organic Reach"
        reason = f"High organic views ({views:,}) with strong like rate ({like_rate}%)"
    else:
        is_boosted = False
        boost_status = "⚪ Standard Organic"
        reason = "Baseline organic collab reach"
        
    # Tier Assignment
    if is_formal_paid and is_boosted:
        tier = 1
        tier_name = "Tier 1: Toggle ON + Boosted"
    elif is_formal_paid and not is_boosted:
        tier = 2
        tier_name = "Tier 2: Toggle ON + Organic"
    elif not is_formal_paid and is_boosted:
        tier = 3
        tier_name = "Tier 3: Toggle OFF + Boosted"
    else:
        tier = 4
        tier_name = "Tier 4: Toggle OFF + Organic (Noise)"
        
    p["tier"] = tier
    p["tier_name"] = tier_name
    p["is_boosted"] = is_boosted
    p["boost_status"] = boost_status
    p["boost_reason"] = reason

# Sort strictly by Tier ascending, then Views descending
all_collab_posts.sort(key=lambda x: (x["tier"], -x["views"]))

with open("footwear_sneakers_4tier_dataset.json", "w", encoding="utf-8") as f:
    json.dump(all_collab_posts, f, indent=2)

print(f"\n{'='*75}", flush=True)
print(f"Scrape Complete! {len(all_collab_posts)} creator posts extracted across 3 brands.", flush=True)
print(f"{'='*75}\n", flush=True)

for b in BRANDS:
    b_name = b["name"]
    b_posts = [p for p in all_collab_posts if p["brand"] == b_name]
    u_c = len(set(p["handle"].lower() for p in b_posts))
    t1 = sum(1 for p in b_posts if p["tier"] == 1)
    t2 = sum(1 for p in b_posts if p["tier"] == 2)
    t3 = sum(1 for p in b_posts if p["tier"] == 3)
    t4 = sum(1 for p in b_posts if p["tier"] == 4)
    high_intent = t1 + t2 + t3
    
    print(f"👟 {b_name} ({b['state']}):")
    print(f"   • Total Collab Posts: {len(b_posts)}")
    print(f"   • Unique Creators:   {u_c}")
    print(f"   • 🟢 Tier 1 (Toggle ON + Boosted): {t1}")
    print(f"   • 🟢 Tier 2 (Toggle ON + Organic): {t2}")
    print(f"   • 🚀 Tier 3 (Toggle OFF + Boosted): {t3}")
    print(f"   • ⚪ Tier 4 (Noise / Unboosted):    {t4}")
    print(f"   • 💎 Total High-Intent Paid:        {high_intent} ({high_intent/len(b_posts)*100:.1f}%)" if b_posts else "   • Total High-Intent: 0")
    print()
