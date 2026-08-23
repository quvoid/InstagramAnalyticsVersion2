"""
Optimized Deep Scraper for GIVA (@giva.co) and Palmonas (@palmonas_official)
"""

import sys, os, json, time, re, csv
from datetime import datetime, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")

session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

BRANDS = [
    {"name": "GIVA Jewellery", "username": "giva.co", "pk": 13618562336},
    {"name": "Palmonas", "username": "palmonas_official", "pk": 49372589192},
]

NON_CREATOR_HANDLES = {
    "giva.co", "giva_cares", "giva_men", "giva_fine_jewellery",
    "palmonas_official", "palmonasofficial", "palmonas_men", "palmonas_silver",
    "mohadikar.pallavi", "amol_57007" # Palmonas founders
}

def fetch_feed(user_id, max_pages=8):
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
                print(f"  [Feed] Page {p}: {len(f_items)} items fetched (total: {len(items)})", flush=True)
                if not max_id or len(f_items) == 0:
                    break
                time.sleep(0.4)
            else:
                break
        except Exception:
            break
    return items

def fetch_clips(user_id, max_pages=8):
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
                print(f"  [Clips] Page {p}: {len(clips)} reels fetched (total: {len(items)})", flush=True)
                if not paging.get("more_available") or not max_id:
                    break
                time.sleep(0.4)
            else:
                break
        except Exception:
            break
    return items

def resolve_follower_count(handle):
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
print("SCRAPING GIVA & PALMONAS CREATOR COLLABORATIONS", flush=True)
print("="*75, flush=True)

all_collab_posts = []

for b in BRANDS:
    b_name = b["name"]
    b_user = b["username"]
    b_pk = b["pk"]
    
    print(f"\n[+] Scraping {b_name} (@{b_user}, ID: {b_pk})...", flush=True)
    f_items = fetch_feed(b_pk, max_pages=8)
    c_items = fetch_clips(b_pk, max_pages=8)
    
    seen_ids = set()
    unique_items = []
    for it in f_items + c_items:
        pk = str(it.get("pk") or it.get("id"))
        if pk and pk not in seen_ids:
            seen_ids.add(pk)
            unique_items.append(it)
            
    print(f"  -> Total unique posts fetched: {len(unique_items)}", flush=True)
    
    brand_posts = []
    for it in unique_items:
        owner = it.get("user", {}).get("username", "").lower()
        coauthors = [c.get("username", "").lower() for c in it.get("coauthor_producers", [])]
        is_paid = bool(it.get("is_paid_partnership", False))
        code = it.get("code") or ""
        post_url = f"https://www.instagram.com/p/{code}/"
        
        cap_obj = it.get("caption") or {}
        cap_text = cap_obj.get("text", "") if isinstance(cap_obj, dict) else str(cap_obj)
        taken_at = it.get("taken_at")
        date_str = datetime.fromtimestamp(taken_at, tz=timezone.utc).strftime("%Y-%m-%d") if taken_at else "N/A"
        
        play_count = it.get("play_count") or it.get("view_count") or 0
        like_count = it.get("like_count") or 0
        comment_count = it.get("comment_count") or 0
        
        if not play_count and like_count:
            play_count = int(like_count * 20)
            
        creator_handle = ""
        # Check if creator-owned or coauthor
        if owner != b_user.lower():
            if owner not in NON_CREATOR_HANDLES:
                creator_handle = f"@{owner}"
        elif coauthors:
            ext = [c for c in coauthors if c not in NON_CREATOR_HANDLES and c != b_user.lower()]
            if ext:
                creator_handle = f"@{ext[0]}"
                
        if not creator_handle:
            continue
            
        brand_posts.append({
            "brand": b_name,
            "handle": creator_handle,
            "raw_handle": creator_handle.replace("@", ""),
            "url": post_url,
            "shortcode": code,
            "media_id": it.get("pk"),
            "date": date_str,
            "is_paid_partnership": is_paid,
            "views": play_count,
            "likes": like_count,
            "comments": comment_count,
            "via": "Post owned by partner (collab)",
            "caption": cap_text[:250].replace("\n", " ").replace("\r", " ")
        })
        
    print(f"  -> Extracted {len(brand_posts)} creator collaborations for {b_name}!", flush=True)
    all_collab_posts.extend(brand_posts)

# Resolve followers for unique handles
unique_handles = list(set(p["raw_handle"] for p in all_collab_posts))
print(f"\n[+] Resolving followers for {len(unique_handles)} creators...", flush=True)

followers_map = {}
with ThreadPoolExecutor(max_workers=12) as ex:
    f_map = {ex.submit(resolve_follower_count, h): h for h in unique_handles}
    for f in as_completed(f_map):
        h = f_map[f]
        cnt = f.result()
        followers_map[h] = cnt

# Assign metrics and boost status
for p in all_collab_posts:
    raw_h = p["raw_handle"]
    fols = followers_map.get(raw_h, 0)
    p["followers"] = fols
    
    views = p["views"]
    likes = p["likes"]
    comments = p["comments"]
    
    er = round(((likes + comments) / fols) * 100, 2) if fols > 0 else 0.0
    like_rate = round((likes / views) * 100, 2) if views > 0 else 0.0
    view_mult = round(views / fols, 2) if fols > 0 else 0.0
    
    p["er"] = er
    p["like_rate"] = like_rate
    p["view_multiplier"] = view_mult
    p["toggle_str"] = "🟢 ON (Formal Label)" if p["is_paid_partnership"] else "⚪ OFF (Collab Only)"
    
    # Boost detection
    if views >= 500000 and like_rate < 0.35:
        p["boost_status"] = "🚀 Heavily Boosted (Paid Ad Spend)"
        p["is_boosted"] = True
    elif view_mult >= 5.0 and like_rate < 0.70:
        p["boost_status"] = "🚀 Boosted (Paid Ad Spend)"
        p["is_boosted"] = True
    elif view_mult >= 3.0 and like_rate < 1.00 and views > 80000:
        p["boost_status"] = "🔍 Likely Boosted (Targeted Ad)"
        p["is_boosted"] = True
    elif view_mult >= 4.0 and like_rate >= 2.00:
        p["boost_status"] = "📈 Viral Organic Reach"
        p["is_boosted"] = False
    else:
        p["boost_status"] = "⚪ Standard Organic"
        p["is_boosted"] = False

all_collab_posts.sort(key=lambda x: x["date"], reverse=True)

with open("giva_palmonas_scraped.json", "w", encoding="utf-8") as f:
    json.dump(all_collab_posts, f, indent=2)

print(f"\n{'='*75}", flush=True)
print(f"Scrape Complete! {len(all_collab_posts)} creator posts extracted.", flush=True)
print(f"{'='*75}\n", flush=True)

for b in BRANDS:
    b_name = b["name"]
    posts = [p for p in all_collab_posts if p["brand"] == b_name]
    u_c = len(set(p["handle"] for p in posts))
    t_on = sum(1 for p in posts if p["is_paid_partnership"])
    t_off = len(posts) - t_on
    boosted = sum(1 for p in posts if p["is_boosted"])
    
    print(f"💎 {b_name}:")
    print(f"   • Total Collab Posts: {len(posts)}")
    print(f"   • Unique Creators:   {u_c}")
    print(f"   • Toggle ON:         {t_on}")
    print(f"   • Toggle OFF:        {t_off}")
    print(f"   • Boosted Posts:     {boosted}")
    print()
