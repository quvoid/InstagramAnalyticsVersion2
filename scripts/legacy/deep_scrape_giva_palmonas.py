"""
Deep Scraper for GIVA (@giva.co) and Palmonas (@palmonas_official)
Extracts all creator collaborations, checks Paid Partnership toggle,
and analyzes views, engagement, and boost status.
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

# Internal non-creator accounts to filter out
NON_CREATOR_HANDLES = {
    "giva.co", "giva_cares", "giva_men", "giva_fine_jewellery",
    "palmonas_official", "palmonasofficial", "palmonas_men", "palmonas_silver",
    "mohadikar.pallavi", "amol_57007" # Founders of Palmonas
}

def fetch_feed_pages(user_id, max_pages=15):
    items = []
    max_id = ""
    for page in range(1, max_pages + 1):
        url = f"https://i.instagram.com/api/v1/feed/user/{user_id}/"
        if max_id:
            url += f"?max_id={max_id}"
        try:
            r = session.get(url, headers=mob_hdrs, cookies=COOKIES, timeout=15)
            if r.status_code == 200:
                data = r.json()
                feed_items = data.get("items", [])
                items.extend(feed_items)
                max_id = data.get("next_max_id")
                print(f"    Page {page:>2}: Fetched {len(feed_items)} items (Total: {len(items)})")
                if not max_id or len(feed_items) == 0:
                    break
                time.sleep(0.6)
            else:
                print(f"    Page {page}: HTTP {r.status_code}")
                break
        except Exception as e:
            print(f"    Page {page}: Error {e}")
            break
    return items

def fetch_clips_pages(user_id, max_pages=15):
    items = []
    max_id = ""
    for page in range(1, max_pages + 1):
        url = "https://i.instagram.com/api/v1/clips/user/"
        payload = {"target_user_id": str(user_id), "page_size": 30}
        if max_id:
            payload["max_id"] = str(max_id)
        try:
            r = session.post(url, headers=mob_hdrs, data=payload, cookies=COOKIES, timeout=15)
            if r.status_code == 200:
                data = r.json()
                clips = [item.get("media") for item in data.get("items", []) if item.get("media")]
                items.extend(clips)
                paging = data.get("paging_info", {})
                max_id = paging.get("max_id")
                print(f"    Clips Page {page:>2}: Fetched {len(clips)} reels (Total: {len(items)})")
                if not paging.get("more_available") or not max_id:
                    break
                time.sleep(0.6)
            else:
                break
        except Exception:
            break
    return items

def get_user_followers(user_id_or_name):
    # Try search endpoint to get accurate follower count
    try:
        sess = make_session()
        r = sess.get(f"https://i.instagram.com/api/v1/users/search/?q={user_id_or_name}", headers=mob_hdrs, cookies=COOKIES, timeout=10)
        if r.status_code == 200:
            for u in r.json().get("users", []):
                if u.get("username", "").lower() == str(user_id_or_name).lower():
                    # Get info
                    pk = u.get("pk")
                    r_inf = sess.get(f"https://i.instagram.com/api/v1/users/{pk}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=10)
                    if r_inf.status_code == 200:
                        u_d = r_inf.json().get("user", {})
                        return u_d.get("follower_count") or u.get("follower_count") or 0
                    return u.get("follower_count", 0)
    except Exception:
        pass
    return 0

print("="*75)
print("DEEP SCRAPING: GIVA (@giva.co) & PALMONAS (@palmonas_official)")
print("="*75)

all_collab_posts = []

for brand in BRANDS:
    b_name = brand["name"]
    b_user = brand["username"]
    b_pk = brand["pk"]
    
    print(f"\n[+] Processing {b_name} (User ID: {b_pk})...")
    print("  Fetching Feed...")
    feed_items = fetch_feed_pages(b_pk, max_pages=15)
    print("  Fetching Clips (Reels)...")
    clips_items = fetch_clips_pages(b_pk, max_pages=15)
    
    # Deduplicate items by media id / pk
    seen_ids = set()
    unique_items = []
    for item in feed_items + clips_items:
        pk = str(item.get("pk") or item.get("id"))
        if pk and pk not in seen_ids:
            seen_ids.add(pk)
            unique_items.append(item)
            
    print(f"  Total unique media items for {b_name}: {len(unique_items)}")
    
    brand_collabs = []
    for item in unique_items:
        owner = item.get("user", {}).get("username", "").lower()
        coauthors = [c.get("username", "").lower() for c in item.get("coauthor_producers", [])]
        is_paid = bool(item.get("is_paid_partnership", False))
        code = item.get("code") or ""
        post_url = f"https://www.instagram.com/p/{code}/"
        
        caption_obj = item.get("caption") or {}
        caption_text = caption_obj.get("text", "") if isinstance(caption_obj, dict) else str(caption_obj)
        taken_at = item.get("taken_at")
        date_str = datetime.fromtimestamp(taken_at, tz=timezone.utc).strftime("%Y-%m-%d") if taken_at else "N/A"
        
        play_count = item.get("play_count") or item.get("view_count") or 0
        like_count = item.get("like_count") or 0
        comment_count = item.get("comment_count") or 0
        
        if not play_count and like_count:
            play_count = int(like_count * 20)
            
        # Determine creator handle
        # Case A: Post published by partner
        if owner != b_user.lower():
            if owner in NON_CREATOR_HANDLES:
                continue
            creator_handle = f"@{owner}"
            via = "Post owned by partner (collab)"
        # Case B: Post published by brand with partner as coauthor
        elif coauthors:
            external_coauthors = [c for c in coauthors if c not in NON_CREATOR_HANDLES and c != b_user.lower()]
            if not external_coauthors:
                continue
            creator_handle = f"@{external_coauthors[0]}"
            via = "Post owned by partner (collab)"
        else:
            continue
            
        brand_collabs.append({
            "brand": b_name,
            "handle": creator_handle,
            "raw_handle": creator_handle.replace("@", ""),
            "url": post_url,
            "shortcode": code,
            "media_id": item.get("pk"),
            "date": date_str,
            "is_paid_partnership": is_paid,
            "views": play_count,
            "likes": like_count,
            "comments": comment_count,
            "via": via,
            "caption": caption_text[:250].replace("\n", " ").replace("\r", " ")
        })
        
    print(f"  -> Extracted {len(brand_collabs)} creator collaboration posts for {b_name}!")
    all_collab_posts.extend(brand_collabs)

# Fetch creator follower counts
unique_handles = list(set(p["raw_handle"] for p in all_collab_posts))
print(f"\n[+] Fetching follower counts for {len(unique_handles)} unique creators...")

followers_cache = {}
with ThreadPoolExecutor(max_workers=10) as executor:
    future_to_h = {executor.submit(get_user_followers, h): h for h in unique_handles}
    for fut in as_completed(future_to_h):
        h = future_to_h[fut]
        f_count = fut.result()
        followers_cache[h] = f_count

# Calculate metrics and boost flags
for p in all_collab_posts:
    raw_h = p["raw_handle"]
    followers = followers_cache.get(raw_h, 0)
    p["followers"] = followers
    
    views = p["views"]
    likes = p["likes"]
    comments = p["comments"]
    
    er = round(((likes + comments) / followers) * 100, 2) if followers > 0 else 0.0
    like_rate = round((likes / views) * 100, 2) if views > 0 else 0.0
    view_mult = round(views / followers, 2) if followers > 0 else 0.0
    
    p["er"] = er
    p["like_rate"] = like_rate
    p["view_multiplier"] = view_mult
    p["toggle_str"] = "🟢 ON (Formal Label)" if p["is_paid_partnership"] else "⚪ OFF (Collab Only)"
    
    # Boost decision
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

# Sort by Date descending
all_collab_posts.sort(key=lambda x: x["date"], reverse=True)

with open("giva_palmonas_full_dataset.json", "w", encoding="utf-8") as f:
    json.dump(all_collab_posts, f, indent=2)

print(f"\n{'='*75}")
print(f"DEEP SCRAPE COMPLETE!")
print(f"Total Creator Collabs Extracted: {len(all_collab_posts)}")
print(f"{'='*75}\n")

for brand in BRANDS:
    b_name = brand["name"]
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
    print("   • Top 5 Sample Posts:")
    for sp in posts[:5]:
        print(f"     - {sp['date']} | {sp['handle']:<22} | Toggle: {sp['toggle_str']} | Views: {sp['views']:>8,} | Likes: {sp['likes']:>6,} | {sp['boost_status']}")
        print(f"       URL: {sp['url']}")
    print()
