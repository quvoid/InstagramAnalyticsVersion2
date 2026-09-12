"""
Scrape GIVA (@giva.co) and Palmonas (@palmonas_official)
Extract all creator collaboration posts, check Meta paid partnership toggle status,
compute engagement metrics & boost detection, and prepare updates for master analysis.
"""

import sys, os, json, time, re, csv
from concurrent.futures import ThreadPoolExecutor, as_completed
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")

session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

web_hdrs = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "x-ig-app-id": "936619743392459",
    "accept": "*/*",
}

BRANDS = [
    {"name": "GIVA Jewellery", "username": "giva.co"},
    {"name": "Palmonas", "username": "palmonas_official"},
]

def get_user_info(username):
    url = f"https://i.instagram.com/api/v1/users/web_profile_info/?username={username}"
    try:
        r = session.get(url, headers=web_hdrs, cookies=COOKIES, timeout=15)
        if r.status_code == 200:
            user_data = r.json().get("data", {}).get("user", {})
            return {
                "id": user_data.get("id"),
                "full_name": user_data.get("full_name"),
                "followers": user_data.get("edge_followed_by", {}).get("count", 0),
                "media_count": user_data.get("edge_owner_to_timeline_media", {}).get("count", 0),
                "timeline_edges": user_data.get("edge_owner_to_timeline_media", {}).get("edges", []),
                "reels_edges": user_data.get("edge_felix_video_timeline", {}).get("edges", [])
            }
    except Exception as e:
        print(f"Error fetching web_profile_info for {username}: {e}")
    return None

def fetch_user_feed_mobile(user_id, max_pages=6):
    items = []
    max_id = ""
    for page in range(max_pages):
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
                if not max_id:
                    break
                time.sleep(0.5)
            else:
                break
        except Exception:
            break
    return items

def fetch_user_clips_mobile(user_id, max_pages=6):
    items = []
    max_id = ""
    for page in range(max_pages):
        url = f"https://i.instagram.com/api/v1/clips/user/"
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
                if not paging.get("more_available") or not max_id:
                    break
                time.sleep(0.5)
            else:
                break
        except Exception:
            break
    return items

def get_creator_followers(handle):
    h_clean = handle.replace("@", "").strip()
    url = f"https://i.instagram.com/api/v1/users/web_profile_info/?username={h_clean}"
    try:
        r = session.get(url, headers=web_hdrs, cookies=COOKIES, timeout=10)
        if r.status_code == 200:
            return r.json().get("data", {}).get("user", {}).get("edge_followed_by", {}).get("count", 0)
    except Exception:
        pass
    return 0

print("="*70)
print("SCRAPING GIVA & PALMONAS CREATOR COLLABORATIONS")
print("="*70)

all_extracted_posts = []

NON_CREATOR_SET = {
    "@giva.co", "@giva_cares", "@giva_men", "@giva_fine_jewellery",
    "@palmonas_official", "@palmonasofficial", "@palmonas_men", "@palmonas_silver",
    "@shraddhakapoor" # Shraddha is co-founder/partner for Palmonas, we will track her specifically!
}

for b_info in BRANDS:
    b_name = b_info["name"]
    b_user = b_info["username"]
    print(f"\n[+] Fetching profile for {b_name} (@{b_user})...")
    
    u_info = get_user_info(b_user)
    if not u_info or not u_info.get("id"):
        print(f"[-] Could not resolve user ID for {b_user}")
        continue
        
    user_id = u_info["id"]
    print(f"    User ID: {user_id} | Followers: {u_info['followers']:,} | Total Posts: {u_info['media_count']:,}")
    
    print(f"    Fetching feed & clips...")
    feed_items = fetch_user_feed_mobile(user_id, max_pages=8)
    clips_items = fetch_user_clips_mobile(user_id, max_pages=8)
    
    # Deduplicate items by pk/id
    seen_ids = set()
    unique_items = []
    for item in feed_items + clips_items:
        pk = str(item.get("pk") or item.get("id"))
        if pk and pk not in seen_ids:
            seen_ids.add(pk)
            unique_items.append(item)
            
    print(f"    Total unique posts/reels fetched: {len(unique_items)}")
    
    # Filter for creator collaborations
    brand_collabs = []
    for item in unique_items:
        owner_username = item.get("user", {}).get("username", "").lower()
        coauthors = [c.get("username", "").lower() for c in item.get("coauthor_producers", [])]
        is_paid = bool(item.get("is_paid_partnership", False))
        code = item.get("code") or ""
        post_url = f"https://www.instagram.com/p/{code}/" if code else ""
        
        caption_obj = item.get("caption") or {}
        caption_text = caption_obj.get("text", "") if isinstance(caption_obj, dict) else str(caption_obj)
        taken_at = item.get("taken_at")
        date_str = time.strftime("%Y-%m-%d", time.gmtime(taken_at)) if taken_at else "N/A"
        
        play_count = item.get("play_count") or item.get("view_count") or 0
        like_count = item.get("like_count") or 0
        comment_count = item.get("comment_count") or 0
        
        if not play_count and like_count:
            play_count = int(like_count * 20)
            
        # Determine if it's a partner-owned post or co-author
        # Case A: Post is owned by partner
        if owner_username != b_user.lower():
            creator_handle = f"@{owner_username}"
            via = "Post owned by partner (collab)"
        # Case B: Post is owned by brand, but has external co-authors
        elif coauthors and any(ca != b_user.lower() for ca in coauthors):
            ext_authors = [ca for ca in coauthors if ca != b_user.lower()]
            creator_handle = f"@{ext_authors[0]}"
            via = "Post owned by partner (collab)" # Joint collab
        else:
            continue
            
        # Exclude brand's internal accounts
        if creator_handle.lower() in [f"@{b_user.lower()}", "@giva_cares", "@giva_men"]:
            continue
            
        brand_collabs.append({
            "brand": b_name,
            "handle": creator_handle,
            "url": post_url,
            "shortcode": code,
            "media_id": item.get("pk"),
            "date": date_str,
            "is_paid_partnership": is_paid,
            "views": play_count,
            "likes": like_count,
            "comments": comment_count,
            "via": via,
            "caption": caption_text[:250].replace("\n", " ")
        })
        
    print(f"    -> Found {len(brand_collabs)} creator collab posts for {b_name}!")
    all_extracted_posts.extend(brand_collabs)

# Fetch creator follower counts in parallel
unique_handles = list(set(p["handle"] for p in all_extracted_posts))
print(f"\n[+] Fetching follower counts for {len(unique_handles)} unique creators...")

creator_followers_map = {}
with ThreadPoolExecutor(max_workers=10) as executor:
    future_to_handle = {executor.submit(get_creator_followers, h): h for h in unique_handles}
    for future in as_completed(future_to_handle):
        h = future_to_handle[future]
        fol = future.result()
        creator_followers_map[h] = fol

# Assign metrics and boost detection
for p in all_extracted_posts:
    h = p["handle"]
    followers = creator_followers_map.get(h, 0)
    p["followers"] = followers
    
    views = p["views"]
    likes = p["likes"]
    comments = p["comments"]
    
    er = round(((likes + comments) / followers) * 100, 2) if followers > 0 else 0.0
    like_rate = round((likes / views) * 100, 2) if views > 0 else 0.0
    view_multiplier = round(views / followers, 2) if followers > 0 else 0.0
    
    p["er"] = er
    p["like_rate"] = like_rate
    p["view_multiplier"] = view_multiplier
    
    # Boost detection
    if views >= 500000 and like_rate < 0.35:
        p["boost_status"] = "🚀 Heavily Boosted (Paid Ad Spend)"
        p["is_boosted"] = True
    elif view_multiplier >= 5.0 and like_rate < 0.70:
        p["boost_status"] = "🚀 Boosted (Paid Ad Spend)"
        p["is_boosted"] = True
    elif view_multiplier >= 3.0 and like_rate < 1.00 and views > 80000:
        p["boost_status"] = "🔍 Likely Boosted (Targeted Ad)"
        p["is_boosted"] = True
    elif view_multiplier >= 4.0 and like_rate >= 2.00:
        p["boost_status"] = "📈 Viral Organic Reach"
        p["is_boosted"] = False
    else:
        p["boost_status"] = "⚪ Standard Organic"
        p["is_boosted"] = False
        
    p["toggle_str"] = "🟢 ON (Formal Label)" if p["is_paid_partnership"] else "⚪ OFF (Collab Only)"

with open("giva_palmonas_scraped.json", "w", encoding="utf-8") as f:
    json.dump(all_extracted_posts, f, indent=2)

print(f"\n{'='*70}")
print(f"SCRAPING COMPLETE: {len(all_extracted_posts)} Total Creator Posts Extracted")
print(f"{'='*70}\n")

# Print summary table
for b_info in BRANDS:
    b_name = b_info["name"]
    b_posts = [p for p in all_extracted_posts if p["brand"] == b_name]
    u_creators = len(set(p["handle"] for p in b_posts))
    on_cnt = sum(1 for p in b_posts if p["is_paid_partnership"])
    boosted_cnt = sum(1 for p in b_posts if p["is_boosted"])
    print(f"📊 {b_name}:")
    print(f"   • Total Paid Collab Posts: {len(b_posts)}")
    print(f"   • Unique Creators: {u_creators}")
    print(f"   • Toggle ON Posts: {on_cnt}")
    print(f"   • Boosted Posts Detected: {boosted_cnt}")
    print()
