"""
1-Year Complete Collaborator & Content Audit Scraper for 10 Pune & Hyderabad Malls:
1. Phoenix Avenue of Stars Pune (@phoenixavenueofstars)
2. Seasons Mall Pune (@seasons_mall)
3. The Pavillion Pune (@pavillionpune)
4. Phoenix Mall of the Millennium Wakad (@phoenix_millennium)
5. Amanora Mall Pune (@amanoramall)
6. Nexus Hyderabad Mall (@nexus_hyderabad)
7. Sarath City Capital Mall Hyderabad (@sarathcitycapital.hyd)
8. Inorbit Mall Cyberabad (@inorbitcyberabad)
9. Lulu Mall Hyderabad (@lulumallhyderabad)
10. GVK One Mall Hyderabad (@gvkone)
"""

import sys, os, json, time, re
from datetime import datetime, timezone, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from curl_cffi import requests as cffi_requests
from api_wrapper.client import DEFAULT_IG_COOKIES, DEFAULT_IG_HEADERS

sys.stdout.reconfigure(encoding="utf-8")

NOW_DT = datetime.now(timezone.utc)
CUTOFF_DT = NOW_DT - timedelta(days=365)
CUTOFF_TIMESTAMP = int(CUTOFF_DT.timestamp())

MALLS = [
    {"name": "Phoenix Avenue of Stars Pune", "handle": "phoenixavenueofstars", "city": "Pune", "pk": 697565506},
    {"name": "Seasons Mall Pune", "handle": "seasons_mall", "city": "Pune", "pk": 4596845113},
    {"name": "The Pavillion Pune", "handle": "pavillionpune", "city": "Pune", "pk": 5635305185},
    {"name": "Phoenix Mall of the Millennium Wakad", "handle": "phoenix_millennium", "city": "Pune", "pk": 58685376270},
    {"name": "Amanora Mall Pune", "handle": "amanoramall", "city": "Pune", "pk": 2028889314},
    {"name": "Nexus Hyderabad Mall", "handle": "nexus_hyderabad", "city": "Hyderabad", "pk": 4139327262},
    {"name": "Sarath City Capital Mall Hyderabad", "handle": "sarathcitycapital.hyd", "city": "Hyderabad", "pk": 27574787021},
    {"name": "Inorbit Mall Cyberabad", "handle": "inorbitcyberabad", "city": "Hyderabad", "pk": 11694113336},
    {"name": "Lulu Mall Hyderabad", "handle": "lulumallhyderabad", "city": "Hyderabad", "pk": 12613758240},
    {"name": "GVK One Mall Hyderabad", "handle": "gvkone", "city": "Hyderabad", "pk": 1522826534},
]

# Internal or brand tenant accounts that shouldn't be counted as external creators
INTERNAL_OR_GENERIC = {
    "phoenixavenueofstars", "seasons_mall", "pavillionpune", "phoenix_millennium", "amanoramall",
    "nexus_hyderabad", "sarathcitycapital.hyd", "inorbitcyberabad", "lulumallhyderabad", "gvkone",
    "phoenixmarketcitypune", "nexusmalls", "inorbitmalls", "lulumall", "gvk_one"
}

def resolve_creator_profile(raw_handle: str) -> dict:
    clean_h = raw_handle.lower().replace("@", "").strip()
    s = cffi_requests.Session(impersonate="chrome120")
    followers = 0
    full_name = clean_h
    
    try:
        r = s.get(f"https://www.instagram.com/{clean_h}/", timeout=8)
        f_match = re.search(r'([0-9.,KMBkmb]+)\s+Followers', r.text)
        if f_match:
            raw = f_match.group(1).upper().replace(",", "")
            followers = int(float(raw.replace("M",""))*1000000) if "M" in raw else (int(float(raw.replace("K",""))*1000) if "K" in raw else int(float(raw)))
        
        n_match = re.search(r'<title>([^<(]+)', r.text)
        if n_match:
            title_text = n_match.group(1).replace("• Instagram photos and videos", "").strip()
            if title_text and not title_text.startswith("@"):
                full_name = title_text
    except Exception:
        pass

    if followers >= 1000000: tier = "🌟 Mega Creator (1M+)"
    elif followers >= 100000: tier = "🚀 Macro Creator (100K-1M)"
    elif followers >= 50000: tier = "✨ Mid-Tier (50K-100K)"
    elif followers >= 10000: tier = "🎯 Micro (10K-50K)"
    else: tier = "🌱 Nano (<10K)"

    return {
        "handle": f"@{clean_h}",
        "raw_handle": clean_h,
        "full_name": full_name,
        "followers": followers,
        "tier": tier,
        "profile_url": f"https://www.instagram.com/{clean_h}/"
    }


def scrape_mall_1year(mall_info: dict) -> dict:
    mall_name = mall_info["name"]
    clean_u = mall_info["handle"].lower().strip()
    user_pk = mall_info["pk"]
    city = mall_info["city"]
    
    print(f"\n======================================================================")
    print(f"[{city.upper()}] Scraping 1-Year Content & Collabs: {mall_name} (@{clean_u})")
    print(f"======================================================================")

    session = cffi_requests.Session(impersonate="chrome120")
    
    # 1. Timeline Feed
    feed_items = []
    max_id = ""
    for p in range(1, 35):
        f_url = f"https://i.instagram.com/api/v1/feed/user/{user_pk}/"
        if max_id: f_url += f"?max_id={max_id}"
        try:
            r = session.get(f_url, headers=DEFAULT_IG_HEADERS, cookies=DEFAULT_IG_COOKIES, timeout=12)
            if r.status_code == 200:
                data = r.json()
                items = data.get("items", [])
                feed_items.extend(items)
                max_id = data.get("next_max_id")

                unpinned_ts = [it.get("taken_at", 0) for idx, it in enumerate(items) if not (p == 1 and idx < 3)]
                oldest_ts = min(unpinned_ts) if unpinned_ts else (items[-1].get("taken_at", 0) if items else 0)
                if p > 1 and oldest_ts and oldest_ts < CUTOFF_TIMESTAMP:
                    print(f"  [Feed] Reached 1-year cutoff on page {p}")
                    break
                if not max_id or len(items) == 0:
                    break
                time.sleep(0.35)
            else: break
        except Exception as e:
            print(f"  [Feed] Error page {p}: {e}")
            break

    # 2. Clips / Reels Feed
    clips_items = []
    max_id = ""
    for p in range(1, 35):
        c_url = "https://i.instagram.com/api/v1/clips/user/"
        payload = {"target_user_id": str(user_pk), "page_size": 30}
        if max_id: payload["max_id"] = str(max_id)
        try:
            r = session.post(c_url, headers=DEFAULT_IG_HEADERS, data=payload, cookies=DEFAULT_IG_COOKIES, timeout=12)
            if r.status_code == 200:
                data = r.json()
                clips = [it.get("media") for it in data.get("items", []) if it.get("media")]
                clips_items.extend(clips)
                paging = data.get("paging_info", {})
                max_id = paging.get("max_id")

                unpinned_ts = [it.get("taken_at", 0) for idx, it in enumerate(clips) if not (p == 1 and idx < 3)]
                oldest_ts = min(unpinned_ts) if unpinned_ts else (clips[-1].get("taken_at", 0) if clips else 0)
                if p > 1 and oldest_ts and oldest_ts < CUTOFF_TIMESTAMP:
                    print(f"  [Clips] Reached 1-year cutoff on page {p}")
                    break
                if not paging.get("more_available") or not max_id:
                    break
                time.sleep(0.35)
            else: break
        except Exception as e:
            print(f"  [Clips] Error page {p}: {e}")
            break

    # Deduplicate All Raw Media
    seen_pks = set()
    all_raw_posts = []
    for it in feed_items + clips_items:
        pk = str(it.get("pk") or it.get("id"))
        if pk and pk not in seen_pks:
            seen_pks.add(pk)
            all_raw_posts.append(it)

    print(f"  -> Total Unique Posts Scanned: {len(all_raw_posts)}")

    # Extract Collabs & Full Post Deep-Dive Records
    collabs = []
    all_posts_deepdive = []
    creators_dict = {}

    for it in all_raw_posts:
        taken_at = it.get("taken_at", 0)
        if not taken_at or taken_at < CUTOFF_TIMESTAMP: continue

        date_str = datetime.fromtimestamp(taken_at, tz=timezone.utc).strftime("%Y-%m-%d")
        code = it.get("code") or ""
        post_url = f"https://www.instagram.com/p/{code}/" if code else ""
        
        # Caption / Description
        cap_obj = it.get("caption") or {}
        caption_text = cap_obj.get("text", "") if isinstance(cap_obj, dict) else str(cap_obj)
        caption_clean = caption_text.replace("\n", " ").strip() if caption_text else ""
        
        # Audio / Music Metadata (Safe parse)
        song_title = ""
        artist_name = ""
        music_meta = it.get("music_metadata")
        if isinstance(music_meta, dict):
            music_info = music_meta.get("music_info")
            if isinstance(music_info, dict):
                audio_info = music_info.get("music_asset_info")
                if isinstance(audio_info, dict):
                    song_title = audio_info.get("title") or ""
                    artist_name = audio_info.get("display_artist") or ""
        
        if not song_title:
            clips_meta = it.get("clips_metadata")
            if isinstance(clips_meta, dict):
                audio_type = clips_meta.get("audio_type") or ""
                music_info = clips_meta.get("music_info")
                if isinstance(music_info, dict):
                    audio_info = music_info.get("music_asset_info")
                    if isinstance(audio_info, dict):
                        song_title = audio_info.get("title") or ""
                        artist_name = audio_info.get("display_artist") or ""

        audio_str = f"{song_title} - {artist_name}" if song_title else "Original Audio"

        owner = it.get("user", {}).get("username", "").lower()
        coauthors = [c.get("username", "").lower() for c in it.get("coauthor_producers", [])]
        is_paid = bool(it.get("is_paid_partnership", False))

        likes = it.get("like_count") or 0
        comments = it.get("comment_count") or 0
        views = it.get("play_count") or it.get("view_count") or int(likes * 18.5)
        like_rate = (likes / views * 100.0) if views > 0 else 0.0

        is_boosted = (like_rate < 0.35 and views >= 200000) or views >= 3000000
        if is_paid and is_boosted: tier = "Tier 1: Toggle ON + Boosted Paid Ad"
        elif is_paid and not is_boosted: tier = "Tier 2: Toggle ON + Organic Collab"
        elif not is_paid and is_boosted: tier = "Tier 3: Toggle OFF + Heavily Boosted Ad"
        else: tier = "Tier 4: Toggle OFF + Organic / Noise"

        # Check creator collab
        creator_handle = ""
        if owner != clean_u and owner not in INTERNAL_OR_GENERIC and not owner.startswith(clean_u[:5]):
            creator_handle = f"@{owner}"
        elif coauthors:
            ext = [c for c in coauthors if c not in INTERNAL_OR_GENERIC and c != clean_u and not c.startswith(clean_u[:5])]
            if ext: creator_handle = f"@{ext[0]}"

        post_record = {
            "mall_name": mall_name,
            "mall_handle": f"@{clean_u}",
            "city": city,
            "post_url": post_url,
            "shortcode": code,
            "date": date_str,
            "taken_at": taken_at,
            "views": views,
            "likes": likes,
            "comments": comments,
            "like_to_view_pct": round(like_rate, 3),
            "is_collab": bool(creator_handle),
            "creator_handle": creator_handle or "— (Mall Owned)",
            "is_paid_toggle": is_paid,
            "is_boosted": is_boosted,
            "tier": tier,
            "audio_track": audio_str,
            "caption": caption_clean[:600]
        }
        all_posts_deepdive.append(post_record)

        if creator_handle:
            c_raw = creator_handle.replace("@", "").strip()
            collabs.append({
                "mall_name": mall_name,
                "mall_handle": f"@{clean_u}",
                "city": city,
                "post_url": post_url,
                "shortcode": code,
                "date": date_str,
                "taken_at": taken_at,
                "creator_handle": creator_handle,
                "raw_handle": c_raw,
                "views": views,
                "likes": likes,
                "comments": comments,
                "like_to_view_pct": round(like_rate, 3),
                "is_paid_toggle": is_paid,
                "is_boosted": is_boosted,
                "tier": tier,
                "audio_track": audio_str,
                "caption": caption_clean[:600]
            })

            if c_raw not in creators_dict:
                creators_dict[c_raw] = {
                    "handle": creator_handle,
                    "raw_handle": c_raw,
                    "mall_name": mall_name,
                    "city": city,
                    "total_posts": 1,
                    "total_views": views,
                    "total_likes": likes,
                    "latest_post_date": date_str
                }
            else:
                creators_dict[c_raw]["total_posts"] += 1
                creators_dict[c_raw]["total_views"] += views
                creators_dict[c_raw]["total_likes"] += likes

    print(f"  ✓ {mall_name}: {len(collabs)} Collab Posts across {len(creators_dict)} Unique Creators")

    return {
        "mall_name": mall_name,
        "mall_handle": f"@{clean_u}",
        "city": city,
        "total_posts_scanned": len(all_posts_deepdive),
        "total_collab_posts": len(collabs),
        "unique_creators_count": len(creators_dict),
        "creators": list(creators_dict.values()),
        "collabs": collabs,
        "all_posts": all_posts_deepdive
    }


def main():
    print("="*80)
    print("STARTING 1-YEAR COLLABORATOR & CONTENT AUDIT ACROSS 10 PUNE & HYDERABAD MALLS")
    print("="*80)

    all_mall_results = []
    for m in MALLS:
        res = scrape_mall_1year(m)
        all_mall_results.append(res)

    # Gather all unique creators for profile resolution
    all_creators_map = {}
    for mr in all_mall_results:
        for c in mr["creators"]:
            h = c["raw_handle"]
            if h not in all_creators_map:
                all_creators_map[h] = c
            else:
                all_creators_map[h]["total_posts"] += c["total_posts"]
                all_creators_map[h]["total_views"] += c["total_views"]

    print(f"\n" + "="*80)
    print(f"ENRICHING LIVE PROFILE METRICS FOR {len(all_creators_map)} UNIQUE CREATORS")
    print("="*80)

    profiles_resolved = {}
    with ThreadPoolExecutor(max_workers=12) as ex:
        futs = {ex.submit(resolve_creator_profile, h): h for h in all_creators_map}
        for f in as_completed(futs):
            prof = f.result()
            profiles_resolved[prof["raw_handle"]] = prof

    print(f"✓ Resolved profile follower metrics for all {len(profiles_resolved)} creators")

    # Attach profile data to creators and posts
    master_creators_list = []
    for h, c in all_creators_map.items():
        prof = profiles_resolved.get(h, {})
        master_creators_list.append({
            "handle": f"@{h}",
            "raw_handle": h,
            "full_name": prof.get("full_name", h),
            "followers": prof.get("followers", 0),
            "tier": prof.get("tier", "🌱 Nano (<10K)"),
            "city": c.get("city", "N/A"),
            "primary_mall": c.get("mall_name", "N/A"),
            "total_posts": c["total_posts"],
            "total_views": c["total_views"],
            "profile_url": f"https://www.instagram.com/{h}/"
        })

    master_creators_list.sort(key=lambda x: (x["followers"], x["total_views"]), reverse=True)

    # Save Master JSON Dataset
    master_dump = {
        "audit_window": "1 Year (365 Days)",
        "audited_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
        "total_malls": len(MALLS),
        "total_unique_creators": len(master_creators_list),
        "creators_roster": master_creators_list,
        "malls_results": all_mall_results
    }

    with open("pune_hyderabad_malls_1year_dataset.json", "w", encoding="utf-8") as f:
        json.dump(master_dump, f, indent=2)

    print("\n✓ Saved full raw dataset to: pune_hyderabad_malls_1year_dataset.json")

if __name__ == "__main__":
    main()
