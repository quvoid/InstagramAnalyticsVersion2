"""
Scrape 1-Year Complete Collaborator & Content Audit for Lake Shore Brand Accounts:
1. KOPA Mall Pune (@kopapune, PK: 62256024228)
2. Lake Shore Y Junction Hyderabad (@lakeshoreyjunction, PK: 74989964663)

Integrate with existing 10 competitor malls dataset, resolve new creators, and update master datasets.
"""

import sys, os, json, time, re
from collections import defaultdict
from datetime import datetime, timezone, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from curl_cffi import requests as cffi_requests
from api_wrapper.client import DEFAULT_IG_COOKIES, DEFAULT_IG_HEADERS

sys.stdout.reconfigure(encoding="utf-8")

NOW_DT = datetime.now(timezone.utc)
CUTOFF_DT = NOW_DT - timedelta(days=365)
CUTOFF_TIMESTAMP = int(CUTOFF_DT.timestamp())

LAKESHORE_MALLS = [
    {"name": "KOPA Mall Pune (Lake Shore)", "handle": "kopapune", "city": "Pune", "pk": 62256024228, "is_client": True},
    {"name": "Lake Shore Y Junction (Hyderabad)", "handle": "lakeshoreyjunction", "city": "Hyderabad", "pk": 74989964663, "is_client": True}
]

INTERNAL_OR_GENERIC = {
    "kopapune", "lakeshoreyjunction", "lakeshoreindia", "lakeshore",
    "phoenixavenueofstars", "seasons_mall", "pavillionpune", "phoenix_millennium", "amanoramall",
    "nexus_hyderabad", "sarathcitycapital.hyd", "inorbitcyberabad", "lulumallhyderabad", "gvkone"
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
    print(f"[LAKE SHORE CLIENT] Scraping 1-Year Audit: {mall_name} (@{clean_u})")
    print(f"======================================================================")

    session = cffi_requests.Session(impersonate="chrome120")
    
    # 1. Timeline Feed
    feed_items = []
    max_id = ""
    for p in range(1, 35):
        f_url = f"https://i.instagram.com/api/v1/feed/user/{user_pk}/"
        if max_id:
            f_url += f"?max_id={max_id}"
        try:
            r = session.get(f_url, headers=DEFAULT_IG_HEADERS, cookies=DEFAULT_IG_COOKIES, timeout=12)
            if r.status_code == 200:
                d = r.json()
                items = d.get("items", [])
                if not items: break
                
                oldest_in_batch = items[-1].get("taken_at", 0)
                feed_items.extend(items)
                print(f"  • Feed Page {p:02d}: Got {len(items)} items | Total so far: {len(feed_items)}")
                
                if oldest_in_batch < CUTOFF_TIMESTAMP:
                    print(f"  ✓ Reached 1-Year Cutoff in Feed ({datetime.fromtimestamp(oldest_in_batch, tz=timezone.utc).strftime('%Y-%m-%d')})")
                    break
                    
                max_id = d.get("next_max_id")
                if not max_id: break
                time.sleep(1.0)
            else:
                print(f"  ⚠ Feed API returned {r.status_code}")
                break
        except Exception as e:
            print(f"  ⚠ Feed exception: {e}")
            break

    # 2. Reels Grid
    reels_items = []
    max_cursor = ""
    for p in range(1, 35):
        r_url = "https://i.instagram.com/api/v1/clips/user/"
        payload = {
            "target_user_id": user_pk,
            "page_size": 24,
            "max_id": max_cursor
        }
        try:
            r = session.post(r_url, data=payload, headers=DEFAULT_IG_HEADERS, cookies=DEFAULT_IG_COOKIES, timeout=12)
            if r.status_code == 200:
                d = r.json()
                items = [x.get("media") for x in d.get("items", []) if x.get("media")]
                if not items: break
                
                oldest_in_batch = items[-1].get("taken_at", 0)
                reels_items.extend(items)
                print(f"  • Reels Page {p:02d}: Got {len(items)} reels | Total so far: {len(reels_items)}")
                
                if oldest_in_batch < CUTOFF_TIMESTAMP:
                    print(f"  ✓ Reached 1-Year Cutoff in Reels ({datetime.fromtimestamp(oldest_in_batch, tz=timezone.utc).strftime('%Y-%m-%d')})")
                    break
                    
                max_cursor = d.get("paging_info", {}).get("max_id")
                if not max_cursor or not d.get("paging_info", {}).get("more_available"): break
                time.sleep(1.0)
            else:
                print(f"  ⚠ Reels API returned {r.status_code}")
                break
        except Exception as e:
            print(f"  ⚠ Reels exception: {e}")
            break

    # 3. Deduplicate
    seen_ids = set()
    unique_posts = []
    for it in (feed_items + reels_items):
        if not it: continue
        pid = it.get("pk") or it.get("id")
        if pid not in seen_ids:
            seen_ids.add(pid)
            unique_posts.append(it)

    print(f"\n→ Total Unique Posts/Reels Harvested: {len(unique_posts)}")

    # 4. Filter strictly to Last 1 Year & Extract Collaborators
    scanned_1year = []
    collabs_1year = []
    
    for it in unique_posts:
        ts = it.get("taken_at", 0)
        if ts < CUTOFF_TIMESTAMP: continue
        
        post_date = datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")
        sc = it.get("code") or it.get("shortcode", "")
        post_url = f"https://www.instagram.com/reel/{sc}/" if it.get("media_type") == 2 else f"https://www.instagram.com/p/{sc}/"
        
        views = it.get("play_count") or it.get("view_count") or 0
        likes = it.get("like_count") or 0
        comments = it.get("comment_count") or 0
        
        # Caption
        cap_dict = it.get("caption")
        caption = cap_dict.get("text", "") if isinstance(cap_dict, dict) else ""
        
        # Music info
        music_meta = it.get("music_metadata") or {}
        clips_meta = it.get("clips_metadata") or {}
        music_info = (music_meta.get("music_info") or {}) if isinstance(music_meta, dict) else {}
        music_asset = (music_info.get("music_asset_info") or {}) if isinstance(music_info, dict) else {}
        audio_track = music_asset.get("title") or clips_meta.get("audio_type") or "Original Audio"
        audio_artist = music_asset.get("display_artist") or ""
        if audio_artist:
            audio_track = f"{audio_track} - {audio_artist}"

        # Coauthor / Tagged Collaborator extraction
        coauthors = it.get("coauthor_producers", []) or []
        tagged_users = it.get("usertags", {}).get("in", []) if isinstance(it.get("usertags"), dict) else []
        
        collab_candidates = []
        for ca in coauthors:
            if isinstance(ca, dict):
                u = ca.get("username", "").lower()
                if u and u != clean_u and u not in INTERNAL_OR_GENERIC:
                    collab_candidates.append(u)
                    
        for tu in tagged_users:
            if isinstance(tu, dict):
                u = tu.get("user", {}).get("username", "").lower()
                if u and u != clean_u and u not in INTERNAL_OR_GENERIC and u not in collab_candidates:
                    collab_candidates.append(u)

        is_collab = len(collab_candidates) > 0
        primary_collab = collab_candidates[0] if is_collab else "—"
        all_collabs_str = ", ".join([f"@{x}" for x in collab_candidates]) if is_collab else "—"

        post_record = {
            "mall_name": mall_name,
            "mall_handle": f"@{clean_u}",
            "city": city,
            "is_client": True,
            "post_id": str(it.get("pk")),
            "shortcode": sc,
            "post_url": post_url,
            "date": post_date,
            "timestamp": ts,
            "media_type": "Video/Reel" if it.get("media_type") == 2 else "Image/Carousel",
            "views": views,
            "likes": likes,
            "comments": comments,
            "like_to_view_ratio": f"{(likes/views*100):.2f}%" if views > 0 else "N/A",
            "is_collab": is_collab,
            "creator_handle": f"@{primary_collab}" if is_collab else "—",
            "all_creators_tagged": all_collabs_str,
            "audio_track": audio_track,
            "caption": caption
        }
        
        scanned_1year.append(post_record)
        
        if is_collab:
            for ch in collab_candidates:
                collabs_1year.append({
                    "mall_name": mall_name,
                    "mall_handle": f"@{clean_u}",
                    "city": city,
                    "is_client": True,
                    "post_id": str(it.get("pk")),
                    "shortcode": sc,
                    "post_url": post_url,
                    "date": post_date,
                    "timestamp": ts,
                    "views": views,
                    "likes": likes,
                    "comments": comments,
                    "like_to_view_ratio": f"{(likes/views*100):.2f}%" if views > 0 else "N/A",
                    "raw_handle": ch,
                    "creator_handle": f"@{ch}",
                    "caption": caption,
                    "audio_track": audio_track
                })

    print(f"✓ 1-Year Audit: {len(scanned_1year)} Total Posts Scanned | {len(collabs_1year)} Collaborative Posts Found")
    
    return {
        "mall_name": mall_name,
        "handle": f"@{clean_u}",
        "raw_handle": clean_u,
        "city": city,
        "is_client": True,
        "total_posts_1year": len(scanned_1year),
        "total_collabs_1year": len(collabs_1year),
        "all_posts": scanned_1year,
        "collabs": collabs_1year
    }

def main():
    print("="*80)
    print("STARTING 1-YEAR SCRAPING AUDIT FOR LAKE SHORE ACCOUNTS")
    print("="*80)

    lakeshore_results = []
    new_creators_to_resolve = set()

    for m in LAKESHORE_MALLS:
        res = scrape_mall_1year(m)
        lakeshore_results.append(res)
        for c in res["collabs"]:
            new_creators_to_resolve.add(c["raw_handle"])

    # Load existing dataset
    with open("pune_hyderabad_malls_1year_dataset.json", encoding="utf-8") as f:
        existing_data = json.load(f)

    existing_malls = existing_data.get("malls_results", [])
    existing_roster = {c["raw_handle"]: c for c in existing_data.get("creators_roster", [])}

    # Resolve newly discovered creators
    unresolved = [h for h in new_creators_to_resolve if h not in existing_roster]
    print(f"\nResolving {len(unresolved)} Newly Discovered Lake Shore Creators...")

    with ThreadPoolExecutor(max_workers=10) as ex:
        futs = {ex.submit(resolve_creator_profile, h): h for h in unresolved}
        for f in as_completed(futs):
            prof = f.result()
            existing_roster[prof["raw_handle"]] = prof
            print(f"  ✓ Resolved Creator: {prof['handle']:<24} | {prof['followers']:>9,} Followers | Tier: {prof['tier']}")

    # Merge results
    # Remove existing entry if already present to avoid duplicates
    merged_malls = [m for m in existing_malls if (m.get("raw_handle") or m.get("handle", "").replace("@", "")) not in ("kopapune", "lakeshoreyjunction")]
    merged_malls.extend(lakeshore_results)

    # Re-calculate 4-tier paid classifications for all collabs
    all_collabs_flat = []
    for mr in merged_malls:
        for c in mr["collabs"]:
            h = c["raw_handle"]
            prof = existing_roster.get(h, {})
            c["full_name"] = prof.get("full_name", h)
            c["followers"] = prof.get("followers", 0)
            c["tier"] = prof.get("tier", "🌱 Nano (<10K)")
            
            # 4-tier heuristic
            cap_l = c["caption"].lower()
            if any(k in cap_l for k in ["#ad", "#sponsored", "paid partnership", "collab with", "in collaboration", "#paid"]):
                paid_tier = "Tier 1: Paid Partnership Toggle Active (Dark Ads)"
                paid_badge = "🔴 T1: Paid Toggle ON (Dark Ads Active)"
            elif any(k in cap_l for k in ["invite", "courtesy", "experience", "hosted", "gifted", "pr"]):
                paid_tier = "Tier 2: Direct Creator Collab (Organic Grid + Potential Dark Boosting)"
                paid_badge = "🟠 T2: Direct Collab (Dark Boost Potential)"
            elif c["views"] >= 50000 or (c["likes"] >= 2500 and c["followers"] < 20000):
                paid_tier = "Tier 3: Boosted Organic Grid Post (High View Spike)"
                paid_badge = "🟡 T3: Boosted Organic Reel (Paid Footfall Spike)"
            else:
                paid_tier = "Tier 4: Pure Organic Barter Collab"
                paid_badge = "🟢 T4: Pure Organic Barter Collab"

            c["paid_classification"] = paid_tier
            c["paid_badge"] = paid_badge
            all_collabs_flat.append(c)

    # Build final creator roster
    creator_malls_map = defaultdict(set)
    creator_views_map = defaultdict(int)
    for c in all_collabs_flat:
        creator_malls_map[c["raw_handle"]].add(c["mall_name"])
        creator_views_map[c["raw_handle"]] += c["views"]

    final_creators_list = []
    for h, prof in existing_roster.items():
        if h in creator_malls_map:
            final_creators_list.append({
                "handle": prof["handle"],
                "raw_handle": h,
                "full_name": prof.get("full_name", h),
                "followers": prof.get("followers", 0),
                "tier": prof.get("tier", "🌱 Nano (<10K)"),
                "profile_url": prof.get("profile_url", f"https://www.instagram.com/{h}/"),
                "primary_mall": list(creator_malls_map[h])[0],
                "all_malls_collaborated": ", ".join(creator_malls_map[h]),
                "total_collabs_done": sum(1 for c in all_collabs_flat if c["raw_handle"] == h),
                "total_views_generated": creator_views_map[h]
            })

    final_creators_list.sort(key=lambda x: (x["followers"], x["total_views_generated"]), reverse=True)

    # Save to JSON
    merged_data = {
        "audited_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
        "total_malls": len(merged_malls),
        "total_collabs_across_all_malls": len(all_collabs_flat),
        "total_unique_creators": len(final_creators_list),
        "creators_roster": final_creators_list,
        "malls_results": merged_malls
    }

    with open("pune_hyderabad_malls_1year_dataset.json", "w", encoding="utf-8") as f:
        json.dump(merged_data, f, indent=2)

    print(f"\n" + "="*80)
    print(f"LAKE SHORE INTEGRATION COMPLETE!")
    print(f"Total Malls: {len(merged_malls)} (including KOPA Pune & Lake Shore Y Junction)")
    print(f"Total Collabs: {len(all_collabs_flat)}")
    print(f"Total Unique Creators: {len(final_creators_list)}")
    print(f"Updated: pune_hyderabad_malls_1year_dataset.json")
    print("="*80)

if __name__ == "__main__":
    main()
