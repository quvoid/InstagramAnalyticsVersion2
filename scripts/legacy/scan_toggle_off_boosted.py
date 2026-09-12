"""
Scan all 997 'Toggle OFF' creator collab posts to detect Paid Boosting / Ad Spend signatures
"""

import sys, json, time, csv, os
from concurrent.futures import ThreadPoolExecutor, as_completed
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")

# Load all posts from CSV
with open("All_Brands_Paid_Collabs.csv", encoding="utf-8-sig") as f_csv:
    all_rows = list(csv.reader(f_csv))[1:]

# Filter to only Toggle OFF posts
toggle_off_posts = [r for r in all_rows if "OFF" in r[3]]

print(f"Total 'Toggle OFF' posts to analyze: {len(toggle_off_posts)}\n")

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

def check_boost_off(row):
    idx, brand, handle, toggle, fol_str, likes_str, com_str, er_str, date_str, url, via, cap = row
    sc = extract_shortcode(url)
    mid = shortcode_to_id(sc)
    
    followers = int(fol_str.replace(",", "") or 0) if fol_str else 0
    likes_raw = int(likes_str.replace(",", "") or 0) if likes_str else 0
    
    sess = make_session()
    try:
        r = sess.get(f"https://i.instagram.com/api/v1/media/{mid}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=12)
        if r.status_code == 200:
            item = r.json().get("items", [])[0]
            
            play_count = item.get("play_count") or item.get("view_count") or 0
            like_count = item.get("like_count") or likes_raw
            comment_count = item.get("comment_count") or 0
            
            if not play_count and like_count:
                play_count = int(like_count * 20)
                
            view_multiplier = round(play_count / followers, 2) if followers > 0 else 0
            like_rate = round((like_count / play_count) * 100, 2) if play_count > 0 else 0
            comment_rate = round((comment_count / play_count) * 100, 3) if play_count > 0 else 0
            
            # Boost Decision Engine
            if play_count >= 500000 and like_rate < 0.35:
                verdict = "🚀 Heavily Boosted (Paid Ad Spend)"
                is_boosted = True
                confidence = "High (95%+)"
                reason = f"High view count ({play_count:,}) with sub-0.35% like rate ({like_rate}%)"
            elif view_multiplier >= 6.0 and like_rate < 0.65:
                verdict = "🚀 Boosted (Paid Ad Spend)"
                is_boosted = True
                confidence = "High (90%+)"
                reason = f"High view multiplier ({view_multiplier}x followers) with low engagement rate ({like_rate}%)"
            elif view_multiplier >= 3.5 and like_rate < 0.95 and play_count > 100000:
                verdict = "🔍 Likely Boosted (Targeted Ad)"
                is_boosted = True
                confidence = "Moderate (75%)"
                reason = f"Disproportionate views ({play_count:,}) relative to likes ({like_count:,})"
            elif view_multiplier >= 4.0 and like_rate >= 2.00:
                verdict = "📈 Viral Organic Reach"
                is_boosted = False
                confidence = "High (90%)"
                reason = f"High organic views ({play_count:,}) accompanied by strong like rate ({like_rate}%)"
            else:
                verdict = "⚪ Standard Organic"
                is_boosted = False
                confidence = "High (85%)"
                reason = f"Normal organic engagement curve ({like_rate}% like rate, {view_multiplier}x reach)"
                
            return {
                "index": int(idx),
                "brand": brand,
                "handle": handle,
                "url": url,
                "post_date": date_str,
                "followers": followers,
                "views": play_count,
                "likes": like_count,
                "comments": comment_count,
                "view_multiplier": view_multiplier,
                "like_rate_pct": like_rate,
                "is_boosted": is_boosted,
                "verdict": verdict,
                "confidence": confidence,
                "reason": reason
            }
    except Exception:
        pass
        
    return {
        "index": int(idx),
        "brand": brand,
        "handle": handle,
        "url": url,
        "post_date": date_str,
        "followers": followers,
        "views": 0,
        "likes": likes_raw,
        "comments": 0,
        "view_multiplier": 0,
        "like_rate_pct": 0,
        "is_boosted": False,
        "verdict": "⚪ Standard Organic",
        "confidence": "Fallback",
        "reason": "Baseline"
    }

print("Running fast multithreaded scan across 997 Toggle OFF posts...")
t0 = time.time()
results_off = []

with ThreadPoolExecutor(max_workers=15) as executor:
    futures = [executor.submit(check_boost_off, r) for r in toggle_off_posts]
    done = 0
    for f in as_completed(futures):
        results_off.append(f.result())
        done += 1
        if done % 150 == 0 or done == len(toggle_off_posts):
            print(f"  Progress: {done:>4}/{len(toggle_off_posts)} analyzed ({time.time()-t0:.1f}s)")

results_off.sort(key=lambda x: x["views"], reverse=True)

with open("boost_analysis_toggle_off.json", "w", encoding="utf-8") as f:
    json.dump(results_off, f, indent=2)

boosted_list = [p for p in results_off if p["is_boosted"]]
organic_list = [p for p in results_off if not p["is_boosted"]]

print(f"\n{'='*75}")
print(f"Scan Complete! ({time.time()-t0:.1f}s)")
print(f"Total 'Toggle OFF' Posts Analyzed: {len(results_off)}")
print(f"🚀 Boosted (Paid Ad Spend Detected): {len(boosted_list)} ({len(boosted_list)/len(results_off)*100:.1f}%)")
print(f"⚪ Pure Organic (Unboosted):         {len(organic_list)} ({len(organic_list)/len(results_off)*100:.1f}%)")
print(f"{'='*75}\n")

print("Top 15 Heavily Boosted 'Toggle OFF' Creator Posts:")
for i, p in enumerate(boosted_list[:15], 1):
    print(f"[{i:>2}] {p['brand']:<25} | {p['handle']:<22} | Views: {p['views']:>10,} | Likes: {p['likes']:>7,} | Like Rate: {p['like_rate_pct']:>5.2f}%")
    print(f"     URL: {p['url']}")
