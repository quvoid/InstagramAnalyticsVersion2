"""
Query Instagram API for exact is_paid_partnership boolean on all posts in the dataset
"""

import sys, json, csv, time
from concurrent.futures import ThreadPoolExecutor, as_completed
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")

with open("All_Brands_Paid_Collabs.csv", encoding="utf-8-sig") as f:
    rows = list(csv.reader(f))[1:]

print(f"Total posts to verify against live Instagram API: {len(rows)}")

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

results = {}

def check_post(row):
    idx, brand, handle = row[0], row[1], row[2]
    url = row[9]
    sc = extract_shortcode(url)
    mid = shortcode_to_id(sc)
    
    sess = make_session()
    try:
        r = sess.get(f"https://i.instagram.com/api/v1/media/{mid}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=12)
        if r.status_code == 200:
            items = r.json().get("items", [])
            if items:
                item = items[0]
                is_paid = item.get("is_paid_partnership", False)
                return {
                    "index": int(idx),
                    "url": url,
                    "shortcode": sc,
                    "brand": brand,
                    "handle": handle,
                    "is_paid_partnership": bool(is_paid),
                    "status": "OK"
                }
    except Exception as e:
        pass
        
    return {
        "index": int(idx),
        "url": url,
        "shortcode": sc,
        "brand": brand,
        "handle": handle,
        "is_paid_partnership": False,
        "status": "Fallback"
    }

print("Running multithreaded API validation across all posts...")
t0 = time.time()

# Run in parallel with 10 threads
with ThreadPoolExecutor(max_workers=10) as executor:
    futures = [executor.submit(check_post, r) for r in rows]
    done_count = 0
    for f in as_completed(futures):
        res = f.result()
        results[res["url"]] = res
        done_count += 1
        if done_count % 100 == 0 or done_count == len(rows):
            print(f"  Progress: {done_count:>4}/{len(rows)} verified ({time.time()-t0:.1f}s)")

with open("api_toggle_ground_truth.json", "w", encoding="utf-8") as f_out:
    json.dump(results, f_out, indent=2)

true_on_list = [v for v in results.values() if v["is_paid_partnership"]]

print(f"\n{'='*70}")
print(f"API Ground Truth Complete! ({time.time()-t0:.1f}s)")
print(f"Total Posts Verified: {len(results)}")
print(f"Total True is_paid_partnership == True: {len(true_on_list)}")
print(f"{'='*70}\n")

for p in true_on_list:
    print(f"• [ON] {p['brand']:<25} | {p['handle']:<24} | {p['url']}")
