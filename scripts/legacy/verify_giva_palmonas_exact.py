"""
Exact Ground-Truth Verification for GIVA & Palmonas Posts using raw media_id (pk)
"""

import sys, json, time
from scrape_bulk import make_session, COOKIES

sys.stdout.reconfigure(encoding="utf-8")
session = make_session()

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

with open("giva_palmonas_scraped.json", encoding="utf-8") as f:
    posts = json.load(f)

print(f"Verifying exact ground truth for all {len(posts)} GIVA & Palmonas posts...\n")

verified_posts = []

for idx, p in enumerate(posts, 1):
    mid = p["media_id"]
    u = p["url"]
    b = p["brand"]
    h = p["handle"]
    
    url_api = f"https://i.instagram.com/api/v1/media/{mid}/info/"
    try:
        r = session.get(url_api, headers=mob_hdrs, cookies=COOKIES, timeout=12)
        if r.status_code == 200:
            item = r.json().get("items", [])[0]
            
            # Check ALL possible paid partnership flags
            is_paid_root = bool(item.get("is_paid_partnership", False))
            sponsor_tags = item.get("sponsor_tags", [])
            has_sponsor = bool(sponsor_tags and len(sponsor_tags) > 0)
            clips_meta = item.get("clips_metadata", {})
            clips_paid = bool(clips_meta.get("is_paid_partnership", False)) if clips_meta else False
            
            # The ONLY condition where Instagram renders "Paid partnership" text under the handle:
            # item has `sponsor_tags` OR `is_paid_partnership` is explicitly True WITH a tagged sponsor
            # or `is_paid_partnership == True` at root
            
            # Let's inspect the exact values
            exact_toggle_on = is_paid_root or has_sponsor
            
            # Check co-authors
            coauthors = [c.get("username") for c in item.get("coauthor_producers", [])]
            
            p["verified_toggle_on"] = exact_toggle_on
            p["sponsor_tags"] = sponsor_tags
            p["is_paid_root"] = is_paid_root
            p["has_sponsor"] = has_sponsor
            
            status_str = "🟢 ON (Paid Label)" if exact_toggle_on else "⚪ OFF (Collab Only)"
            print(f"[{idx:>3}/112] {b:<15} | {h:<22} -> Toggle: {status_str} (is_paid={is_paid_root}, sponsor={has_sponsor})")
            
        else:
            print(f"[{idx:>3}/112] {b:<15} | {h:<22} -> HTTP {r.status_code}")
            p["verified_toggle_on"] = False
    except Exception as e:
        print(f"[{idx:>3}/112] Error: {e}")
        p["verified_toggle_on"] = False
        
    time.sleep(0.3)

with open("giva_palmonas_verified_ground_truth.json", "w", encoding="utf-8") as f:
    json.dump(posts, f, indent=2)

true_on_cnt = sum(1 for p in posts if p.get("verified_toggle_on"))
true_off_cnt = len(posts) - true_on_cnt
print(f"\n=======================================================")
print(f"Verified Ground Truth Summary for GIVA & Palmonas:")
print(f"• Total Posts: {len(posts)}")
print(f"• 🟢 Verified Toggle ON (with 'Paid partnership' tag): {true_on_cnt}")
print(f"• ⚪ Verified Toggle OFF (Co-author / Collab Invite):    {true_off_cnt}")
print(f"=======================================================\n")
