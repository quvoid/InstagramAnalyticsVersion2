"""
Download all 70 verified Toggle ON creator creatives (videos/images) into the 'video' folder
"""

import sys, os, json, time, re
from curl_cffi import requests
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")

OUTPUT_DIR = os.path.join(os.getcwd(), "video")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Load ground truth API results
with open("api_toggle_ground_truth.json", encoding="utf-8") as f:
    api_results = json.load(f)

# Filter to only the 70 verified Toggle ON posts
verified_posts = [v for v in api_results.values() if v.get("is_paid_partnership")]

print(f"Total Verified Creatives to download: {len(verified_posts)}")
print(f"Target Directory: {OUTPUT_DIR}\n")

mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

session = make_session()

download_summary = []

for idx, p in enumerate(verified_posts, 1):
    url = p["url"]
    brand = p["brand"]
    handle = p["handle"].replace("@", "")
    sc = extract_shortcode(url)
    mid = shortcode_to_id(sc)
    
    brand_clean = re.sub(r'[^a-zA-Z0-9]', '_', brand).strip('_')
    handle_clean = re.sub(r'[^a-zA-Z0-9]', '_', handle).strip('_')
    
    print(f"[{idx:>2}/{len(verified_posts)}] Fetching media for {brand} | @{handle} ({sc}) ...", end=" ", flush=True)
    
    try:
        r = session.get(f"https://i.instagram.com/api/v1/media/{mid}/info/", headers=mob_hdrs, cookies=COOKIES, timeout=15)
        if r.status_code == 200:
            item = r.json().get("items", [])[0]
            
            media_type = item.get("media_type") # 1 = Image, 2 = Video, 8 = Carousel
            video_versions = item.get("video_versions", [])
            
            # 1. Direct Video
            if video_versions:
                video_url = video_versions[0]["url"]
                filename = f"{idx:02d}_{brand_clean}_{handle_clean}_{sc}.mp4"
                filepath = os.path.join(OUTPUT_DIR, filename)
                
                # Download stream
                r_dl = session.get(video_url, timeout=30)
                if r_dl.status_code == 200:
                    with open(filepath, "wb") as f_out:
                        f_out.write(r_dl.content)
                    sz_mb = len(r_dl.content) / (1024 * 1024)
                    print(f"✓ Video saved ({sz_mb:.1f} MB) -> {filename}")
                    download_summary.append({
                        "index": idx, "brand": brand, "handle": handle, "type": "Video (.mp4)",
                        "filename": filename, "size_mb": round(sz_mb, 2), "url": url
                    })
                else:
                    print(f"⚠ Video DL HTTP {r_dl.status_code}")
                    
            # 2. Carousel
            elif "carousel_media" in item:
                c_media = item.get("carousel_media", [])
                saved_sub = 0
                for c_idx, child in enumerate(c_media, 1):
                    c_vids = child.get("video_versions", [])
                    if c_vids:
                        v_url = c_vids[0]["url"]
                        filename = f"{idx:02d}_{brand_clean}_{handle_clean}_{sc}_slide{c_idx}.mp4"
                        filepath = os.path.join(OUTPUT_DIR, filename)
                        r_dl = session.get(v_url, timeout=30)
                        if r_dl.status_code == 200:
                            with open(filepath, "wb") as f_out:
                                f_out.write(r_dl.content)
                            saved_sub += 1
                    else:
                        c_imgs = child.get("image_versions2", {}).get("candidates", [])
                        if c_imgs:
                            img_url = c_imgs[0]["url"]
                            filename = f"{idx:02d}_{brand_clean}_{handle_clean}_{sc}_slide{c_idx}.jpg"
                            filepath = os.path.join(OUTPUT_DIR, filename)
                            r_dl = session.get(img_url, timeout=30)
                            if r_dl.status_code == 200:
                                with open(filepath, "wb") as f_out:
                                    f_out.write(r_dl.content)
                                saved_sub += 1
                print(f"✓ Carousel saved ({saved_sub} slides)")
                download_summary.append({
                    "index": idx, "brand": brand, "handle": handle, "type": f"Carousel ({saved_sub} slides)",
                    "filename": f"{idx:02d}_{brand_clean}_{handle_clean}_{sc}_*", "size_mb": 0, "url": url
                })
                
            # 3. Static Photo
            else:
                imgs = item.get("image_versions2", {}).get("candidates", [])
                if imgs:
                    img_url = imgs[0]["url"]
                    filename = f"{idx:02d}_{brand_clean}_{handle_clean}_{sc}.jpg"
                    filepath = os.path.join(OUTPUT_DIR, filename)
                    r_dl = session.get(img_url, timeout=30)
                    if r_dl.status_code == 200:
                        with open(filepath, "wb") as f_out:
                            f_out.write(r_dl.content)
                        sz_kb = len(r_dl.content) / 1024
                        print(f"✓ Photo saved ({sz_kb:.0f} KB) -> {filename}")
                        download_summary.append({
                            "index": idx, "brand": brand, "handle": handle, "type": "Photo (.jpg)",
                            "filename": filename, "size_mb": round(sz_kb / 1024, 2), "url": url
                        })
        else:
            print(f"⚠ API HTTP {r.status_code}")
    except Exception as e:
        print(f"⚠ Error: {e}")
        
    time.sleep(0.5)

print(f"\n{'='*75}")
print(f"Download Finished! Total Creatives Processed: {len(download_summary)}")
print(f"Target Directory: {OUTPUT_DIR}")
print(f"{'='*75}\n")

with open(os.path.join(OUTPUT_DIR, "manifest.json"), "w", encoding="utf-8") as f_m:
    json.dump(download_summary, f_m, indent=2)
