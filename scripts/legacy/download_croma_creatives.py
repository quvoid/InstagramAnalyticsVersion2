"""
Download Toggle ON videos for Croma & Electronics Competitors into video_croma/
"""

import sys, json, os, re
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")

VIDEO_DIR = os.path.join(os.getcwd(), "video_croma")
os.makedirs(VIDEO_DIR, exist_ok=True)

with open("croma_4tier_master_dataset.json", encoding="utf-8") as f:
    posts = json.load(f)

toggle_on_posts = [p for p in posts if p["tier"] in (1, 2)]
print(f"Total Toggle ON posts to download: {len(toggle_on_posts)}\n")

session = make_session()
mob_hdrs = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

for idx, p in enumerate(toggle_on_posts, 1):
    u = p["url"]
    sc = extract_shortcode(u)
    b_clean = re.sub(r'[^a-zA-Z0-9]', '_', p["brand"]).strip('_')
    h_clean = re.sub(r'[^a-zA-Z0-9]', '_', p["handle"].replace("@", "")).strip('_')
    
    print(f"[{idx:>2}/{len(toggle_on_posts)}] {p['brand']} | {p['handle']} ({sc})...", end=" ")
    try:
        # Try fetching web page video stream directly
        r_page = session.get(u, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}, cookies=COOKIES, timeout=12)
        if r_page.status_code == 200:
            v_matches = re.findall(r'"video_url":"([^"]+)"', r_page.text)
            if v_matches:
                v_url = v_matches[0].replace(r"\u0026", "&")
                fn = f"{idx:02d}_{b_clean}_{h_clean}_{sc}.mp4"
                fp = os.path.join(VIDEO_DIR, fn)
                r_dl = session.get(v_url, timeout=30)
                if r_dl.status_code == 200:
                    with open(fp, "wb") as f_out:
                        f_out.write(r_dl.content)
                    print(f"✓ Video saved ({len(r_dl.content)/(1024*1024):.1f} MB) -> {fn}")
                    continue
            # Try image candidate
            img_matches = re.findall(r'"display_url":"([^"]+)"', r_page.text)
            if img_matches:
                img_url = img_matches[0].replace(r"\u0026", "&")
                fn = f"{idx:02d}_{b_clean}_{h_clean}_{sc}.jpg"
                fp = os.path.join(VIDEO_DIR, fn)
                r_dl = session.get(img_url, timeout=30)
                if r_dl.status_code == 200:
                    with open(fp, "wb") as f_out:
                        f_out.write(r_dl.content)
                    print(f"✓ Image saved -> {fn}")
                    continue
        print("⚠ Media stream not publicly accessible")
    except Exception as e:
        print(f"⚠ Error: {e}")

print(f"\n✓ Download process complete. Target: {VIDEO_DIR}")
