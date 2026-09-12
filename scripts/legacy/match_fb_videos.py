"""
Match Facebook Video URLs with their captions, dates, and campaign themes
"""

import json, sys

sys.stdout.reconfigure(encoding="utf-8")

with open("grt_facebook_all_posts.json", encoding="utf-8") as f:
    posts = json.load(f)

# Extract video IDs and matching captions
video_map = {}
for p in posts:
    url = p.get("url", "")
    pid = p.get("post_id", "")
    vw = p.get("video_views", 0)
    rc = p.get("reactions", 0)
    dt = p.get("date", "N/A")
    cap = p.get("caption", "").strip().replace("\n", " ")
    
    # Check if this post has a video ID or VK parameter
    vid = None
    if "posts/" in url:
        v_part = url.split("posts/")[-1]
        if v_part.isdigit() and len(v_part) > 5:
            vid = v_part
    elif "videos/" in url:
        v_part = url.split("videos/")[-1].strip("/")
        if v_part.isdigit():
            vid = v_part
            
    if vid:
        if vid not in video_map:
            video_map[vid] = {"id": vid, "views": vw, "reactions": rc, "date": dt, "caption": cap}
        else:
            if vw > video_map[vid]["views"]:
                video_map[vid]["views"] = vw
            if rc > video_map[vid]["reactions"]:
                video_map[vid]["reactions"] = rc
            if dt != "N/A":
                video_map[vid]["date"] = dt
            if cap and len(cap) > len(video_map[vid]["caption"]):
                video_map[vid]["caption"] = cap

# Also check base64 permalinks containing captions
import base64
for p in posts:
    url = p.get("url", "")
    cap = p.get("caption", "").strip().replace("\n", " ")
    if "Uzpf" in url and cap:
        b64_part = url.split("Uzpf")[-1]
        try:
            dec = base64.b64decode(b64_part + "==").decode("utf-8", errors="ignore")
            # find digit IDs inside dec
            import re
            ids = re.findall(r"\d{6,30}", dec)
            for d in ids:
                if d in video_map and not video_map[d]["caption"]:
                    video_map[d]["caption"] = cap
        except Exception:
            pass

print(f"Total Video Posts: {len(video_map)}")
results = list(video_map.values())
results.sort(key=lambda x: x["views"], reverse=True)

for i, v in enumerate(results, 1):
    direct_fb_url = f"https://www.facebook.com/grtjewellers/videos/{v['id']}/"
    print(f"[{i:>2}] Date: {v['date']} | Views: {v['views']:>10,} | Reacts: {v['reactions']:>6,} | URL: {direct_fb_url}")
    if v['caption']:
        print(f"     Caption: {v['caption'][:120]}")
