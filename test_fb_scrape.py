"""
Parse and extract Facebook Page Posts and Partnerships for GRT Jewellers
"""

import sys, re, json, time
from datetime import datetime, timezone
from curl_cffi import requests

sys.stdout.reconfigure(encoding="utf-8")

session = requests.Session(impersonate="chrome120")
headers = {
    "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
    "accept-language": "en-US,en;q=0.9",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
}

print("Fetching Facebook page for @grtjewellers...")
r = session.get("https://www.facebook.com/grtjewellers/", headers=headers, timeout=20)
html = r.text

print(f"Status: {r.status_code}, Length: {len(html)}")

# Extract Page ID
page_id = None
for pat in [r'"pageID":"(\d+)"', r'"page_id":"(\d+)"', r'"delegate_page":\{"id":"(\d+)"', r'"targetID":"(\d+)"', r'fb://page/(\d+)']:
    m = re.search(pat, html)
    if m:
        page_id = m.group(1)
        break

print("Detected Facebook Page ID:", page_id)

# Extract scripts
scripts = re.findall(r'<script type="application/json"[^>]*>(.*?)</script>', html, re.DOTALL)
print(f"Total script JSON tags: {len(scripts)}")

posts = []

def extract_stories(obj):
    if isinstance(obj, dict):
        # check if this is a story node
        if "story" in obj and isinstance(obj["story"], dict):
            extract_story_details(obj["story"])
        elif "node" in obj and isinstance(obj["node"], dict) and ("comet_sections" in obj["node"] or "creation_time" in obj["node"] or "story" in obj["node"]):
            extract_story_details(obj["node"])
        for k, v in obj.items():
            extract_stories(v)
    elif isinstance(obj, list):
        for item in obj:
            extract_stories(item)

def extract_story_details(node):
    post_id = node.get("post_id") or node.get("id") or (node.get("feedback") or {}).get("id")
    ts = node.get("creation_time") or (node.get("comet_sections") or {}).get("context_layout", {}).get("story", {}).get("comet_sections", {}).get("metadata", [{}])[0].get("story", {}).get("creation_time", 0)
    
    # Text / Message
    msg = ""
    msg_obj = node.get("message") or (node.get("comet_sections") or {}).get("content", {}).get("story", {}).get("message")
    if isinstance(msg_obj, dict):
        msg = msg_obj.get("text", "")
    elif isinstance(msg_obj, str):
        msg = msg_obj

    # Feedback / Reactions / Comments / Shares
    feedback = node.get("feedback") or {}
    react_count = (feedback.get("reaction_count") or {}).get("count", 0) if isinstance(feedback.get("reaction_count"), dict) else feedback.get("reaction_count", 0)
    com_count = (feedback.get("comments_count") or {}).get("total_count", 0) if isinstance(feedback.get("comments_count"), dict) else feedback.get("comment_count", 0)
    share_count = (feedback.get("share_count") or {}).get("count", 0) if isinstance(feedback.get("share_count"), dict) else feedback.get("share_count", 0)
    
    # Check for sponsored / branded content / co-authors
    actors = node.get("actors", []) or (node.get("comet_sections") or {}).get("header", {}).get("story", {}).get("actors", [])
    actor_names = [a.get("name") for a in actors if isinstance(a, dict) and a.get("name")]
    
    # Sponsors / Branded Content
    sponsors = node.get("sponsor_tags") or node.get("branded_content_sponsor_relationship") or []
    
    if msg or react_count or ts:
        date_str = datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d") if ts else "N/A"
        url = node.get("url") or f"https://www.facebook.com/grtjewellers/posts/{post_id}" if post_id else ""
        
        posts.append({
            "post_id": post_id,
            "date": date_str,
            "url": url,
            "caption": msg,
            "reactions": react_count,
            "comments": com_count,
            "shares": share_count,
            "actors": actor_names,
            "sponsors": sponsors,
        })

for s in scripts:
    try:
        data = json.loads(s)
        extract_stories(data)
    except Exception:
        pass

print(f"Extracted {len(posts)} raw post objects")
# Deduplicate
unique_posts = {}
for p in posts:
    pid = p.get("post_id") or p.get("url")
    if pid and pid not in unique_posts:
        unique_posts[pid] = p

print(f"Unique posts extracted: {len(unique_posts)}")
for i, (pid, p) in enumerate(list(unique_posts.items())[:10], 1):
    print(f"\n[{i}] Date: {p['date']} | Reactions: {p['reactions']} | Comments: {p['comments']} | Shares: {p['shares']}")
    print(f"    URL: {p['url']}")
    print(f"    Actors: {p['actors']}")
    print(f"    Text: {p['caption'][:120]}...")
