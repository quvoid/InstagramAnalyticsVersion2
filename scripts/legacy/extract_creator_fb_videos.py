"""
Extract Facebook Video URLs from Creator Facebook Pages
"""

import sys, re, json
from datetime import datetime, timezone
from curl_cffi import requests

sys.stdout.reconfigure(encoding="utf-8")

session = requests.Session(impersonate="chrome120")
headers = {
    "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "accept-language": "en-US,en;q=0.9",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
}

creators = [
    ("Athulya Ravi", "AthulyaRaviOfficial"),
    ("R. Ashwin", "AshwinRavi99"),
    ("Roshni Haripriyan", "RoshniHaripriyanOfficial"),
    ("Prithi Ashwin", "PrithiNarayananOfficial"),
    ("Ritu Varma", "RituVarmaOfficial"),
    ("Niharika Konidela", "IamNiharikaKonidela"),
    ("Faria Abdullah", "fariaabdullahofficial"),
    ("Siri Hanmanth", "SiriHanmanthOfficial"),
    ("Megha Shetty", "Meghashettyofficial"),
    ("Tamil Rithika", "TamilRithikaOfficial"),
    ("Chaitra Reddy", "ChaitraReddyOfficial"),
    ("Gayathri Yuvraaj", "GayathriYuvraajOfficial"),
    ("Deepika Das", "DeepikaDasOfficial"),
    ("Janani Ashok Kumar", "JananiAshokKumarOfficial"),
]

results = []

for name, fb_user in creators:
    url = f"https://www.facebook.com/{fb_user}/videos/"
    r = session.get(url, headers=headers, timeout=15)
    if r.status_code == 200:
        html = r.text
        scripts = re.findall(r'<script type="application/json"[^>]*>(.*?)</script>', html, re.DOTALL)
        
        def scan(obj):
            if isinstance(obj, dict):
                post_id = obj.get("post_id") or obj.get("id") or obj.get("video_id")
                ts = obj.get("creation_time") or obj.get("publish_time") or 0
                msg_obj = obj.get("message") or obj.get("savable_description") or obj.get("name")
                msg = ""
                if isinstance(msg_obj, dict):
                    msg = msg_obj.get("text", "")
                elif isinstance(msg_obj, str):
                    msg = msg_obj
                
                feedback = obj.get("feedback") or {}
                reacts = (feedback.get("reaction_count") or {}).get("count", 0) if isinstance(feedback.get("reaction_count"), dict) else (feedback.get("reaction_count") or 0)
                views = obj.get("video_view_count") or 0
                v_url = obj.get("url") or (f"https://www.facebook.com/{fb_user}/videos/{post_id}/" if post_id else "")
                
                if (msg or views or reacts or ts) and (post_id or v_url):
                    msg_l = msg.lower()
                    if any(k in msg_l for k in ["grt", "jewel", "gold", "silver", "diamond", "platinum", "necklace"]):
                        dt = datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d") if ts else "N/A"
                        results.append({
                            "creator": name,
                            "fb_username": fb_user,
                            "date": dt,
                            "url": v_url,
                            "views": views,
                            "reactions": reacts,
                            "caption": msg[:150].replace("\n", " ")
                        })
                
                for k, v in obj.items():
                    scan(v)
            elif isinstance(obj, list):
                for item in obj:
                    scan(item)
                    
        for s in scripts:
            try:
                data = json.loads(s)
                scan(data)
            except Exception:
                pass

print(f"Total Creator Videos Found: {len(results)}")
for i, r in enumerate(results, 1):
    print(f"[{i:>2}] {r['creator']:<20} | Date: {r['date']} | Views: {r['views']:>8,} | URL: {r['url']}")
    if r['caption']:
        print(f"     Text: {r['caption'][:120]}...")
