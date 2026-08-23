"""
Scan individual creator Facebook pages for their GRT Jewellers paid partnership posts
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

# List of top creators who partner with GRT Jewellers
CREATORS_FB = [
    ("R. Ashwin", "AshwinRavi99"),
    ("Prithi Ashwin", "PrithiNarayananOfficial"),
    ("Athulya Ravi", "AthulyaRaviOfficial"),
    ("Ritu Varma", "RituVarmaOfficial"),
    ("Niharika Konidela", "IamNiharikaKonidela"),
    ("Faria Abdullah", "fariaabdullahofficial"),
    ("Siri Hanmanth", "SiriHanmanthOfficial"),
    ("Roshni Haripriyan", "RoshniHaripriyanOfficial"),
    ("Megha Shetty", "Meghashettyofficial"),
    ("Bhumika Basavaraj", "BhumikaBasavarajOfficial"),
    ("Tamil Rithika", "TamilRithikaOfficial"),
    ("Chaitra Reddy", "ChaitraReddyOfficial"),
    ("Gayathri Yuvraaj", "GayathriYuvraajOfficial"),
    ("Deepika Das", "DeepikaDasOfficial"),
    ("Anjana Rangan", "AnjanaRanganOfficial"),
    ("Shanvi Srivastava", "ShanviSrivastavaOfficial"),
    ("Milana Nagaraj", "MilanaNagarajOfficial"),
    ("Janani Ashok Kumar", "JananiAshokKumarOfficial"),
    ("Teju Ashwini", "TejuAshwiniOfficial"),
    ("Vithika Sheru", "VithikaSheruOfficial"),
    ("Prasanna", "ActorPrasanna"),
    ("Ranjani Raghavan", "RanjaniRaghavanOfficial"),
    ("Gouri Priya", "GouriPriyaOfficial"),
    ("Divya Uruduga", "DivyaUrudugaOfficial"),
    ("Simran Choudhary", "SimranChoudharyOfficial"),
    ("Varshini Sounderajan", "VarshiniSounderajanOfficial"),
    ("Poornima Ravi", "PoornimaRaviOfficial"),
    ("Varsha Dsouza", "VarshaDsouzaOfficial"),
    ("Pujita Ponnada", "PujitaPonnadaOfficial"),
    ("Delna Davis", "DelnaDavisOfficial"),
    ("Ramya Subramanian", "RamyaSubramanianOfficial"),
]

found_creator_posts = []

def extract_posts_from_html(html_text, creator_name, fb_username):
    scripts = re.findall(r'<script type="application/json"[^>]*>(.*?)</script>', html_text, re.DOTALL)
    
    def scan_node(obj):
        if isinstance(obj, dict):
            post_id = obj.get("post_id") or obj.get("id") or obj.get("video_id")
            ts = obj.get("creation_time") or obj.get("publish_time") or 0
            msg_obj = obj.get("message") or obj.get("savable_description") or obj.get("name")
            
            if not ts and "comet_sections" in obj:
                try:
                    ts = obj["comet_sections"]["context_layout"]["story"]["comet_sections"]["metadata"][0]["story"]["creation_time"]
                except Exception:
                    pass
            if not msg_obj and "comet_sections" in obj:
                try:
                    msg_obj = obj["comet_sections"]["content"]["story"]["message"]
                except Exception:
                    pass
                    
            msg = ""
            if isinstance(msg_obj, dict):
                msg = msg_obj.get("text", "")
            elif isinstance(msg_obj, str):
                msg = msg_obj
                
            feedback = obj.get("feedback") or {}
            reacts = 0
            if isinstance(feedback, dict):
                r_c = feedback.get("reaction_count")
                reacts = r_c.get("count", 0) if isinstance(r_c, dict) else (r_c or 0)
                
            url = obj.get("url") or obj.get("permalink_url") or ""
            
            # Check if this creator post mentions GRT or jewellery
            msg_lower = msg.lower()
            if any(k in msg_lower for k in ["grt", "grtjewellers", "grt jewellers", "grtdiamonds", "grt silver"]):
                dt = datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d") if ts else "N/A"
                if not url and post_id:
                    url = f"https://www.facebook.com/{fb_username}/posts/{post_id}"
                
                found_creator_posts.append({
                    "creator": creator_name,
                    "fb_username": fb_username,
                    "post_id": str(post_id),
                    "date": dt,
                    "url": url,
                    "reactions": reacts,
                    "caption": msg
                })
            
            for k, v in obj.items():
                scan_node(v)
        elif isinstance(obj, list):
            for item in obj:
                scan_node(item)

    for s in scripts:
        try:
            d = json.loads(s)
            scan_node(d)
        except Exception:
            pass

print("Scanning Creator Facebook Pages for GRT Jewellers Posts...\n")

for name, fb_user in CREATORS_FB:
    page_url = f"https://www.facebook.com/{fb_user}/"
    print(f"Checking {name:<22} ({page_url}) ...", end=" ", flush=True)
    try:
        r = session.get(page_url, headers=headers, timeout=15)
        if r.status_code == 200:
            before_len = len(found_creator_posts)
            extract_posts_from_html(r.text, name, fb_user)
            after_len = len(found_creator_posts)
            found_count = after_len - before_len
            if found_count > 0:
                print(f"✓ Found {found_count} GRT partnership post(s)!")
            else:
                print("✓ Page checked (No GRT post on initial feed)")
        else:
            print(f"⚠ HTTP {r.status_code}")
    except Exception as e:
        print(f"⚠ Error: {e}")
    time.sleep(1.0)

print(f"\n{'='*75}")
print(f"Total Creator-Owned Facebook Posts for GRT: {len(found_creator_posts)}")
print(f"{'='*75}\n")

with open("grt_creator_facebook_posts.json", "w", encoding="utf-8") as f:
    json.dump(found_creator_posts, f, ensure_ascii=False, indent=2)
