"""
Search and extract more creator video posts on Facebook for GRT
"""

import sys, re, json, base64

with open("extract_creator_fb_videos.py", "r") as f:
    pass

from curl_cffi import requests
session = requests.Session(impersonate="chrome120")
headers = {
    "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
}

more_creators = [
    ("Niharika Konidela", "IamNiharikaKonidela"),
    ("Ritu Varma", "RituVarmaOfficial"),
    ("Faria Abdullah", "fariaabdullahofficial"),
    ("Siri Hanmanth", "SiriHanmanthOfficial"),
    ("Bhumika Basavaraj", "BhumikaBasavarajOfficial"),
    ("Gayathri Yuvraaj", "GayathriYuvraajOfficial"),
    ("Deepika Das", "DeepikaDasOfficial"),
    ("Janani Ashok Kumar", "JananiAshokKumarOfficial"),
    ("Teju Ashwini", "TejuAshwiniOfficial"),
    ("Divya Uruduga", "DivyaUrudugaOfficial"),
    ("Shanvi Srivastava", "ShanviSrivastavaOfficial"),
    ("Milana Nagaraj", "MilanaNagarajOfficial"),
    ("Prasanna", "ActorPrasanna"),
    ("Ranjani Raghavan", "RanjaniRaghavanOfficial"),
    ("Gouri Priya", "GouriPriyaOfficial"),
]

for name, fb_user in more_creators:
    for sub in ["videos/", "posts/"]:
        u = f"https://www.facebook.com/{fb_user}/{sub}"
        r = session.get(u, headers=headers, timeout=15)
        if r.status_code == 200:
            for m in re.finditer(r'\{"node":\{"__typename":"Video".*?\}\}', r.text):
                node_str = m.group(0)
                if "grt" in node_str.lower() or "jewel" in node_str.lower():
                    print(f"Found node for {name}: {node_str[:150]}")
