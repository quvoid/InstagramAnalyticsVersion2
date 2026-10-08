# -*- coding: utf-8 -*-
"""Print each brand's transcribed collab reels compactly, for per-reel analysis.
Usage: dump_reels.py <brand_handle> [start] [count]"""
import json, glob, sys, re
from datetime import datetime, timezone

ROOT = "C:/Users/omkar/OneDrive/Desktop/InstagramAnalytics"
cache = json.load(open(ROOT + "/brand_scan_cache.json", encoding="utf-8"))["2026-04-01_2026-10-07_own"]
T = {}
try:
    T.update({k: v for k, v in json.load(open(ROOT + "/jewellery_audio_transcripts.json", encoding="utf-8")).items()
              if v.get("model") == "large-v3-turbo"})
except Exception:
    pass
for f in sorted(glob.glob(ROOT + "/jewellery_audio_transcripts_groq*.json")):
    T.update(json.load(open(f, encoding="utf-8")))
sc = lambda u: ([x for x in str(u).split("/") if x] or [""])[-1]

b = sys.argv[1]; start = int(sys.argv[2]) if len(sys.argv) > 2 else 0; cnt = int(sys.argv[3]) if len(sys.argv) > 3 else 999
reels = [p for p in cache[b]["posts"] if p.get("kind") == "Reel"]
reels.sort(key=lambda p: -(p.get("taken_at") or 0))
print(f"{b}: {len(reels)} reels")
for p in reels[start:start + cnt]:
    c = sc(p["post_url"]); t = T.get(c, {})
    txt = (t.get("text") or "").replace("\n", " ")
    cap = re.sub(r"\s+", " ", p.get("caption") or "")
    print(f"\n[{c}] {datetime.fromtimestamp(p['taken_at'], tz=timezone.utc):%Y-%m-%d} @{p['creator']} "
          f"| {p.get('followers')} fol | {p.get('views')} views | paid={p.get('is_paid_partnership')} | {t.get('lang','')} {t.get('dur','')}s")
    print("  CAP:", cap[:260])
    print("  TR :", txt[:600] if txt else "(none)")
