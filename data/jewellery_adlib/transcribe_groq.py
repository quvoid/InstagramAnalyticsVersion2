# -*- coding: utf-8 -*-
"""Transcribe the jewellery collab reels with Groq's hosted whisper-large-v3.

Runs alongside the local CPU run: Groq works from the LARGEST brands down, the local
run from the smallest up, and each skips reels the other has already finished, so they
meet in the middle. Results go to their own file (no write races) and are merged later.

Free-tier limits (verified 2026-10-08): 20 requests/min, 7,200 audio-seconds/hour,
28,800 audio-seconds/day. The script paces itself under those and honours 429s.
"""
import json, os, re, subprocess, sys, threading, time
from collections import deque
from curl_cffi import requests, CurlMime

ROOT = "C:/Users/omkar/OneDrive/Desktop/InstagramAnalytics"
SP = "C:/Users/omkar/AppData/Local/Temp/claude/C--Users-omkar-OneDrive-Desktop-InstagramAnalytics/5179662b-3af2-43fe-969b-ca0e773e6ab6/scratchpad"
OUT = ROOT + "/" + os.environ.get("GROQ_OUT", "jewellery_audio_transcripts_groq.json")
LOCAL = ROOT + "/jewellery_audio_transcripts.json"
AUD = SP + "/groq_audio_" + os.environ.get("GROQ_OUT", "a").replace(".json", "")   # own folder per run
KEY = re.search(r"^" + os.environ.get("GROQ_KEY_VAR", "GROQ_API_KEY") + r"=(.*)$", open(ROOT + "/.env", encoding="utf-8").read(), re.M).group(1).strip()
URL = "https://api.groq.com/openai/v1/audio/transcriptions"
MODEL = "whisper-large-v3"
BATCH = 10
ASH_CAP, RPM_CAP = 7000, 18          # stay just under 7,200 s/hour and 20 req/min
MUSIC = "no speech - background music only"
os.makedirs(AUD, exist_ok=True)

BRANDS = ["malabargoldanddiamonds", "joyalukkas", "tanishqjewellery", "sencogoldanddiamonds", "josalukkas",
          "bluestone_jewellery", "kalyanjewellers_official", "thangamayiljewellery", "pc_jeweller", "caratlane"]
cache = json.load(open(ROOT + "/brand_scan_cache.json", encoding="utf-8"))["2026-04-01_2026-10-07_own"]
sc = lambda u: ([x for x in str(u).split("/") if x] or [""])[-1]
jobs, seen = [], set()
for b in BRANDS:
    for p in cache.get(b, {}).get("posts", []):
        if p.get("kind") == "Reel" and sc(p["post_url"]) not in seen:
            seen.add(sc(p["post_url"]))
            jobs.append({"brand": b, "creator": p["creator"], "url": p["post_url"], "date": p.get("taken_at"),
                         "code": sc(p["post_url"])})
if os.environ.get("JOBS_FROM"):                                   # e.g. diwali2025_reels.json
    jobs, seen = [], set()
    _src = {}
    for _fn in [f for f in os.environ["JOBS_FROM"].split(",") if os.path.exists(ROOT + "/" + f)]:
        for _b, _r in json.load(open(ROOT + "/" + _fn, encoding="utf-8")).items():
            _src.setdefault(_b, {}).update({k: v for k, v in _r.items() if v})
    for b, res in _src.items():
        for p in res.get("reels", []) + [g for g in (res.get("grid") or {}).get("posts", []) if g.get("kind") == "Reel"]:
            if p["code"] not in seen:
                seen.add(p["code"])
                jobs.append({"brand": b, "creator": p["creator"], "url": p["post_url"], "date": p.get("taken_at"), "code": p["code"]})
size = {}
for j in jobs:
    size[j["brand"]] = size.get(j["brand"], 0) + 1
jobs.sort(key=lambda j: (-size[j["brand"]], j["brand"]))          # largest brand first
if os.environ.get("GROQ_ORDER") == "asc":                         # second instance works from the other end
    jobs.reverse()
elif os.environ.get("GROQ_ORDER") == "mid":                       # third instance starts in the middle
    jobs = jobs[len(jobs) // 2:] + jobs[:len(jobs) // 2]
print("groq order:", " > ".join(f"{b}({n})" for b, n in sorted(size.items(), key=lambda kv: -kv[1])), flush=True)

def load(p):
    try:
        return json.load(open(p, encoding="utf-8"))
    except Exception:
        return {}

done = load(OUT)
win_audio, win_req = deque(), deque()     # (t, seconds) for the hourly audio cap; t for the per-minute cap

def pace(dur):
    while True:
        now = time.time()
        while win_audio and now - win_audio[0][0] > 3600: win_audio.popleft()
        while win_req and now - win_req[0] > 60: win_req.popleft()
        used = sum(s for _, s in win_audio)
        if used + max(dur, 10) <= ASH_CAP and len(win_req) < RPM_CAP:
            return
        wait = 5 if len(win_req) >= RPM_CAP else max(5, 3600 - (now - win_audio[0][0]))
        print(f"   pacing: {used:.0f}s audio this hour, {len(win_req)} req this minute - waiting {wait:.0f}s", flush=True)
        time.sleep(min(wait, 120))

def dur_of(path):
    try:
        return float(subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", path],
                                    capture_output=True, text=True, timeout=30).stdout.strip())
    except Exception:
        return 40.0

def fetch(batch):
    need = [j for j in batch if not os.path.exists(f"{AUD}/{j['code']}.mp3")]
    if need:
        try:
            subprocess.run(["yt-dlp", "-q", "--no-warnings", "--ignore-errors", "-N", "4", "-x", "--audio-format", "mp3",
                            "-o", AUD + "/%(id)s.%(ext)s"] + [j["url"] for j in need], capture_output=True, timeout=900)
        except Exception as e:
            print("   download batch failed:", type(e).__name__, flush=True)

# Each Groq model has its OWN free quota (7,200 audio-s/hour each), so when large-v3 is
# exhausted, turbo keeps going instead of everything waiting. Waits come from Groq's own
# retry-after header, not a local estimate, so the quota is used right up to the limit.
MODELS = ["whisper-large-v3", "whisper-large-v3-turbo"]
blocked_until = {m: 0.0 for m in MODELS}

def transcribe(path, dur):
    for attempt in range(40):
        now = time.time()
        while win_req and now - win_req[0] > 60: win_req.popleft()
        if len(win_req) >= RPM_CAP:
            time.sleep(3); continue
        ready = [m for m in MODELS if blocked_until[m] <= now]
        if not ready:
            wait = min(blocked_until.values()) - now
            print(f"   both models at their limit - waiting {wait:.0f}s", flush=True)
            time.sleep(min(max(wait, 5), 300)); continue
        model = ready[0]
        if not os.path.exists(path):
            return {"error": "audio file vanished before upload"}
        mp = CurlMime()
        mp.addpart(name="file", filename=os.path.basename(path), local_path=path, content_type="audio/mpeg")
        for k, v in {"model": model, "response_format": "verbose_json", "temperature": "0"}.items():
            mp.addpart(name=k, data=v)
        try:
            r = requests.post(URL, headers={"Authorization": "Bearer " + KEY}, multipart=mp, timeout=180)
        except Exception:
            time.sleep(10); continue
        win_req.append(time.time())
        if r.status_code == 200:
            j = r.json(); j["_model"] = model
            return j
        if r.status_code == 429:
            ra = float(r.headers.get("retry-after") or 60)
            blocked_until[model] = time.time() + ra + 1
            print(f"   {model} at its limit - blocked {ra:.0f}s, switching", flush=True)
            continue
        return {"error": f"HTTP {r.status_code}: {r.text[:120]}"}
    return {"error": "gave up after retries"}


def clean(j):
    """Drop segments Whisper itself marks as probably-not-speech: on music it invents
    'Thank you.' and lyrics; no_speech_prob / avg_logprob flag those."""
    segs = j.get("segments") or []
    keep = [s["text"].strip() for s in segs
            if not (s.get("no_speech_prob", 0) > 0.6 and s.get("avg_logprob", 0) < -0.8)]
    txt = " ".join(t for t in keep if t).strip() if segs else (j.get("text") or "").strip()
    if not txt or re.fullmatch(r"(thank you\.?|thanks for watching!?|music|\W*)", txt, re.I):
        return MUSIC
    return txt

# Drop finished reels BEFORE batching, so nothing is downloaded twice.
_local = load(LOCAL)
REDO_LOCAL = os.environ.get("REDO_LOCAL") == "1"     # second pass: replace the weaker local-CPU transcripts
jobs = [j for j in jobs if j["code"] not in done
        and (REDO_LOCAL or not (j["code"] in _local and _local[j["code"]].get("model") == "large-v3-turbo"))]
print(f"groq: {len(jobs)} reels still to do", flush=True)
nxt = threading.Thread(target=fetch, args=(jobs[:BATCH],), daemon=True); nxt.start()
t0, n = time.time(), 0
for b in range(0, len(jobs), BATCH):
    nxt.join()
    nxt = threading.Thread(target=fetch, args=(jobs[b + BATCH:b + 2 * BATCH],), daemon=True); nxt.start()
    for j in jobs[b:b + BATCH]:
        local = load(LOCAL)
        others = {}
        for f in __import__("glob").glob(ROOT + "/jewellery_audio_transcripts_groq*.json"):
            if os.path.abspath(f) != os.path.abspath(OUT):
                others.update(load(f))
        if j["code"] in others:
            continue
        if j["code"] in done or (not REDO_LOCAL and j["code"] in local and local[j["code"]].get("model") == "large-v3-turbo"):
            continue                                  # finished here, or by the local run
        n += 1
        mp3 = f"{AUD}/{j['code']}.mp3"
        rec = dict(j, model="groq " + MODEL)
        if not os.path.exists(mp3):
            rec.update(text="audio not retrievable - reel private, deleted or region-locked", lang="", dur=0)
        else:
            d = dur_of(mp3)
            res = transcribe(mp3, d)
            if res.get("error"):
                rec.update(text="transcription failed - " + res["error"], lang="", dur=round(d, 1))
            else:
                rec.update(text=clean(res), lang=res.get("language", ""), dur=round(res.get("duration") or d, 1),
                           model="groq " + res.get("_model", MODEL))
            try: os.remove(mp3)
            except OSError: pass
        done[j["code"]] = rec
        if n % 5 == 0:
            json.dump(done, open(OUT + ".tmp", "w", encoding="utf-8"), ensure_ascii=False); os.replace(OUT + ".tmp", OUT)
            el = time.time() - t0
            print(f"  groq {n} done | total cached {len(done)} | {el:.0f}s ({el / n:.1f}s per reel)", flush=True)
json.dump(done, open(OUT + ".tmp", "w", encoding="utf-8"), ensure_ascii=False); os.replace(OUT + ".tmp", OUT)
print("groq finished", len(done), flush=True)
