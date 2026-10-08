# -*- coding: utf-8 -*-
"""Audio -> text for the jewellery brands' collab reels (brand's own page, reels only).

Batched yt-dlp downloads, a separate whisper worker with a wall-clock cap per clip,
results cached per reel so the run resumes. Usage: transcribe_jewellery.py <model>
"""
import json, os, re, subprocess, sys, threading, time

ROOT = "C:/Users/omkar/OneDrive/Desktop/InstagramAnalytics"
SP = "C:/Users/omkar/AppData/Local/Temp/claude/C--Users-omkar-OneDrive-Desktop-InstagramAnalytics/5179662b-3af2-43fe-969b-ca0e773e6ab6/scratchpad"
OUT = ROOT + "/jewellery_audio_transcripts.json"
AUD = SP + "/jaudio"
WORKER = SP + "/whisper_worker_hq.py"
MODEL = sys.argv[1] if len(sys.argv) > 1 else "large-v3-turbo"
BATCH = 10
MUSIC = "no speech - background music only"
os.makedirs(AUD, exist_ok=True)

TARGETS = [("2026-04-01_2026-10-07_own", ["malabargoldanddiamonds", "joyalukkas", "tanishqjewellery",
                                           "sencogoldanddiamonds", "josalukkas", "bluestone_jewellery",
                                           "kalyanjewellers_official", "thangamayiljewellery", "pc_jeweller", "caratlane"])]
cache = json.load(open(ROOT + "/brand_scan_cache.json", encoding="utf-8"))
sc = lambda u: ([x for x in str(u).split("/") if x] or [""])[-1]
jobs = []
for key, brands in TARGETS:
    for b in brands:
        res = cache.get(key, {}).get(b)
        if not res:
            print("NOT SCANNED:", b, flush=True); continue
        for p in res["posts"]:
            if p.get("kind") == "Reel":
                jobs.append({"brand": b, "creator": p["creator"], "url": p["post_url"],
                             "date": p.get("taken_at"), "code": sc(p["post_url"])})
seen, uniq = set(), []
for j in jobs:
    if j["code"] not in seen:
        seen.add(j["code"]); uniq.append(j)
jobs = uniq

done = json.load(open(OUT, encoding="utf-8")) if os.path.exists(OUT) else {}
todo = [j for j in jobs if j["code"] not in done or done[j["code"]].get("model") != MODEL]
print(f"reels {len(jobs)} | to transcribe {len(todo)} | model {MODEL}", flush=True)

HINT_WORDS = ["tanishq", "malabar", "joyalukkas", "senco", "jos alukkas", "bluestone", "kalyan",
              "thangamayil", "khazana", "pc jeweller", "caratlane"]

def echoed_hint(t):
    """Whisper can parrot the spelling hint on music-only audio; 4+ brand names in one
    short clip is that echo, not speech."""
    low = t.lower()
    return sum(1 for w in HINT_WORDS if w in low) >= 4

W = {"p": None}
def spawn():
    if W["p"]:
        try: W["p"].kill()
        except Exception: pass
    W["p"] = subprocess.Popen(["python", "-u", WORKER, MODEL], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                              stderr=subprocess.DEVNULL, text=True, encoding="utf-8")
    first = W["p"].stdout.readline()
    if not first.strip():
        raise SystemExit("worker failed to start (model load) - likely out of memory")
    print("   worker up", flush=True)

def run_clip(path, cap):
    res = {}
    def go():
        try:
            W["p"].stdin.write(path + "\n"); W["p"].stdin.flush()
            line = W["p"].stdout.readline()
            res.update(json.loads(line) if line.strip() else {"error": "no output"})
        except Exception as e:
            res.update({"error": type(e).__name__})
    t = threading.Thread(target=go, daemon=True); t.start(); t.join(cap)
    return None if t.is_alive() else res

def dur_of(path):
    try:
        return float(subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", path],
                                    capture_output=True, text=True, timeout=30).stdout.strip())
    except Exception:
        return 60.0

spawn()
t0, n = time.time(), 0
for b in range(0, len(todo), BATCH):
    chunk = [j for j in todo[b:b + BATCH] if not os.path.exists(f"{AUD}/{j['code']}.mp3")]
    if chunk:
        try:
            subprocess.run(["yt-dlp", "-q", "--no-warnings", "--ignore-errors", "-N", "4", "-x", "--audio-format", "mp3",
                            "-o", AUD + "/%(id)s.%(ext)s"] + [j["url"] for j in chunk],
                           capture_output=True, text=True, timeout=900)
        except Exception as e:
            print("   download batch failed:", type(e).__name__, flush=True)
    for j in todo[b:b + BATCH]:
        n += 1
        mp3 = f"{AUD}/{j['code']}.mp3"
        rec = dict(j, model=MODEL)
        if not os.path.exists(mp3):
            rec.update(text="audio not retrievable - reel private, deleted or region-locked", lang="", dur=0)
        else:
            d = dur_of(mp3)
            r = run_clip(mp3, max(240, d * 10))
            if r is None:
                rec.update(text="transcription abandoned - decoder ran past its time limit", lang="", dur=round(d, 1))
                spawn()
            elif r.get("error"):
                rec.update(text="transcription failed - " + r["error"], lang="", dur=round(d, 1))
            else:
                txt = r["text"]
                if not txt or echoed_hint(txt):
                    txt = MUSIC
                rec.update(text=txt, lang=r["lang"], lang_prob=r.get("lang_prob"), dur=r["dur"])
            try: os.remove(mp3)
            except OSError: pass
        done[j["code"]] = rec
        if n % 5 == 0 or n == len(todo):
            json.dump(done, open(OUT + ".tmp", "w", encoding="utf-8"), ensure_ascii=False)
            os.replace(OUT + ".tmp", OUT)
            el = time.time() - t0
            spoke = sum(1 for v in done.values() if not v["text"].startswith(("no speech", "audio not", "transcription")))
            print(f"  {n}/{len(todo)} | with speech {spoke} | {el:.0f}s ({el / n:.1f}s per reel)", flush=True)

json.dump(done, open(OUT + ".tmp", "w", encoding="utf-8"), ensure_ascii=False); os.replace(OUT + ".tmp", OUT)
W["p"].kill()
print("finished", len(done), flush=True)
