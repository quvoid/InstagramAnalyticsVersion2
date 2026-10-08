# -*- coding: utf-8 -*-
"""Re-transcribe the reference Ad Library videos (client sheet 1tK_Hi...) with Groq whisper-large-v3.
The sheet's own transcripts came from whisper-base and many are loops/garbage.
Input grt_adslibrarytranscript.csv (+ raw for Creative URL); output adlib_ref_transcripts.json by Library ID.
  GROQ_KEY_VAR=GROQ_API_KEY_2 SHARD=0/3 python adlib_ref_transcribe.py
"""
import csv, json, os, re, subprocess, time, sys
from curl_cffi import requests, CurlMime

HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(os.path.dirname(HERE))
KEY = re.search(r"^" + os.environ.get("GROQ_KEY_VAR", "GROQ_API_KEY") + r"=(.*)$", open(ROOT + "/.env", encoding="utf-8").read(), re.M).group(1).strip()
i, n = map(int, os.environ.get("SHARD", "0/1").split("/"))
OUT = HERE + f"/adlib_ref_transcripts_{i}.json"
AUD = HERE + f"/_adref_audio_{i}"; os.makedirs(AUD, exist_ok=True)
done = json.load(open(OUT, encoding="utf-8")) if os.path.exists(OUT) else {}
jobs = {}
for f, idc, urlc in [("grt_adslibrarytranscript.csv", "Library ID", "Video URL"), ("grt_adslibraryraw.csv", "Library ID", "Creative URL")]:
    for r in csv.DictReader(open(HERE + "/" + f, encoding="utf-8")):
        lid = re.sub(r"\D", "", r["Library link"].split("id=")[-1]) or r[idc]
        u = r.get(urlc) or ""
        if ".mp4" in u and lid not in jobs:
            jobs[lid] = u
todo = [(k, v) for j, (k, v) in enumerate(sorted(jobs.items())) if j % n == i and k not in done]
print(len(jobs), "videos total;", len(todo), "in this shard", flush=True)
for k, (lid, url) in enumerate(todo, 1):
    mp4, mp3 = f"{AUD}/{lid}.mp4", f"{AUD}/{lid}.mp3"
    try:
        r = requests.get(url, timeout=60, impersonate="chrome120")
        if r.status_code != 200:
            done[lid] = {"error": f"video HTTP {r.status_code} (link expired?)"}; continue
        open(mp4, "wb").write(r.content)
        subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", mp4, "-vn", "-ac", "1", "-ar", "16000", "-b:a", "48k", mp3], check=True)
        if not os.path.exists(mp3) or os.path.getsize(mp3) < 2000:
            done[lid] = {"text": "no audio track", "lang": ""}; continue
        for model in ["whisper-large-v3", "whisper-large-v3-turbo"] * 3:
            mp = CurlMime()
            mp.addpart(name="file", filename=lid + ".mp3", local_path=mp3, content_type="audio/mpeg")
            for kk, vv in {"model": model, "response_format": "verbose_json", "temperature": "0"}.items():
                mp.addpart(name=kk, data=vv)
            g = requests.post("https://api.groq.com/openai/v1/audio/transcriptions", timeout=120,
                              headers={"Authorization": "Bearer " + KEY}, multipart=mp)
            if g.status_code == 429:
                time.sleep(min(60, float(g.headers.get("retry-after", 15)) + 1)); continue
            break
        if g.status_code != 200:
            done[lid] = {"error": f"groq {g.status_code}"}; continue
        j = g.json()
        segs = [s for s in j.get("segments", []) if s.get("no_speech_prob", 0) < 0.6]
        txt = " ".join(s["text"].strip() for s in segs).strip()
        hook = " ".join(s["text"].strip() for s in segs if s.get("start", 0) < 3.5).strip()
        done[lid] = {"text": txt or "no speech - background music only", "hook3s": hook, "lang": j.get("language", ""),
                     "dur": j.get("duration"), "model": "groq " + model}
    except Exception as e:
        done[lid] = {"error": type(e).__name__ + ": " + str(e)[:120]}
    finally:
        for p in (mp4, mp3):
            if os.path.exists(p): os.remove(p)
        json.dump(done, open(OUT + ".tmp", "w", encoding="utf-8"), ensure_ascii=False, indent=0); os.replace(OUT + ".tmp", OUT)
    if k % 10 == 0:
        print(f"{k}/{len(todo)}", flush=True)
print("done", len(done), flush=True)
