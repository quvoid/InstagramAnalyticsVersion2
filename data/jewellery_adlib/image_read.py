# -*- coding: utf-8 -*-
"""Read the image creatives (photos / carousels) from diwali2025_reels.json with a vision model.

For every non-video post: fetch the post page, take up to 4 frames (carousel slides), and ask
Groq's Llama-4 vision model for (1) every word printed on the creative, verbatim, and
(2) a one-line description of what is shown (product, model, setting, offer badge).
Results go to diwali_image_reads.json keyed by shortcode, saved after every post. Rerunnable.

  GROQ_KEY_VAR=GROQ_API_KEY_3 python data/jewellery_adlib/image_read.py
"""
import json, os, re, sys, time
from curl_cffi import requests

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
from core import ig_web

SRC, OUT = ROOT + "/diwali2025_reels.json", ROOT + "/diwali_image_reads.json"
KEY = re.search(r"^" + os.environ.get("GROQ_KEY_VAR", "GROQ_API_KEY_3") + r"=(.*)$",
                open(ROOT + "/.env", encoding="utf-8").read(), re.M).group(1).strip()
MODEL = os.environ.get("VISION_MODEL", "qwen/qwen3.8-27b")
PROMPT = ("This is an Instagram ad creative from an Indian jewellery brand. Reply in JSON with keys "
          "\"text\" (every word printed on the image, verbatim, keep the original language and add an English "
          "translation in brackets if it is not English; \"\" if none), \"offer\" (any price, discount or offer "
          "shown, else \"\"), and \"visual\" (one sentence: what is shown - product, model/celebrity, setting, festival cues).")


def load(p):
    try:
        return json.load(open(p, encoding="utf-8"))
    except Exception:
        return {}


def frames(code):
    it = ig_web.post_web_info(code) or {}
    items = it.get("carousel_media") or [it]
    urls = []
    for m in items[:4]:
        c = ((m.get("image_versions2") or {}).get("candidates") or [])
        if c:
            urls.append(sorted(c, key=lambda x: x.get("width", 0))[-1]["url"])
    return urls


def ask(urls):
    content = [{"type": "text", "text": PROMPT}] + [{"type": "image_url", "image_url": {"url": u}} for u in urls]
    for attempt in range(6):
        r = requests.post("https://api.groq.com/openai/v1/chat/completions", timeout=120,
                          headers={"Authorization": "Bearer " + KEY},
                          json={"model": MODEL, "messages": [{"role": "user", "content": content}],
                                "response_format": {"type": "json_object"}, "temperature": 0})
        if r.status_code == 429:
            time.sleep(float(r.headers.get("retry-after", 20)) + 1); continue
        if r.status_code != 200:
            return {"error": f"HTTP {r.status_code}: {r.text[:150]}"}
        try:
            return json.loads(r.json()["choices"][0]["message"]["content"])
        except Exception as e:
            return {"error": "bad json: " + str(e)}
    return {"error": "rate limited"}


if __name__ == "__main__":
    done = load(OUT)
    jobs = []
    for b, res in list(load(SRC).items()) + list(load(ROOT + "/diwali2025_grid.json").items()):
        for p in (res.get("grid") or {}).get("posts", []):
            if p.get("kind") != "Reel" and p["code"] not in done:
                jobs.append((b, p["code"]))
    print(len(jobs), "image creatives to read", flush=True)
    for i, (b, code) in enumerate(jobs, 1):
        try:
            urls = frames(code)
            res = ask(urls) if urls else {"error": "no image urls"}
        except Exception as e:
            print(code, "network error, will retry next pass:", type(e).__name__, flush=True)
            time.sleep(10); continue
        res["brand"], res["frames"] = b, len(urls)
        done[code] = res
        json.dump(done, open(OUT + ".tmp", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        os.replace(OUT + ".tmp", OUT)
        if i % 10 == 0:
            print(f"{i}/{len(jobs)}", flush=True)
        time.sleep(2)
    print("done", len(done), flush=True)
