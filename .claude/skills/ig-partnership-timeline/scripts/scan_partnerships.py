#!/usr/bin/env python3
"""
scan_partnerships.py -- Scan a brand's own Instagram feed for "partnership"
posts (organic creator collabs) and dump the results as JSON.

A post counts as a partnership signal if ANY of these are true (checked in
this priority order, since #1 is by far the dominant real-world signal --
Instagram's collab feature makes the creator's OWN post appear on the
brand's timeline with the creator, not the brand, as `item.user`):
  1. item['user']['username'] != target            (owner-collab)
  2. someone else is in item['coauthor_producers']  (coauthor-tag)
  3. item['is_paid_partnership'] is true
  4. caption matches #ad / #paidpartnership / #sponsored / #collab /
     #partnership -- in which case @mentions are pulled out as the
     likely partner (caption-hashtag, weakest signal, keep as fallback only)

Usage:
    python scan_partnerships.py <brand_username> [--max-posts 400] \
        [--project-dir .] [--out scratch_<brand>.json]

Requires a scrape_profiles.py in --project-dir (default: current directory)
that exposes a live COOKIES dict (sessionid, csrftoken, mid, ds_user_id) and
a BASE_HEADERS dict, matching the pattern already used across this project.
Install dependency: pip install curl_cffi
"""
import sys
import os
import re
import json
import time
import random
import argparse
from datetime import datetime, timezone


HASHTAG_PATTERN = re.compile(
    r'#ad\b|#paidpartnership|#sponsored|#collab|#partnership|paid partnership with',
    re.IGNORECASE,
)
MENTION_PATTERN = re.compile(r'@([A-Za-z0-9_.]+)')


def resolve_user_id(session, base_headers, cookies, username):
    """Resolve an IG username to a numeric user_id.

    web_profile_info is preferred but has a known bug on some business
    accounts (400 error mentioning ig_business_category_subvertical) --
    fall back to scraping the plain profile HTML and regexing
    "profilePage_(\\d+)" out of it, which works even when the REST
    endpoint doesn't.
    """
    hdrs = {**base_headers, "referer": f"https://www.instagram.com/{username}/"}
    try:
        url = f"https://www.instagram.com/api/v1/users/web_profile_info/?username={username}"
        r = session.get(url, headers=hdrs, cookies=cookies, timeout=15)
        if r.status_code == 200:
            user_data = r.json().get("data", {}).get("user") or {}
            if user_data.get("id"):
                return user_data["id"]
        else:
            print(f"  web_profile_info returned {r.status_code} "
                  f"(known IG-side bug on some business accounts) -- falling back to HTML scrape")
    except Exception as e:
        print(f"  web_profile_info error: {e} -- falling back to HTML scrape")

    try:
        r = session.get(f"https://www.instagram.com/{username}/", headers=hdrs, cookies=cookies, timeout=15)
        if r.status_code == 200:
            m = re.search(r'"profilePage_(\d+)"', r.text)
            if m:
                return m.group(1)
    except Exception as e:
        print(f"  HTML fallback error: {e}")
    return None


def extract_partnership_info(item, target):
    """Return a hit dict if this feed item looks like a partnership post, else None."""
    partners = {}

    owner = (item.get("user", {}) or {}).get("username", "").strip()
    if owner and owner.lower() != target.lower():
        partners[owner.lower()] = ("owner-collab", owner)

    for c in item.get("coauthor_producers") or []:
        u = (c.get("username") or "").strip()
        if u and u.lower() != target.lower():
            partners.setdefault(u.lower(), ("coauthor-tag", u))

    is_paid = bool(item.get("is_paid_partnership"))

    cap_obj = item.get("caption")
    caption = (cap_obj.get("text", "") if isinstance(cap_obj, dict) else "") or ""
    if HASHTAG_PATTERN.search(caption):
        for m in MENTION_PATTERN.findall(caption):
            if m.lower() != target.lower():
                partners.setdefault(m.lower(), ("caption-hashtag", m))

    if not partners and not is_paid:
        return None

    code = item.get("code", item.get("shortcode", ""))
    ts = item.get("taken_at", 0)
    date_str = (
        datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")
        if ts else "N/A"
    )
    return {
        "url": f"https://www.instagram.com/p/{code}/",
        "date": date_str,
        "partners": partners,
        "caption": caption,
    }


def scan(target, max_posts, project_dir):
    sys.path.insert(0, os.path.abspath(project_dir))
    try:
        from scrape_profiles import COOKIES, BASE_HEADERS
    except ImportError:
        print(f"ERROR: could not import COOKIES/BASE_HEADERS from scrape_profiles.py in {project_dir}")
        print("Make sure a scrape_profiles.py with a live COOKIES dict exists there.")
        sys.exit(1)

    try:
        from curl_cffi import requests as cffi_requests
    except ImportError:
        print("ERROR: pip install curl_cffi")
        sys.exit(1)

    s = cffi_requests.Session(impersonate="chrome120")

    print(f"Resolving user_id for @{target}...")
    user_id = resolve_user_id(s, BASE_HEADERS, COOKIES, target)
    if not user_id:
        print(f"ERROR: could not resolve user_id for @{target}")
        sys.exit(1)
    print(f"  user_id = {user_id}")

    unique = {}
    hits = []
    scanned = 0
    max_id = None

    while scanned < max_posts:
        url = f"https://www.instagram.com/api/v1/feed/user/{user_id}/?count=12"
        if max_id:
            url += f"&max_id={max_id}"
        r = s.get(url, headers=BASE_HEADERS, cookies=COOKIES, timeout=15)
        if r.status_code != 200:
            print(f"  stopped, status {r.status_code}: {r.text[:150]}")
            break
        j = r.json()
        items = j.get("items", [])
        if not items:
            break
        for item in items:
            scanned += 1
            info = extract_partnership_info(item, target)
            if info:
                hits.append(info)
                for uname_lower, (method, disp) in info["partners"].items():
                    e = unique.setdefault(uname_lower, {"display": disp, "count": 0})
                    e["count"] += 1
            if scanned >= max_posts:
                break
        print(f"  [{target}] scanned={scanned} hits={len(hits)}")
        max_id = j.get("next_max_id")
        if not j.get("more_available") or not max_id:
            break
        time.sleep(random.uniform(1.2, 2.2))

    # Flag a likely self-referential sister-brand pair: if one "partner"
    # dominates the hit list, it's often a sister/sub-brand account cross-
    # tagging the parent brand rather than a real third-party creator.
    # (e.g. jockeywomanindia <-> jockeyindia mutually tagging each other.)
    # This can't be fully automated -- surface it so a human/agent can judge.
    if unique and hits:
        top_partner, top_info = max(unique.items(), key=lambda kv: kv[1]["count"])
        share = top_info["count"] / len(hits)
        if share > 0.5:
            print(f"  NOTE: '{top_info['display']}' appears in {share:.0%} of hits -- check whether this "
                  f"is a sister/sub-brand account cross-tagging rather than a real third-party creator "
                  f"before using this list for outreach or metrics scraping.")

    return {"target": target, "user_id": user_id, "scanned": scanned, "hits": hits, "unique": unique}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("username", help="Brand's Instagram handle (no @)")
    ap.add_argument("--max-posts", type=int, default=400)
    ap.add_argument("--project-dir", default=".", help="Directory containing scrape_profiles.py (for COOKIES)")
    ap.add_argument("--out", default=None, help="Output JSON path (default: scratch_<username>.json)")
    args = ap.parse_args()

    target = args.username.strip().lstrip("@")
    out_path = args.out or f"scratch_{target}.json"

    result = scan(target, args.max_posts, args.project_dir)

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False)

    print(f"\nDone. scanned={result['scanned']} hits={len(result['hits'])} unique_partners={len(result['unique'])}")
    print(f"Saved: {out_path}")


if __name__ == "__main__":
    main()
