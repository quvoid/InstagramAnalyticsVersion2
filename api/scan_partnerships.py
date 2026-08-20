"""
scan_partnerships.py -- core partnership-scan logic, used by the API (main.py).

This is a self-contained copy of the same logic bundled in the
ig-partnership-timeline Claude skill (scripts/scan_partnerships.py), adapted
to be imported as a function rather than run from the CLI. Keep the two in
sync if the detection logic changes.

A post counts as a partnership signal if ANY of these are true (checked in
this priority order -- #1 is the dominant real-world signal, since
Instagram's collab feature makes the creator's OWN post appear on the
brand's timeline with the creator, not the brand, as `item.user`):
  1. item['user']['username'] != target            (owner-collab)
  2. someone else is in item['coauthor_producers']  (coauthor-tag)
  3. item['is_paid_partnership'] is true
  4. caption matches #ad / #paidpartnership / #sponsored / #collab /
     #partnership -- @mentions pulled out as the likely partner
     (caption-hashtag, weakest signal, kept as fallback only)
"""
import re
import time
import random
from datetime import datetime, timezone

HASHTAG_PATTERN = re.compile(
    r'#ad\b|#paidpartnership|#sponsored|#collab|#partnership|paid partnership with',
    re.IGNORECASE,
)
MENTION_PATTERN = re.compile(r'@([A-Za-z0-9_.]+)')


class ScanError(Exception):
    pass


def resolve_user_id(session, base_headers, cookies, username):
    """Resolve an IG username to a numeric user_id.

    web_profile_info is preferred but has a known bug on some business
    accounts (400 error mentioning ig_business_category_subvertical) --
    fall back to scraping the plain profile HTML and regexing
    "profilePage_(\\d+)" out of it.
    """
    hdrs = {**base_headers, "referer": f"https://www.instagram.com/{username}/"}
    try:
        url = f"https://www.instagram.com/api/v1/users/web_profile_info/?username={username}"
        r = session.get(url, headers=hdrs, cookies=cookies, timeout=15)
        if r.status_code == 200:
            user_data = r.json().get("data", {}).get("user") or {}
            if user_data.get("id"):
                return user_data["id"]
    except Exception:
        pass

    try:
        r = session.get(f"https://www.instagram.com/{username}/", headers=hdrs, cookies=cookies, timeout=15)
        if r.status_code == 200:
            m = re.search(r'"profilePage_(\d+)"', r.text)
            if m:
                return m.group(1)
    except Exception:
        pass
    return None


def extract_partnership_info(item, target):
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


def run_scan(target, cookies, base_headers, max_posts=400, progress_cb=None):
    """Blocking. Run in a threadpool from async code -- it uses time.sleep()
    between paginated requests to stay within Instagram's rate limits, same
    as every other scraper in this project."""
    try:
        from curl_cffi import requests as cffi_requests
    except ImportError:
        raise ScanError("curl_cffi is not installed (pip install curl_cffi)")

    target = target.strip().lstrip("@")
    s = cffi_requests.Session(impersonate="chrome120")

    user_id = resolve_user_id(s, base_headers, cookies, target)
    if not user_id:
        raise ScanError(f"could not resolve user_id for @{target} (private, doesn't exist, or blocked)")

    unique = {}
    hits = []
    scanned = 0
    max_id = None

    while scanned < max_posts:
        url = f"https://www.instagram.com/api/v1/feed/user/{user_id}/?count=12"
        if max_id:
            url += f"&max_id={max_id}"
        r = s.get(url, headers=base_headers, cookies=cookies, timeout=15)
        if r.status_code != 200:
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
        if progress_cb:
            progress_cb(scanned, len(hits))
        max_id = j.get("next_max_id")
        if not j.get("more_available") or not max_id:
            break
        time.sleep(random.uniform(1.2, 2.2))

    dominant_partner_warning = None
    if unique and hits:
        top_partner, top_info = max(unique.items(), key=lambda kv: kv[1]["count"])
        share = top_info["count"] / len(hits)
        if share > 0.5:
            dominant_partner_warning = (
                f"'{top_info['display']}' appears in {share:.0%} of hits -- likely a sister/sub-brand "
                f"account cross-tagging rather than a genuine third-party creator. Verify before using it."
            )

    return {
        "target": target,
        "user_id": user_id,
        "scanned": scanned,
        "hits": hits,
        "unique": unique,
        "warning": dominant_partner_warning,
    }
