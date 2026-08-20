"""
scan_profile_metrics.py -- core profile-metrics logic, used by the API (main.py).

Given a username, resolves it to a numeric user_id (with the same
web_profile_info -> HTML-fallback strategy used everywhere else in this
project, since some business accounts 400 on the structured endpoint) and
pulls: follower/following/post counts, verified/business flags, and average
likes/comments/ER% over the account's most recent posts.

Note: Instagram's feed/user endpoint caps a single-page request at ~12 items
regardless of the requested `count` -- so "last N posts" here really means
"last ~12 posts" in practice. This is consistent across every account, not a
bug, just worth knowing when reading the numbers.
"""
import re
from datetime import datetime, timezone


class ProfileScanError(Exception):
    pass


def _parse_post(item: dict) -> dict:
    likes = item.get("like_count", 0) or 0
    comments = item.get("comment_count", 0) or 0
    return {"likes": likes, "comments": comments}


def scrape_profile(username: str, cookies: dict, base_headers: dict, posts_sample: int = 20) -> dict:
    """Blocking. Run in a threadpool from async code."""
    try:
        from curl_cffi import requests as cffi_requests
    except ImportError:
        raise ProfileScanError("curl_cffi is not installed (pip install curl_cffi)")

    username = username.strip().lstrip("@")
    s = cffi_requests.Session(impersonate="chrome120")

    hdrs = {**base_headers, "referer": f"https://www.instagram.com/{username}/"}
    profile = None
    user_id = None

    # Attempt 1: structured web_profile_info
    try:
        url = f"https://www.instagram.com/api/v1/users/web_profile_info/?username={username}"
        r = s.get(url, headers=hdrs, cookies=cookies, timeout=15)
        if r.status_code == 200:
            u = (r.json().get("data", {}) or {}).get("user")
            if u and u.get("id"):
                user_id = u["id"]
                profile = {
                    "username": username,
                    "full_name": u.get("full_name", "N/A"),
                    "followers": u.get("edge_followed_by", {}).get("count", 0),
                    "following": u.get("edge_follow", {}).get("count", 0),
                    "total_posts": u.get("edge_owner_to_timeline_media", {}).get("count", 0),
                    "verified": u.get("is_verified", False),
                    "is_business": u.get("is_business_account", False),
                    "user_id": user_id,
                }
    except Exception:
        pass

    # Attempt 2: HTML profilePage_<id> scrape + flat /users/{id}/info/ endpoint
    # (works around IG's 400 on web_profile_info for some business accounts)
    if not profile:
        try:
            r = s.get(f"https://www.instagram.com/{username}/", headers=hdrs, cookies=cookies, timeout=15)
            m = re.search(r'"profilePage_(\d+)"', r.text) if r.status_code == 200 else None
            uid = m.group(1) if m else None
            if uid:
                r2 = s.get(f"https://www.instagram.com/api/v1/users/{uid}/info/", headers=base_headers, cookies=cookies, timeout=15)
                if r2.status_code == 200:
                    uu = r2.json().get("user", {})
                    user_id = uid
                    profile = {
                        "username": uu.get("username", username),
                        "full_name": uu.get("full_name", "N/A"),
                        "followers": uu.get("follower_count", 0),
                        "following": uu.get("following_count", 0),
                        "total_posts": uu.get("media_count", 0),
                        "verified": uu.get("is_verified", False),
                        "is_business": uu.get("is_business", False),
                        "user_id": uid,
                    }
        except Exception:
            pass

    if not profile:
        raise ProfileScanError(f"could not resolve @{username} (private, doesn't exist, or blocked)")

    followers = profile["followers"]
    total_likes = total_comments = n = 0
    if user_id:
        try:
            purl = f"https://www.instagram.com/api/v1/feed/user/{user_id}/?count={posts_sample}"
            r3 = s.get(purl, headers=base_headers, cookies=cookies, timeout=15)
            if r3.status_code == 200:
                for item in r3.json().get("items", [])[:posts_sample]:
                    p = _parse_post(item)
                    total_likes += p["likes"]
                    total_comments += p["comments"]
                    n += 1
        except Exception:
            pass

    profile["posts_sampled"] = n
    profile["avg_likes"] = (total_likes // n) if n else 0
    profile["avg_comments"] = (total_comments // n) if n else 0
    profile["avg_er"] = round((total_likes + total_comments) / n / followers * 100, 4) if (n and followers) else 0
    return profile
