"""
scan_page.py -- Facebook Page metrics + recent content, the counterpart to
../api/scan_profile_metrics.py on the Instagram side.

Everything here works with NO cookies (verified 2026-08-22 on zivame, nike,
cocacola, MamaEarth.India). Cookies are accepted and passed through if you
have them configured, but they are not required.

What Facebook actually gives up anonymously, and what it doesn't:

  page id            yes -- "userID" in the page's Relay payload
  ad-library page id yes -- "delegate_page".id, which is the id the Ad Library
                     uses as view_all_page_id (verified: Zivame's page HTML
                     gives 234603919914240, and that is exactly the page_id
                     the Ad Library returns for Zivame's ads)
  likes              yes, exact, from the og:description meta tag
  talking about      yes, exact, same place
  followers          ROUNDED ONLY -- Facebook renders "794K followers" and
                     never puts the exact number in the anonymous payload.
                     We return both the raw string and a parsed integer, and
                     flag it as approximate. Don't report it as exact.
  category           yes
  recent posts       ONE post from the timeline payload. That's not a bug in
                     this parser -- Facebook server-renders a single story and
                     lazy-loads the rest.
  recent videos      ~6 per request from the /videos/ tab, WITH play counts and
                     reaction counts, plus a cursor for more. This is the best
                     anonymous engagement sample available, so /posts leans on
                     it for video-led brand pages.

Note on "followers vs likes": these are different numbers on Facebook (a Page
can be liked without being followed). The Instagram side has one follower
count; here you get both, so don't compare them across the two APIs blindly.
"""
import re
import json
import html as html_lib
from datetime import datetime, timezone

from fb_session import (
    FB_BASE,
    PageScanError,
    balanced_json,
    find_first,
    get_html,
    new_session,
)

NOT_AVAILABLE_MARKERS = (
    "content isn&#039;t available",
    "content isn't available",
    "This content isn",
)


def _iso(ts):
    if not ts:
        return None
    return datetime.fromtimestamp(int(ts), tz=timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _parse_compact_number(text):
    """'794K' -> 794000, '1.2M' -> 1200000, '39,561,807' -> 39561807."""
    if not text:
        return None
    text = text.strip().replace(",", "")
    m = re.match(r"^([\d.]+)\s*([KMB]?)$", text, re.IGNORECASE)
    if not m:
        return None
    value = float(m.group(1))
    suffix = m.group(2).upper()
    multiplier = {"": 1, "K": 1_000, "M": 1_000_000, "B": 1_000_000_000}[suffix]
    return int(value * multiplier)


def _clean(value):
    return html_lib.unescape(value).strip() if isinstance(value, str) else value


def scrape_page(handle: str, cookies: dict = None, session=None,
                check_verified: bool = False) -> dict:
    """Blocking. Run in a threadpool from async code.

    Deliberately fetches LOGGED OUT even when a session cookie is configured.
    Facebook serves two different renders: the logged-out one carries the
    `og:` meta tags (exact likes, talking-about, name, bio), and the logged-in
    Comet render does not -- sending cookies here silently nulls half the
    response. Verified 2026-08-22 on /IamTridhaOfficial: anonymous gives
    "3,259,343 likes . 28,713 talking about this" and the bio; the same URL
    with cookies gives neither.

    `check_verified` is the one thing worth a second, cookie'd request: the
    blue-badge state only appears in the logged-in render.
    """
    handle = handle.strip().lstrip("@").rstrip("/")
    if handle.startswith("http"):
        handle = handle.rstrip("/").rsplit("/", 1)[-1]
    s = session or new_session()

    _, page_html = get_html(s, f"{FB_BASE}/{handle}/")

    if any(marker in page_html for marker in NOT_AVAILABLE_MARKERS):
        raise PageScanError(
            f"Facebook has no public page at /{handle} ('this content isn't available'). "
            f"Check the vanity handle -- brand pages often differ from the Instagram "
            f"username (e.g. Mamaearth is /MamaEarth.India, not /mamaearth.in)."
        )

    page_id = find_first(r'"userID":"(\d+)"', page_html)
    if not page_id:
        raise PageScanError(f"could not resolve a page id for /{handle} (page may be gated or deleted)")

    og_desc = _clean(find_first(r'<meta property="og:description" content="([^"]*)"', page_html, default=""))
    og_title = _clean(find_first(r'<meta property="og:title" content="([^"]*)"', page_html, default=""))

    likes = find_first(r"([\d,]+)\s+likes", og_desc)
    talking_about = find_first(r"([\d,]+)\s+talking about this", og_desc)
    were_here = find_first(r"([\d,]+)\s+were here", og_desc)

    followers_display = find_first(r'"text":"([\d.,]+[KMB]?) followers"', page_html)

    # og:description is "<Name>. <n> likes . <n> talking about this. <bio...>" --
    # everything after the last stats clause is the page's own blurb.
    bio = None
    if og_desc:
        parts = re.split(r"talking about this\.?|were here\.?|likes\.?", og_desc)
        if parts:
            tail = parts[-1].strip(" .·")
            bio = tail or None

    # Categories come out of raw JSON, so they still carry escapes
    # ("Baby goods\/kids goods") -- decode them as the JSON strings they are.
    categories = []
    for raw in dict.fromkeys(re.findall(r'"category_name":"([^"]+)"', page_html)):
        try:
            categories.append(json.loads(f'"{raw}"'))
        except json.JSONDecodeError:
            categories.append(raw.replace("\\/", "/"))

    # Optional second pass, logged in: the verified badge and, on pages whose
    # logged-out render is thin, a name to fall back on.
    verified = None
    display_name = og_title or None
    if check_verified and cookies:
        try:
            _, member_html = get_html(s, f"{FB_BASE}/{handle}/", cookies=cookies)
            verified = 'title="Verified account"' in member_html
            if not display_name:
                display_name = find_first(r'"owning_profile":\{"__typename":"User","name":"([^"]+)"',
                                          member_html)
        except Exception:
            pass  # a stale session shouldn't cost you the rest of the page

    return {
        "handle": handle,
        "name": display_name,
        "page_id": page_id,
        # The id the Ad Library indexes this page under. Feed it straight into
        # /api/v1/adlibrary/{brand}?page_id=... for an exact, noise-free scan.
        "ad_library_page_id": find_first(r'"delegate_page":\{[^}]*"id":"(\d+)"', page_html) or page_id,
        "url": _clean(find_first(r'<link rel="canonical" href="([^"]*)"', page_html)) or f"{FB_BASE}/{handle}/",
        "categories": categories,
        "category": categories[0] if categories else None,
        "likes": int(likes.replace(",", "")) if likes else None,
        "talking_about": int(talking_about.replace(",", "")) if talking_about else None,
        "were_here": int(were_here.replace(",", "")) if were_here else None,
        "followers_display": followers_display,
        "followers_approx": _parse_compact_number(followers_display),
        "followers_is_approximate": True,
        "bio": bio,
        # null unless check_verified ran. The anonymous payload has no reliable
        # badge field -- its `is_verified` hits belong to other entities nested
        # in the same blob (nike reads false there while being verified), so we
        # only report this from the logged-in render's actual badge element.
        "verified": verified,
    }


def _parse_videos(videos_html: str) -> list:
    """Pull the /videos/ tab's `latest_videos` connection into flat dicts."""
    connection, _ = balanced_json(videos_html, '"latest_videos":')
    if not connection:
        return []
    out = []
    for edge in connection.get("edges", []):
        node = edge.get("node") or {}
        feedback = node.get("feedback") or {}
        reactions = (feedback.get("reaction_count") or {}).get("count")
        story = node.get("creation_story") or {}
        message = (story.get("message") or {}).get("text")
        out.append({
            "type": "video",
            "id": node.get("id"),
            "url": node.get("canonical_uri_with_fallback"),
            "published_at": _iso(node.get("publish_time")),
            "published_ts": node.get("publish_time"),
            "title": (node.get("savable_title") or {}).get("text"),
            "message": message,
            "plays": node.get("play_count"),
            "post_plays": node.get("post_play_count"),
            "reactions": reactions,
            "comments": None,   # not present in the anonymous videos payload
            "shares": None,
        })
    return out


def _parse_timeline(page_html: str) -> list:
    """Pull whatever stories Facebook server-rendered into the page payload.

    In practice this is one story. Kept as a list because a logged-in cookie
    jar occasionally yields more, and because the shape shouldn't change if it
    does.
    """
    connection, _ = balanced_json(page_html, '"timeline_list_feed_units":')
    if not connection:
        return []
    out = []
    for edge in connection.get("edges", []):
        node = edge.get("node") or {}
        if not node.get("post_id"):
            continue
        blob = _find_engagement(node)
        out.append({
            "type": "post",
            "id": node.get("post_id"),
            "url": blob.get("url"),
            "published_at": _iso(node.get("creation_time")),
            "published_ts": node.get("creation_time"),
            "title": None,
            "message": blob.get("message"),
            "plays": blob.get("video_view_count"),
            "post_plays": None,
            "reactions": blob.get("reactions"),
            "comments": blob.get("comments"),
            "shares": blob.get("shares"),
        })
    return out


def _find_engagement(node: dict) -> dict:
    """Walk a story node for the engagement counters, wherever Comet put them.

    The counters live at different depths depending on which renderer strategy
    Facebook picked for that story, so a targeted walk beats hardcoded paths.
    """
    found = {"reactions": None, "comments": None, "shares": None,
             "video_view_count": None, "url": None, "message": None}

    def visit(obj, depth=0):
        if depth > 25 or all(v is not None for v in found.values()):
            return
        if isinstance(obj, dict):
            for key, value in obj.items():
                if key == "reaction_count" and isinstance(value, dict) and found["reactions"] is None:
                    found["reactions"] = value.get("count")
                elif key in ("comment_count", "total_comment_count") and found["comments"] is None:
                    found["comments"] = value.get("count") if isinstance(value, dict) else value
                elif key == "share_count" and isinstance(value, dict) and found["shares"] is None:
                    found["shares"] = value.get("count")
                elif key == "video_view_count" and found["video_view_count"] is None:
                    found["video_view_count"] = value
                elif (key == "url" and found["url"] is None and isinstance(value, str)
                        and any(part in value for part in ("/reel/", "/posts/", "/videos/"))):
                    found["url"] = value
                elif key == "text" and found["message"] is None and isinstance(value, str) and len(value) > 15:
                    found["message"] = value
                visit(value, depth + 1)
        elif isinstance(obj, list):
            for item in obj:
                visit(item, depth + 1)

    visit(node)
    return found


def scrape_page_posts(handle: str, limit: int = 12, cookies: dict = None,
                      followers_hint: int = None, session=None) -> dict:
    """Recent content + an engagement summary for a page.

    Combines the one server-rendered timeline story with the ~6 videos the
    /videos/ tab renders, because neither surface alone is a usable sample.
    """
    handle = handle.strip().lstrip("@").rstrip("/")
    s = session or new_session()

    # Logged out, same reason as scrape_page: the anonymous render is the one
    # with parseable story/video payloads.
    _, page_html = get_html(s, f"{FB_BASE}/{handle}/")
    if any(marker in page_html for marker in NOT_AVAILABLE_MARKERS):
        raise PageScanError(f"Facebook has no public page at /{handle}")

    items = _parse_timeline(page_html)

    try:
        _, videos_html = get_html(s, f"{FB_BASE}/{handle}/videos/",
                                  referer=f"{FB_BASE}/{handle}/")
        items.extend(_parse_videos(videos_html))
    except Exception:
        # Pages with no video tab (or a gated one) still get their timeline story.
        pass

    # A reel shows up twice -- once as the timeline story, once as a video --
    # under two different ids, so dedupe on content too: same text posted
    # within a couple of minutes is the same thing seen from two tabs.
    seen_ids = set()
    seen_content = set()
    unique = []
    for item in sorted(items, key=lambda i: i.get("published_ts") or 0, reverse=True):
        item_id = item.get("id")
        if item_id in seen_ids:
            continue
        text = (item.get("message") or item.get("title") or "").strip()
        bucket = (text[:120], int((item.get("published_ts") or 0) / 120))
        if text and bucket in seen_content:
            continue
        seen_ids.add(item_id)
        seen_content.add(bucket)
        unique.append(item)
    unique = unique[:limit]

    reaction_values = [i["reactions"] for i in unique if isinstance(i.get("reactions"), int)]
    play_values = [i["plays"] for i in unique if isinstance(i.get("plays"), int)]
    avg_reactions = int(sum(reaction_values) / len(reaction_values)) if reaction_values else 0
    avg_plays = int(sum(play_values) / len(play_values)) if play_values else 0

    engagement_rate = None
    if followers_hint and reaction_values:
        engagement_rate = round(avg_reactions / followers_hint * 100, 4)

    return {
        "handle": handle,
        "items_returned": len(unique),
        "avg_reactions": avg_reactions,
        "avg_plays": avg_plays,
        "engagement_rate_pct": engagement_rate,
        "engagement_rate_basis": (
            "avg reactions / follower estimate -- Facebook only exposes a rounded "
            "follower count anonymously, so treat this as indicative, not exact"
            if engagement_rate is not None else None
        ),
        "items": unique,
    }
