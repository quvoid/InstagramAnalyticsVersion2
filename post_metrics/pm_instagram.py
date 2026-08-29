"""
Instagram post-metrics collector for a single handle.

Uses the private mobile API with a logged-in session cookie, which is the only
way to read public metrics (likes / comments / plays) at volume. Two endpoints
are combined because neither one is complete on its own:

  feed/user/{id}   -- timeline posts (photos, carousels, and the reels the
                      account chose to show on its grid)
  clips/user/      -- reels, including ones hidden from the grid

Results are merged and de-duplicated by shortcode.
"""

import re
import sys
import time
from datetime import datetime, timedelta, timezone

import pm_config as cfg

try:
    from curl_cffi import requests as curl_requests
    _HAS_CURL_CFFI = True
except ImportError:
    _HAS_CURL_CFFI = False
import requests as plain_requests

MOBILE_HEADERS = {
    "User-Agent": (
        "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; "
        "OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)"
    ),
    "x-ig-app-id": "936619743392459",
}

WEB_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "x-ig-app-id": "936619743392459",
    "accept": "*/*",
}

MEDIA_TYPE_NAMES = {1: "Photo", 2: "Video", 8: "Carousel"}


class IGAuthError(RuntimeError):
    """Raised when Instagram rejects the session cookie."""


class IGNotFound(RuntimeError):
    """Raised when the handle cannot be resolved."""


def make_session():
    """A TLS-fingerprint-matching session when curl_cffi is available."""
    if _HAS_CURL_CFFI:
        return curl_requests.Session(impersonate="chrome120")
    return plain_requests.Session()


def parse_handle(value: str) -> str:
    """Accept a bare handle, an @handle, or any instagram.com profile URL."""
    v = (value or "").strip()
    if not v:
        return ""
    m = re.search(r"instagram\.com/([A-Za-z0-9._]+)", v)
    if m:
        v = m.group(1)
    return v.lstrip("@").strip("/").strip().lower()


class PostMetricsCollector:
    def __init__(self):
        self.session = make_session()
        self.cookies = cfg.get_cookies()
        if not self.cookies.get("sessionid"):
            raise IGAuthError("No IG_SESSIONID configured. Set it in Streamlit secrets.")

    # -- profile --------------------------------------------------------------
    def resolve_profile(self, handle: str) -> dict:
        """Resolve a handle to a numeric id plus follower count.

        Tries the mobile search endpoint first, then falls back to scraping the
        profile HTML, which still works for business accounts whose
        web_profile_info response is broken on Instagram's side.
        """
        handle = parse_handle(handle)
        if not handle:
            raise IGNotFound("Empty handle.")

        url = "https://i.instagram.com/api/v1/users/search/?q=" + handle
        r = self._get(url, MOBILE_HEADERS)
        if r is not None and r.status_code == 200:
            try:
                for u in r.json().get("users", []):
                    if (u.get("username") or "").lower() == handle:
                        return {
                            "user_id": u.get("pk"),
                            "handle": u.get("username"),
                            "full_name": u.get("full_name") or "",
                            "is_verified": bool(u.get("is_verified")),
                            "follower_count": u.get("follower_count") or 0,
                        }
            except ValueError:
                pass
        elif r is not None and r.status_code in (401, 403):
            raise IGAuthError(
                "Instagram rejected the session (HTTP %s). The sessionid has "
                "expired -- refresh it in secrets." % r.status_code
            )

        return self._resolve_via_html(handle)

    def _resolve_via_html(self, handle: str) -> dict:
        r = self._get("https://www.instagram.com/%s/" % handle, WEB_HEADERS)
        if r is None or r.status_code != 200:
            raise IGNotFound("Could not resolve @%s." % handle)
        m = re.search(r'"profilePage_(\d+)"', r.text)
        if not m:
            m = re.search(r'"profile_id"\s*:\s*"(\d+)"', r.text)
        if not m:
            raise IGNotFound("Could not find a user id for @%s." % handle)
        followers = 0
        fm = re.search(r'"edge_followed_by"\s*:\s*\{\s*"count"\s*:\s*(\d+)', r.text)
        if fm:
            followers = int(fm.group(1))
        return {
            "user_id": int(m.group(1)),
            "handle": handle,
            "full_name": "",
            "is_verified": False,
            "follower_count": followers,
        }

    # -- collection -----------------------------------------------------------
    def collect(self, handle: str, max_posts: int, since=None, log=None) -> tuple:
        """Return (profile, [normalised post dicts]).

        `since` stops pagination once a whole page is older than that instant,
        which is what makes the daily delta cheap. `max_posts` is a hard cap.
        """
        say = log or (lambda m: None)
        profile = self.resolve_profile(handle)
        say("Resolved @%s -> id %s (%s followers)" % (
            profile["handle"], profile["user_id"],
            format(profile["follower_count"], ",")))

        raw = {}
        for item in self._paginate_feed(profile["user_id"], max_posts, since, say):
            code = item.get("code")
            if code:
                raw.setdefault(code, item)
        say("Timeline feed: %d posts" % len(raw))

        before = len(raw)
        for item in self._paginate_clips(profile["user_id"], max_posts, since, say):
            code = item.get("code")
            if code:
                raw.setdefault(code, item)
        say("Clips feed: +%d additional reels" % (len(raw) - before))

        posts = [self.normalise(i, profile) for i in raw.values()]
        posts.sort(key=lambda p: p["taken_at"], reverse=True)
        return profile, posts

    def _page_is_stale(self, items, since) -> bool:
        """True when no item on the page is newer than `since`.

        Pinned posts are excluded from the test -- they surface at the top of
        the feed with an old timestamp and would otherwise stop pagination on
        the very first page.
        """
        if since is None:
            return False
        cutoff = since.timestamp()
        for it in items:
            if it.get("timeline_pinned_user_ids"):
                continue
            if (it.get("taken_at") or 0) > cutoff:
                return False
        return True

    def _paginate_feed(self, user_id, max_posts, since, say):
        collected, max_id, page = [], "", 0
        while len(collected) < max_posts:
            page += 1
            url = "https://i.instagram.com/api/v1/feed/user/%s/?count=33" % user_id
            if max_id:
                url += "&max_id=" + str(max_id)
            r = self._get(url, MOBILE_HEADERS)
            if r is None or r.status_code != 200:
                if r is not None and r.status_code in (401, 403):
                    raise IGAuthError("Session rejected on feed page %d." % page)
                break
            try:
                data = r.json()
            except ValueError:
                break
            items = data.get("items") or []
            if not items:
                break
            collected.extend(items)
            if self._page_is_stale(items, since):
                say("Reached the refresh cutoff on feed page %d; stopping." % page)
                break
            max_id = data.get("next_max_id")
            if not max_id or not data.get("more_available", True):
                break
            time.sleep(cfg.PAGE_SLEEP_SECONDS)
        return collected[:max_posts]

    def _paginate_clips(self, user_id, max_posts, since, say):
        collected, max_id, page = [], "", 0
        while len(collected) < max_posts:
            page += 1
            payload = {"target_user_id": str(user_id), "page_size": "30"}
            if max_id:
                payload["max_id"] = str(max_id)
            r = self._post("https://i.instagram.com/api/v1/clips/user/",
                           MOBILE_HEADERS, payload)
            if r is None or r.status_code != 200:
                break
            try:
                data = r.json()
            except ValueError:
                break
            items = [it.get("media") for it in (data.get("items") or []) if it.get("media")]
            if not items:
                break
            collected.extend(items)
            if self._page_is_stale(items, since):
                break
            paging = data.get("paging_info") or {}
            max_id = paging.get("max_id")
            if not max_id or not paging.get("more_available"):
                break
            time.sleep(cfg.PAGE_SLEEP_SECONDS)
        return collected[:max_posts]

    # -- normalisation --------------------------------------------------------
    @staticmethod
    def normalise(item: dict, profile: dict) -> dict:
        """Flatten one raw API item into the row shape written to the sheet."""
        code = item.get("code") or ""
        taken_at = int(item.get("taken_at") or 0)
        caption = ((item.get("caption") or {}).get("text") or "").replace("\n", " ").strip()

        media_type = item.get("media_type")
        product_type = item.get("product_type") or ""
        if product_type == "clips":
            kind = "Reel"
        elif product_type == "igtv":
            kind = "IGTV"
        else:
            kind = MEDIA_TYPE_NAMES.get(media_type, "Unknown")

        likes = item.get("like_count")
        comments = item.get("comment_count")
        views = (item.get("play_count") or item.get("ig_play_count")
                 or item.get("view_count") or 0)

        coauthors = [
            (c.get("username") or "")
            for c in (item.get("coauthor_producers") or [])
            if c.get("username")
        ]

        followers = profile.get("follower_count") or 0
        interactions = (likes or 0) + (comments or 0)
        engagement_rate = round(interactions / followers * 100, 4) if followers else 0.0

        return {
            "shortcode": code,
            "handle": profile.get("handle", ""),
            "post_url": ("https://www.instagram.com/p/%s/" % code) if code else "",
            "taken_at": taken_at,
            "posted_at": (datetime.fromtimestamp(taken_at, timezone.utc).isoformat()
                          if taken_at else ""),
            "media_kind": kind,
            "carousel_count": len(item.get("carousel_media") or []),
            "owner": ((item.get("user") or {}).get("username") or ""),
            "caption": caption[:4000],
            "like_count": likes if likes is not None else "",
            "comment_count": comments if comments is not None else "",
            "view_count": views or "",
            "video_duration": round(item.get("video_duration") or 0, 2),
            "is_paid_partnership": bool(item.get("is_paid_partnership")),
            "coauthors": ", ".join(coauthors),
            "counts_hidden": bool(item.get("like_and_view_counts_disabled")),
            "follower_count_at_scrape": followers,
            "engagement_rate_pct": engagement_rate,
            "scraped_at": datetime.now(timezone.utc).isoformat(),
        }

    # -- transport ------------------------------------------------------------
    def _get(self, url, headers):
        try:
            return self.session.get(url, headers=headers, cookies=self.cookies,
                                    timeout=cfg.REQUEST_TIMEOUT)
        except Exception as e:
            print("[warn] GET %s failed: %s" % (url[:70], e), file=sys.stderr)
            return None

    def _post(self, url, headers, data):
        try:
            return self.session.post(url, headers=headers, data=data,
                                     cookies=self.cookies, timeout=cfg.REQUEST_TIMEOUT)
        except Exception as e:
            print("[warn] POST %s failed: %s" % (url[:70], e), file=sys.stderr)
            return None


def refresh_cutoff(days=None):
    """The instant before which posts are considered settled."""
    days = cfg.REFRESH_WINDOW_DAYS if days is None else days
    return datetime.now(timezone.utc) - timedelta(days=days)
