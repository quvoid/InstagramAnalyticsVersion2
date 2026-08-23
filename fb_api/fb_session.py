"""
fb_session.py -- shared HTTP plumbing for the Facebook side of this project.

Everything Facebook-facing in here talks to the same two surfaces:

  1. Public Page HTML  (https://www.facebook.com/<handle>/)
     Works ANONYMOUSLY -- verified 2026-08-22 against zivame / nike / cocacola /
     MamaEarth.India. No cookies needed at all, which is the big difference from
     the Instagram side of this project (see ../api/, which cannot do anything
     without a live IG sessionid).

  2. Ad Library        (https://www.facebook.com/ads/library/?...)
     Does NOT work anonymously. Meta answers an anonymous request with HTTP 403
     and a tiny page whose script POSTs to a one-off `/__rd_verify_...` URL --
     an anti-bot gate. This module deliberately does NOT try to defeat that
     gate; it detects it and raises ChallengeError telling you to supply
     cookies from a browser where you normally use the Ad Library. Same deal
     as the Instagram scrapers: your own session, pasted into config.

Both surfaces embed their data as a Relay payload inside the server-rendered
HTML, which is why the balanced-JSON extractor below exists -- the payload is
not a tidy <script type="application/json"> blob you can json.loads() whole,
it's one JSON object buried in a much larger script body.
"""
import os
import re
import json
import threading

try:
    from curl_cffi import requests as cffi_requests
except ImportError:  # pragma: no cover - dependency check, not logic
    cffi_requests = None

FB_BASE = "https://www.facebook.com"

BASE_HEADERS = {
    "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "accept-language": "en-US,en;q=0.9",
    "sec-fetch-dest": "document",
    "sec-fetch-mode": "navigate",
    "sec-fetch-site": "none",
    "sec-fetch-user": "?1",
    "upgrade-insecure-requests": "1",
}

CHALLENGE_MARKER = "__rd_verify_"


class FacebookError(Exception):
    """Base for everything this package raises."""


class ChallengeError(FacebookError):
    """Meta served its anti-bot challenge instead of the page we asked for."""


class PageScanError(FacebookError):
    pass


class AdLibraryError(FacebookError):
    pass


def parse_cookie_string(raw: str) -> dict:
    """Turn a browser 'Cookie:' header value into a dict.

    Accepts exactly what you get from DevTools -> Network -> any facebook.com
    request -> Request Headers -> cookie, i.e.
        'datr=abc; sb=def; c_user=123; xs=...'
    A leading 'cookie:' / 'Cookie:' label is tolerated, as are surrounding
    quotes, because that's how it comes off the clipboard half the time.
    Empty/whitespace input gives an empty dict (anonymous mode).
    """
    raw = (raw or "").strip().strip('"').strip("'")
    if raw[:7].lower() == "cookie:":
        raw = raw[7:].strip()
    cookies = {}
    for part in raw.replace("\n", ";").split(";"):
        part = part.strip()
        if not part or "=" not in part:
            continue
        name, _, value = part.partition("=")
        cookies[name.strip()] = value.strip()
    return cookies


def parse_cookie_file(path: str) -> dict:
    """Read cookies from a file, in whichever of the three usual shapes it is.

    1. A raw cookie header on one line (same thing FB_COOKIE takes).
    2. JSON exported by a cookie-editor extension: a list of
       {"name": ..., "value": ...} objects, or {"cookies": [...]}.
    3. Netscape cookies.txt (what curl/wget and most exporters emit):
       tab-separated, comment lines start with '#'.

    Useful because a full Facebook cookie header is long enough that pasting it
    into .env on one line is genuinely annoying.
    """
    with open(path, "r", encoding="utf-8") as fh:
        content = fh.read().strip()
    if not content:
        return {}

    if content[0] in "[{":
        try:
            data = json.loads(content)
        except json.JSONDecodeError:
            data = None
        if isinstance(data, dict):
            data = data.get("cookies", [])
        if isinstance(data, list):
            return {
                c["name"]: c["value"]
                for c in data
                if isinstance(c, dict) and c.get("name") and c.get("value") is not None
            }

    if "\t" in content:
        cookies = {}
        for line in content.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            fields = line.split("\t")
            if len(fields) >= 7:
                cookies[fields[5]] = fields[6]
        if cookies:
            return cookies

    return parse_cookie_string(content)


def load_cookies(cookie_string: str = "", cookie_file: str = "") -> dict:
    """FB_COOKIE / FB_COOKIE_FILE, whichever is set. File wins if both are."""
    if cookie_file:
        expanded = os.path.expanduser(os.path.expandvars(cookie_file))
        if not os.path.exists(expanded):
            raise FacebookError(f"FB_COOKIE_FILE points at a file that doesn't exist: {expanded}")
        return parse_cookie_file(expanded)
    return parse_cookie_string(cookie_string)


# Cookies that actually matter, and what each one is for. Used by the
# /api/v1/session/check diagnostic so a bad paste is obvious immediately
# rather than three endpoints later.
SESSION_COOKIES = {
    "c_user": "your numeric user id -- present only when logged in",
    "xs": "the session secret; without this you are anonymous no matter what else is set",
    "datr": "browser identifier Meta ties the session to",
    "sb": "secure browser id, usually issued alongside datr",
    "fr": "ads/session cookie, normally present on a logged-in browser",
}


def describe_cookies(cookies: dict) -> dict:
    """Non-sensitive summary of a cookie jar -- names only, never values."""
    names = sorted(cookies or {})
    return {
        "cookie_count": len(names),
        "cookie_names": names,
        "logged_in": bool(cookies.get("c_user") and cookies.get("xs")),
        "missing_session_cookies": [c for c in ("c_user", "xs", "datr") if c not in (cookies or {})],
    }


def new_session():
    if cffi_requests is None:
        raise FacebookError("curl_cffi is not installed (pip install curl_cffi)")
    # Same impersonation the Instagram scrapers use -- Meta fingerprints TLS,
    # and a plain `requests` handshake gets noticeably more friction.
    return cffi_requests.Session(impersonate="chrome")


def get_html(session, url, cookies=None, params=None, timeout=45, referer=None):
    """GET a Facebook page and hand back the HTML, or raise ChallengeError."""
    headers = dict(BASE_HEADERS)
    if referer:
        headers["referer"] = referer
    r = session.get(url, params=params, headers=headers, cookies=cookies or {},
                    timeout=timeout, allow_redirects=True)
    html = r.text
    if r.status_code == 403 and CHALLENGE_MARKER in html:
        sent = describe_cookies(cookies or {})
        if not sent["cookie_count"]:
            detail = "No cookies were sent -- this request was anonymous, which Meta always challenges."
        elif not sent["logged_in"]:
            detail = (f"Cookies were sent ({', '.join(sent['cookie_names'])}) but they are not a "
                      f"logged-in session; missing {', '.join(sent['missing_session_cookies'])}.")
        else:
            detail = ("A logged-in cookie jar was sent but Meta still challenged it -- the session is "
                      "probably stale. Re-copy the cookie header from a live browser tab.")
        raise ChallengeError(
            f"Meta served its anti-bot challenge for this URL (HTTP 403). {detail} "
            f"Copy the FULL cookie header from DevTools -> Network -> any facebook.com request -> "
            f"Request Headers -> cookie, and put it in FB_COOKIE (or a file named by FB_COOKIE_FILE) "
            f"in fb_api/.env. Check it with GET /api/v1/session/check."
        )
    if CHALLENGE_MARKER in html:
        raise ChallengeError("Meta served its anti-bot challenge instead of page content.")
    return r, html


BS = chr(92)  # backslash, spelled this way so the escaping below stays readable


def balanced_json(html: str, key: str, start: int = 0):
    """Extract the balanced {...} object that follows `key` in `html`.

    Facebook inlines its Relay payloads inside larger script bodies, so a regex
    can't grab them -- you have to walk braces while respecting strings and
    escapes. Returns (parsed_object, index_after) or (None, -1).
    """
    i = html.find(key, start)
    if i < 0:
        return None, -1
    i = html.find("{", i)
    if i < 0:
        return None, -1
    depth = 0
    in_str = False
    esc = False
    for j in range(i, len(html)):
        c = html[j]
        if esc:
            esc = False
        elif c == BS:
            esc = True
        elif c == '"':
            in_str = not in_str
        elif not in_str:
            if c == "{":
                depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0:
                    blob = html[i:j + 1]
                    try:
                        return json.loads(blob), j + 1
                    except json.JSONDecodeError:
                        return None, j + 1
    return None, -1


def find_first(pattern: str, html: str, group: int = 1, default=None):
    m = re.search(pattern, html)
    return m.group(group) if m else default


def find_lsd(html: str):
    """The LSD (cross-site request) token Facebook embeds in every page.

    Required on the /api/graphql/ POST used for Ad Library pagination.
    """
    return find_first(r'"LSD",\[\],\{"token":"([^"]+)"', html)


# --- GraphQL persisted-query (doc_id) discovery -----------------------------
# Facebook's GraphQL endpoint only accepts persisted queries: you send a
# numeric doc_id, not a query string. Each Relay operation's doc_id lives in
# the JS bundle as:
#     __d("<OperationName>_facebookRelayOperation",[],(function(a,b,c,d,e,f){e.exports="24922295957467452"}),null)
# The id rotates whenever Meta ships that bundle, so we discover it from the
# page we just loaded rather than hardcoding it -- but we seed the cache with
# the value verified on 2026-08-22 so the common path costs zero extra
# requests, and we only go hunting again if a request actually fails.
_DOC_ID_CACHE = {"AdLibrarySearchPaginationQuery": "24922295957467452"}
_DOC_ID_LOCK = threading.Lock()

_DOC_ID_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".doc_id_cache.json")


def _load_doc_id_cache():
    try:
        with open(_DOC_ID_CACHE_FILE, "r", encoding="utf-8") as fh:
            _DOC_ID_CACHE.update(json.load(fh))
    except Exception:
        pass


def _save_doc_id_cache():
    try:
        with open(_DOC_ID_CACHE_FILE, "w", encoding="utf-8") as fh:
            json.dump(_DOC_ID_CACHE, fh)
    except Exception:
        pass


_load_doc_id_cache()


def get_cached_doc_id(operation: str):
    """Cached doc_id, unless AD_DOC_ID in .env overrides the pagination query.

    The override is the escape hatch for the day Meta rotates the id and you
    don't want to install Playwright just to rediscover it: open the Ad
    Library in a browser, search the JS bundles for
    "AdLibrarySearchPaginationQuery_facebookRelayOperation", and paste the
    number next to it into AD_DOC_ID.
    """
    if operation == "AdLibrarySearchPaginationQuery":
        override = os.environ.get("AD_DOC_ID", "").strip()
        if override:
            return override
    with _DOC_ID_LOCK:
        return _DOC_ID_CACHE.get(operation)


def discover_doc_id(session, html: str, operation: str, cookies=None, max_bundles: int = 60):
    """Find `operation`'s doc_id by grepping the page's own JS bundles.

    Sequential with an early exit -- the bundle that carries an Ad Library
    operation is usually among the first handful of Ad Library scripts, and
    bailing on the first hit keeps this to a couple of extra requests.
    """
    needle = f'__d("{operation}_facebookRelayOperation"'
    pattern = re.compile(
        re.escape(operation) + r'_facebookRelayOperation",\[\],\(function\([^)]*\)\{[a-z]\.exports="(\d+)"'
    )
    srcs = re.findall(r'<script[^>]+src="([^"]+\.js[^"]*)"', html)
    seen = set()
    for src in srcs[:max_bundles]:
        src = src.replace("&amp;", "&")
        if src in seen:
            continue
        seen.add(src)
        try:
            r = session.get(src, headers={"accept": "*/*"}, cookies=cookies or {}, timeout=30)
        except Exception:
            continue
        if r.status_code != 200:
            continue
        body = r.text
        if needle not in body and operation not in body:
            continue
        m = pattern.search(body)
        if m:
            with _DOC_ID_LOCK:
                _DOC_ID_CACHE[operation] = m.group(1)
                _save_doc_id_cache()
            return m.group(1)
    return None


def graphql_post(session, variables: dict, doc_id: str, friendly_name: str,
                 lsd: str, cookies=None, timeout=45, user="0", dtsg=None):
    """POST a persisted Relay query to /api/graphql/ and return parsed JSON.

    Verified request shape (2026-08-22): form-encoded body with lsd +
    fb_api_req_friendly_name + variables + doc_id, `x-fb-lsd` header set.
    Responses are sometimes prefixed with the `for (;;);` XSSI guard.

    `user`/`dtsg` come from a logged-in session; both are omitted for anonymous
    calls, which the endpoint also accepts.
    """
    body = {
        "av": user,
        "__user": user,
        "__a": "1",
        "__comet_req": "1",
        "lsd": lsd,
        "fb_api_caller_class": "RelayModern",
        "fb_api_req_friendly_name": friendly_name,
        "variables": json.dumps(variables, separators=(",", ":")),
        "server_timestamps": "true",
        "doc_id": doc_id,
    }
    if dtsg:
        body["fb_dtsg"] = dtsg
    headers = {
        "content-type": "application/x-www-form-urlencoded",
        "accept": "*/*",
        "origin": FB_BASE,
        "referer": FB_BASE + "/ads/library/",
        "x-fb-lsd": lsd,
        "x-fb-friendly-name": friendly_name,
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-origin",
    }
    r = session.post(FB_BASE + "/api/graphql/", data=body, headers=headers,
                     cookies=cookies or {}, timeout=timeout)
    text = r.text
    if CHALLENGE_MARKER in text:
        raise ChallengeError("Meta served its anti-bot challenge on the GraphQL endpoint.")
    text = re.sub(r"^for\s*\(;;\);", "", text.strip())
    # Comet sometimes streams several JSON documents newline-separated; the
    # first one carries the connection we asked for.
    first_line = text.split("\n", 1)[0]
    try:
        return json.loads(first_line)
    except json.JSONDecodeError:
        raise AdLibraryError(f"unparseable GraphQL response (HTTP {r.status_code}): {text[:200]}")
