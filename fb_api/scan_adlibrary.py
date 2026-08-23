"""
scan_adlibrary.py -- Meta Ad Library scan, i.e. "Method B" from
PROJECT_CONTEXT.md, finally available as something other than hand-driven
browser scrolling.

Why this exists
---------------
The old Method B workflow scrolled the Ad Library UI and scraped rendered
text. PROJECT_CONTEXT.md §4 records what that cost: for Underneat we captured
108 creator ads and called it complete, and TrendTrack showed 208. The
scroll-plateau was a FALSE plateau, and the gap was unpredictable per brand.

This module doesn't scroll. It reads the same data the UI reads:

  * Meta server-renders the first page of results into the HTML as a Relay
    payload under "search_results_connection", which carries `count` --
    Meta's OWN total for the query. So we always know how many ads should
    exist, and can say plainly whether we got them all.
  * Further pages come from the same persisted GraphQL query the UI fires on
    scroll (AdLibrarySearchPaginationQuery), driven by page_info.end_cursor.

Every response therefore carries reported_total / captured / complete. Do not
strip those out when building deliverables -- they are the fix for the
"confirmed complete" claim that had to be walked back last time.

Branded-content detection is also structural now, not string-matching. The old
approach keyed off the advertiser label reading "<Creator> with <Brand>". The
payload actually contains, verified 2026-08-22 on Zivame:

    result.page_id / page_name          -> the BRAND paying for the ad
    result.snapshot.page_name           -> the CREATOR's page
    result.snapshot.page_profile_uri    -> the CREATOR's profile URL
    result.snapshot.page_like_count     -> the CREATOR's page likes
    result.snapshot.branded_content     -> {page_id, page_name, page_profile_uri}
                                           of the BRAND being promoted

`snapshot.branded_content` being non-null IS the paid-partnership signal.

Transports
----------
"browser" (default, verified working): drives a headless Chromium via
Playwright. Anonymous plain-HTTP requests to /ads/library/ are answered with
HTTP 403 and Meta's anti-bot challenge page; a real browser runs the page's
own script and loads normally, exactly like you loading the Ad Library by
hand. This module does not attempt to defeat that challenge itself.

"http": plain curl_cffi with whatever cookies you configured in FB_COOKIE.
Cheaper and faster when your session cookies are accepted; raises
ChallengeError when they aren't, and you should fall back to "browser".
"""
import time
import json
import uuid
import random
import threading
from datetime import datetime, timezone

from fb_session import (
    FB_BASE,
    AdLibraryError,
    ChallengeError,
    balanced_json,
    discover_doc_id,
    find_first,
    find_lsd,
    get_cached_doc_id,
    get_html,
    graphql_post,
    new_session,
)

PAGINATION_OPERATION = "AdLibrarySearchPaginationQuery"
PAGE_SIZE = 30                 # what Meta serves per page, observed consistently
MAX_PAGES_SAFETY = 400         # 400 * 30 = 12k ads; nothing here should get near it


def build_search_url(brand=None, page_id=None, country="IN", active_status="active",
                     ad_type="all", media_type="all", search_type=None,
                     start_date_min=None, start_date_max=None) -> str:
    """Build an Ad Library URL.

    Two modes, and the choice matters a lot:

      page mode (page_id given) -- exact. Only ads from that one Page.
      keyword mode (brand given) -- fuzzy. PROJECT_CONTEXT.md §4 warns that
        generic brand names (Jockey, Clovia, Lyra, Amante) drag in unrelated
        advertisers; a Zivame keyword search returns ~350 ads against ~69 for
        the actual Zivame page. Prefer page mode whenever you can resolve the
        page id (GET /api/v1/page/{handle} gives you ad_library_page_id).
    """
    params = [
        ("active_status", active_status),
        ("ad_type", ad_type),
        ("country", country),
        ("media_type", media_type),
    ]
    if page_id:
        params.append(("view_all_page_id", str(page_id)))
        params.append(("search_type", search_type or "page"))
    else:
        if not brand:
            raise AdLibraryError("either brand (keyword) or page_id is required")
        params.append(("q", brand))
        params.append(("search_type", search_type or "keyword_unordered"))
    if start_date_min:
        params.append(("start_date[min]", start_date_min))
    if start_date_max:
        params.append(("start_date[max]", start_date_max))

    from urllib.parse import urlencode
    return f"{FB_BASE}/ads/library/?" + urlencode(params)


# --- payload parsing --------------------------------------------------------

def extract_connection(html: str):
    """The server-rendered first page of results."""
    connection, _ = balanced_json(html, '"search_results_connection":')
    return connection


def extract_root_variables(html: str):
    """The exact GraphQL variables the UI used for this search.

    Reusing Meta's own variables (rather than reconstructing them) is what
    keeps pagination honest: filters, session id, sort order and country all
    carry over untouched, so page 2 is genuinely page 2 of the same query.
    """
    marker = "adp_AdLibraryFoundationRootQueryRelayPreloader"
    idx = html.find(marker)
    if idx < 0:
        return None
    segment = html[idx:idx + 6000]
    # This blob sits inside a JSON string in some renders and raw in others.
    for candidate in (segment, segment.replace('\\"', '"')):
        variables, _ = balanced_json(candidate, '"variables":')
        if variables and "activeStatus" in variables:
            return variables
    return None


def _iso_date(ts):
    if not ts:
        return None
    return datetime.fromtimestamp(int(ts), tz=timezone.utc).strftime("%Y-%m-%d")


def _body_text(snapshot):
    body = snapshot.get("body")
    if isinstance(body, dict):
        return body.get("text")
    if isinstance(body, str):
        # Some renders hand back the JSON-encoded body as a string.
        if body.startswith('{"text"'):
            try:
                return json.loads(body).get("text")
            except json.JSONDecodeError:
                return body
        return body
    return None


def normalize_ad(result: dict) -> dict:
    snapshot = result.get("snapshot") or {}
    branded = snapshot.get("branded_content")
    archive_id = result.get("ad_archive_id")

    creator = None
    if branded:
        creator = {
            "name": (snapshot.get("page_name") or "").strip() or None,
            "profile_url": snapshot.get("page_profile_uri"),
            "page_id": snapshot.get("page_id"),
            "page_likes": snapshot.get("page_like_count"),
            "categories": snapshot.get("page_categories") or [],
        }

    return {
        "ad_archive_id": archive_id,
        "ad_library_url": f"{FB_BASE}/ads/library/?id={archive_id}" if archive_id else None,
        "advertiser_page_id": result.get("page_id"),
        "advertiser_page_name": result.get("page_name"),
        "is_active": result.get("is_active"),
        "started_running": _iso_date(result.get("start_date")),
        "ended_running": _iso_date(result.get("end_date")),
        "start_ts": result.get("start_date"),
        "publisher_platforms": result.get("publisher_platform") or [],
        "display_format": snapshot.get("display_format"),
        "title": snapshot.get("title"),
        "body": _body_text(snapshot),
        "caption": snapshot.get("caption"),
        "cta_text": snapshot.get("cta_text"),
        "link_url": snapshot.get("link_url"),
        "collation_count": result.get("collation_count"),
        "collation_id": result.get("collation_id"),
        "categories": result.get("categories") or [],
        "currency": result.get("currency") or None,
        "spend": result.get("spend"),
        "reach_estimate": result.get("reach_estimate"),
        "is_branded_content": bool(branded),
        "creator": creator,
        "branded_content_brand": {
            "page_id": branded.get("page_id"),
            "page_name": branded.get("page_name"),
            "profile_url": branded.get("page_profile_uri"),
        } if branded else None,
    }


def iter_results(connection: dict):
    for edge in connection.get("edges") or []:
        node = edge.get("node") or {}
        for result in (node.get("collated_results") or [node]):
            if result.get("ad_archive_id"):
                yield result


def _is_self_partnership(ad: dict) -> bool:
    """Is this 'branded content' just the brand tagging itself?

    Verified on Underneat: 63 of its branded-content ads name Underneat's own
    page as the creator. Same trap PROJECT_CONTEXT.md §3 flags on the organic
    side with sister/sub-brand accounts -- counting these as creator
    partnerships inflates the partner list with the advertiser itself.
    """
    creator = ad.get("creator") or {}
    if not creator:
        return False
    if creator.get("page_id") and creator["page_id"] == ad.get("advertiser_page_id"):
        return True
    creator_name = (creator.get("name") or "").strip().lower()
    advertiser_name = (ad.get("advertiser_page_name") or "").strip().lower()
    return bool(creator_name) and creator_name == advertiser_name


def summarize_partners(ads: list) -> list:
    """Unique creator partners across the captured ads.

    Keyed on the creator's profile URL when present (page names are not unique
    and often carry trailing whitespace straight from Meta). Self-partnerships
    are dropped here -- see _is_self_partnership.
    """
    partners = {}
    for ad in ads:
        creator = ad.get("creator")
        if not creator or ad.get("is_self_partnership"):
            continue
        key = (creator.get("profile_url") or creator.get("name") or "").lower().strip()
        if not key:
            continue
        entry = partners.setdefault(key, {
            "name": creator.get("name"),
            "profile_url": creator.get("profile_url"),
            "page_id": creator.get("page_id"),
            "page_likes": creator.get("page_likes"),
            "categories": creator.get("categories"),
            "ad_count": 0,
            "first_seen": None,
            "last_seen": None,
            "ad_library_urls": [],
        })
        entry["ad_count"] += 1
        started = ad.get("started_running")
        if started:
            if not entry["first_seen"] or started < entry["first_seen"]:
                entry["first_seen"] = started
            if not entry["last_seen"] or started > entry["last_seen"]:
                entry["last_seen"] = started
        if ad.get("ad_library_url") and len(entry["ad_library_urls"]) < 25:
            entry["ad_library_urls"].append(ad["ad_library_url"])
    return sorted(partners.values(), key=lambda p: p["ad_count"], reverse=True)


# --- transports -------------------------------------------------------------

class _GraphQLTransport:
    """Cookies only. No page load, no browser. This is the default.

    The trick: Meta's anti-bot challenge sits on the /ads/library/ HTML route,
    not on /api/graphql/. Verified 2026-08-22 -- a logged-in session still gets
    a 403 challenge on the HTML page, while the same cookies query the GraphQL
    endpoint fine. So this transport never touches that route at all: it takes
    the LSD/DTSG tokens from the ordinary facebook.com homepage (not
    challenged) and queries AdLibrarySearchPaginationQuery directly, building
    the variables itself instead of lifting them out of a rendered page.

    The one thing it gives up: the HTML payload's `count` field, Meta's own
    total for the query, which the GraphQL response doesn't carry. Completeness
    is proven a different way here -- by paginating until Meta says
    has_next_page is false. That's a stronger signal than a count comparison
    anyway; it's the cursor running out, not a number matching.
    """

    name = "graphql"

    # Tokens are good for a while; refetching a 2MB homepage per scan is silly.
    _token_cache = {}
    _token_lock = threading.Lock()
    TOKEN_TTL = 600

    def __init__(self, cookies=None):
        self.cookies = cookies or {}
        self.session = new_session()
        self.lsd = None
        self.dtsg = None
        self.user = "0"

    def _fetch_tokens(self):
        cache_key = self.cookies.get("c_user") or "anon"
        with self._token_lock:
            cached = self._token_cache.get(cache_key)
            if cached and (time.time() - cached[0]) < self.TOKEN_TTL:
                self.lsd, self.dtsg, self.user = cached[1]
                return

        _, html = get_html(self.session, f"{FB_BASE}/", cookies=self.cookies)
        self.lsd = find_lsd(html)
        self.dtsg = find_first(r'"DTSGInitialData",\[\],\{"token":"([^"]+)"', html)
        self.user = find_first(r'"USER_ID":"(\d+)"', html, default="0")
        if not self.lsd:
            raise AdLibraryError("could not read an LSD token from facebook.com -- cookies rejected?")
        if self.user == "0":
            raise ChallengeError(
                "facebook.com loaded but the session is logged out (no USER_ID). Re-copy the "
                "cookie header from a logged-in tab into FB_COOKIE."
            )
        with self._token_lock:
            self._token_cache[cache_key] = (time.time(), (self.lsd, self.dtsg, self.user))

    def build_variables(self, brand=None, page_id=None, country="IN", active_status="active",
                        media_type="all", ad_type="ALL", start_date_min=None, start_date_max=None):
        """Reconstruct the variables the Ad Library UI would have sent.

        Shape lifted verbatim from a real AdLibraryFoundationRootQuery
        preloader payload, so the names and nulls match what Meta expects.
        """
        start_date = None
        if start_date_min or start_date_max:
            start_date = {"min": start_date_min, "max": start_date_max}
        return {
            "activeStatus": active_status,
            "adType": ad_type.upper(),
            "audienceTimeframe": "LAST_7_DAYS",
            "bylines": [],
            "collationToken": None,
            "contentLanguages": [],
            "countries": [country.upper()],
            "country": country.upper(),
            "excludedIDs": None,
            "isAboutTab": False,
            "isAudienceTab": False,
            "isLandingPage": False,
            "isTargetedCountry": False,
            "hasDeeplinkAdID": False,
            "location": None,
            "mediaType": media_type,
            "multiCountryFilterMode": None,
            "pageIDs": [],
            "potentialReachInput": None,
            "publisherPlatforms": [],
            "queryString": "" if page_id else (brand or ""),
            "regions": None,
            "searchType": "page" if page_id else "keyword_unordered",
            "sessionID": str(uuid.uuid4()),
            "source": None,
            "sortData": None,
            "startDate": start_date,
            "fetchPageInfo": False,
            "fetchSharedDisclaimers": False,
            "viewAllPageID": str(page_id) if page_id else "0",
            "v": "29f071",
        }

    def first_page(self, **query):
        self._fetch_tokens()
        variables = self.build_variables(**query)
        return self.paginate(variables, cursor=None), variables, None

    def paginate(self, variables, cursor, count=PAGE_SIZE):
        if not self.lsd:
            self._fetch_tokens()
        payload = dict(variables)
        payload["count"] = count
        payload["cursor"] = cursor
        doc_id = get_cached_doc_id(PAGINATION_OPERATION)
        data = graphql_post(self.session, payload, doc_id, PAGINATION_OPERATION, self.lsd,
                            cookies=self.cookies, user=self.user, dtsg=self.dtsg)
        connection = _connection_from_graphql(data)
        if connection is None:
            raise AdLibraryError(
                f"GraphQL returned no results connection. Meta most likely rotated the "
                f"{PAGINATION_OPERATION} doc_id (currently {doc_id}). Set AD_DOC_ID in .env to the "
                f"new one, or run this scan once with ?transport=browser to rediscover it "
                f"automatically. Response: {json.dumps(data)[:200]}"
            )
        return connection

    def close(self):
        pass


class _HttpTransport:
    """Plain curl_cffi against the /ads/library/ HTML route.

    Kept because this route's payload carries Meta's own `count` for the
    query, which the GraphQL transport can't see. In practice Meta challenges
    this route even for logged-in sessions, so expect ChallengeError and use
    the graphql transport -- but if Meta ever relaxes that, this is the
    richest source.
    """

    name = "http"

    def __init__(self, cookies=None):
        self.cookies = cookies or {}
        self.session = new_session()
        self.html = None
        self.lsd = None

    def first_page(self, **query):
        html = self.load(build_search_url(**query))
        connection = extract_connection(html)
        if connection is None:
            raise AdLibraryError("no search_results_connection in the Ad Library response")
        return connection, extract_root_variables(html), connection.get("count")

    def load(self, url):
        _, html = get_html(self.session, url, cookies=self.cookies,
                           referer=f"{FB_BASE}/ads/library/")
        self.html = html
        self.lsd = find_lsd(html)
        return html

    def paginate(self, variables, cursor, count=PAGE_SIZE):
        if not self.lsd:
            raise AdLibraryError("no LSD token found in the Ad Library page -- cannot paginate")
        doc_id = get_cached_doc_id(PAGINATION_OPERATION)
        payload = dict(variables)
        payload["count"] = count
        payload["cursor"] = cursor
        data = graphql_post(self.session, payload, doc_id, PAGINATION_OPERATION,
                            self.lsd, cookies=self.cookies)
        connection = _connection_from_graphql(data)
        if connection is None:
            # doc_id rotates whenever Meta ships the bundle -- rediscover once.
            fresh = discover_doc_id(self.session, self.html, PAGINATION_OPERATION,
                                    cookies=self.cookies)
            if not fresh or fresh == doc_id:
                raise AdLibraryError(f"GraphQL pagination failed and no fresh doc_id found: "
                                     f"{json.dumps(data)[:200]}")
            data = graphql_post(self.session, payload, fresh, PAGINATION_OPERATION,
                                self.lsd, cookies=self.cookies)
            connection = _connection_from_graphql(data)
            if connection is None:
                raise AdLibraryError(f"GraphQL pagination failed: {json.dumps(data)[:200]}")
        return connection

    def close(self):
        pass


class _BrowserTransport:
    """Headless Chromium via Playwright.

    Needed because /ads/library/ answers plain HTTP requests with an anti-bot
    challenge. A browser runs the page's own script and continues normally --
    the same thing that happens when you open the Ad Library yourself. All the
    parsing still happens on the payload, not on rendered pixels.
    """

    name = "browser"

    UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
          "(KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36")

    def __init__(self, cookies=None, headless=True, page_timeout_ms=60000):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            raise AdLibraryError(
                "the 'browser' transport needs Playwright: pip install playwright "
                "&& python -m playwright install chromium"
            )
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(
            headless=headless, args=["--disable-blink-features=AutomationControlled"]
        )
        self._ctx = self._browser.new_context(
            user_agent=self.UA, locale="en-US", viewport={"width": 1280, "height": 900}
        )
        if cookies:
            self._ctx.add_cookies([
                {"name": k, "value": v, "domain": ".facebook.com", "path": "/"}
                for k, v in cookies.items()
            ])
        self.page = self._ctx.new_page()
        self.timeout = page_timeout_ms
        self.html = None

    def first_page(self, **query):
        html = self.load(build_search_url(**query))
        connection = extract_connection(html)
        if connection is None:
            raise AdLibraryError("no search_results_connection in the Ad Library response")
        return connection, extract_root_variables(html), connection.get("count")

    def load(self, url):
        self.page.goto(url, wait_until="domcontentloaded", timeout=self.timeout)
        # The challenge page is ~500 bytes and reloads itself; give it a few
        # beats to land on the real thing before deciding it failed.
        deadline = time.time() + 30
        while time.time() < deadline:
            html = self.page.content()
            if '"search_results_connection"' in html:
                self.html = html
                return html
            if "__rd_verify_" not in html and len(html) > 200000:
                self.html = html
                return html
            self.page.wait_for_timeout(2000)
        raise ChallengeError(
            "Ad Library did not finish loading in the browser transport (still on Meta's "
            "challenge page after 30s). Retry, or run with headless=false to see what it wants."
        )

    def paginate(self, variables, cursor, count=PAGE_SIZE):
        """Fire the UI's own pagination query from inside the page.

        Same-origin, same cookies, same LSD token the app itself uses -- this
        is exactly the request the page makes when you scroll.
        """
        payload = dict(variables)
        payload["count"] = count
        payload["cursor"] = cursor
        doc_id = get_cached_doc_id(PAGINATION_OPERATION)
        result = self.page.evaluate(
            """async ([vars, docId, friendly]) => {
                const html = document.documentElement.innerHTML;
                const lsd = (html.match(/"LSD",\\[\\],\\{"token":"([^"]+)"/) || [])[1];
                if (!lsd) return {error: 'no lsd token in page'};
                const body = new URLSearchParams({
                    av: '0', __user: '0', __a: '1', __comet_req: '1', lsd: lsd,
                    fb_api_caller_class: 'RelayModern',
                    fb_api_req_friendly_name: friendly,
                    variables: JSON.stringify(vars),
                    server_timestamps: 'true',
                    doc_id: docId,
                });
                const r = await fetch('/api/graphql/', {
                    method: 'POST', credentials: 'include',
                    headers: {'content-type': 'application/x-www-form-urlencoded',
                              'x-fb-lsd': lsd, 'x-fb-friendly-name': friendly},
                    body,
                });
                const txt = await r.text();
                try { return {json: JSON.parse(txt.replace(/^for\\s*\\(;;\\);/, '').split('\\n')[0])}; }
                catch (e) { return {error: 'unparseable', head: txt.slice(0, 300)}; }
            }""",
            [payload, doc_id, PAGINATION_OPERATION],
        )
        if result.get("error"):
            raise AdLibraryError(f"pagination failed: {result.get('error')} {result.get('head', '')}")
        connection = _connection_from_graphql(result.get("json") or {})
        if connection is None:
            fresh = self._discover_doc_id_in_page()
            if not fresh or fresh == doc_id:
                raise AdLibraryError("pagination returned no connection and no fresh doc_id found")
            return self.paginate(variables, cursor, count)
        return connection

    def _discover_doc_id_in_page(self):
        """Re-read the doc_id out of the page's own JS bundles."""
        found = self.page.evaluate(
            """async (operation) => {
                const srcs = [...document.querySelectorAll('script[src]')].map(s => s.src);
                const re = new RegExp(operation + '_facebookRelayOperation",\\\\[\\\\],\\\\(function\\\\([^)]*\\\\)\\\\{[a-z]\\\\.exports="(\\\\d+)"');
                for (const u of srcs) {
                    let txt;
                    try { txt = await (await fetch(u)).text(); } catch (e) { continue; }
                    if (!txt.includes(operation)) continue;
                    const m = txt.match(re);
                    if (m) return m[1];
                }
                return null;
            }""",
            PAGINATION_OPERATION,
        )
        if found:
            from fb_session import _DOC_ID_CACHE, _DOC_ID_LOCK, _save_doc_id_cache
            with _DOC_ID_LOCK:
                _DOC_ID_CACHE[PAGINATION_OPERATION] = found
                _save_doc_id_cache()
        return found

    def close(self):
        for closer in (getattr(self, "_ctx", None), getattr(self, "_browser", None)):
            try:
                closer.close()
            except Exception:
                pass
        try:
            self._pw.stop()
        except Exception:
            pass


def _connection_from_graphql(data: dict):
    """Dig search_results_connection out of a GraphQL response.

    Meta has moved this around before, so walk for it instead of hardcoding
    data.ad_library_main.search_results_connection.
    """
    if not isinstance(data, dict):
        return None
    stack = [data]
    while stack:
        current = stack.pop()
        if isinstance(current, dict):
            if "search_results_connection" in current:
                return current["search_results_connection"]
            stack.extend(current.values())
        elif isinstance(current, list):
            stack.extend(current)
    return None


def _make_transport(transport, cookies, headless):
    if transport == "graphql":
        return _GraphQLTransport(cookies=cookies)
    if transport == "http":
        return _HttpTransport(cookies=cookies)
    if transport == "browser":
        return _BrowserTransport(cookies=cookies, headless=headless)
    raise AdLibraryError(f"unknown transport '{transport}' (use 'graphql', 'http' or 'browser')")


def run_ad_scan(brand=None, page_id=None, country="IN", active_status="active",
                media_type="all", ad_type="all", max_ads=2000, transport="graphql",
                cookies=None, headless=True, start_date_min=None, start_date_max=None,
                progress_cb=None) -> dict:
    """Blocking. Run in a threadpool from async code.

    Returns every ad Meta will hand over for the query, plus the completeness
    metadata that makes the number trustworthy.
    """
    query = dict(brand=brand, page_id=page_id, country=country,
                 active_status=active_status, ad_type=ad_type, media_type=media_type,
                 start_date_min=start_date_min, start_date_max=start_date_max)
    url = build_search_url(**query)

    client = _make_transport(transport, cookies, headless)
    ads = []
    seen = set()
    reported_total = None
    pages_fetched = 0
    pagination_error = None
    exhausted = False

    try:
        connection, variables, reported_total = client.first_page(**query)

        while True:
            pages_fetched += 1
            for result in iter_results(connection):
                archive_id = result.get("ad_archive_id")
                if archive_id in seen:
                    continue
                seen.add(archive_id)
                ad = normalize_ad(result)
                ad["is_self_partnership"] = _is_self_partnership(ad)
                ads.append(ad)

            if progress_cb:
                progress_cb(len(ads), reported_total)

            page_info = connection.get("page_info") or {}
            cursor = page_info.get("end_cursor")
            if not page_info.get("has_next_page") or not cursor:
                # Meta says there is nothing after this page. On the GraphQL
                # transport, which gets no total to compare against, this is
                # what "complete" means.
                exhausted = True
                break
            if len(ads) >= max_ads or pages_fetched >= MAX_PAGES_SAFETY:
                break
            if not variables:
                pagination_error = ("could not read the search's GraphQL variables from the page, "
                                    "so only the first rendered page was captured")
                break

            # Same courtesy delay every other scraper in this project uses.
            time.sleep(random.uniform(0.8, 1.6))
            try:
                connection = client.paginate(variables, cursor)
            except (AdLibraryError, ChallengeError) as exc:
                pagination_error = str(exc)
                break
    finally:
        client.close()

    partners = summarize_partners(ads)

    # Two independent ways to know a scan finished, depending on transport:
    #   cursor exhausted -- Meta said has_next_page: false (graphql + others)
    #   count match      -- we captured at least Meta's own reported total
    if pagination_error:
        complete, basis = False, "incomplete"
    elif exhausted:
        complete, basis = True, "cursor exhausted (Meta reported no further pages)"
    elif reported_total is not None and len(ads) >= reported_total:
        complete, basis = True, "captured >= Meta's reported total"
    else:
        complete, basis = False, "stopped before the cursor ran out"

    warning = pagination_error
    if warning is None and not exhausted and len(ads) >= max_ads:
        total = f"Meta reports {reported_total} ads for this query" if reported_total else \
                "there are more pages available"
        warning = f"stopped at max_ads={max_ads}; {total}. Raise max_ads for the full set."
    elif warning is None and not exhausted and pages_fetched >= MAX_PAGES_SAFETY:
        warning = (f"stopped at the {MAX_PAGES_SAFETY}-page safety limit with more pages still "
                   f"available -- this is not a full capture.")
    elif warning is None and reported_total and len(ads) < reported_total:
        gap = reported_total - len(ads)
        if gap <= max(2, reported_total * 0.02):
            warning = (f"captured {len(ads)} of the {reported_total} ads Meta reports -- a gap this "
                       f"small is normally Meta collating duplicate versions of the same ad.")
        else:
            warning = (f"captured only {len(ads)} of the {reported_total} ads Meta reports for this "
                       f"query ({gap} missing). Do not report this as a full capture.")

    return {
        "query": brand or f"page:{page_id}",
        "search_mode": "page" if page_id else "keyword",
        "page_id": str(page_id) if page_id else None,
        "country": country,
        "active_status": active_status,
        "search_url": url,
        "transport": client.name,
        "reported_total": reported_total,
        "captured": len(ads),
        "complete": complete,
        "completeness_basis": basis,
        "pages_fetched": pages_fetched,
        "branded_content_ads": sum(1 for a in ads if a["is_branded_content"]),
        "self_partnership_ads": sum(1 for a in ads if a.get("is_self_partnership")),
        "partner_count": len(partners),
        "partners": partners,
        "ads": ads,
        "warning": warning,
    }
