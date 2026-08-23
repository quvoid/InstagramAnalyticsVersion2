"""
Facebook scan API -- the Facebook-side counterpart to ../api/ (the Instagram
partnership + profile-metrics API), built the same way so a website can call
either one the same way.

Run:
    cd fb_api
    pip install -r requirements.txt
    copy .env.example .env      (then set API_KEY and FB_COOKIE)
    uvicorn main:app --reload --port 8001

Call:
    curl -H "X-API-Key: <your key>" "http://localhost:8001/api/v1/session/check"
    curl -H "X-API-Key: <your key>" "http://localhost:8001/api/v1/page/zivame"
    curl -H "X-API-Key: <your key>" "http://localhost:8001/api/v1/adlibrary/zivame?page_id=234603919914240"

Three surfaces, and they behave very differently:

  /api/v1/page/*       Public Page data. No cookies, no token, ~2s per page.
  /api/v1/adlibrary/*  Meta Ad Library, via your FB session cookie. No browser:
                       it queries Meta's GraphQL endpoint directly, because the
                       /ads/library/ HTML route serves an anti-bot challenge
                       even to logged-in sessions. ~9s for a 70-ad brand, ~36s
                       for a 380-ad one.
  /api/v1/official/*   Meta's sanctioned Graph API, token required. Read the
                       header comment in graph_official.py before using it --
                       outside the EU it only carries political ads.

Your session goes in fb_api/.env as FB_COOKIE (or FB_COOKIE_FILE). Check it
with /api/v1/session/check before running a batch.

Like the Instagram API, everything here hits Facebook LIVE per request. The
only caching is a short per-query cooldown that collapses accidental
double-calls into one scan; set SCAN_COOLDOWN_SECONDS=0 to disable it.
"""
import os
import time
import uuid
import random
import secrets
import functools
import threading
from typing import List, Optional

from fastapi import Body, FastAPI, Header, HTTPException, Query
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

load_dotenv()

from fb_session import (  # noqa: E402
    ChallengeError,
    FacebookError,
    PageScanError,
    describe_cookies,
    load_cookies,
)
from scan_page import scrape_page, scrape_page_posts  # noqa: E402
from scan_adlibrary import run_ad_scan  # noqa: E402
import graph_official  # noqa: E402

API_KEY = os.environ.get("API_KEY", "")
ALLOWED_ORIGINS = [o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "*").split(",")]
SCAN_COOLDOWN_SECONDS = int(os.environ.get("SCAN_COOLDOWN_SECONDS", "60"))

# Your Facebook session lives in fb_api/.env -- either FB_COOKIE (the cookie
# header on one line) or FB_COOKIE_FILE (a path to a file holding it, which is
# easier to live with because the header is long). Nothing else reads it, and
# .env is gitignored.
FB_COOKIES = load_cookies(os.environ.get("FB_COOKIE", ""), os.environ.get("FB_COOKIE_FILE", ""))

DEFAULT_COUNTRY = os.environ.get("DEFAULT_COUNTRY", "IN")
DEFAULT_MAX_ADS = int(os.environ.get("DEFAULT_MAX_ADS", "2000"))
# Cookies by default. 'graphql' queries Meta's GraphQL endpoint directly with
# FB_COOKIES and never loads the /ads/library/ HTML route -- which matters,
# because that route serves an anti-bot challenge even to logged-in sessions.
# 'http' (same cookies, HTML route) and 'browser' (Playwright) stay available.
AD_TRANSPORT = os.environ.get("AD_TRANSPORT", "graphql")
HEADLESS = os.environ.get("HEADLESS", "true").lower() != "false"

BULK_MAX_ITEMS = int(os.environ.get("BULK_MAX_ITEMS", "500"))
BULK_DELAY_MIN = float(os.environ.get("BULK_DELAY_MIN", "0.8"))
BULK_DELAY_MAX = float(os.environ.get("BULK_DELAY_MAX", "1.8"))

if not API_KEY:
    raise RuntimeError(
        "API_KEY is not set. Copy .env.example to .env and set a real secret key before running this API."
    )

app = FastAPI(
    title="Facebook Page + Ad Library API",
    description="Live Facebook Page metrics and Meta Ad Library scans for this project's brands.",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


def check_api_key(x_api_key: Optional[str]):
    if not x_api_key or not secrets.compare_digest(x_api_key, API_KEY):
        raise HTTPException(status_code=401, detail="Missing or invalid X-API-Key header")


def _as_http_error(exc: Exception):
    """Map scraper failures onto status codes a caller can act on."""
    if isinstance(exc, ChallengeError):
        # 503: not the caller's fault and worth retrying / switching transport.
        return HTTPException(status_code=503, detail=str(exc))
    if isinstance(exc, PageScanError):
        return HTTPException(status_code=404, detail=str(exc))
    if isinstance(exc, FacebookError):
        return HTTPException(status_code=502, detail=str(exc))
    return HTTPException(status_code=500, detail=f"unexpected: {exc}")


# --- per-query cooldown, NOT a data cache ----------------------------------
# Identical idea to the Instagram API: stops two overlapping requests for the
# same thing from firing two live scans at once. An Ad Library scan spawns a
# browser, so this matters more here than it does over there.
_lock = threading.Lock()
_last_result: dict = {}
_key_locks: dict = {}


def _get_key_lock(key: str) -> threading.Lock:
    with _lock:
        if key not in _key_locks:
            _key_locks[key] = threading.Lock()
        return _key_locks[key]


def _cooldown_hit(key: str):
    cached = _last_result.get(key)
    if SCAN_COOLDOWN_SECONDS > 0 and cached and (time.time() - cached[0]) < SCAN_COOLDOWN_SECONDS:
        return {**cached[1], "served_from_cooldown_window": True}
    return None


def _remember(key: str, result: dict):
    _last_result[key] = (time.time(), result)
    return {**result, "served_from_cooldown_window": False}


# Ad Library scans each drive their own Chromium. Running several at once is a
# quick way to eat the machine's RAM, so they queue.
_ad_scan_semaphore = threading.Semaphore(int(os.environ.get("MAX_CONCURRENT_AD_SCANS", "1")))


@app.get("/health")
def health():
    return {"status": "ok", "ad_transport": AD_TRANSPORT, "cookies_configured": bool(FB_COOKIES)}


@app.get("/api/v1/session/check")
async def session_check(x_api_key: Optional[str] = Header(default=None)):
    """Is the configured Facebook session actually usable? Call this first.

    Reports which cookies loaded (names only, never values), whether they add
    up to a logged-in session, and then does the one thing that really settles
    it: a live Ad Library request over the cookie transport. If Meta
    challenges it, the session is anonymous or stale and every /adlibrary/*
    call will fail the same way.
    """
    check_api_key(x_api_key)
    summary = describe_cookies(FB_COOKIES)
    summary["source"] = (
        "FB_COOKIE_FILE" if os.environ.get("FB_COOKIE_FILE")
        else ("FB_COOKIE" if os.environ.get("FB_COOKIE") else "none configured")
    )
    summary["ad_transport"] = AD_TRANSPORT

    def _probe():
        # Cheapest real check available: one Ad Library page, no pagination.
        return run_ad_scan(brand="nike", country=DEFAULT_COUNTRY, max_ads=1,
                           transport=AD_TRANSPORT if AD_TRANSPORT != "browser" else "graphql",
                           cookies=FB_COOKIES)

    try:
        probe = await run_in_threadpool(_probe)
        summary["ad_library_over_cookies"] = "ok"
        summary["probe"] = {"captured": probe["captured"], "transport": probe["transport"]}
    except ChallengeError as exc:
        summary["ad_library_over_cookies"] = "challenged"
        summary["detail"] = str(exc)
    except Exception as exc:
        summary["ad_library_over_cookies"] = "error"
        summary["detail"] = str(exc)
    return summary


# --- Pages ------------------------------------------------------------------

@app.get("/api/v1/page/{handle}")
async def get_page(
    handle: str,
    check_verified: bool = Query(
        default=True,
        description="Make a second, logged-in request to read the blue-badge state. "
                    "Needs FB_COOKIE. Set false to halve the request count in bulk work.",
    ),
    x_api_key: Optional[str] = Header(default=None),
):
    """A brand's public Page: ids, category, likes, talking-about, follower
    estimate, bio.

    The `ad_library_page_id` in the response is the id to hand to
    /api/v1/adlibrary/{brand}?page_id=... for an exact ad scan -- much cleaner
    than a keyword search, which drags in unrelated advertisers.

    Followers are APPROXIMATE (Facebook renders "794K", never the exact
    number, to logged-out callers). Likes and talking-about are exact.
    """
    check_api_key(x_api_key)
    handle = handle.strip().lstrip("@")
    key = f"page:{handle.lower()}:{check_verified}"

    lock = _get_key_lock(key)
    if not lock.acquire(blocking=False):
        raise HTTPException(status_code=429, detail=f"a scan for '{handle}' is already running")
    try:
        hit = _cooldown_hit(key)
        if hit:
            return hit
        try:
            result = await run_in_threadpool(
                functools.partial(scrape_page, handle, FB_COOKIES, check_verified=check_verified))
        except Exception as exc:
            raise _as_http_error(exc)
        return _remember(key, result)
    finally:
        lock.release()


@app.get("/api/v1/page/{handle}/posts")
async def get_page_posts(
    handle: str,
    limit: int = Query(default=12, ge=1, le=50),
    followers_hint: Optional[int] = Query(
        default=None,
        description="Follower count to compute engagement rate against. Omit and the "
                    "endpoint uses the page's own rounded follower estimate.",
    ),
    x_api_key: Optional[str] = Header(default=None),
):
    """Recent content + engagement for a Page.

    Honest scope: Facebook server-renders ONE timeline story to logged-out
    callers, and about six videos (with play and reaction counts) on the
    /videos/ tab. This endpoint returns the union of those, newest first. It
    is a sample, not a feed -- there is no anonymous equivalent of the
    Instagram side's 12-posts-per-request.
    """
    check_api_key(x_api_key)
    handle = handle.strip().lstrip("@")
    key = f"posts:{handle.lower()}:{limit}"

    lock = _get_key_lock(key)
    if not lock.acquire(blocking=False):
        raise HTTPException(status_code=429, detail=f"a scan for '{handle}' is already running")
    try:
        hit = _cooldown_hit(key)
        if hit:
            return hit

        followers = followers_hint
        if followers is None:
            try:
                followers = (await run_in_threadpool(scrape_page, handle, FB_COOKIES)).get("followers_approx")
            except Exception:
                followers = None
        try:
            result = await run_in_threadpool(scrape_page_posts, handle, limit, FB_COOKIES, followers)
        except Exception as exc:
            raise _as_http_error(exc)
        return _remember(key, result)
    finally:
        lock.release()


# --- Ad Library -------------------------------------------------------------

@app.get("/api/v1/adlibrary/{brand}")
async def get_ads(
    brand: str,
    page_id: Optional[str] = Query(
        default=None,
        description="Facebook Page id (from /api/v1/page/{handle} -> ad_library_page_id). "
                    "Strongly preferred: exact, no cross-advertiser noise.",
    ),
    country: str = Query(default=DEFAULT_COUNTRY, min_length=2, max_length=2),
    active_status: str = Query(default="active", pattern="^(active|inactive|all)$"),
    media_type: str = Query(default="all", pattern="^(all|image|meme|video|none)$"),
    max_ads: int = Query(default=DEFAULT_MAX_ADS, ge=1, le=10000),
    include_ads: bool = Query(default=True, description="Set false for just the partner summary."),
    transport: Optional[str] = Query(
        default=None, pattern="^(graphql|http|browser)$",
        description="Override AD_TRANSPORT for this call: 'graphql' (cookies, no browser), "
                    "'http' (cookies against the HTML route, usually challenged), or "
                    "'browser' (Playwright, no cookies needed).",
    ),
    x_api_key: Optional[str] = Header(default=None),
):
    """Every ad Meta will show for a brand, with completeness metadata.

    The response always carries `reported_total` (Meta's own count for the
    query), `captured`, and `complete`. Use them. The old scroll-the-UI method
    this replaces produced a false plateau that undercounted one brand by
    roughly half while looking finished -- see PROJECT_CONTEXT.md §4.

    `partners` lists the creators behind branded-content ads, taken from the
    structured branded_content field rather than by pattern-matching an
    advertiser label. Ads where the brand tagged itself are excluded from
    `partners` and counted in `self_partnership_ads`.
    """
    check_api_key(x_api_key)
    brand = brand.strip().lstrip("@")
    transport = transport or AD_TRANSPORT
    key = f"ads:{page_id or brand.lower()}:{country}:{active_status}:{media_type}:{max_ads}"

    lock = _get_key_lock(key)
    if not lock.acquire(blocking=False):
        raise HTTPException(status_code=429, detail=f"an Ad Library scan for '{brand}' is already running")
    try:
        hit = _cooldown_hit(key)
        if hit:
            return hit if include_ads else {**hit, "ads": []}

        def _scan():
            # The semaphore only exists to stop several Chromiums running at
            # once; the cookie transport is just HTTP and doesn't need it.
            if transport == "browser":
                with _ad_scan_semaphore:
                    return _run(transport)
            return _run(transport)

        def _run(which):
            return run_ad_scan(
                brand=None if page_id else brand,
                page_id=page_id,
                country=country.upper(),
                active_status=active_status,
                media_type=media_type,
                max_ads=max_ads,
                transport=which,
                cookies=FB_COOKIES,
                headless=HEADLESS,
            )

        try:
            result = await run_in_threadpool(_scan)
        except Exception as exc:
            raise _as_http_error(exc)
        stored = _remember(key, result)
        return stored if include_ads else {**stored, "ads": []}
    finally:
        lock.release()


@app.get("/api/v1/adlibrary/{brand}/partners")
async def get_ad_partners(
    brand: str,
    page_id: Optional[str] = Query(default=None),
    country: str = Query(default=DEFAULT_COUNTRY, min_length=2, max_length=2),
    active_status: str = Query(default="active", pattern="^(active|inactive|all)$"),
    max_ads: int = Query(default=DEFAULT_MAX_ADS, ge=1, le=10000),
    transport: Optional[str] = Query(default=None, pattern="^(graphql|http|browser)$"),
    x_api_key: Optional[str] = Header(default=None),
):
    """Just the creator partners -- the Facebook analogue of the Instagram
    API's /partnerships/{brand}/usernames."""
    full = await get_ads(
        brand=brand, page_id=page_id, country=country, active_status=active_status,
        media_type="all", max_ads=max_ads, include_ads=False, transport=transport,
        x_api_key=x_api_key,
    )
    return {
        "query": full["query"],
        "country": full["country"],
        "reported_total_ads": full["reported_total"],
        "captured_ads": full["captured"],
        "complete": full["complete"],
        "completeness_basis": full["completeness_basis"],
        "branded_content_ads": full["branded_content_ads"],
        "self_partnership_ads": full["self_partnership_ads"],
        "partner_count": full["partner_count"],
        "partners": full["partners"],
        "warning": full.get("warning"),
        "served_from_cooldown_window": full.get("served_from_cooldown_window", False),
    }


# --- background bulk jobs ---------------------------------------------------
# Same reasoning as the Instagram API: a long list can't be served inside one
# HTTP request. Page lookups are ~2s each, Ad Library scans 10-60s each, so a
# few hundred brands is comfortably past any proxy's patience. Start a job,
# poll it.

_bulk_lock = threading.Lock()
_bulk_jobs: dict = {}
_bulk_active_job_id: Optional[str] = None


def _start_bulk(kind: str, items: List[str], worker, extra: dict):
    global _bulk_active_job_id
    with _bulk_lock:
        if _bulk_active_job_id is not None:
            raise HTTPException(
                status_code=429,
                detail=f"a bulk job is already running (job_id={_bulk_active_job_id})",
            )
        job_id = uuid.uuid4().hex[:12]
        _bulk_jobs[job_id] = {
            "job_id": job_id,
            "kind": kind,
            "status": "running",
            "total": len(items),
            "completed": 0,
            "started_at": time.time(),
            "finished_at": None,
            "results": {},
        }
        _bulk_active_job_id = job_id
        thread = threading.Thread(target=_run_bulk, args=(job_id, items, worker, extra), daemon=True)
        thread.start()
    return job_id


def _run_bulk(job_id: str, items: List[str], worker, extra: dict):
    global _bulk_active_job_id
    job = _bulk_jobs[job_id]
    for item in items:
        try:
            job["results"][item] = worker(item, extra)
        except FacebookError as exc:
            job["results"][item] = {"error": str(exc)}
        except Exception as exc:
            job["results"][item] = {"error": f"unexpected: {exc}"}
        job["completed"] += 1
        time.sleep(random.uniform(BULK_DELAY_MIN, BULK_DELAY_MAX))
    job["status"] = "done"
    job["finished_at"] = time.time()
    with _bulk_lock:
        if _bulk_active_job_id == job_id:
            _bulk_active_job_id = None


@app.post("/api/v1/page/bulk")
def start_bulk_pages(
    handles: List[str] = Body(..., embed=True),
    include_posts: bool = Body(default=False, embed=True),
    x_api_key: Optional[str] = Header(default=None),
):
    """Metrics for many Pages. Returns a job_id immediately."""
    check_api_key(x_api_key)
    handles = [h.strip().lstrip("@") for h in handles if h and h.strip()]
    if not handles:
        raise HTTPException(status_code=400, detail="handles must be a non-empty list")
    if len(handles) > BULK_MAX_ITEMS:
        raise HTTPException(status_code=400, detail=f"too many handles (max {BULK_MAX_ITEMS})")

    def worker(handle, extra):
        page = scrape_page(handle, FB_COOKIES)
        if extra.get("include_posts"):
            try:
                page["recent"] = scrape_page_posts(handle, 12, FB_COOKIES, page.get("followers_approx"))
            except FacebookError as exc:
                page["recent"] = {"error": str(exc)}
        return page

    job_id = _start_bulk("page", handles, worker, {"include_posts": include_posts})
    return {"job_id": job_id, "total": len(handles), "status": "running"}


@app.post("/api/v1/adlibrary/bulk")
def start_bulk_adlibrary(
    brands: List[str] = Body(..., embed=True),
    country: str = Body(default=DEFAULT_COUNTRY, embed=True),
    active_status: str = Body(default="active", embed=True),
    max_ads: int = Body(default=DEFAULT_MAX_ADS, embed=True),
    include_ads: bool = Body(default=False, embed=True),
    x_api_key: Optional[str] = Header(default=None),
):
    """Ad Library scans for many brands.

    Each entry may be a keyword or a numeric Page id -- a value that's all
    digits is treated as a Page id (exact), anything else as a keyword search.
    `include_ads=false` (the default) keeps the job payload to the partner
    summaries, which is usually what you want for a few hundred brands.
    """
    check_api_key(x_api_key)
    brands = [b.strip().lstrip("@") for b in brands if b and b.strip()]
    if not brands:
        raise HTTPException(status_code=400, detail="brands must be a non-empty list")
    if len(brands) > BULK_MAX_ITEMS:
        raise HTTPException(status_code=400, detail=f"too many brands (max {BULK_MAX_ITEMS})")

    def worker(brand, extra):
        is_page_id = brand.isdigit()
        scan = functools.partial(
            run_ad_scan,
            brand=None if is_page_id else brand,
            page_id=brand if is_page_id else None,
            country=extra["country"].upper(),
            active_status=extra["active_status"],
            max_ads=extra["max_ads"],
            transport=AD_TRANSPORT,
            cookies=FB_COOKIES,
            headless=HEADLESS,
        )
        if AD_TRANSPORT == "browser":
            with _ad_scan_semaphore:
                result = scan()
        else:
            result = scan()
        if not extra["include_ads"]:
            result.pop("ads", None)
        return result

    job_id = _start_bulk("adlibrary", brands, worker, {
        "country": country, "active_status": active_status,
        "max_ads": max_ads, "include_ads": include_ads,
    })
    return {"job_id": job_id, "total": len(brands), "status": "running"}


@app.get("/api/v1/jobs/{job_id}")
def get_job(
    job_id: str,
    include_results: bool = Query(default=True),
    x_api_key: Optional[str] = Header(default=None),
):
    """Poll any bulk job. Safe to call repeatedly; a running job returns
    whatever has finished so far."""
    check_api_key(x_api_key)
    job = _bulk_jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail=f"no job found with id {job_id}")
    out = {k: v for k, v in job.items() if k != "results"}
    if include_results:
        out["results"] = job["results"]
    return out


# --- official Graph API -----------------------------------------------------

@app.get("/api/v1/official/ads")
async def official_ads(
    search_terms: Optional[str] = Query(default=None),
    page_ids: Optional[str] = Query(default=None, description="Comma-separated Page ids, max 10."),
    countries: str = Query(default=DEFAULT_COUNTRY, description="Comma-separated ISO codes."),
    ad_type: str = Query(default="ALL"),
    ad_active_status: str = Query(default="ACTIVE"),
    limit: int = Query(default=100, ge=1, le=500),
    max_pages: int = Query(default=10, ge=1, le=100),
    x_api_key: Optional[str] = Header(default=None),
):
    """Meta's sanctioned ads_archive endpoint (needs FB_GRAPH_TOKEN).

    Outside the EU/UK this only returns political and social-issue ads -- that
    is Meta's documented scope, not a limitation of this wrapper. For Indian
    commercial brands use /api/v1/adlibrary/{brand} instead.
    """
    check_api_key(x_api_key)
    call = functools.partial(
        graph_official.ads_archive,
        search_terms=search_terms,
        search_page_ids=[p.strip() for p in page_ids.split(",")] if page_ids else None,
        countries=[c.strip().upper() for c in countries.split(",")],
        ad_type=ad_type,
        ad_active_status=ad_active_status,
        limit=limit,
        max_pages=max_pages,
    )
    try:
        return await run_in_threadpool(call)
    except Exception as exc:
        raise _as_http_error(exc)


@app.get("/api/v1/official/page/{page_id}")
async def official_page(
    page_id: str,
    fields: Optional[str] = Query(default=None, description="Comma-separated Page fields."),
    x_api_key: Optional[str] = Header(default=None),
):
    """Page node via the Graph API (needs FB_GRAPH_TOKEN, and app review for
    most public-page fields). /api/v1/page/{handle} needs neither."""
    check_api_key(x_api_key)
    call = functools.partial(
        graph_official.page_fields,
        page_id,
        fields=[f.strip() for f in fields.split(",")] if fields else None,
    )
    try:
        return await run_in_threadpool(call)
    except Exception as exc:
        raise _as_http_error(exc)
