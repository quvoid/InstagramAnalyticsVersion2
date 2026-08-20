"""
Partnership scan API -- exposes the organic Instagram partnership scan
(Method A from this project) as a small local HTTP API, so another website
can request it instead of you running the script by hand.

Run:
    pip install -r requirements.txt
    copy .env.example .env      (then fill in API_KEY and PROJECT_DIR)
    uvicorn main:app --reload --port 8000

Call:
    curl -H "X-API-Key: <your key>" "http://localhost:8000/api/v1/partnerships/officialzivame"

Scope: this covers Method A (organic feed scan) only. Method B (Meta Ad
Library live-ad cross-check) is not included -- it depends on driving an
actual browser interactively and isn't something this kind of API can do
without a much heavier headless-browser setup (see PROJECT_CONTEXT.md).

Every scan here hits Instagram LIVE using the real session cookie in your
scrape_profiles.py, per your choice to keep this live-per-request rather
than serving cached data. The one safety net kept in is SCAN_COOLDOWN_SECONDS
below -- it only collapses two requests for the SAME brand arriving within
that window into one actual scan (protects against accidental double-calls
double-hitting Instagram at once); it does not make responses stale beyond
that window, and you can set it to 0 to disable it entirely.
"""
import os
import sys
import time
import uuid
import random
import secrets
import threading
from typing import Optional, List

from fastapi import FastAPI, HTTPException, Header, Query, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.concurrency import run_in_threadpool
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.environ.get("API_KEY", "")
PROJECT_DIR = os.environ.get("PROJECT_DIR", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
ALLOWED_ORIGINS = [o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "*").split(",")]
SCAN_COOLDOWN_SECONDS = int(os.environ.get("SCAN_COOLDOWN_SECONDS", "60"))
DEFAULT_MAX_POSTS = int(os.environ.get("DEFAULT_MAX_POSTS", "400"))

if not API_KEY:
    raise RuntimeError(
        "API_KEY is not set. Copy .env.example to .env and set a real secret key before running this API."
    )

sys.path.insert(0, PROJECT_DIR)
try:
    from scrape_profiles import COOKIES, BASE_HEADERS
except ImportError as e:
    raise RuntimeError(
        f"Could not import COOKIES/BASE_HEADERS from scrape_profiles.py in PROJECT_DIR={PROJECT_DIR}. "
        f"Set PROJECT_DIR in .env to the folder containing your scrape_profiles.py. ({e})"
    )

from scan_partnerships import run_scan, ScanError  # noqa: E402
from scan_profile_metrics import scrape_profile, ProfileScanError  # noqa: E402

BULK_MAX_USERNAMES = int(os.environ.get("BULK_MAX_USERNAMES", "2000"))
BULK_DELAY_MIN = float(os.environ.get("BULK_DELAY_MIN", "0.8"))
BULK_DELAY_MAX = float(os.environ.get("BULK_DELAY_MAX", "1.8"))

app = FastAPI(
    title="Instagram Partnership Scan API",
    description="Live organic-partnership-scan endpoint for a single project's brands.",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=["GET"],
    allow_headers=["*"],
)

# --- in-memory per-brand cooldown / lock, NOT a data cache -----------------
# Keeps two overlapping requests for the same brand from firing two live
# scans at once; still re-scrapes live on every request once the cooldown
# has passed. Cleared on process restart. Set SCAN_COOLDOWN_SECONDS=0 to
# disable entirely.
_lock = threading.Lock()
_last_result: dict = {}   # brand -> (timestamp, result)
_brand_locks: dict = {}   # brand -> threading.Lock (one scan at a time per brand)


def _get_brand_lock(brand: str) -> threading.Lock:
    with _lock:
        if brand not in _brand_locks:
            _brand_locks[brand] = threading.Lock()
        return _brand_locks[brand]


def check_api_key(x_api_key: Optional[str] = Header(default=None)):
    if not x_api_key or not secrets.compare_digest(x_api_key, API_KEY):
        raise HTTPException(status_code=401, detail="Missing or invalid X-API-Key header")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/api/v1/partnerships/{brand}")
async def get_partnerships(
    brand: str,
    max_posts: int = Query(default=DEFAULT_MAX_POSTS, ge=12, le=1000),
    x_api_key: Optional[str] = Header(default=None),
):
    """Full scan result: every detected partnership post, dates, URLs,
    captions, and the aggregated unique-partner list."""
    check_api_key(x_api_key)
    brand = brand.strip().lstrip("@").lower()

    brand_lock = _get_brand_lock(brand)
    if not brand_lock.acquire(blocking=False):
        raise HTTPException(
            status_code=429,
            detail=f"A scan for '{brand}' is already in progress -- try again shortly.",
        )
    try:
        now = time.time()
        cached = _last_result.get(brand)
        if SCAN_COOLDOWN_SECONDS > 0 and cached and (now - cached[0]) < SCAN_COOLDOWN_SECONDS:
            result = cached[1]
            result = {**result, "served_from_cooldown_window": True}
        else:
            try:
                result = await run_in_threadpool(
                    run_scan, brand, COOKIES, BASE_HEADERS, max_posts
                )
            except ScanError as e:
                raise HTTPException(status_code=502, detail=str(e))
            result["served_from_cooldown_window"] = False
            _last_result[brand] = (now, result)
        return result
    finally:
        brand_lock.release()


@app.get("/api/v1/partnerships/{brand}/usernames")
async def get_partnership_usernames(
    brand: str,
    max_posts: int = Query(default=DEFAULT_MAX_POSTS, ge=12, le=1000),
    x_api_key: Optional[str] = Header(default=None),
):
    """Just the flat list of unique partner @handles -- the format used to
    feed straight into a profile-metrics scraper."""
    full = await get_partnerships(brand, max_posts, x_api_key)
    usernames = [info["display"] for info in full["unique"].values()]
    return {
        "target": full["target"],
        "count": len(usernames),
        "usernames": usernames,
        "warning": full.get("warning"),
        "served_from_cooldown_window": full.get("served_from_cooldown_window", False),
    }


# --- profile metrics --------------------------------------------------------
# Single-username lookups are fast (a couple seconds) and run synchronously.
# Bulk lookups are NOT run synchronously -- this project's own scrape
# tonight took over 10 hours for ~1,800 accounts once Instagram's rate
# limiting kicked in, and no HTTP client/reverse-proxy will hold a request
# open that long. Bulk requests instead start a background job and return
# immediately; poll the job endpoint for progress and partial/final results.

@app.get("/api/v1/profile/{username}")
async def get_profile(
    username: str,
    posts_sample: int = Query(default=20, ge=1, le=50),
    x_api_key: Optional[str] = Header(default=None),
):
    """Single account: followers, verified/business flags, and avg
    likes/comments/ER% over its most recent posts (Instagram caps a single
    feed request at ~12 items regardless of posts_sample requested)."""
    check_api_key(x_api_key)
    try:
        return await run_in_threadpool(scrape_profile, username, COOKIES, BASE_HEADERS, posts_sample)
    except ProfileScanError as e:
        raise HTTPException(status_code=502, detail=str(e))


_bulk_lock = threading.Lock()          # only one bulk job runs at a time
_bulk_jobs: dict = {}                  # job_id -> job state dict
_bulk_active_job_id: Optional[str] = None


def _run_bulk_job(job_id: str, usernames: List[str], posts_sample: int):
    job = _bulk_jobs[job_id]
    consecutive_errors = 0
    for uname in usernames:
        try:
            profile = scrape_profile(uname, COOKIES, BASE_HEADERS, posts_sample)
            job["results"][uname] = profile
            consecutive_errors = 0
        except ProfileScanError as e:
            job["results"][uname] = {"error": str(e)}
            consecutive_errors += 1
        except Exception as e:
            job["results"][uname] = {"error": f"unexpected: {e}"}
            consecutive_errors += 1

        job["completed"] += 1
        if consecutive_errors >= 15:
            # Same heuristic as the standalone bulk-scrape script tonight --
            # 15 failures in a row almost always means we're hard rate-limited.
            job["backing_off_until"] = time.time() + 90
            time.sleep(90)
            consecutive_errors = 0
        else:
            time.sleep(random.uniform(BULK_DELAY_MIN, BULK_DELAY_MAX))

    job["status"] = "done"
    job["finished_at"] = time.time()
    global _bulk_active_job_id
    with _bulk_lock:
        if _bulk_active_job_id == job_id:
            _bulk_active_job_id = None


@app.post("/api/v1/profile/bulk")
def start_bulk_profile_job(
    usernames: List[str] = Body(..., embed=True),
    posts_sample: int = Body(default=20, embed=True),
    x_api_key: Optional[str] = Header(default=None),
):
    """Start a background bulk profile-metrics scrape. Returns a job_id
    immediately -- poll GET /api/v1/profile/bulk/{job_id} for progress.
    Only one bulk job runs at a time (protects the single IG session from
    two overlapping bulk scrapes compounding rate-limit pressure)."""
    check_api_key(x_api_key)
    if not usernames:
        raise HTTPException(status_code=400, detail="usernames must be a non-empty list")
    if len(usernames) > BULK_MAX_USERNAMES:
        raise HTTPException(status_code=400, detail=f"too many usernames (max {BULK_MAX_USERNAMES})")

    global _bulk_active_job_id
    with _bulk_lock:
        if _bulk_active_job_id is not None:
            raise HTTPException(
                status_code=429,
                detail=f"A bulk job is already running (job_id={_bulk_active_job_id}). "
                       f"Wait for it to finish or check its progress first.",
            )
        job_id = uuid.uuid4().hex[:12]
        clean_usernames = [u.strip().lstrip("@") for u in usernames if u.strip()]
        _bulk_jobs[job_id] = {
            "job_id": job_id,
            "status": "running",
            "total": len(clean_usernames),
            "completed": 0,
            "started_at": time.time(),
            "finished_at": None,
            "backing_off_until": None,
            "results": {},
        }
        _bulk_active_job_id = job_id
        t = threading.Thread(target=_run_bulk_job, args=(job_id, clean_usernames, posts_sample), daemon=True)
        t.start()

    return {"job_id": job_id, "total": len(clean_usernames), "status": "running"}


@app.get("/api/v1/profile/bulk/{job_id}")
def get_bulk_profile_job(
    job_id: str,
    include_results: bool = Query(default=True),
    x_api_key: Optional[str] = Header(default=None),
):
    """Poll a bulk job's progress. Results accumulate as the job runs --
    safe to call repeatedly; a running job returns whatever's completed so
    far, not just the final state."""
    check_api_key(x_api_key)
    job = _bulk_jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail=f"no job found with id {job_id}")
    out = {
        "job_id": job["job_id"],
        "status": job["status"],
        "total": job["total"],
        "completed": job["completed"],
        "started_at": job["started_at"],
        "finished_at": job["finished_at"],
        "currently_backing_off": bool(job["backing_off_until"] and time.time() < job["backing_off_until"]),
    }
    if include_results:
        out["results"] = job["results"]
    return out
