# Partnership Scan + Profile Metrics API

A small local HTTP API that wraps two of this project's scrapers — the
organic Instagram partnership scan (Method A), and the profile-metrics
scraper (followers, verified/business flags, avg likes/comments/ER%) — so
another website can request them instead of running the scripts by hand.

**Scope:** organic partnership scan + profile metrics only. Meta Ad Library
cross-checking (Method B) is not included — it needs an interactive browser
and isn't something this kind of API does without a much heavier
headless-browser setup. Ask if you want that added later.

## Setup

```bash
cd api
pip install -r requirements.txt
copy .env.example .env
```

Edit `.env`:
- `API_KEY` — generate a real one: `python -c "import secrets; print(secrets.token_urlsafe(32))"`
- `PROJECT_DIR` — folder containing your `scrape_profiles.py` (defaults to this project's root)

## Run

```bash
uvicorn main:app --reload --port 8000
```

## Endpoints

All require an `X-API-Key` header matching your `.env` value.

**`GET /api/v1/partnerships/{brand}`** — full scan result (every detected
partnership post: date, URL, caption, tagged partner) plus an aggregated
`unique` partner list.

```bash
curl -H "X-API-Key: your-key" "http://localhost:8000/api/v1/partnerships/officialzivame"
```

**`GET /api/v1/partnerships/{brand}/usernames`** — just the flat list of
unique partner handles, for feeding straight into a profile-metrics
scraper.

```bash
curl -H "X-API-Key: your-key" "http://localhost:8000/api/v1/partnerships/officialzivame/usernames"
```

Both accept `?max_posts=400` (default from `.env`, max 1000).

**`GET /api/v1/profile/{username}`** — single account's metrics: followers,
following, total posts, verified/business flags, and avg likes/comments/ER%
over its most recent posts. Fast (a couple seconds) — runs synchronously.

```bash
curl -H "X-API-Key: your-key" "http://localhost:8000/api/v1/profile/joeyking"
```

Accepts `?posts_sample=20` (default 20, max 50) — note Instagram caps a
single feed request at ~12 items regardless of what you ask for, so "last N
posts" in practice means "last ~12 posts". Consistent across every account,
not a bug.

**`POST /api/v1/profile/bulk`** — scrape metrics for many accounts at once.
This does **not** block until finished — a bulk run of hundreds/thousands
of accounts can take hours once Instagram's rate limiting kicks in (we saw
10+ hours for ~1,800 accounts in one run), and no HTTP client or
reverse-proxy will hold a connection open that long. It starts a background
job and returns a `job_id` immediately; poll the job endpoint for progress.
Only one bulk job runs at a time — a second `POST` while one is running
returns `429` with the running job's ID instead of starting a second one
(protects the single IG session from two overlapping bulk scrapes
compounding rate-limit pressure).

```bash
curl -X POST -H "X-API-Key: your-key" -H "Content-Type: application/json" \
  -d '{"usernames": ["joeyking", "rupertfriend", "thedanielbruhl"], "posts_sample": 20}' \
  "http://localhost:8000/api/v1/profile/bulk"
# -> {"job_id": "a1b2c3d4e5f6", "total": 3, "status": "running"}
```

**`GET /api/v1/profile/bulk/{job_id}`** — poll a bulk job. Safe to call
repeatedly; returns whatever's completed so far while `status` is
`"running"`, and everything once `status` is `"done"`. Pass
`?include_results=false` to just check progress without pulling the (large)
results payload each time.

```bash
curl -H "X-API-Key: your-key" "http://localhost:8000/api/v1/profile/bulk/a1b2c3d4e5f6?include_results=false"
# -> {"job_id": "...", "status": "running", "total": 1795, "completed": 812, "currently_backing_off": false, ...}
```

`GET /health` — no auth, returns `{"status": "ok"}`, useful for uptime checks.

## What "live" means here

Every call actually scans Instagram in real time using the session cookie
in your `scrape_profiles.py` — nothing is pre-scraped or served from a
database. A full scan takes roughly 1-2 minutes depending on `max_posts`
(the scraper deliberately sleeps between paginated requests to avoid
tripping Instagram's rate limiting).

The one exception: if two requests for the *same brand* land within
`SCAN_COOLDOWN_SECONDS` (default 60s) of each other, the second one reuses
the first's result instead of starting a second overlapping scan — this
isn't a data cache, it just stops accidental double-calls from hitting
Instagram twice at once. Set it to `0` in `.env` to turn this off entirely.

## Before this is reachable by anything other than you

- Rotate the Instagram session in `scrape_profiles.py` if you haven't
  since it was pushed to GitHub — see `PROJECT_CONTEXT.md`. An API pointed
  at that session makes an exposed credential more consequential, not less.
- Tighten `ALLOWED_ORIGINS` in `.env` to your actual website's domain
  instead of `*`.
- Don't commit `.env` (already gitignored) — it holds your API key.
