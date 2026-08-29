# Daily Post Metrics → Google Sheets

Tracks one Instagram account. Backfills its history once, then every morning at
09:00 IST picks up posts it has not seen and refreshes metrics on anything
posted in the last 30 days.

Self-contained in `post_metrics/`. It shares nothing with the older scripts in
the repo root — no imports, no config files, no output files in common — so
either side can change without touching the other.

## Files

All paths below are relative to `post_metrics/`.

| File | Role |
|---|---|
| `pm_config.py` | Tunables and secret loading. No secret is ever read from source. |
| `pm_instagram.py` | Fetches posts + metrics for one handle (timeline feed + clips feed, merged). |
| `pm_sheets.py` | Talks to the Apps Script web app. All traffic is POST. |
| `pm_ingest.py` | Orchestration: backfill / daily, dedupe, locking, run log. Also a CLI. |
| `pm_app.py` | Streamlit control panel — status, config, manual triggers. |
| `google_apps_script/PostMetrics.gs` | Sheets backend **and** the 09:00 clock. |
| `requirements.txt` | Dependencies for this pipeline only. |

## Why the clock lives in Apps Script

Streamlit Cloud has no scheduler and sleeps idle apps, so an in-process timer
(what `cron_scheduler.py` does today) stops firing without warning. Instead a
time-driven Apps Script trigger wakes the app at 09:00, and the app checks the
sheet for whether today's run is outstanding.

The check is keyed on the **local calendar date**, not a timer. So if a wake-up
does not land, the run is still outstanding and happens the next time the app is
opened. No day is skipped, and no day runs twice.

Streamlit Cloud's filesystem is also wiped on every restart and redeploy, which
is why all pipeline state — the shortcode index, last run date, the lock — lives
in the sheet rather than in local JSON.

## Setup

**1. Sheet + Apps Script**

Create a spreadsheet, then **Extensions → Apps Script**, paste
`post_metrics/google_apps_script/PostMetrics.gs` — note this is a *different*
script from the repo's older `google_apps_script/Code.gs`; use a separate
spreadsheet unless you want both writing to the same one. Edit the three
constants at the top:

- `TOKEN` — a long random string
- `STREAMLIT_APP_URL` — your deployed app URL
- `RUN_HOUR` / `TIMEZONE` — if 09:00 IST is not what you want

Then **Deploy → New deployment → Web app**, *Execute as: Me*, *Who has access:
Anyone*. Copy the `/exec` URL.

Reload the sheet and use the **Post Metrics** menu: *Initialise sheets*, then
*Schedule daily wake-up*.

**2. Secrets**

Streamlit Cloud: **Manage app → Settings → Secrets**. Locally: a `.env` file
inside `post_metrics/` (gitignored).

```toml
IG_SESSIONID = "..."
IG_DS_USER_ID = "..."
IG_CSRFTOKEN = "..."
IG_MID = "..."
SHEETS_WEBHOOK_URL = "https://script.google.com/macros/s/AKfy.../exec"
SHEETS_WEBHOOK_TOKEN = "same value as TOKEN in PostMetrics.gs"
```

Get the cookie values from Chrome → F12 → Application → Cookies → instagram.com.

**3. Backfill once, from your own machine**

Run the backfill locally rather than on Streamlit Cloud — a long history can
outlast a cloud request, and your residential IP is far less likely to trip
Instagram's rate limits than a datacenter one.

```bash
cd post_metrics && pip install -r requirements.txt && python pm_ingest.py --handle https://www.instagram.com/<account>/ --mode backfill
```

**4. Deploy**

Point Streamlit Cloud at `post_metrics/pm_app.py` as the main file. Streamlit
puts the app file's own directory on `sys.path`, so the `pm_*` imports resolve
without any packaging.

One deployment wrinkle: Streamlit Cloud looks for `requirements.txt` at the
**repo root**. The root file here belongs to the older dashboard and does not
list `pandas` or `tzdata`. If the deploy fails on a missing module, either add
those two lines to the root `requirements.txt`, or point the deployment at a
branch where `post_metrics/` is the repo root.

## Sheet tabs

| Tab | Contents |
|---|---|
| `Posts` | One row per post, latest metrics. Keyed on shortcode, so reruns overwrite rather than duplicate. |
| `Metrics History` | One snapshot per post per run — the engagement growth curve. |
| `Run Log` | One row per run: mode, counts, duration, error. |
| `State` | Handle, last run date, backfill flag, run lock. |

## Metrics captured

Likes, comments, views/plays, video duration, media kind (Photo / Carousel /
Reel / IGTV), caption, post URL, owner, co-authors, paid-partnership flag,
follower count at scrape time, and computed engagement rate.

These are the public metrics. Reach, impressions and saves are not available
this way — they need the Graph API on an account you own.

## Known limits

- **The wake-up is best-effort.** An HTTP GET reliably wakes a sleeping
  container, but Streamlit only runs the app script once a browser session
  connects. The date-keyed due check is what makes that safe.
- **The session cookie expires**, typically in weeks. When it does, the run log
  records an auth error and the app shows it. Re-paste `IG_SESSIONID`.
- **Backfill on Streamlit Cloud can time out** on large accounts. Run it locally.
- **`../scrape_bulk.py:31` still contains a live hardcoded `sessionid`** and
  that file is committed. This pipeline does not import it, but that credential
  should be rotated and moved into secrets.

## CLI

```bash
cd post_metrics && python pm_ingest.py --mode daily
```

Handy if you later decide to move the schedule off Streamlit entirely — point
Windows Task Scheduler at that command and the app becomes a pure dashboard.
