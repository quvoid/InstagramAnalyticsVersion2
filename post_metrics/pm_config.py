"""
Configuration + secret loading for the daily post-metrics pipeline.

Secrets are NEVER read from source files. Order of precedence:
  1. Streamlit secrets  (st.secrets)  -- how Streamlit Cloud supplies them
  2. Environment variables            -- how a local/CLI run supplies them
  3. .env file next to this script     -- local convenience only (gitignored)

Required secrets
----------------
  IG_SESSIONID    Instagram session cookie
  IG_DS_USER_ID   Instagram numeric user id that owns the session
  IG_CSRFTOKEN    Instagram csrf token
  IG_MID          Instagram mid cookie (optional but improves stability)
  SHEETS_WEBHOOK_URL    Deployed Apps Script web-app /exec URL
  SHEETS_WEBHOOK_TOKEN  Shared secret, must match TOKEN in PostMetrics.gs
"""

import os

# ── Tunables ──────────────────────────────────────────────────────────────────
TIMEZONE = "Asia/Kolkata"        # the timezone "9 AM" is interpreted in
DAILY_RUN_HOUR = 9               # local hour the daily catch-up becomes due
REFRESH_WINDOW_DAYS = 30         # re-scrape posts younger than this each run
BACKFILL_MAX_POSTS = 3000        # hard ceiling so a backfill can't run forever
DAILY_MAX_POSTS = 400            # safety cap; the refresh window normally
                                 # stops the daily pass long before this
PAGE_SLEEP_SECONDS = 0.4         # politeness delay between feed pages
REQUEST_TIMEOUT = 15
LOCK_TTL_MINUTES = 30            # a crashed run's lock self-expires after this

_DOTENV_LOADED = False


def _load_dotenv_once():
    """Read a local .env into os.environ. No-op if the file is absent."""
    global _DOTENV_LOADED
    if _DOTENV_LOADED:
        return
    _DOTENV_LOADED = True
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def get_secret(name: str, default: str = "") -> str:
    """Fetch one secret from Streamlit secrets, then env, then .env."""
    try:
        import streamlit as st
        if name in st.secrets:
            return str(st.secrets[name]).strip()
    except Exception:
        # Not running under Streamlit, or no secrets.toml present.
        pass
    _load_dotenv_once()
    return os.environ.get(name, default).strip()


def get_cookies() -> dict:
    """Instagram cookie jar assembled from secrets."""
    jar = {
        "sessionid": get_secret("IG_SESSIONID"),
        "ds_user_id": get_secret("IG_DS_USER_ID"),
        "csrftoken": get_secret("IG_CSRFTOKEN"),
        "mid": get_secret("IG_MID"),
    }
    return {k: v for k, v in jar.items() if v}


def get_webhook() -> tuple:
    """(url, token) for the Apps Script web app."""
    return get_secret("SHEETS_WEBHOOK_URL"), get_secret("SHEETS_WEBHOOK_TOKEN")


def missing_secrets() -> list:
    """Names of secrets that must be set before the pipeline can run."""
    required = ["IG_SESSIONID", "IG_DS_USER_ID", "IG_CSRFTOKEN", "SHEETS_WEBHOOK_URL"]
    return [n for n in required if not get_secret(n)]
