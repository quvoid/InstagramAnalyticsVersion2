"""
Google Sheets sink, via the deployed Apps Script web app (PostMetrics.gs).

Everything -- reads included -- goes over POST so the shared token never lands
in a URL, a referrer header, or Google's request logs.

Because Streamlit Cloud's filesystem is ephemeral, the sheet is also the only
durable store for pipeline state: which posts have been scraped, when the last
run happened, and the run lock. Nothing here is cached to disk.
"""

import json
import time

import requests

import pm_config as cfg

CHUNK_ROWS = 200      # rows per request; keeps each payload well inside limits
MAX_ATTEMPTS = 3


class SheetsError(RuntimeError):
    pass


class SheetsClient:
    def __init__(self, url=None, token=None):
        default_url, default_token = cfg.get_webhook()
        self.url = url or default_url
        self.token = token or default_token
        if not self.url:
            raise SheetsError("SHEETS_WEBHOOK_URL is not configured.")

    def _call(self, action, payload=None):
        body = {"token": self.token, "action": action}
        if payload:
            body.update(payload)

        last_error = None
        for attempt in range(1, MAX_ATTEMPTS + 1):
            try:
                # Apps Script answers /exec with a 302 to googleusercontent;
                # requests follows it, but only if the redirect stays a GET,
                # which is why the payload must be sent as the POST body.
                r = requests.post(
                    self.url,
                    data=json.dumps(body),
                    headers={"Content-Type": "application/json"},
                    timeout=120,
                    allow_redirects=True,
                )
                if r.status_code != 200:
                    raise SheetsError("Webhook HTTP %s: %s" % (r.status_code, r.text[:300]))
                try:
                    data = r.json()
                except ValueError:
                    raise SheetsError("Webhook returned non-JSON: %s" % r.text[:300])
                if not data.get("ok"):
                    raise SheetsError("Webhook error: %s" % data.get("error"))
                return data
            except Exception as e:
                last_error = e
                if attempt < MAX_ATTEMPTS:
                    time.sleep(2 * attempt)
        raise SheetsError("'%s' failed after %d attempts: %s"
                          % (action, MAX_ATTEMPTS, last_error))

    # -- state ----------------------------------------------------------------
    def ping(self):
        """Verify the webhook URL and token before doing any real work."""
        return self._call("ping")

    def get_index(self):
        """Everything the pipeline needs to decide what to scrape.

        Returns {"state": {...}, "posts": {shortcode: taken_at_epoch}}.
        The post index carries only shortcodes and timestamps, so it stays
        small even for accounts with thousands of posts.
        """
        data = self._call("index")
        return {
            "state": data.get("state") or {},
            "posts": data.get("posts") or {},
        }

    def set_state(self, values: dict):
        return self._call("set_state", {"values": values})

    # -- locking --------------------------------------------------------------
    def acquire_lock(self, owner: str, ttl_minutes: int = None):
        """True when this caller now holds the run lock.

        Guards against two browser sessions waking the app at once and both
        deciding the daily run is due.
        """
        ttl = cfg.LOCK_TTL_MINUTES if ttl_minutes is None else ttl_minutes
        data = self._call("acquire_lock", {"owner": owner, "ttl_minutes": ttl})
        return bool(data.get("acquired")), data.get("held_by") or ""

    def release_lock(self, owner: str):
        return self._call("release_lock", {"owner": owner})

    # -- writes ---------------------------------------------------------------
    def upsert_posts(self, rows, log=None):
        """Insert new posts and overwrite existing ones, keyed on shortcode."""
        return self._chunked("upsert_posts", rows, log)

    def append_history(self, rows, log=None):
        """Append one immutable metrics snapshot per post per run."""
        return self._chunked("append_history", rows, log)

    def _chunked(self, action, rows, log=None):
        say = log or (lambda m: None)
        total = 0
        for i in range(0, len(rows), CHUNK_ROWS):
            chunk = rows[i:i + CHUNK_ROWS]
            self._call(action, {"rows": chunk})
            total += len(chunk)
            if len(rows) > CHUNK_ROWS:
                say("  %s: %d/%d rows" % (action, total, len(rows)))
        return total

    def log_run(self, record: dict):
        return self._call("log_run", {"record": record})
