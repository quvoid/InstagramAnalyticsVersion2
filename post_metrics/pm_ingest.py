"""
Orchestration for the daily post-metrics run.

Two modes, one code path:

  backfill  One-off. Walks the whole account history (up to BACKFILL_MAX_POSTS)
            and seeds the sheet.
  daily     Everything newer than the refresh cutoff. That single pass covers
            both jobs at once -- posts the sheet has never seen get inserted,
            and posts younger than REFRESH_WINDOW_DAYS get their metrics
            rewritten, because engagement on a post keeps climbing for days
            after it goes up.

Every run also appends an immutable snapshot per post to the history tab, so
growth over time stays visible even though the main tab only holds the latest
value.

The run is idempotent: shortcode is the key, so running it twice in a day
changes nothing except adding a second history snapshot.

CLI:
    python pm_ingest.py --handle instagram --mode backfill
    python pm_ingest.py --mode daily
"""

import argparse
import socket
import sys
import traceback
import uuid
from datetime import datetime, timezone

import pm_config as cfg
from pm_instagram import PostMetricsCollector, parse_handle, refresh_cutoff
from pm_sheets import SheetsClient

try:
    from zoneinfo import ZoneInfo
    _TZ = ZoneInfo(cfg.TIMEZONE)
except Exception:
    # Windows without tzdata installed -- fall back to a fixed IST offset.
    from datetime import timedelta
    _TZ = timezone(timedelta(hours=5, minutes=30))


def local_now():
    return datetime.now(_TZ)


def today_key():
    """The local calendar date the daily run is keyed on."""
    return local_now().strftime("%Y-%m-%d")


def is_due(state: dict, now=None) -> bool:
    """True when today's scheduled run has not happened yet.

    Being date-keyed rather than timer-based is what makes missed days
    self-healing: if the app was asleep at 9 AM, the next wake-up -- whenever
    it lands -- still sees today's run as outstanding and performs it.
    """
    now = now or local_now()
    if str(state.get("enabled", "TRUE")).upper() == "FALSE":
        return False
    if now.hour < cfg.DAILY_RUN_HOUR:
        return False
    return (state.get("last_run_date") or "") != now.strftime("%Y-%m-%d")


def run(handle=None, mode="daily", log=None, client=None):
    """Execute one ingest. Returns a summary dict."""
    say = log or (lambda m: print(m, flush=True))
    started = datetime.now(timezone.utc)
    owner = "%s:%s" % (socket.gethostname()[:20], uuid.uuid4().hex[:8])

    sheets = client or SheetsClient()
    index = sheets.get_index()
    state = index["state"]
    known = index["posts"]

    handle = parse_handle(handle or state.get("handle") or "")
    if not handle:
        raise ValueError("No Instagram handle configured. Set one in the app first.")

    acquired, held_by = sheets.acquire_lock(owner)
    if not acquired:
        say("Another run is already in progress (%s). Skipping." % held_by)
        return {"status": "skipped", "reason": "locked", "held_by": held_by}

    summary = {
        "run_id": owner,
        "mode": mode,
        "handle": handle,
        "status": "error",
        "new_posts": 0,
        "refreshed_posts": 0,
        "scanned_posts": 0,
        "started_at": started.isoformat(),
    }

    try:
        if mode == "backfill":
            max_posts, since = cfg.BACKFILL_MAX_POSTS, None
            say("Backfilling @%s (up to %d posts)..." % (handle, max_posts))
        else:
            max_posts = cfg.DAILY_MAX_POSTS
            since = refresh_cutoff()
            say("Daily run for @%s -- posts since %s"
                % (handle, since.strftime("%Y-%m-%d")))

        collector = PostMetricsCollector()
        profile, posts = collector.collect(handle, max_posts=max_posts,
                                           since=since, log=say)

        new_rows = [p for p in posts if p["shortcode"] not in known]
        refreshed_rows = [p for p in posts if p["shortcode"] in known]
        summary["new_posts"] = len(new_rows)
        summary["refreshed_posts"] = len(refreshed_rows)
        summary["scanned_posts"] = len(posts)
        summary["follower_count"] = profile.get("follower_count", 0)
        say("Scanned %d posts -- %d new, %d refreshed"
            % (len(posts), len(new_rows), len(refreshed_rows)))

        if posts:
            sheets.upsert_posts(posts, log=say)
            sheets.append_history(
                [_history_row(p, summary["run_id"]) for p in posts], log=say)
        else:
            say("Nothing to write.")

        sheets.set_state({
            "handle": handle,
            "user_id": profile.get("user_id", ""),
            "follower_count": profile.get("follower_count", 0),
            "last_run_date": today_key(),
            "last_run_at": datetime.now(timezone.utc).isoformat(),
            "last_run_mode": mode,
            "last_run_status": "ok",
            "last_error": "",
            "total_posts_tracked": len(known) + len(new_rows),
            "backfill_done": "TRUE" if (mode == "backfill"
                                        or state.get("backfill_done") == "TRUE")
                             else "FALSE",
        })
        summary["status"] = "ok"

    except Exception as e:
        summary["error"] = str(e)
        say("FAILED: %s" % e)
        traceback.print_exc()
        try:
            sheets.set_state({
                "last_run_status": "error",
                "last_error": str(e)[:500],
                "last_run_at": datetime.now(timezone.utc).isoformat(),
            })
        except Exception:
            pass
    finally:
        summary["finished_at"] = datetime.now(timezone.utc).isoformat()
        summary["duration_seconds"] = round(
            (datetime.now(timezone.utc) - started).total_seconds(), 1)
        try:
            sheets.log_run(summary)
        except Exception as e:
            say("Could not write the run log: %s" % e)
        try:
            sheets.release_lock(owner)
        except Exception:
            pass

    return summary


def _history_row(post: dict, run_id: str) -> dict:
    """One point on a post's engagement curve."""
    return {
        "run_id": run_id,
        "snapshot_date": today_key(),
        "snapshot_at": post["scraped_at"],
        "shortcode": post["shortcode"],
        "handle": post["handle"],
        "posted_at": post["posted_at"],
        "like_count": post["like_count"],
        "comment_count": post["comment_count"],
        "view_count": post["view_count"],
        "follower_count": post["follower_count_at_scrape"],
        "engagement_rate_pct": post["engagement_rate_pct"],
    }


def run_if_due(log=None, client=None):
    """The entry point the app calls on every page load."""
    say = log or (lambda m: print(m, flush=True))
    sheets = client or SheetsClient()
    state = sheets.get_index()["state"]
    if not is_due(state):
        return {"status": "not_due", "last_run_date": state.get("last_run_date", "")}
    mode = "backfill" if state.get("backfill_done") != "TRUE" else "daily"
    say("Daily run is due -- starting in %s mode." % mode)
    return run(handle=state.get("handle"), mode=mode, log=say, client=sheets)


def main():
    ap = argparse.ArgumentParser(description="Instagram post-metrics ingest.")
    ap.add_argument("--handle", help="Instagram handle or profile URL.")
    ap.add_argument("--mode", choices=["daily", "backfill", "auto"], default="auto",
                    help="'auto' runs only if today's run is still outstanding.")
    args = ap.parse_args()

    missing = cfg.missing_secrets()
    if missing:
        print("Missing secrets: %s" % ", ".join(missing), file=sys.stderr)
        return 2

    result = run_if_due() if args.mode == "auto" else run(args.handle, args.mode)
    print(result)
    return 0 if result.get("status") in ("ok", "not_due", "skipped") else 1


if __name__ == "__main__":
    sys.exit(main())
