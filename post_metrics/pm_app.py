"""
Streamlit control panel for the Instagram post-metrics pipeline.

The app is a dashboard and a trigger -- it is NOT the scheduler. The clock
lives in PostMetrics.gs, which wakes this app at 09:00. On every page load the
app asks the sheet whether today's run has happened yet, and runs it if not.

That design is what makes a missed wake-up harmless: the check is keyed on the
local calendar date, so an outstanding run stays outstanding until it actually
completes, whenever the app next opens.

Run locally:  streamlit run pm_app.py
"""

import streamlit as st

import pm_config as cfg
import pm_ingest
from pm_instagram import parse_handle
from pm_sheets import SheetsClient, SheetsError

st.set_page_config(page_title="Instagram Post Metrics", page_icon="📈",
                   layout="wide")

st.title("📈 Instagram Post Metrics → Google Sheets")
st.caption("Daily at %02d:00 %s · new posts appended · posts under %d days old "
           "refreshed" % (cfg.DAILY_RUN_HOUR, cfg.TIMEZONE, cfg.REFRESH_WINDOW_DAYS))


# ── connection ────────────────────────────────────────────────────────────────
missing = cfg.missing_secrets()
if missing:
    st.error("Missing secrets: **%s**" % ", ".join(missing))
    with st.expander("How to set these"):
        st.markdown(
            "On Streamlit Cloud: **Manage app → Settings → Secrets**. "
            "Locally: a `.env` file next to this script (it is gitignored)."
        )
        st.code(
            'IG_SESSIONID = "..."\n'
            'IG_DS_USER_ID = "..."\n'
            'IG_CSRFTOKEN = "..."\n'
            'IG_MID = "..."\n'
            'SHEETS_WEBHOOK_URL = "https://script.google.com/macros/s/AKfy.../exec"\n'
            'SHEETS_WEBHOOK_TOKEN = "the same string as TOKEN in PostMetrics.gs"',
            language="toml",
        )
    st.stop()

try:
    sheets = SheetsClient()
    index = sheets.get_index()
except SheetsError as e:
    st.error("Cannot reach the Google Sheet webhook: %s" % e)
    st.stop()

state = index["state"]
known_posts = index["posts"]
configured_handle = str(state.get("handle") or "")
backfill_done = str(state.get("backfill_done", "")).upper() == "TRUE"


def live_log():
    """A log sink that streams progress into the page as it happens."""
    lines, box = [], st.empty()
    def say(message):
        lines.append(str(message))
        box.code("\n".join(lines[-20:]))
    return say


def execute(mode):
    with st.spinner("Running %s ingest..." % mode):
        result = pm_ingest.run(handle=configured_handle, mode=mode,
                               log=live_log(), client=sheets)
    if result.get("status") == "ok":
        st.success("Done — %d new, %d refreshed, %d scanned in %ss."
                   % (result["new_posts"], result["refreshed_posts"],
                      result["scanned_posts"], result.get("duration_seconds")))
    elif result.get("status") == "skipped":
        st.warning("Another run holds the lock (%s)." % result.get("held_by"))
    else:
        st.error("Run failed: %s" % result.get("error"))
    st.rerun()


# ── sidebar: configuration ────────────────────────────────────────────────────
with st.sidebar:
    st.header("Account")
    handle_input = st.text_input(
        "Instagram profile URL or handle",
        value=configured_handle,
        placeholder="https://www.instagram.com/nike/",
        help="Paste the profile URL or just the handle. Saved to the sheet, "
             "so the scheduled run uses it too.",
    )
    parsed = parse_handle(handle_input)
    if parsed and parsed != configured_handle:
        st.info("Will track **@%s**" % parsed)
    if st.button("Save account", use_container_width=True, disabled=not parsed):
        changed_account = parsed != configured_handle and configured_handle != ""
        sheets.set_state({
            "handle": parsed,
            # A different account means the existing rows are someone else's
            # history, so the backfill has to run again.
            "backfill_done": "FALSE" if changed_account else state.get("backfill_done", "FALSE"),
        })
        st.success("Saved @%s" % parsed)
        st.rerun()

    st.divider()
    st.header("Schedule")
    enabled = str(state.get("enabled", "TRUE")).upper() != "FALSE"
    new_enabled = st.toggle("Daily run enabled", value=enabled)
    if new_enabled != enabled:
        sheets.set_state({"enabled": "TRUE" if new_enabled else "FALSE"})
        st.rerun()
    st.caption("The 09:00 trigger lives in Apps Script "
               "(**Post Metrics → Schedule daily wake-up**), not in this app. "
               "Streamlit Cloud sleeps when idle and cannot keep its own timer.")


# ── auto-run when due ─────────────────────────────────────────────────────────
# One attempt per browser session; the sheet lock stops concurrent sessions
# from both firing, and the date key stops repeats within the same day.
if configured_handle and not st.session_state.get("due_check_done"):
    st.session_state["due_check_done"] = True
    if pm_ingest.is_due(state):
        st.info("Today's scheduled run has not happened yet — running it now.")
        execute("backfill" if not backfill_done else "daily")


# ── status ────────────────────────────────────────────────────────────────────
if not configured_handle:
    st.warning("Set an Instagram account in the sidebar to get started.")
    st.stop()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Account", "@%s" % configured_handle)
c2.metric("Posts tracked", "{:,}".format(len(known_posts)))
c3.metric("Followers", "{:,}".format(int(state.get("follower_count") or 0)))
c4.metric("Last run", str(state.get("last_run_date") or "never"))

status = str(state.get("last_run_status") or "")
if status == "error":
    st.error("Last run failed: %s" % state.get("last_error"))
elif status == "ok":
    st.success("Last run completed at %s (mode: %s)."
               % (state.get("last_run_at"), state.get("last_run_mode")))

if not backfill_done:
    st.warning("**Backfill has not run yet.** It walks the full history "
               "(up to %s posts) and is the slow one — run it once, from a "
               "machine you can leave open, before relying on the daily job."
               % "{:,}".format(cfg.BACKFILL_MAX_POSTS))

lock_owner = str(state.get("lock_owner") or "")
if lock_owner:
    st.info("A run is currently in progress (%s, expires %s)."
            % (lock_owner, state.get("lock_until")))


# ── manual triggers ───────────────────────────────────────────────────────────
st.subheader("Run now")
b1, b2 = st.columns(2)
with b1:
    if st.button("Daily run", use_container_width=True, type="primary"):
        execute("daily")
    st.caption("New posts since the last run, plus a metrics refresh for "
               "everything posted in the last %d days." % cfg.REFRESH_WINDOW_DAYS)
with b2:
    if st.button("Full backfill", use_container_width=True):
        execute("backfill")
    st.caption("Walks the entire account history. Slow — expect several "
               "minutes, and Streamlit Cloud may time out on large accounts.")

st.divider()
with st.expander("Where the data lands"):
    st.markdown(
        "| Tab | Contents |\n"
        "|---|---|\n"
        "| **Posts** | One row per post, latest metrics, keyed on shortcode |\n"
        "| **Metrics History** | One snapshot per post per run — the growth curve |\n"
        "| **Run Log** | One row per run, with counts, duration and errors |\n"
        "| **State** | Handle, last run date, lock. The pipeline's memory — "
        "it lives here because Streamlit Cloud's disk is wiped on restart |"
    )
