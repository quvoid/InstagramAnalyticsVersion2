/**
 * PostMetrics.gs -- Sheets backend + scheduler for the daily Instagram
 * post-metrics pipeline.
 *
 * This script does two separate jobs:
 *
 *  1. WEBHOOK. doPost() is the only way the Python ingest reads or writes the
 *     spreadsheet. Reads go over POST too, so the shared token never appears
 *     in a URL or a referrer header.
 *
 *  2. CLOCK. Streamlit Cloud has no scheduler and puts idle apps to sleep, so
 *     nothing inside the app can reliably fire at 09:00. A time-driven trigger
 *     here does it instead: it wakes the Streamlit app, which then notices
 *     today's run is outstanding and performs it.
 *
 * SETUP
 *   1. Extensions > Apps Script on the target spreadsheet, paste this file.
 *   2. Set TOKEN below to a long random string. Put the same value in the
 *      Streamlit app's secrets as SHEETS_WEBHOOK_TOKEN.
 *   3. Set STREAMLIT_APP_URL to your deployed app URL.
 *   4. Deploy > New deployment > Web app
 *        Execute as: Me
 *        Who has access: Anyone
 *      Copy the /exec URL into secrets as SHEETS_WEBHOOK_URL.
 *   5. Run initSheets() once, then setupDailyTrigger() once.
 */

var TOKEN = 'CHANGE-ME-to-a-long-random-string';
var STREAMLIT_APP_URL = 'https://your-app.streamlit.app';
var RUN_HOUR = 9;            // local hour of the daily wake-up
var TIMEZONE = 'Asia/Kolkata';

var POSTS_SHEET = 'Posts';
var HISTORY_SHEET = 'Metrics History';
var STATE_SHEET = 'State';
var LOG_SHEET = 'Run Log';

var POST_COLUMNS = [
  'shortcode', 'handle', 'post_url', 'posted_at', 'media_kind',
  'carousel_count', 'owner', 'caption', 'like_count', 'comment_count',
  'view_count', 'video_duration', 'is_paid_partnership', 'coauthors',
  'counts_hidden', 'follower_count_at_scrape', 'engagement_rate_pct',
  'taken_at', 'first_seen_at', 'scraped_at'
];

var HISTORY_COLUMNS = [
  'snapshot_date', 'snapshot_at', 'shortcode', 'handle', 'posted_at',
  'like_count', 'comment_count', 'view_count', 'follower_count',
  'engagement_rate_pct', 'run_id'
];

var LOG_COLUMNS = [
  'started_at', 'finished_at', 'duration_seconds', 'mode', 'handle', 'status',
  'scanned_posts', 'new_posts', 'refreshed_posts', 'follower_count',
  'error', 'run_id'
];

// ── 1. WEBHOOK ────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.token !== TOKEN) {
      return json({ ok: false, error: 'Invalid token' });
    }
    switch (body.action) {
      case 'ping':           return json({ ok: true, sheet: SpreadsheetApp.getActive().getName() });
      case 'index':          return json(handleIndex());
      case 'set_state':      return json(handleSetState(body.values || {}));
      case 'acquire_lock':   return json(handleAcquireLock(body.owner, body.ttl_minutes));
      case 'release_lock':   return json(handleReleaseLock(body.owner));
      case 'upsert_posts':   return json(handleUpsertPosts(body.rows || []));
      case 'append_history': return json(handleAppendHistory(body.rows || []));
      case 'log_run':        return json(handleLogRun(body.record || {}));
      default:               return json({ ok: false, error: 'Unknown action: ' + body.action });
    }
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 2. SHEET SETUP ────────────────────────────────────────────────────────────

function initSheets() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet(ss, POSTS_SHEET, POST_COLUMNS);
  ensureSheet(ss, HISTORY_SHEET, HISTORY_COLUMNS);
  ensureSheet(ss, LOG_SHEET, LOG_COLUMNS);
  ensureSheet(ss, STATE_SHEET, ['key', 'value']);

  var state = ss.getSheetByName(STATE_SHEET);
  if (state.getLastRow() < 2) {
    state.getRange(2, 1, 6, 2).setValues([
      ['handle', ''],
      ['enabled', 'TRUE'],
      ['backfill_done', 'FALSE'],
      ['last_run_date', ''],
      ['last_run_status', ''],
      ['total_posts_tracked', 0]
    ]);
  }
  SpreadsheetApp.getActive().toast('Sheets ready.');
}

function ensureSheet(ss, name, columns) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, columns.length).setValues([columns])
      .setFontWeight('bold').setBackground('#1B4F72').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── 3. READ ───────────────────────────────────────────────────────────────────

/**
 * The whole picture the ingest needs, in one round trip: current state plus a
 * shortcode -> taken_at index. Only those two columns are read, so the
 * response stays small even at thousands of posts.
 */
function handleIndex() {
  var ss = SpreadsheetApp.getActive();
  var posts = {};
  var sheet = ss.getSheetByName(POSTS_SHEET);
  if (sheet && sheet.getLastRow() > 1) {
    var codeCol = POST_COLUMNS.indexOf('shortcode') + 1;
    var takenCol = POST_COLUMNS.indexOf('taken_at') + 1;
    var n = sheet.getLastRow() - 1;
    var codes = sheet.getRange(2, codeCol, n, 1).getValues();
    var taken = sheet.getRange(2, takenCol, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      if (codes[i][0]) posts[String(codes[i][0])] = taken[i][0] || 0;
    }
  }
  return { ok: true, state: readState(), posts: posts };
}

function readState() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(STATE_SHEET);
  var out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0]) out[String(rows[i][0])] = rows[i][1];
  }
  return out;
}

function handleSetState(values) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ensureSheet(ss, STATE_SHEET, ['key', 'value']);
  var existing = {};
  if (sheet.getLastRow() > 1) {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < rows.length; i++) existing[String(rows[i][0])] = i + 2;
  }
  for (var key in values) {
    var value = values[key];
    if (existing[key]) {
      sheet.getRange(existing[key], 2).setValue(value);
    } else {
      sheet.appendRow([key, value]);
      existing[key] = sheet.getLastRow();
    }
  }
  return { ok: true };
}

// ── 4. LOCKING ────────────────────────────────────────────────────────────────

/**
 * Two viewers opening the app at the same moment would both see the daily run
 * as outstanding. The lock makes the second one stand down. It carries an
 * expiry so a run that dies mid-way cannot block tomorrow's.
 */
function handleAcquireLock(owner, ttlMinutes) {
  var guard = LockService.getScriptLock();
  guard.waitLock(20000);
  try {
    var state = readState();
    var until = state.lock_until ? new Date(state.lock_until).getTime() : 0;
    var now = Date.now();
    if (state.lock_owner && until > now && state.lock_owner !== owner) {
      return { ok: true, acquired: false, held_by: state.lock_owner };
    }
    var ttl = (ttlMinutes || 30) * 60 * 1000;
    handleSetState({
      lock_owner: owner,
      lock_until: new Date(now + ttl).toISOString()
    });
    return { ok: true, acquired: true, held_by: owner };
  } finally {
    guard.releaseLock();
  }
}

function handleReleaseLock(owner) {
  var state = readState();
  if (state.lock_owner && state.lock_owner !== owner) {
    return { ok: true, released: false };
  }
  handleSetState({ lock_owner: '', lock_until: '' });
  return { ok: true, released: true };
}

// ── 5. WRITE ──────────────────────────────────────────────────────────────────

/**
 * Insert new posts, overwrite existing ones. Keyed on shortcode, so a rerun on
 * the same day is a no-op rather than a duplicate.
 *
 * first_seen_at is deliberately preserved on update -- it records when the
 * pipeline discovered the post, which is not the same as when it was posted.
 */
function handleUpsertPosts(rows) {
  if (!rows.length) return { ok: true, inserted: 0, updated: 0 };
  var ss = SpreadsheetApp.getActive();
  var sheet = ensureSheet(ss, POSTS_SHEET, POST_COLUMNS);

  var codeCol = POST_COLUMNS.indexOf('shortcode') + 1;
  var firstSeenCol = POST_COLUMNS.indexOf('first_seen_at') + 1;
  var rowByCode = {};
  if (sheet.getLastRow() > 1) {
    var codes = sheet.getRange(2, codeCol, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < codes.length; i++) {
      if (codes[i][0]) rowByCode[String(codes[i][0])] = i + 2;
    }
  }

  var appended = [], updated = 0, now = new Date().toISOString();
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var target = rowByCode[String(row.shortcode)];
    if (target) {
      var keep = sheet.getRange(target, firstSeenCol).getValue();
      row.first_seen_at = keep || now;
      sheet.getRange(target, 1, 1, POST_COLUMNS.length)
        .setValues([toRow(row, POST_COLUMNS)]);
      updated++;
    } else {
      row.first_seen_at = now;
      appended.push(toRow(row, POST_COLUMNS));
    }
  }
  if (appended.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appended.length, POST_COLUMNS.length)
      .setValues(appended);
  }
  sortPostsByDate(sheet);
  return { ok: true, inserted: appended.length, updated: updated };
}

function sortPostsByDate(sheet) {
  var takenCol = POST_COLUMNS.indexOf('taken_at') + 1;
  if (sheet.getLastRow() > 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, POST_COLUMNS.length)
      .sort({ column: takenCol, ascending: false });
  }
}

function handleAppendHistory(rows) {
  if (!rows.length) return { ok: true, appended: 0 };
  var sheet = ensureSheet(SpreadsheetApp.getActive(), HISTORY_SHEET, HISTORY_COLUMNS);
  var values = rows.map(function (row) { return toRow(row, HISTORY_COLUMNS); });
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, HISTORY_COLUMNS.length)
    .setValues(values);
  return { ok: true, appended: values.length };
}

function handleLogRun(record) {
  var sheet = ensureSheet(SpreadsheetApp.getActive(), LOG_SHEET, LOG_COLUMNS);
  sheet.appendRow(toRow(record, LOG_COLUMNS));
  return { ok: true };
}

function toRow(obj, columns) {
  return columns.map(function (c) {
    var v = obj[c];
    if (v === undefined || v === null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return v;
  });
}

// ── 6. THE CLOCK ──────────────────────────────────────────────────────────────

function setupDailyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('wakeStreamlitApp')
    .timeBased().atHour(RUN_HOUR).nearMinute(0).everyDays(1)
    .inTimezone(TIMEZONE).create();
  SpreadsheetApp.getActive().toast('Daily wake-up scheduled for ' + RUN_HOUR + ':00.');
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'wakeStreamlitApp') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Wake the Streamlit app so it runs the ingest.
 *
 * Caveat worth knowing: a plain HTTP GET reliably wakes a sleeping Streamlit
 * Cloud container, but Streamlit only executes the app script once a browser
 * websocket session connects. So this is best-effort, not a guarantee -- which
 * is exactly why the app's due-check is date-keyed. If a wake-up does not take,
 * the run simply happens the next time the app is opened, and no day is
 * skipped or double-counted.
 */
function wakeStreamlitApp() {
  var attempts = [];
  for (var i = 0; i < 3; i++) {
    try {
      var res = UrlFetchApp.fetch(STREAMLIT_APP_URL + '?wake=1', {
        muteHttpExceptions: true,
        followRedirects: true,
        validateHttpsCertificates: true
      });
      attempts.push(res.getResponseCode());
      if (res.getResponseCode() === 200) break;
    } catch (err) {
      attempts.push(String(err));
    }
    Utilities.sleep(20000);   // give a cold container time to boot
  }
  handleSetState({
    last_wake_at: new Date().toISOString(),
    last_wake_result: attempts.join(', ')
  });
}

// ── 7. MENU ───────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Post Metrics')
    .addItem('Initialise sheets', 'initSheets')
    .addItem('Schedule daily wake-up', 'setupDailyTrigger')
    .addItem('Remove wake-up trigger', 'removeTriggers')
    .addItem('Wake the app now', 'wakeStreamlitApp')
    .addToUi();
}
