/**
 * ==============================================================================
 * INSTAGRAM CREATOR & 4-TIER PAID COLLABORATION ANALYTICS ENGINE
 * Google Apps Script (Cron Job & Webhook Automation)
 * ==============================================================================
 * 
 * Features:
 * 1. Custom Menu in Google Sheets ("🚀 Instagram Analytics")
 * 2. Automated Daily / Weekly Cron Trigger ("setupDailyCronJob")
 * 3. Populates 3 Sheets:
 *    - 📊 Executive Summary (High-Intent Paid % & Brand Totals)
 *    - 👥 Creators Profile Metrics (Followers, Pure Tiers, ER%)
 *    - 🏷️ All Collaborations - 4-Tier Hierarchy (Tier 1, 2, 3, 4 + Video Genres)
 * 4. Webhook Receiver (doPost) for Python / GitHub Actions direct sync
 */

// ── CONFIGURATION ─────────────────────────────────────────────────────────────
var CONFIG = {
  // Option A: URL of your hosted FastAPI backend (e.g. Render / Railway / Cloud Run / VPS)
  API_BASE_URL: "https://your-api-domain.com/api/v1/analyze/brand",
  
  // Default list of brand handles to audit if none provided in the "Config" sheet
  DEFAULT_BRANDS: [
    "grtoriana",
    "gullylabs",
    "thecometuniverse",
    "skechersindia",
    "japam.in",
    "divinehindu.in"
  ],
  
  MAX_PAGES: 6
};

// ── 1. CUSTOM MENU CREATION ──────────────────────────────────────────────────
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 Instagram Analytics")
    .addItem("▶ Run Audit for All Configured Brands", "runManualBrandAudit")
    .addItem("⏰ Setup Daily Cron Trigger (9:00 AM)", "setupDailyCronTrigger")
    .addItem("⏱️ Setup Weekly Cron Trigger (Mondays)", "setupWeeklyCronTrigger")
    .addItem("🚫 Remove All Automated Triggers", "removeAutomatedTriggers")
    .addSeparator()
    .addItem("⚙️ Initialize / Reset Config Sheet", "initConfigSheet")
    .addToUi();
}

// ── 2. INITIALIZE CONFIG SHEET ───────────────────────────────────────────────
function initConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Brands Config") || ss.insertSheet("Brands Config", 0);
  
  sheet.clear();
  sheet.getRange("A1:C1").setValues([["#", "Brand Instagram Handle", "Active Status"]]);
  sheet.getRange("A1:C1").setBackground("#1B2631").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  
  var rows = [];
  for (var i = 0; i < CONFIG.DEFAULT_BRANDS.length; i++) {
    rows.push([i + 1, CONFIG.DEFAULT_BRANDS[i], "Active"]);
  }
  
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.autoResizeColumns(1, 3);
  SpreadsheetApp.getUi().alert("✅ 'Brands Config' sheet created! You can add or remove brand handles here.");
}

// ── 3. CRON JOB TRIGGER SETUP ────────────────────────────────────────────────
function setupDailyCronTrigger() {
  removeAutomatedTriggers();
  ScriptApp.newTrigger("runScheduledBrandAudit")
    .timeBased()
    .everyDays(1)
    .atHour(9) // 9:00 AM daily
    .create();
  
  SpreadsheetApp.getUi().alert("⏰ Daily Cron Job Scheduled! The audit will run automatically every day at 9:00 AM.");
}

function setupWeeklyCronTrigger() {
  removeAutomatedTriggers();
  ScriptApp.newTrigger("runScheduledBrandAudit")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9) // Every Monday at 9:00 AM
    .create();
    
  SpreadsheetApp.getUi().alert("⏱️ Weekly Cron Job Scheduled! The audit will run automatically every Monday at 9:00 AM.");
}

function removeAutomatedTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runScheduledBrandAudit") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ── 4. SCHEDULED AUDIT EXECUTOR ──────────────────────────────────────────────
function runScheduledBrandAudit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName("Brands Config");
  var brandsToAudit = [];
  
  if (configSheet) {
    var data = configSheet.getRange(2, 2, configSheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      var handle = String(data[i][0]).replace("@", "").trim();
      var status = String(data[i][1]).trim().toLowerCase();
      if (handle && status === "active") {
        brandsToAudit.push(handle);
      }
    }
  }
  
  if (brandsToAudit.length === 0) {
    brandsToAudit = CONFIG.DEFAULT_BRANDS;
  }
  
  var allPosts = [];
  var allProfilesMap = {};
  var summaryList = [];
  
  for (var b = 0; b < brandsToAudit.length; b++) {
    var brandHandle = brandsToAudit[b];
    try {
      var url = CONFIG.API_BASE_URL + "?username=" + encodeURIComponent(brandHandle) + "&max_pages=" + CONFIG.MAX_PAGES;
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, timeout: 60000 });
      
      if (response.getResponseCode() === 200) {
        var json = JSON.parse(response.getContentText());
        if (json.status === "success") {
          processBrandAuditResult(json, allPosts, allProfilesMap, summaryList);
        }
      }
    } catch (e) {
      Logger.log("Error auditing " + brandHandle + ": " + e.toString());
    }
    Utilities.sleep(1000); // 1-second pause between brands
  }
  
  renderExecutiveSummarySheet(ss, summaryList);
  renderCreatorMetricsSheet(ss, Object.values(allProfilesMap));
  renderMasterHierarchySheet(ss, allPosts);
}

function runManualBrandAudit() {
  var ui = SpreadsheetApp.getUi();
  ui.alert("⏳ Starting Brand Audit. Please allow 1–2 minutes for the data to refresh.");
  runScheduledBrandAudit();
  ui.alert("✅ Audit Complete! All sheets have been updated with the latest 4-Tier data.");
}

// ── 5. DATA INGESTION & WEBHOOK (doPost) ──────────────────────────────────────
/**
 * Allows external Python scripts / GitHub Actions to push JSON directly into Google Sheets:
 * curl -X POST -H "Content-Type: application/json" -d @audit.json <APPS_SCRIPT_WEBAPP_URL>
 */
function doPost(e) {
  try {
    var json = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var allPosts = json.posts || [];
    var profiles = json.profiles || [];
    var summaries = json.summaries || [];
    
    if (summaries.length > 0) renderExecutiveSummarySheet(ss, summaries);
    if (profiles.length > 0) renderCreatorMetricsSheet(ss, profiles);
    if (allPosts.length > 0) renderMasterHierarchySheet(ss, allPosts);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", count: allPosts.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function processBrandAuditResult(json, allPosts, allProfilesMap, summaryList) {
  var brand = json.brand;
  var sum = json.summary;
  
  var t1 = json.tier_1_toggle_on_boosted || [];
  var t2 = json.tier_2_toggle_on_organic || [];
  var t3 = json.tier_3_toggle_off_boosted || [];
  var t4 = json.tier_4_noise || [];
  
  var brandPosts = [].concat(t1, t2, t3, t4);
  for (var i = 0; i < brandPosts.length; i++) {
    allPosts.push(brandPosts[i]);
    var h = brandPosts[i].creator_handle.toLowerCase();
    if (!allProfilesMap[h]) {
      allProfilesMap[h] = {
        handle: brandPosts[i].creator_handle,
        followers: brandPosts[i].creator_followers || 0,
        full_name: brandPosts[i].creator_name || brandPosts[i].creator_handle,
        brands: brand.name,
        creator_tier: getCreatorPureTier(brandPosts[i].creator_followers || 0)
      };
    }
  }
  
  summaryList.push({
    brand: brand.name,
    state: brand.state_origin || "Pan-India",
    posts: sum.total_collab_posts,
    creators: sum.total_unique_creators,
    t1_p: sum.tier_1_toggle_on_boosted.posts_count,
    t2_p: sum.tier_2_toggle_on_organic.posts_count,
    t3_p: sum.tier_3_toggle_off_boosted.posts_count,
    t4_p: sum.tier_4_toggle_off_organic_noise.posts_count,
    high_intent: sum.total_high_intent_paid_ads,
    high_intent_pct: sum.high_intent_paid_rate_pct / 100
  });
}

function getCreatorPureTier(fols) {
  if (fols >= 1000000) return "🌟 Mega Creator (1M+)";
  if (fols >= 100000) return "🚀 Macro Creator (100K - 1M)";
  if (fols >= 50000) return "✨ Mid-Tier Creator (50K - 100K)";
  if (fols >= 10000) return "🎯 Micro Creator (10K - 50K)";
  return "🌱 Nano Creator (<10K)";
}

// ── 6. SHEET RENDERERS ───────────────────────────────────────────────────────
function renderExecutiveSummarySheet(ss, summaryList) {
  var sheet = ss.getSheetByName("Executive Summary") || ss.insertSheet("Executive Summary", 1);
  sheet.clear();
  
  var headers = [
    "#", "Brand Name", "State / Origin (HQ)", "Total Collabs", "Unique Creators",
    "🟢 Tier 1 (Boosted Paid)", "🟢 Tier 2 (Organic Paid)", "🚀 Tier 3 (Boosted Collab)",
    "⚪ Tier 4 (Noise)", "💎 High-Intent Paid Total", "High-Intent Adoption %"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#0B2240").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  
  var rows = [];
  for (var i = 0; i < summaryList.length; i++) {
    var b = summaryList[i];
    rows.push([
      i + 1, b.brand, b.state, b.posts, b.creators,
      b.t1_p > 0 ? b.t1_p + " posts" : "—",
      b.t2_p > 0 ? b.t2_p + " posts" : "—",
      b.t3_p > 0 ? b.t3_p + " posts" : "—",
      b.t4_p > 0 ? b.t4_p + " posts" : "—",
      b.high_intent,
      b.high_intent_pct
    ]);
  }
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 11, rows.length, 1).setNumberFormat("0.0%");
  }
  sheet.autoResizeColumns(1, headers.length);
}

function renderCreatorMetricsSheet(ss, profilesList) {
  var sheet = ss.getSheetByName("Creators Profile Metrics") || ss.insertSheet("Creators Profile Metrics", 2);
  sheet.clear();
  
  var headers = ["#", "Creator Handle", "Creator Tier / Size", "Brand", "Followers", "Profile URL"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#1B4F72").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  
  var rows = [];
  for (var i = 0; i < profilesList.length; i++) {
    var p = profilesList[i];
    rows.push([
      i + 1, p.handle, p.creator_tier, p.brands, p.followers,
      "https://www.instagram.com/" + p.handle.replace("@", "") + "/"
    ]);
  }
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat("#,##0");
  }
  sheet.autoResizeColumns(1, headers.length);
}

function renderMasterHierarchySheet(ss, allPosts) {
  var sheet = ss.getSheetByName("All Brands - Master Hierarchy") || ss.insertSheet("All Brands - Master Hierarchy", 3);
  sheet.clear();
  
  var headers = [
    "#", "Hierarchy Tier", "Brand Name", "Creator Handle", "Followers", "Video Content Genre",
    "Views / Plays", "Likes", "Comments", "Like-to-View %", "Creator ER%", "Post Date", "Instagram URL", "Boost Status & Reason"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#1F2D3D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  
  var rows = [];
  for (var i = 0; i < allPosts.length; i++) {
    var p = allPosts[i];
    rows.push([
      i + 1,
      p.tier_name || ("Tier " + p.tier),
      p.brand,
      p.creator_handle || p.handle,
      p.creator_followers || p.followers || 0,
      p.video_genre || "Lifestyle",
      p.estimated_views || p.views || 0,
      p.likes || 0,
      p.comments || 0,
      (p.like_to_view_pct || p.like_rate_pct || 0) / 100,
      (p.creator_er_pct || p.er_pct || 0) / 100,
      p.post_date || p.date || "",
      p.post_url || p.url || "",
      p.boost_reason || p.boost_status || ""
    ]);
  }
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat("#,##0");
    sheet.getRange(2, 7, rows.length, 3).setNumberFormat("#,##0");
    sheet.getRange(2, 10, rows.length, 2).setNumberFormat("0.00%");
  }
  sheet.autoResizeColumns(1, headers.length);
}
