# 🤖 AI Agent Integration Prompt & Developer README
## Unified Competitor Influencer & Meta Ads Intelligence API (v2.1.0)

> **Copy & Paste this document as a system prompt into ANY other project, AI agent (ChatGPT, Claude, Cursor, Windsurf), or developer workflow to instantly integrate competitor influencer audits with Date Ranges (1w, 1m, 3m, 6m, 1y, 2y).**

---

## 📋 Table of Contents
1. [AI System Prompt (Copy-Paste for Other Agents)](#1-ai-system-prompt-copy-paste-for-other-agents)
2. [Supported Date Ranges & Cohorts](#2-supported-date-ranges--cohorts)
3. [REST Endpoint Reference](#3-rest-endpoint-reference)
4. [Python SDK Integration (Plug & Play)](#4-python-sdk-integration-plug--play)
5. [Node.js / TypeScript / Next.js Integration](#5-nodejs--typescript--nextjs-integration)

---

## 1. AI System Prompt (Copy-Paste for Other Agents)

```markdown
<SYSTEM_INSTRUCTION>
You have access to the Unified Competitor & Paid Media Intelligence API running at http://localhost:8080.

This API provides:
1. Full 360-Degree Competitor Audits across custom date ranges:
   - '1w' / '7d'   -> Last 1 Week
   - '1m' / '30d'  -> Last 1 Month
   - '3m' / '90d'  -> Last 3 Months (Quarter)
   - '6m' / '180d' -> Last 6 Months (Half-Year)
   - '1y' / '365d' -> Last 1 Year (Default)
   - '2y' / '730d' -> Last 2 Years
2. Creator Follower Counts & Audience Sizing Tiers:
   - 🌟 Mega Creator (1M+)
   - 🚀 Macro Creator (100K - 1M)
   - ✨ Mid-Tier (50K - 100K)
   - 🎯 Micro (10K - 50K)
   - 🌱 Nano (<10K)
3. 4-Tier Paid & Boost Classification:
   - Tier 1: Toggle ON + Boosted Paid Ad
   - Tier 2: Toggle ON + Organic Collab
   - Tier 3: Toggle OFF + Heavily Boosted Ad (like-to-view < 0.35% with high views)
   - Tier 4: Toggle OFF + Organic / Noise
4. Meta Ad Library Scanner: Captures active/inactive ad cards, library IDs, and whitelisted creator handles.

Always use GET /api/v1/intelligence/audit/{brand_username}?range={time_window} when the user asks for competitor influencer research or creator metrics.
</SYSTEM_INSTRUCTION>
```

---

## 2. Supported Date Ranges & Cohorts

| Parameter | Meaning | Days Scanned |
|---|---|:---:|
| `?range=1w` | **Last 1 Week** | 7 Days |
| `?range=1m` | **Last 1 Month** | 30 Days |
| `?range=3m` | **Last 3 Months** | 90 Days |
| `?range=6m` | **Last 6 Months** | 180 Days |
| `?range=1y` | **Last 1 Year** *(Default)* | 365 Days |
| `?range=2y` | **Last 2 Years** | 730 Days |

---

## 3. REST Endpoint Reference

### 🟢 Endpoint 1: Full 360° Competitor Audit (With Date Range)
* **Method**: `GET`
* **Route**: `/api/v1/intelligence/audit/{brand_username}?range={time_window}&export_excel={true|false}`
* **cURL Examples**:
  ```bash
  # Audit for Last 1 Week
  curl -X GET "http://localhost:8080/api/v1/intelligence/audit/palmonas_official?range=1w"

  # Audit for Last 3 Months
  curl -X GET "http://localhost:8080/api/v1/intelligence/audit/palmonas_official?range=3m"

  # Audit for Last 6 Months
  curl -X GET "http://localhost:8080/api/v1/intelligence/audit/palmonas_official?range=6m"

  # Audit for Last 1 Year
  curl -X GET "http://localhost:8080/api/v1/intelligence/audit/palmonas_official?range=1y"
  ```
* **Sample JSON Response**:
  ```json
  {
    "brand": "palmonas_official",
    "time_window": "3m",
    "days_audited": 90,
    "date_range_cohorts": {
      "last_7d": 4,
      "last_30d": 18,
      "last_90d": 52,
      "last_180d": 89,
      "total_in_window": 52
    },
    "total_unique_creators": 64,
    "instagram_grid_creators": 48,
    "meta_adlibrary_dark_creators": 16,
    "total_collab_posts": 52,
    "total_meta_ads": 233,
    "unified_creators": [
      {
        "handle": "@shraddhakapoor",
        "full_name": "Shraddha Kapoor",
        "followers": 93000000,
        "creator_tier": "🌟 Mega Creator (1M+)",
        "presence_platform": "💎 Both (IG Grid + Meta Ads)",
        "total_grid_posts": 4,
        "total_grid_views": 18450000,
        "active_meta_ads": 19,
        "sample_ad_url": "https://www.facebook.com/ads/library/?id=616056441550193"
      }
    ],
    "excel_file": "palmonas_official_3m_creator_audit.xlsx"
  }
  ```

---

### 🟢 Endpoint 2: Instagram Collabs (With Date Range)
* **Route**: `GET /api/v1/instagram/partnerships/{brand_username}?range=1w`
* **cURL**:
  ```bash
  curl -X GET "http://localhost:8080/api/v1/instagram/partnerships/giva.co?range=1m"
  ```

---

### 🟢 Endpoint 3: Live Creator Profile & Sizing Tier
* **Route**: `GET /api/v1/instagram/profile/{username}`
* **cURL**:
  ```bash
  curl -X GET "http://localhost:8080/api/v1/instagram/profile/nikitadhongdi"
  ```

---

### 🟢 Endpoint 4: Meta Ad Library Search
* **Route**: `GET /api/v1/meta/ads?query={brand}&active_only=false`
* **cURL**:
  ```bash
  curl -X GET "http://localhost:8080/api/v1/meta/ads?query=Palmonas"
  ```

---

### 🟢 Endpoint 5: Facebook Page Metrics & Delegate ID
* **Route**: `GET /api/v1/facebook/page/{page_handle}`
* **cURL**:
  ```bash
  curl -X GET "http://localhost:8080/api/v1/facebook/page/zivame"
  ```
