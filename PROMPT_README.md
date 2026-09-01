# 🤖 AI Agent Integration Prompt & Developer README
## Unified Competitor Influencer & Meta Ads Intelligence API

> **Copy & Paste this document as a system prompt into ANY other project, AI agent (ChatGPT, Claude, Cursor, Windsurf), or developer workflow to instantly integrate competitor influencer audits, Instagram scrapers, and Meta Ad Library intelligence.**

---

## 📋 Table of Contents
1. [AI System Prompt (Copy-Paste for Other Agents)](#1-ai-system-prompt-copy-paste-for-other-agents)
2. [API Architecture & Base URL](#2-api-architecture--base-url)
3. [REST Endpoint Reference](#3-rest-endpoint-reference)
4. [Python SDK Integration (Plug & Play)](#4-python-sdk-integration-plug--play)
5. [Node.js / TypeScript / Next.js Integration](#5-nodejs--typescript--nextjs-integration)
6. [Pre-Built AI Prompt Templates for Common Tasks](#6-pre-built-ai-prompt-templates-for-common-tasks)

---

## 1. AI System Prompt (Copy-Paste for Other Agents)

```markdown
<SYSTEM_INSTRUCTION>
You have access to the Unified Competitor & Paid Media Intelligence API running at http://localhost:8080 (or your deployed server).

This API provides:
1. Full 360-Degree Competitor Audits (combines Instagram 1-year organic posts with Meta Ad Library dark ads).
2. Pure Creator Profile Metrics (live follower counts, audience sizing tiers: Mega 1M+, Macro 100K-1M, Mid 50K-100K, Micro 10K-50K, Nano <10K).
3. 4-Tier Boost Classification for Instagram videos:
   - Tier 1: Toggle ON + Boosted Paid Ad
   - Tier 2: Toggle ON + Organic Collab
   - Tier 3: Toggle OFF + Heavily Boosted Ad (like-to-view < 0.35% with high views)
   - Tier 4: Toggle OFF + Organic / Noise
4. Meta Ad Library Scanner: Captures active/inactive ad cards, library IDs, runtime dates, and whitelisted creator handles.

Always use these endpoints when the user asks for competitor influencer research, creator metrics, or Meta ad library intelligence.
</SYSTEM_INSTRUCTION>
```

---

## 2. API Architecture & Base URL

* **Local Development Base URL**: `http://localhost:8080`
* **Interactive Swagger UI Documentation**: `http://localhost:8080/docs`
* **Python SDK Package Location**: `from api_wrapper import CompetitorIntelligenceClient`

---

## 3. REST Endpoint Reference

### 🟢 Endpoint 1: Full-Funnel 360° Competitor Audit (Recommended)
Executes the sequential pipeline: Instagram 1-Year Grid + Profile Metric Resolution + Meta Ad Library + Dark Ads Resolution + Cross-Platform Deduplication + Master Excel Export.

* **Method**: `GET`
* **Route**: `/api/v1/intelligence/audit/{brand_username}`
* **Query Parameters**:
  * `page_id` *(optional, string)*: Meta Facebook Page ID (e.g. `100076111693972`)
  * `days_back` *(optional, default: `365`)*: Number of days to audit
  * `export_excel` *(optional, default: `true`)*: Whether to build `.xlsx` file
* **cURL**:
  ```bash
  curl -X GET "http://localhost:8080/api/v1/intelligence/audit/palmonas_official?days_back=365&export_excel=true"
  ```
* **Sample JSON Response**:
  ```json
  {
    "brand": "palmonas_official",
    "total_unique_creators": 154,
    "instagram_grid_creators": 122,
    "meta_adlibrary_dark_creators": 32,
    "both_platforms_creators": 53,
    "total_collab_posts": 143,
    "total_meta_ads": 233,
    "unified_creators": [
      {
        "handle": "@shraddhakapoor",
        "full_name": "Shraddha Kapoor",
        "followers": 93000000,
        "creator_tier": "🌟 Mega Creator (1M+)",
        "presence_platform": "💎 Both (IG Grid + Meta Ads)",
        "total_grid_posts": 12,
        "total_grid_views": 22233142,
        "active_meta_ads": 19,
        "total_meta_ads": 23,
        "sample_grid_url": "https://www.instagram.com/p/C-h902-sR0z/",
        "sample_ad_url": "https://www.facebook.com/ads/library/?id=616056441550193"
      }
    ],
    "excel_file": "palmonas_official_complete_creator_audit.xlsx"
  }
  ```

---

### 🟢 Endpoint 2: Instagram 1-Year Partnerships & Boost Hierarchy
* **Method**: `GET`
* **Route**: `/api/v1/instagram/partnerships/{brand_username}?days_back=365`
* **cURL**:
  ```bash
  curl -X GET "http://localhost:8080/api/v1/instagram/partnerships/giva.co?days_back=365"
  ```

---

### 🟢 Endpoint 3: Live Creator Profile Lookup & Sizing Tier
* **Method**: `GET`
* **Route**: `/api/v1/instagram/profile/{username}`
* **cURL**:
  ```bash
  curl -X GET "http://localhost:8080/api/v1/instagram/profile/nikitadhongdi"
  ```
* **Sample Response**:
  ```json
  {
    "username": "nikitadhongdi",
    "handle": "@nikitadhongdi",
    "followers": 833000,
    "tier": "🚀 Macro Creator (100K-1M)",
    "profile_url": "https://www.instagram.com/nikitadhongdi/"
  }
  ```

---

### 🟢 Endpoint 4: Meta Ad Library Search & Whitelists
* **Method**: `GET`
* **Route**: `/api/v1/meta/ads?query={brand}&active_only=false`
* **cURL**:
  ```bash
  curl -X GET "http://localhost:8080/api/v1/meta/ads?query=Palmonas"
  ```

---

### 🟢 Endpoint 5: Facebook Page Information & Delegate ID
* **Method**: `GET`
* **Route**: `/api/v1/facebook/page/{page_handle}`
* **cURL**:
  ```bash
  curl -X GET "http://localhost:8080/api/v1/facebook/page/zivame"
  ```

---

## 4. Python SDK Integration (Plug & Play)

If you are writing Python scripts in another folder or project:

```python
import sys
# Point to the InstagramAnalytics root directory
sys.path.append(r"c:\Users\omkar\OneDrive\Desktop\InstagramAnalytics")

from api_wrapper import CompetitorIntelligenceClient

client = CompetitorIntelligenceClient()

# 1. Run Complete Competitor Audit
audit_data = client.audit_brand(
    target_brand="giva.co",
    days_back=365,
    export_excel=True
)

print(f"Total Creators: {audit_data['total_unique_creators']}")
print(f"Excel Deliverable: {audit_data['excel_file']}")
```

---

## 5. Node.js / TypeScript / Next.js Integration

```typescript
// services/competitorApi.ts
const API_BASE = "http://localhost:8080";

export interface CreatorAuditResponse {
  brand: string;
  total_unique_creators: number;
  instagram_grid_creators: number;
  meta_adlibrary_dark_creators: number;
  unified_creators: Array<{
    handle: string;
    full_name: string;
    followers: number;
    creator_tier: string;
    presence_platform: string;
    total_grid_posts: number;
    total_grid_views: number;
    active_meta_ads: number;
    sample_ad_url: string;
  }>;
  excel_file: string;
}

export async function fetchCompetitorAudit(brand: string, days = 365): Promise<CreatorAuditResponse> {
  const res = await fetch(`${API_BASE}/api/v1/intelligence/audit/${brand}?days_back=${days}&export_excel=true`);
  if (!res.ok) throw new Error(`Audit failed for ${brand}`);
  return res.json();
}
```

---

## 6. Pre-Built AI Prompt Templates for Common Tasks

### 🎨 Prompt 1: Build a Streamlit Dashboard for Competitor Research
```text
I want you to build a Streamlit dashboard (app.py) that connects to our local Competitor Intelligence API at http://localhost:8080.

Features to include:
1. Search Bar for Brand Username (e.g. palmonas_official, giva.co, snitch.co.in) and Days slider (30 to 730 days).
2. KPI Cards: Total Unique Creators, Total Instagram Views, Active Meta Ads, and % Dark Ads vs Organic.
3. Top Creator Table with interactive sorting by Followers, Video Views, and Active Ads.
4. Download Button for the generated Master Excel file.
5. Filter by Creator Scale Tier (Mega, Macro, Mid, Micro, Nano) and Platform Presence.
```

---

### ⏰ Prompt 2: Setup a Daily Competitor Monitoring Cron Job
```text
Create a Python background script (competitor_cron.py) that runs daily at 8:00 AM.

For a list of competitor brands ['palmonas_official', 'giva.co', 'bluestone', 'mia_by_tanishq']:
1. Call http://localhost:8080/api/v1/intelligence/audit/{brand}?days_back=30.
2. Detect any newly launched Meta Ads or new creator collaborations launched in the last 24 hours.
3. If new ads/creators are detected, send a formatted Slack / Discord webhook alert with creator name, follower count, and ad link.
```
