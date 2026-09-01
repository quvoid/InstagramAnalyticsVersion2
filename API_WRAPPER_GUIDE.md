# Unified Competitor & Paid Media Intelligence API Wrapper
## Official Python SDK, CLI & FastAPI Microservice

> **Production Reference & Developer Manual**  
> Unified Python Wrapper combining **Instagram Private API**, **Facebook Page API**, and **Meta Ad Library (Playwright + GraphQL)** into a single clean interface.

---

## 📑 Quick Navigation
1. [Architecture Overview](#1-architecture-overview)
2. [Python SDK Usage (In Any Script or Jupyter Notebook)](#2-python-sdk-usage-in-any-script-or-jupyter-notebook)
3. [Running the Unified FastAPI REST Server](#3-running-the-unified-fastapi-rest-server)
4. [Command-Line Interface (CLI) Usage](#4-command-line-interface-cli-usage)
5. [API Endpoint Reference & Schemas](#5-api-endpoint-reference--schemas)

---

## 1. Architecture Overview

The `api_wrapper` package lives in `c:\Users\omkar\OneDrive\Desktop\InstagramAnalytics\api_wrapper\` and provides 3 interfaces:

```
                                  api_wrapper PACKAGE
                                           │
            ┌──────────────────────────────┼──────────────────────────────┐
            ▼                              ▼                              ▼
    1. PYTHON SDK / CLIENT         2. FASTAPI REST SERVER        3. TERMINAL CLI
    from api_wrapper import        uvicorn api_wrapper.server    python -m api_wrapper.cli
    CompetitorIntelligenceClient   --port 8080                   audit --brand giva.co
```

---

## 2. Python SDK Usage (In Any Script or Jupyter Notebook)

### Example 1: Full-Funnel 360° Competitor Audit (Instagram + Meta Ad Library)
```python
from api_wrapper import CompetitorIntelligenceClient

client = CompetitorIntelligenceClient()

# Runs Instagram 1-Year Scrape + Meta Ad Library Scrape, deduplicates, and generates Excel
audit_result = client.audit_brand(
    target_brand="palmonas_official",
    fb_page_id="100076111693972",  # Optional: speeds up Meta Ad Library
    days_back=365,
    export_excel=True
)

print(f"Total Unique Creators Found: {audit_result['total_unique_creators']}")
print(f"Instagram Collab Posts: {audit_result['total_grid_posts']}")
print(f"Meta Ads Captured: {audit_result['total_meta_ads']}")
print(f"Excel Deliverable Saved: {audit_result['excel_file']}")
```

---

### Example 2: Instagram-Only 1-Year Partnerships & Boost Hierarchy
```python
from api_wrapper import CompetitorIntelligenceClient

client = CompetitorIntelligenceClient()

# Scrapes 365 days of co-authors, reels, and boosts
res = client.instagram.get_partnerships("giva.co", days_back=365)

print(f"Total Collabs: {res['total_collab_posts']}")
for c in res['collabs'][:5]:
    print(f"- {c['creator_handle']} | Views: {c['views']:,} | {c['partnership_tier']}")
```

---

### Example 3: Meta Ad Library Scraper (Active & Inactive Dark Ads)
```python
from api_wrapper import CompetitorIntelligenceClient

client = CompetitorIntelligenceClient()

# Search ads by brand name or keyword
ads_data = client.ad_library.search_ads(
    query="Palmonas",
    page_id="100076111693972",
    active_only=False,
    max_scrolls=30
)

print(f"Captured {ads_data['total_ads_captured']} Ads")
print(f"Identified {ads_data['unique_creators_count']} Whitelist Creator Partners")
for partner in ads_data['creators']:
    print(f"• {partner['name']} (Active Ads: {partner['active_ads']}) -> {partner['sample_ad_url']}")
```

---

### Example 4: Facebook Page Metrics & Delegate Page ID
```python
from api_wrapper import CompetitorIntelligenceClient

client = CompetitorIntelligenceClient()

fb_info = client.facebook.get_page_info("zivame")
print(fb_info)
# Output: {'page_handle': 'zivame', 'page_id': '234603919914240', 'followers_str': '794K', ...}
```

---

## 3. Running the Unified FastAPI REST Server

You can run the wrapper as a standalone HTTP microservice with interactive OpenAPI Swagger documentation.

### Start the Server:
```bash
uvicorn api_wrapper.server:app --reload --port 8080
```

### Interactive Swagger Docs:
Open `http://localhost:8080/docs` in your browser.

### Key REST Endpoints:
| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/api/v1/intelligence/audit/{brand_username}?page_id={id}&days_back=365` | **Full-Funnel Audit** (IG Grid + Meta Ad Library + Excel Export) |
| `GET` | `/api/v1/instagram/partnerships/{brand_username}?days_back=365` | Instagram 1-Year Collabs & 4-Tier Boosts |
| `GET` | `/api/v1/instagram/profile/{username}` | Live Instagram Profile Metrics & Audience Tier |
| `GET` | `/api/v1/meta/ads?query={brand}&page_id={id}&active_only=false` | Meta Ad Library Scraper & Whitelists |
| `GET` | `/api/v1/facebook/page/{page_handle}` | Facebook Page Followers & Delegate Page ID |

---

## 4. Command-Line Interface (CLI) Usage

The wrapper includes a CLI tool for fast terminal operations:

```bash
# 1. Run full 360-degree competitor audit
python -m api_wrapper.cli audit --brand palmonas_official --page-id 100076111693972 --days 365

# 2. Search Meta Ad Library
python -m api_wrapper.cli adlib --query Palmonas --page-id 100076111693972

# 3. Scan Instagram Partnerships
python -m api_wrapper.cli insta --brand giva.co --days 365

# 4. Fetch Facebook Page Info
python -m api_wrapper.cli fb --page zivame
```

---

## 5. API Endpoint Reference & Schemas

### Audit Response Schema (`GET /api/v1/intelligence/audit/{brand_username}`):
```json
{
  "brand": "palmonas_official",
  "total_unique_creators": 154,
  "instagram_grid_creators": 122,
  "meta_adlibrary_creators": 32,
  "total_grid_posts": 143,
  "total_meta_ads": 233,
  "unified_creators": [
    {
      "handle": "@shraddhakapoor",
      "name": "shraddhakapoor",
      "on_instagram_grid": true,
      "on_meta_adlibrary": true,
      "total_grid_posts": 12,
      "total_grid_views": 22233142,
      "active_meta_ads": 19,
      "total_meta_ads": 23,
      "sample_grid_url": "https://www.instagram.com/p/C-h902-sR0z/",
      "sample_ad_url": "https://www.facebook.com/ads/library/?id=616056441550193"
    }
  ],
  "excel_file": "palmonas_official_unified_competitor_master.xlsx"
}
```
