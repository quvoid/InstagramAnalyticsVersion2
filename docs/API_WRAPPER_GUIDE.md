# Unified Competitor & Paid Media Intelligence API Wrapper
## Official Python SDK, CLI & FastAPI Microservice (v2.1.0)

> **Production Reference & Developer Manual**  
> Unified Python Wrapper combining **Instagram Private API**, **Facebook Page API**, and **Meta Ad Library (Playwright + GraphQL)** into a single clean interface with **Flexible Date Ranges & Cohort Analysis**.

---

## 📑 Quick Navigation
1. [Supported Date Ranges & Cohorts](#1-supported-date-ranges--cohorts)
2. [Python SDK Usage (In Any Script or Jupyter Notebook)](#2-python-sdk-usage-in-any-script-or-jupyter-notebook)
3. [Running the Unified FastAPI REST Server](#3-running-the-unified-fastapi-rest-server)
4. [Command-Line Interface (CLI) Usage](#4-command-line-interface-cli-usage)
5. [API Endpoint Reference & Schemas](#5-api-endpoint-reference--schemas)

---

## 1. Supported Date Ranges & Cohorts

You can pass human-readable string presets or exact day counts to any endpoint or Python function:

| Preset String | Named Range | Exact Days Audited |
|---|---|:---:|
| `"1w"` or `"7d"` or `"1week"` | **Last 1 Week** | 7 Days |
| `"1m"` or `"30d"` or `"1month"` | **Last 1 Month** | 30 Days |
| `"3m"` or `"90d"` or `"3months"` | **Last 3 Months (Quarter)** | 90 Days |
| `"6m"` or `"180d"` or `"6months"` | **Last 6 Months (Half-Year)** | 180 Days |
| `"1y"` or `"365d"` or `"1year"` | **Last 1 Year** *(Default)* | 365 Days |
| `"2y"` or `"730d"` or `"2years"` | **Last 2 Years** | 730 Days |

Every audit automatically breaks down results into **Historical Date Cohorts**:
* `last_7d`: Posts & collabs published in the last 7 days
* `last_30d`: Posts & collabs published in the last 30 days
* `last_90d`: Posts & collabs published in the last 90 days
* `last_180d`: Posts & collabs published in the last 180 days

---

## 2. Python SDK Usage (In Any Script or Jupyter Notebook)

### Example 1: Full-Funnel 360° Competitor Audit (3 Months / 6 Months / 1 Year)
```python
from api_wrapper import CompetitorIntelligenceClient

client = CompetitorIntelligenceClient()

# Audit for Last 3 Months (or pass '1w', '1m', '6m', '1y')
audit_result = client.audit_brand(
    target_brand="palmonas_official",
    fb_page_id="100076111693972",  # Optional
    time_window="3m",              # '1w', '1m', '3m', '6m', '1y'
    export_excel=True
)

print(f"Total Unique Creators Found: {audit_result['total_unique_creators']}")
print(f"Collabs in Last 7 Days: {audit_result['date_range_cohorts']['last_7d']}")
print(f"Collabs in Last 30 Days: {audit_result['date_range_cohorts']['last_30d']}")
print(f"Excel Deliverable Saved: {audit_result['excel_file']}")
```

---

### Example 2: Instagram Collabs for Last 1 Week
```python
from api_wrapper import CompetitorIntelligenceClient

client = CompetitorIntelligenceClient()

# Scrapes last 7 days of collabs & boosts
res = client.instagram.get_partnerships("giva.co", time_window="1w")

print(f"Collabs in Last Week: {res['total_collab_posts']}")
for c in res['collabs']:
    print(f"- {c['creator_handle']} | Views: {c['views']:,} | {c['partnership_tier']}")
```

---

### Example 3: Meta Ad Library Scraper (Active & Inactive Dark Ads)
```python
from api_wrapper import CompetitorIntelligenceClient

client = CompetitorIntelligenceClient()

ads_data = client.ad_library.search_ads(
    query="Palmonas",
    page_id="100076111693972",
    active_only=False,
    max_scrolls=30
)

print(f"Captured {ads_data['total_ads_captured']} Ads")
print(f"Identified {ads_data['unique_creators_count']} Whitelist Creator Partners")
```

---

## 3. Running the Unified FastAPI REST Server

```bash
uvicorn api_wrapper.server:app --reload --port 8080
```

### Key REST Endpoints with Date Presets:
| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/api/v1/intelligence/audit/{brand_username}?range=3m&export_excel=true` | **Full-Funnel Audit** (`range=1w`, `1m`, `3m`, `6m`, `1y`, `2y`) |
| `GET` | `/api/v1/instagram/partnerships/{brand_username}?range=6m` | Instagram Collabs & Boosts in date window |
| `GET` | `/api/v1/instagram/profile/{username}` | Live Instagram Profile Metrics & Audience Tier |
| `GET` | `/api/v1/meta/ads?query={brand}&active_only=false` | Meta Ad Library Scraper & Whitelists |
| `GET` | `/api/v1/facebook/page/{page_handle}` | Facebook Page Followers & Delegate Page ID |

---

## 4. Command-Line Interface (CLI) Usage

```bash
# 1. Run 360-degree audit for Last 3 Months
python -m api_wrapper.cli audit --brand palmonas_official --range 3m

# 2. Scan Instagram Partnerships for Last 1 Week
python -m api_wrapper.cli insta --brand giva.co --range 1w

# 3. Scan Instagram Partnerships for Last 6 Months
python -m api_wrapper.cli insta --brand giva.co --range 6m

# 4. Search Meta Ad Library
python -m api_wrapper.cli adlib --query Palmonas

# 5. Fetch Facebook Page Info
python -m api_wrapper.cli fb --page zivame
```
