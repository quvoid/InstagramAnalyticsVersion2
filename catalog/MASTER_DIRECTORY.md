# 📚 MASTER DIRECTORY & PROJECT REGISTRY

Welcome to the **Instagram & Omnichannel Marketing Intelligence Suite**. This repository contains enterprise-grade scrapers, 4-tier collaboration analyzers, creator discovery engines, and executive workbook generators across diverse Indian industries.

---

## 🧭 Repository Quick Map

```
InstagramAnalytics/
├── run.py                                    # 🚀 Natural Language CLI & Dispatcher (Primary Entrypoint)
├── AGENTS.md                                 # 🤖 Comprehensive Guide for AI Coding Agents & LLMs
├── catalog/
│   ├── MASTER_DIRECTORY.md                   # 📖 Full Asset & File Inventory (You are here)
│   └── REGISTRY.json                         # ⚙️ Machine-Readable Intent-to-Script Mapping
├── core/
│   ├── kolkata_engine.py                     # 🏙️ Kolkata & Bengal Creator Discovery Engine (10K+ filter)
│   ├── fintech_engine.py                     # 💳 FinTech 53-Brand 2-Year 4-Tier Collaboration Engine
│   ├── profile_auditor.py                    # 🔍 Live Instagram Profile & Follower Verification
│   └── ad_boost_engine.py                    # 🚀 Boosted Reels & Paid Media Ad Spend Classifier
├── api_wrapper/
│   ├── client.py                             # 🐍 Core Python SDK & Instagram Service
│   ├── cli.py                                # 💻 Native Subcommand CLI
│   └── server.py                             # 🌐 FastAPI REST API Server
├── competitor/                               # 🏢 Competitor Intelligence Sub-Suite (Data, Docs, Scripts)
│   ├── data/
│   ├── docs/
│   ├── scripts/
│   └── workbooks/
└── [Root Datasets & Master Workbooks]        # 📊 Master Deliverable Excel Files & JSON Caches
```

---

## 🎯 The Core Business Verticals

### 1. Kolkata & Bengal Creator Ecosystem
* **Purpose**: Discover, live-audit, and tier authentic Kolkata & West Bengal creators (10K+ followers) for brand partnerships, festive Durga Puja campaigns, food trails, and regional media buying.
* **Key Deliverables**:
  - `bengal_kolkata_authentic_creators_all_inclusive_175.xlsx`: Consolidated 175-creator master workbook across Mega, Macro, Mid, and Micro tiers.
  - `bengal_kolkata_chef_creators_deep_scan_master.xlsx`: 21 verified culinary icons with cross-platform (IG, YT, FB) audience sizing.
  - `kolkata_durga_pujo_1000_regional_creators_master.xlsx`: High-volume Durga Puja campaign roster.
* **Key Scripts**:
  - `core/kolkata_engine.py`: Multi-source discovery (hashtags, hubs, brand collabs, live Playwright audit).
  - `build_kolkata_creator_network_engine.py`: Google Trends validation and regional graph modeling.
  - `deep_scan_kolkata_creators.py`: Deep profile auditing with reel view sampling.

---

### 2. FinTech 4-Tier Partnership & Media Spend Intelligence
* **Purpose**: Scan 53 major Indian FinTech brands across 6 segments (Payments, WealthTech, Neobanks, Digital Lending, InsurTech, B2B SaaS) over a 2-year window (2024–2026).
* **Classification Framework**:
  - **Tier 1**: Meta Paid Partnership Toggle ON + Boosted Ad Spend
  - **Tier 2**: Meta Paid Partnership Toggle ON + Organic Reach
  - **Tier 3**: Meta Paid Partnership Toggle OFF + Heavy Paid Ad Spend (High View/Like Discrepancy)
  - **Tier 4**: Meta Paid Partnership Toggle OFF + Organic / Barter
* **Key Deliverables**:
  - `India_Fintech_4Tier_Partnership_Master.xlsx`: 53 brands benchmarked, 179 collab posts, 143 creators (Zero emojis, 3 tabs).
  - `fintech_collabs_cache.json`: Cached partnership records for instant querying.
* **Key Scripts**:
  - `core/fintech_engine.py`: 2-Year modal stepping scraper with exact datetime cutoff.
  - `scrape_fintech_segment_partnerships.py`: Segment-wise partnership extractor and workbook exporter.

---

### 3. Retail, Jewellery & Consumer Electronics
* **Purpose**: Benchmarking direct-to-consumer and heritage brands on paid media adoption, celebrity ambassadors, and creator tiers.
* **Brands Covered**:
  - *Jewellery*: GRT Jewellers, Oriana, Palmonas, GIVA, Kalyan Jewellers, Tanishq, CaratLane.
  - *Electronics*: Croma, Reliance Digital, Vijay Sales.
  - *Footwear*: Skechers, Puma, Bata, Metro, Campus, Red Chief.
* **Key Deliverables**:
  - `jewellery_brands_consolidated_analysis.xlsx`: Full omnichannel analysis of Indian jewellery brands.
  - `Palmonas_All_4Tier_Master.csv` & `GRT_Oriana_All_4Tier_Master.csv`: Brand-specific 4-tier collaboration logs.
  - `croma_electronics_master_analysis.xlsx`: Comprehensive Croma competitor benchmark.
  - `footwear_sneaker_brands_master_analysis.xlsx`: Sneaker & athletic footwear creator ecosystem.
* **Key Scripts**:
  - `scrape_giva_palmonas.py`, `scrape_grt_oriana.py`, `scrape_footwear_brands.py`
  - `playwright_adlibrary_graphql_scraper.py`: Meta Ad Library extraction.

---

### 4. Hyperlocal Malls, Real Estate & Community Sentiment
* **Purpose**: Footfall intelligence, tenant directory mining, customer review sentiment, and Instagram geotag analysis.
* **Locations**: Pune, Hyderabad, Bangalore, Mumbai (Phoenix Marketcity, Inorbit, Kopa / Lakeshore).
* **Key Deliverables**:
  - `pune_hyderabad_malls_master_analysis.xlsx`: 1-year omnichannel mall performance.
  - `Lake_Shore_Media_Buyer_Intelligence_Engine.xlsx`: Kopa Mall media buying playbook.
* **Key Scripts**:
  - `scrape_malls_1year.py`: 1-year geotagged and brand-tagged content scan.
  - `scrape_popular_times_busyness.py`: Google Maps footfall and busyness index.
  - `scrape_google_maps_qna.py`: Community reviews and friction points.

---

### 5. Education, FMCG & Niche Markets
* **Purpose**: Deep sentiment, student discussions, parental friction, and niche product creator networks.
* **Sectors**:
  - *Higher Education*: Manav Rachna International Institute (`manav_rachna_university_seo_sentiment_master.xlsx`).
  - *Kids Nutrition*: Complan, Pediasure, Bournvita, Little Joys (`kids_nutrition_raw_master_data.xlsx`).
  - *Spiritual Lifestyle*: Rudralife & Rudraksha (`spiritual_rudraksha_brands_master_analysis.xlsx`).
* **Key Scripts**:
  - `build_kids_nutrition_intelligence_engine.py`
  - `scrape_and_build_universities_excel.py`
  - `scrape_spiritual_brands_2years.py`

---

## 🛠️ Core Technology & Architecture

| Component | Technology | Primary Function |
|---|---|---|
| **Headless Browser** | Playwright (Chromium) | Renders dynamic feeds, profiles, modals, and reels grids using authenticated user session cookies. |
| **Fast Profile Scraper** | `curl_cffi` (Chrome 120 impersonation) | Ultra-fast follower resolution and meta tag parsing without browser overhead. |
| **Workbook Styler** | `openpyxl` | Generates enterprise Excel workbooks with custom typography, auto-fit columns, and zero-emoji compliance. |
| **API & CLI** | FastAPI + Typer | Programmatic REST API endpoints and command-line execution. |
| **Natural Language CLI** | `run.py` | Fuzzy intent matching that executes any repository capability via everyday natural language. |

---

## 🔒 Session Cookies & Authentication

All Instagram scraping requires valid session cookies injected into Playwright contexts or `curl_cffi` sessions.
The canonical active cookie parameters are stored in:
- `api_wrapper/client.py` (`DEFAULT_IG_COOKIES`)
- `core/profile_auditor.py`
- `core/kolkata_engine.py`

```python
COOKIES = {
    "sessionid": "25113411270%3AyQFaao428g6Xb9%3A0%3AAYjwS_M5UIDcS-i41tvdVFu8WKSqfzfgEhlZLGMvFg",
    "csrftoken": "3gJbkGDZp99lA8QQ0brobyoHzOreuu8f",
    "mid": "afyCbwALAAFRStE-k17-dfO5_jfa",
    "ds_user_id": "25113411270",
}
```

---

## 📖 How to Run via Natural Language

Simply invoke `run.py` with your prompt in quotes:

```bash
# Discover Kolkata creators with 10K+ followers
python run.py "find kolkata creators with 10k+ followers"

# Audit a specific creator or brand profile
python run.py "audit profile @rjpraveen"

# Run the 2-year deep scan for all 53 FinTech brands
python run.py "scan fintech brands 2 years"

# Detect boosted paid ad spend on a reel
python run.py "detect boost on reel https://www.instagram.com/p/DF22uT5SQ5V/"

# List all master Excel workbooks
python run.py "list master workbooks"
```

For complete instructions on building agents or automating workflows, consult [AGENTS.md](../AGENTS.md).
