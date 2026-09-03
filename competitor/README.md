# 🏢 Lake Shore & Competitor Intelligence Engine

> **Client Brand**: **Lake Shore India Advisory** (`@kopapune` — KOPA Mall Pune & `@lakeshoreyjunction` — Lake Shore Y Junction Hyderabad)  
> **Target Markets**: Pune & Hyderabad (12 Major Shopping Centers & Regional Malls Audited)  
> **Purpose**: End-to-End Competitor Intelligence, Real-Time Footfall Modeling, Customer Sentiment Mining, and Agency-Grade Paid Media Blueprint.

---

## 📁 Repository & Folder Hierarchy

```text
competitor/
├── README.md                              <- Full documentation, scraping protocols & architecture
├── data/                                  <- Raw JSON & Excel databases
│   ├── all_malls_reviews.xlsx             <- 23,218 Google Maps Verified Reviews across 12 malls
│   ├── youtube_mall_master_analysis.xlsx  <- 180 YouTube Videos, Shorts & Transcripts
│   ├── pune_hyderabad_malls_1year_dataset.json <- 4,076 Instagram Posts, 1,848 Collabs, 586 Creators
│   ├── real_mall_comments_dataset.json    <- 1,524 Scraped Real Instagram Comments with NLP Tags
│   ├── real_mall_meta_ads_dataset.json    <- 425 Scraped Meta Ad Library Live/Inactive Cards
│   ├── google_autocomplete_intent_dataset.json <- 1,324 Google Autocomplete Search Queries
│   ├── instagram_hashtag_ugc_dataset.json <- 12-Mall Hashtag Volume & UGC Share of Voice
│   ├── google_ads_transparency_intelligence.json <- Competitor Search Text Ads & Extensions
│   ├── google_popular_times_busyness_dataset.json <- 24h x 7-Day Capacity Profiles & Ad Triggers
│   ├── mall_tenant_brand_directories_dataset.json <- 800+ Stores Roster & KOPA Exclusivity Tags
│   ├── google_maps_community_qna_dataset.json <- Google Community Q&A & Search Ad Mappings
│   ├── google_place_ids_master.json       <- Verified Google Place IDs (ChIJ...)
│   └── master_mall_place_ids.json         <- Master Coordinate & Address Directory
├── workbooks/                             <- Master Deliverables & Multi-Tab Analysis Workbooks
│   ├── pune_hyderabad_malls_master_analysis.xlsx <- 27-Tab Master Database with 7 Native Charts
│   ├── lakeshore_agency_media_buying_blueprint.xlsx <- 12-Tab Agency Media Buying Engine
│   ├── lakeshore_executive_learning_and_insights.xlsx <- 3-Tab Executive Leadership Summary
│   ├── lakeshore_advanced_competitor_intelligence.xlsx <- 3-Tab Autocomplete, Hashtag & Google Ads
│   ├── lakeshore_realtime_operational_intelligence.xlsx <- 3-Tab Popular Times, Tenants & Q&A
│   ├── lakeshore_omnichannel_competitor_master.xlsx <- 5-Tab Omnichannel Competitor Overview
│   ├── lakeshore_real_competitor_research_master.xlsx <- 5-Tab Qualitative Comments & Meta Ads
│   ├── lakeshore_kopa_paid_media_intelligence.xlsx <- 7-Tab Strategic Paid Intelligence
│   └── Lake_Shore_Media_Buyer_Intelligence_Engine.xlsx <- Performance Media Planning Sheet
├── scripts/                               <- Production Scraping & Excel Builder Pipeline
│   ├── scrape_and_integrate_lakeshore.py  <- Ingests 1-year Instagram posts & collabs
│   ├── scrape_and_mine_real_comments.py   <- Scrapes & NLP-classifies Instagram comments
│   ├── scrape_malls_meta_ads_real.py      <- Playwright scraper for Meta Ad Library
│   ├── scrape_malls_youtube.py            <- YouTube transcript & Shorts scraper (yt-dlp)
│   ├── scrape_google_autocomplete.py      <- Google search autocomplete intent miner
│   ├── scrape_instagram_hashtags.py       <- Instagram branded hashtag & SOV auditor
│   ├── scrape_google_ads_intelligence.py  <- Google Ads Transparency Center analyzer
│   ├── scrape_popular_times_busyness.py   <- Google Popular Times 24h busyness modeler
│   ├── scrape_mall_tenant_directories.py  <- 12-mall store directory & exclusivity parser
│   ├── scrape_google_maps_qna.py          <- Google Community Q&A scraper & ad generator
│   ├── build_malls_master_excel.py        <- Builds base 16-tab master workbook
│   ├── build_agency_media_blueprint.py    <- Builds 6-tab agency media blueprint
│   ├── expand_agency_blueprint.py         <- Expands agency blueprint to 12 tabs
│   ├── build_executive_learning_excel.py  <- Builds 3-tab executive dashboard
│   ├── build_advanced_intelligence_excel.py <- Compiles Autocomplete & Google Ads suite
│   ├── build_realtime_operational_excel.py <- Compiles Popular Times & Q&A suite
│   └── enhance_master_workbook.py         <- Ingests all data and appends 11 tabs with charts
└── docs/                                  <- Master AI Agent Prompts & Briefs
    ├── MEDIA_BUYER_AGENT_PROMPT.md        <- Master System Prompt for Paid Media Buyers
    └── YOUTUBE_AGENT_PROMPT.md            <- Master System Prompt for YouTube Video Intelligence
```

---

## 🛠️ How We Scraped the Data: Authentication, APIs & Login Protocols

### 1. 📱 Instagram 1-Year Feed, Collabs & Real Comments Scraping:
* **Tool / Method**: Direct Instagram Mobile GraphQL & REST Endpoints (`api/v1/feed/user/{user_id}/`, `api/v1/media/{media_id}/comments/`).
* **Authentication / Session Protocol**:
  - Uses authenticated Instagram web/mobile session headers (`x-ig-app-id: 936619743392459`, `User-Agent: Instagram 278.0.0.19.115`, `sessionid` cookie).
  - Shortcodes (e.g. `DbcYysaCmFH`) are converted to integer Media IDs via base64 decoding:
    $$\text{Media ID} = \sum_{i=0}^{n-1} (\text{AlphabetIndex}(s[i]) \times 64^{n-1-i})$$
  - Pagination is handled via `max_id` query parameters until 365 days of post history is collected.

### 2. 🛡️ Meta Ad Library (425 Real Ads Extracted):
* **Tool / Method**: Headless Browser Automation via **Playwright** (`playwright.chromium`).
* **Scraping Protocol**:
  - Target URL: `https://www.facebook.com/ads/library/?active_status=all&ad_type=all&country=IN&q={mall_name}`
  - Bypasses Meta dynamic obfuscation by targeting parent container blocks splitting on regex `Library ID:\s*(\d+)`.
  - Extracts Advertiser Name, Start Date, Active / Inactive Status, Primary Copy, CTA Button text, and Media Asset types without requiring Meta Graph API developer tokens.

### 3. 📍 Google Maps 23,218 Reviews & Place IDs:
* **Tool / Method**: Google Places REST API (`places.googleapis.com/v1/places/{placeId}`) + ScrapePlaces review extraction pipeline.
* **Scraping Protocol**:
  - Coordinates & Place IDs (`ChIJ...`) verified across Pune & Hyderabad.
  - Extracted 22 data fields per review: Star rating (1★–5★), verbatim text, reviewer Local Guide badge level, owner responses, and 15,000+ customer photo URLs.
  - Sentiment & NLP topic classification mapped into 6 buckets: Parking, Dining, Hygiene, Crowd, Retail Variety, and Cinema.

### 4. 🔍 Google Autocomplete & Search Intent Mining (1,324 Queries):
* **Tool / Method**: Google Query Suggestion Engine (`https://suggestqueries.google.com/complete/search?client=chrome&q=...`).
* **Scraping Protocol**:
  - Iterated 21 intent-modifying seeds (`parking charges`, `valet`, `brands`, `restaurants`, `movie`, `timings`, `entry fee`, `zara`, `sephora`, `owner`, `how to reach`) across all 12 mall entities.
  - Mined 1,324 distinct autocomplete suggestions categorized into high-purchase intent, conquest search terms, and strict negative keywords.

### 5. 🎥 YouTube Videos, Shorts & Transcripts (180 Videos):
* **Tool / Method**: YouTube Data v3 API & `yt-dlp` Speech-to-Text Transcript Extractor.
* **Scraping Protocol**:
  - Queried 1-year video rosters and Shorts across 12 malls.
  - Extracted video duration, view velocity, like counts, channel subscriber scale, and full spoken transcripts to detect tenant brand mentions (Zara, Sephora, Armani, PVR Director's Cut).

### 6. 🕐 Google Popular Times & Hourly Capacity Modeling:
* **Tool / Method**: Google Maps Live Busyness Graph Extraction.
* **Protocol**:
  - Extracted Monday-to-Sunday 24-hour capacity percentages (0%–100%) and dwell times to establish real-time ad triggers when competitors hit $>80\%$ capacity on weekends.

---

## 📊 Summary of Master Workbooks Generated

| Workbook File | Total Tabs | Key Target Use Case |
|---|:---:|---|
| [`pune_hyderabad_malls_master_analysis.xlsx`](workbooks/pune_hyderabad_malls_master_analysis.xlsx) | **27 Tabs** | Complete raw data warehouse + 7 color-coded native interactive Excel charts. |
| [`lakeshore_agency_media_buying_blueprint.xlsx`](workbooks/lakeshore_agency_media_buying_blueprint.xlsx) | **12 Tabs** | Agency execution blueprint: Taxonomy, Geofences, Ad Copy Scripts, Top 15 Audio Tracks, Tenant Co-Op Matching, ManyChat DM Funnels. |
| [`lakeshore_executive_learning_and_insights.xlsx`](workbooks/lakeshore_executive_learning_and_insights.xlsx) | **3 Tabs** | Clean executive digest for leadership: Key takeaways, competitor cheat sheets, 4-step execution roadmap. |
| [`lakeshore_advanced_competitor_intelligence.xlsx`](workbooks/lakeshore_advanced_competitor_intelligence.xlsx) | **3 Tabs** | Autocomplete intent engine, Instagram hashtag UGC share of voice, and Google Ads Transparency audit. |
| [`lakeshore_realtime_operational_intelligence.xlsx`](workbooks/lakeshore_realtime_operational_intelligence.xlsx) | **3 Tabs** | Popular Times 24h capacity curves, 800+ store tenant directory with KOPA exclusivity tags, and Google Community Q&A. |

---

## 🚀 How to Re-Run Any Scraper or Pipeline

From the root project directory, run:

```bash
# 1. Scrape Google Autocomplete Search Queries
python competitor/scripts/scrape_google_autocomplete.py

# 2. Audit Instagram Hashtag Volume & UGC SOV
python competitor/scripts/scrape_instagram_hashtags.py

# 3. Analyze Google Ads Transparency Center
python competitor/scripts/scrape_google_ads_intelligence.py

# 4. Generate Google Popular Times Capacity Curves
python competitor/scripts/scrape_popular_times_busyness.py

# 5. Extract Mall Tenant Brand Directories
python competitor/scripts/scrape_mall_tenant_directories.py

# 6. Scrape Google Community Q&A
python competitor/scripts/scrape_google_maps_qna.py

# 7. Rebuild the 12-Tab Agency Media Buying Blueprint
python competitor/scripts/expand_agency_blueprint.py

# 8. Rebuild the 27-Tab Master Multi-Mall Analysis Workbook
python competitor/scripts/enhance_master_workbook.py
```
