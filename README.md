# 📊 Instagram Paid Partnerships, Boosted Ads & Creator Intelligence API

An enterprise-grade Python SDK, REST API (FastAPI), and CLI engine to scrape Instagram brand collaboration campaigns, verify Meta Paid Partnership toggles, detect paid media ad spend (boosted reels), classify creator tiers, analyze video content genres, and export executive-ready Excel workbooks and CSV reports.

---

## 📑 Table of Contents
- [✨ Core Capabilities](#-core-capabilities)
- [🏛️ The 4-Tier Collaboration Hierarchy](#️-the-4-tier-collaboration-hierarchy)
- [📈 All Data Points & Information Provided](#-all-data-points--information-provided)
- [🎬 Video Content Genre Taxonomy](#-video-content-genre-taxonomy)
- [👥 Creator Tier Classification](#-creator-tier-classification)
- [🧠 Boost & Ad Spend Detection Engine](#-boost--ad-spend-detection-engine)
- [🛠️ API Methods & SDK Reference](#️-api-methods--sdk-reference)
- [🚀 3 Ways to Use the Engine](#-3-ways-to-use-the-engine)
  - [1. Python SDK](#1-python-sdk)
  - [2. REST API Web Server (FastAPI)](#2-rest-api-web-server-fastapi)
  - [3. Command Line Interface (CLI)](#3-command-line-interface-cli)
- [📋 Complete Output JSON Schema](#-complete-output-json-schema)
- [📗 Master Excel & CSV Export Capabilities](#-master-excel--csv-export-capabilities)

---

## ✨ Core Capabilities

1. **Automatic Collaboration Extraction**: Deeply scans brand feeds and reels streams (`/feed/user/` and `/clips/user/`) to detect co-author producers (`coauthor_producers`), partner-owned reels, and tagged creator campaigns.
2. **Meta Disclosure Verification**: Verifies whether creators activated Meta's official `"Paid partnership"` toggle (`Toggle ON` vs `Toggle OFF`).
3. **Paid Media Ad Spend Detection**: Evaluates live view-to-like ratios, play counts, and follower multipliers to detect boosted influencer ads (identifying campaigns with 1M to 300M+ paid views).
4. **4-Tier Strategic Classification**: Automatically separates high-intent paid campaigns from organic noise.
5. **Creator Profile Enrichment**: Resolves creator follower counts, following, total posts, verified badges, full names, and engagement rates (ER%).
6. **NLP Video Content Genre Detection**: Automatically categorizes reels into 9 creative video formats (Styling, Unboxing, Comedy Skits, Sports Performance, Brand Drops, etc.).
7. **Executive Deliverables**: Generates multi-sheet Excel workbooks with colored tier banners, clickable URLs, executive summary cards, and flat CSV datasets.

---

## 🏛️ The 4-Tier Collaboration Hierarchy

Every collaboration reel is evaluated and categorized into one of four structured tiers:

```
┌───────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                4-TIER COLLABORATION HIERARCHY                                     │
├───────────────────────────────────────────────────────────────────────────────────────────────────┤
│ 🟢 Tier 1: Toggle ON + 🚀 Boosted  │ Formal Paid Partnership Label + Active Paid Ad Spend         │
│ 🟢 Tier 2: Toggle ON + ⚪ Organic  │ Formal Paid Partnership Label + Natural Organic Reach Only   │
│ 🚀 Tier 3: Toggle OFF + 🚀 Boosted │ Co-Author Collab (Toggle OFF) + Heavy Paid Media Ad Spend    │
│ ⚪ Tier 4: Toggle OFF + ⚪ Organic │ Standard Organic Collab / Barter / Low Reach / Noise         │
└───────────────────────────────────────────────────────────────────────────────────────────────────┘
```

* **💎 Total High-Intent Paid Ads**: Sum of **Tier 1 + Tier 2 + Tier 3**
* **📈 High-Intent Paid Adoption Rate %**: `(High-Intent Ads / Total Collabs) * 100`

---

## 📈 All Data Points & Information Provided

For every brand, creator, and collaboration video, the API extracts and computes:

### 1. Brand Profile & Regional Metadata
* `brand.name`: Full display name of the brand.
* `brand.handle`: Instagram handle (e.g. `@skechersindia`, `@giva.co`, `@croma`).
* `brand.state_origin`: Regional headquarters and state of origin (e.g. *Pan-India / Karnataka (HQ: Bengaluru, D2C)*).
* `brand.followers`: Live follower count.
* `brand.pk`: Numeric Instagram user ID.

### 2. Creator Profile & Audience Sizing
* `creator_handle`: Instagram username (e.g. `@kartikaaryan`, `@thevishnukaushal`).
* `full_name`: Creator's full legal/display name.
* `creator_followers`: Exact live follower count.
* `creator_following`: Following count.
* `total_posts`: Total lifetime posts published.
* `is_verified`: Boolean badge status (`True` / `False`).
* `is_business`: Account type (Business / Professional / Personal).
* `creator_tier`: Standardized audience size tier (Mega, Macro, Mid, Micro, Nano).
* `profile_url`: Direct link to the creator's profile (`https://www.instagram.com/{handle}/`).

### 3. Post & Creative Metadata
* `post_date`: Exact publication date (`YYYY-MM-DD`).
* `post_url`: Direct clickable Instagram URL (`https://www.instagram.com/p/{shortcode}/`).
* `shortcode`: Instagram alpha-numeric shortcode.
* `media_id`: Numeric media ID.
* `caption`: Full caption text and hashtags (detecting `#ad`, `#collab`, etc.).
* `video_genre`: Detected content format (e.g. `Styling & OOTD`, `Unboxing & Review`).

### 4. Engagement & Performance Metrics
* `views` / `play_count`: Live video plays or estimated reach.
* `likes`: Exact post like count.
* `comments`: Exact comment count.
* `like_to_view_pct`: Ratio of likes to total views (`(likes / views) * 100`).
* `view_to_follower_multiplier`: Multiple of views relative to creator followers (`views / followers`).
* `creator_er_pct`: Engagement Rate percentage (`((likes + comments) / followers) * 100`).

### 5. Boost Classification & Diagnostic Reasoning
* `is_boosted`: Boolean flag (`True` / `False`).
* `boost_status`: Visual badge (`🚀 Heavily Boosted`, `🔍 Likely Boosted`, `📈 Viral Organic`, `⚪ Standard Organic`).
* `boost_reason`: Plain-English explanation detailing why ad spend was detected *(e.g. "High view count (38,520,394) with sub-0.35% like rate (0.05%) indicates paid video ads campaign")*.

---

## 🎬 Video Content Genre Taxonomy

The engine uses NLP keyword parsing, creator niches, and collaboration context to assign each video into one of 9 formats:

| # | Video Genre / Format | Key Signals & Content Type |
|---|---|---|
| **1** | 🌟 **Celebrity Ambassador Campaign** | High-production brand TVC/DVC cuts, major celebrity ambassador drops. |
| **2** | ⚡ **Athlete & Sports Performance** | Workouts, running drills, cricket pitch tests, athlete endurance demonstrations. |
| **3** | 👗 **Styling & OOTD / GRWM** | "Get Ready With Me", fit-checks, lookbooks, outfit styling guides. |
| **4** | 📦 **Unboxing & Sneaker Review** | Packaging unboxings, material teardowns, on-feet first impressions. |
| **5** | 🎬 **Comedy & Relatable Skit** | Humorous POV skits, couple/friend sketches, viral comedic audio acting. |
| **6** | 🎨 **Design & Craft Storytelling** | Story behind the product silhouette, embroidery, handmade craftsmanship. |
| **7** | 🤝 **Brand Collab Drop / Co-Creation** | Official brand-to-brand limited edition capsule drops (e.g., Uno x Comet). |
| **8** | 🎉 **Event & Pop-Up Activation** | Sneaker festivals, store launch parties, booth walkthroughs. |
| **9** | 👟 **Streetwear Lifestyle & Culture** | Urban streetwear flex, aesthetic b-roll, community culture reels. |

---

## 👥 Creator Tier Classification

Every creator is classified based on pure follower volume:

* 🌟 **Mega Creator / Celebrity (1M+)**: $\ge 1,000,000$ followers
* 🚀 **Macro Creator (100K - 1M)**: $100,000 - 1,000,000$ followers
* ✨ **Mid-Tier Creator (50K - 100K)**: $50,000 - 100,000$ followers
* 🎯 **Micro Creator (10K - 50K)**: $10,000 - 50,000$ followers
* 🌱 **Nano Creator (< 10K)**: $< 10,000$ followers

---

## 🧠 Boost & Ad Spend Detection Engine

Organic Instagram reels follow a natural engagement curve where the **Like-to-View Ratio is typically 2.0% to 8.0%**. 

When a brand runs paid Meta Ads (Dark Ads or Partnership Boosts) behind a creator reel:
1. **Views Skyrocket to Millions**: Media spend pushes the video into non-follower feeds.
2. **Engagement Dilutes Dramatically**: Passive ad viewers watch the video without liking, causing the like-to-view ratio to collapse to **0.01% – 0.50%**.

### Detection Rules:
* **🚀 Heavily Boosted**: $\text{Views} \ge 500,000$ AND $\text{Like-to-View Ratio} < 0.35\%$
* **🚀 Boosted (High Multiplier)**: $\text{View Multiplier} \ge 5.0\text{x}$ AND $\text{Like-to-View Ratio} < 0.70\%$
* **🔍 Likely Boosted**: $\text{View Multiplier} \ge 3.0\text{x}$ AND $\text{Like-to-View Ratio} < 1.00\%$ (with $\text{Views} \ge 80,000$)
* **📈 Viral Organic**: $\text{View Multiplier} \ge 4.0\text{x}$ AND $\text{Like-to-View Ratio} \ge 2.00\%$

---

## 🛠️ API Methods & SDK Reference

The core engine is implemented in [`instagram_paid_collabs_api.py`](instagram_paid_collabs_api.py).

### Class: `InstagramPaidPartnershipEngine(max_workers=15)`

| Method | Parameters | Returns | Description |
|---|---|---|---|
| `analyze_brand()` | `username: str, max_pages: int = 6` | `dict` | Full end-to-end brand audit returning 4 tiers, creator profiles, and summaries. |
| `get_user_id_and_info()` | `username: str` | `dict` | Resolves numeric user ID (`pk`), followers, and profile details. |
| `fetch_brand_feed()` | `user_id: int, max_pages: int = 6` | `list` | Paginates timeline feed items. |
| `fetch_brand_clips()` | `user_id: int, max_pages: int = 6` | `list` | Paginates reels/clips stream. |
| `get_creator_followers()` | `handle: str` | `int` | Resolves live follower count for any creator handle. |
| `evaluate_boost_and_tier()` | `is_paid_toggle, views, likes, comments, followers, caption` | `dict` | Calculates ER%, detects boost status, and assigns Tier 1, 2, 3, or 4. |
| `analyze_posts_custom_list()`| `posts_list: list` | `dict` | Classifies an arbitrary list of post dictionaries. |

---

## 🚀 3 Ways to Use the Engine

### 1. Python SDK

```python
from instagram_paid_collabs_api import InstagramPaidPartnershipEngine

# Initialize engine
engine = InstagramPaidPartnershipEngine(max_workers=15)

# Audit a brand handle
audit = engine.analyze_brand("skechersindia", max_pages=6)

print("Brand:", audit["brand"]["name"])
print("Total Collabs:", audit["summary"]["total_collab_posts"])
print("High-Intent Paid Ads:", audit["summary"]["total_high_intent_paid_ads"])

# Access posts by Tier:
tier_1 = audit["tier_1_toggle_on_boosted"]
tier_2 = audit["tier_2_toggle_on_organic"]
tier_3 = audit["tier_3_toggle_off_boosted"]
tier_4 = audit["tier_4_noise"]

for post in tier_1 + tier_2 + tier_3:
    print(f"[{post['tier_name']}] {post['creator_handle']} | {post['post_date']} | Views: {post['estimated_views']:,} | URL: {post['post_url']}")
```

---

### 2. REST API Web Server (FastAPI)

Launch the REST server:
```bash
python instagram_paid_collabs_api.py --server --port 8000
```

#### Available Endpoints:
* `GET /api/v1/health`: Healthcheck endpoint.
* `GET /api/v1/analyze/brand?username={handle}&max_pages={pages}`: Audit a brand via GET request.
* `POST /api/v1/analyze/brand`: Audit a brand via JSON body `{ "username": "giva.co", "max_pages": 6 }`.
* `POST /api/v1/analyze/posts`: Bulk classify custom post URLs.

#### Sample cURL Request:
```bash
curl -X GET "http://localhost:8000/api/v1/analyze/brand?username=skechersindia&max_pages=5"
```

---

### 3. Command Line Interface (CLI)

Run audits directly in your terminal:

```bash
# Audit a brand and print summary in terminal
python instagram_paid_collabs_api.py --brand giva.co --pages 5

# Audit and export full results to JSON file
python instagram_paid_collabs_api.py --brand skechersindia --output skechers_audit.json

# Start the web server
python instagram_paid_collabs_api.py --server --port 8000
```

---

## 📋 Complete Output JSON Schema

```json
{
  "status": "success",
  "audit_timestamp": "2026-08-25T13:45:00Z",
  "brand": {
    "name": "Skechers India",
    "handle": "@skechersindia",
    "state_origin": "Pan-India / Maharashtra (HQ: Mumbai)",
    "followers": 412000
  },
  "summary": {
    "total_collab_posts": 17,
    "total_unique_creators": 11,
    "tier_1_toggle_on_boosted": { "posts_count": 0, "unique_creators": 0 },
    "tier_2_toggle_on_organic": { "posts_count": 6, "unique_creators": 5 },
    "tier_3_toggle_off_boosted": { "posts_count": 0, "unique_creators": 0 },
    "tier_4_toggle_off_organic_noise": { "posts_count": 11, "unique_creators": 9 },
    "total_high_intent_paid_ads": 6,
    "high_intent_paid_rate_pct": 35.29
  },
  "tier_2_toggle_on_organic": [
    {
      "brand": "Skechers India",
      "brand_handle": "@skechersindia",
      "creator_handle": "@surya_14kumar",
      "creator_followers": 20366023,
      "post_url": "https://www.instagram.com/p/DXZGA7DNlRa/",
      "shortcode": "DXZGA7DNlRa",
      "post_date": "2026-04-21",
      "is_paid_partnership": true,
      "is_boosted": false,
      "tier": 2,
      "tier_name": "🟢 Tier 2: Toggle ON + Organic",
      "tier_description": "Formal Paid Partnership Label + Organic Reach",
      "boost_status": "⚪ Standard Organic",
      "boost_reason": "Baseline organic collab reach",
      "estimated_views": 38520394,
      "likes": 702391,
      "comments": 4120,
      "like_to_view_pct": 1.82,
      "creator_er_pct": 3.47,
      "caption": "Glad to be part of the Skechers family. Playing my game, my way..."
    }
  ]
}
```

---

## 📗 Master Excel & CSV Export Capabilities

The repository includes pre-built scripts to generate formatted Excel spreadsheets and CSV exports with visual tier styling:

1. **`Executive Summary` Tab**: High-level matrix showing State, Total Posts, Unique Creators, Tier 1/2/3/4 counts, High-Intent Total, Avg Views, and Avg ER%.
2. **`Creators Profile Metrics` Tab**: Deduped creator database with Full Name, Verified status, Followers, Following, Posts, Avg Likes, Avg Comments, and Pure Size Tiers.
3. **`All Brands - Master Hierarchy` Tab**: All posts structured into Tier 1 $\rightarrow$ Tier 2 $\rightarrow$ Tier 3 $\rightarrow$ Tier 4 with colored section banners.
4. **Individual Brand Tabs**: Dedicated sheets per brand with top metadata header blocks.

### Pre-Generated Datasets Available in Repo:
* 📗 **Footwear & Sneakers**: [`footwear_sneaker_brands_master_analysis.xlsx`](footwear_sneaker_brands_master_analysis.xlsx)
* 📗 **Jewellery Brands**: [`jewellery_brands_master_analysis.xlsx`](jewellery_brands_master_analysis.xlsx)
* 📗 **Electronics Retail**: [`croma_electronics_master_analysis.xlsx`](croma_electronics_master_analysis.xlsx)
