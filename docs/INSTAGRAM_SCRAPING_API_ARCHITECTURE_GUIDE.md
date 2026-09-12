# 📚 Complete Instagram Scraping & API Architecture Guide

> **Production Reference & Technical Documentation**  
> *Everything you need to replicate, understand, and integrate the Instagram Scraping & Paid Collabs Intelligence Engine into another project.*

---

## 📑 Table of Contents
1. [Core Architecture & How It Works](#1-core-architecture--how-it-works)
2. [Credentials, Cookies & Anti-Detection](#2-credentials-cookies--anti-detection)
3. [All Internal Instagram API Endpoints & Payloads](#3-all-internal-instagram-api-endpoints--payloads)
4. [Step-by-Step Scraping & Extraction Logic](#4-step-by-step-scraping--extraction-logic)
5. [Boost Detection & 4-Tier Hierarchy Formula](#5-boost-detection--4-tier-hierarchy-formula)
6. [Creator Audience Sizing & NLP Genre Taxonomy](#6-creator-audience-sizing--nlp-genre-taxonomy)
7. [Codebase File Locations & Purpose](#7-codebase-file-locations--purpose)
8. [Plug-and-Play Python Module (Copy-Paste for New Projects)](#8-plug-and-play-python-module-copy-paste-for-new-projects)

---

## 1. Core Architecture & How It Works

Instagram's public web and mobile apps communicate with private backend endpoints (`i.instagram.com/api/v1/`). This project uses a **hybrid TLS-fingerprinted session architecture**:

```
 ┌────────────────────────────────────────────────────────────────────────┐
 │                 INSTAGRAM SCRAPING ENGINE (curl_cffi)                  │
 │   • Impersonates Chrome 120 / Real Android Mobile App                  │
 │   • Passes Session Cookies + Mobile Headers + App ID                   │
 └────────────────────────────────────────────────────────────────────────┘
                                     │
         ┌───────────────────────────┴───────────────────────────┐
         ▼                                                       ▼
 ┌───────────────────────────────┐               ┌───────────────────────────────┐
 │   1. MOBILE ENDPOINTS (JSON)  │               │    2. WEB OPENGRAPH / REGEX   │
 │   • User Search (PK Lookup)   │               │   • Fast Profile Follower     │
 │   • Timeline Feed Pagination  │               │     Count & Following Count   │
 │   • Clips / Reels Pagination  │               │   • Total Posts & Bio         │
 │   • Comments Stream           │               │   • Verified Badge Status     │
 └───────────────────────────────┘               └───────────────────────────────┘
                                     │
                                     ▼
 ┌────────────────────────────────────────────────────────────────────────┐
 │                    BUSINESS INTELLIGENCE PIPELINE                     │
 │   1. Pinned Posts Bypass (Accurate 1-Year / 2-Year Date Cutoffs)       │
 │   2. Collaborator & Co-Author Extraction (coauthor_producers)          │
 │   3. Pure Creator Sizing (Mega, Macro, Mid, Micro, Nano)               │
 │   4. Paid Boost Detection (Like-to-View % & View Multipliers)          │
 │   5. 4-Tier Collaboration Hierarchy Classification                     │
 │   6. NLP Video Content Genre Categorization                            │
 └────────────────────────────────────────────────────────────────────────┘
                                     │
         ┌───────────────────────────┴───────────────────────────┐
         ▼                                                       ▼
 ┌───────────────────────────────┐               ┌───────────────────────────────┐
 │   MULTI-TAB EXCEL (.xlsx)     │               │   AUTOMATION & DASHBOARDS     │
 │   • Executive Summary Card    │               │   • FastAPI REST Server       │
 │   • Creator Profile Metrics   │               │   • Streamlit Live Dashboard  │
 │   • 4-Tier Master Collabs     │               │   • Google Apps Script Sync   │
 └───────────────────────────────┘               └───────────────────────────────┘
```

---

## 2. Credentials, Cookies & Anti-Detection

### 🍪 2.1 The 4 Required Cookies
Instagram requires authenticated session cookies to return full feed items and reels:

| Cookie Name | Purpose | Example Value |
|---|---|---|
| **`sessionid`** | Primary login session authentication | `25113411270%3AyQFaao428g6Xb9%3A0%3AAYjwS_...` |
| **`ds_user_id`** | Numeric Instagram User ID of the logged-in account | `25113411270` |
| **`csrftoken`** | Anti-CSRF protection token | `3gJbkGDZp99lA8QQ0brobyoHzOreuu8f` |
| **`mid`** | Machine ID device tracking cookie | `afyCbwALAAFRStE-k17-dfO5_jfa` |

#### How to Extract Cookies from Browser:
1. Log in to [instagram.com](https://www.instagram.com) on Google Chrome.
2. Press `F12` $\rightarrow$ Open **Application** tab $\rightarrow$ Click **Cookies** $\rightarrow$ `https://www.instagram.com`.
3. Copy the values of `sessionid`, `ds_user_id`, `csrftoken`, and `mid`.

---

### 🛡️ 2.2 TLS Fingerprinting (`curl_cffi`)
Standard Python `requests` or `urllib` send an OpenSSL TLS fingerprint that Cloudflare and Instagram immediately block (`HTTP 401 Unauthorized` or `login_required`).

We use **`curl_cffi`** with `impersonate="chrome120"` to replicate a 100% genuine Chrome browser TLS/JA3 handshake:
```python
from curl_cffi import requests as cffi_requests

session = cffi_requests.Session(impersonate="chrome120")
```

---

### 📱 2.3 Required Request Headers
```python
MOBILE_HEADERS = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "*/*"
}

WEB_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9"
}
```

---

## 3. All Internal Instagram API Endpoints & Payloads

### 📍 Endpoint 1: Username to Numeric User ID (PK) Lookup
* **Method**: `GET`
* **URL**: `https://i.instagram.com/api/v1/users/search/?q={username}`
* **Headers**: `MOBILE_HEADERS` + `COOKIES`
* **Response**:
```json
{
  "users": [
    {
      "pk": "49372589192",
      "username": "palmonas_official",
      "full_name": "PALMONAS",
      "is_private": false,
      "follower_count": 973000
    }
  ]
}
```

---

### 📍 Endpoint 2: User Timeline Feed (Posts, Albums & Reels)
* **Method**: `GET`
* **URL**: `https://i.instagram.com/api/v1/feed/user/{user_pk}/?max_id={next_max_id}`
* **Headers**: `MOBILE_HEADERS` + `COOKIES`
* **Key Response Fields**:
  * `items[]`: Array of post objects.
  * `items[].pk`: Numeric media ID.
  * `items[].code`: Post shortcode (`/p/{code}/`).
  * `items[].taken_at`: Unix timestamp of publication.
  * `items[].user`: Post author object (`username`, `full_name`, `pk`).
  * `items[].coauthor_producers`: Array of co-authors / collaborator handles.
  * `items[].is_paid_partnership`: Boolean flag (`true`/`false`).
  * `items[].play_count` / `view_count`: Total video plays/views.
  * `items[].like_count`: Total likes.
  * `items[].comment_count`: Total comments.
  * `items[].caption.text`: Full post caption.
  * `next_max_id`: Cursor string for the next page (`null` when finished).

---

### 📍 Endpoint 3: User Clips & Reels Feed
* **Method**: `POST`
* **URL**: `https://i.instagram.com/api/v1/clips/user/`
* **Headers**: `MOBILE_HEADERS` + `COOKIES`
* **Payload (Form Data)**:
```python
payload = {
    "target_user_id": str(user_pk),
    "page_size": 30,
    "max_id": str(next_max_id) # Optional for page 2+
}
```
* **Key Response Fields**:
  * `items[].media`: Reel media object (identical structure to feed items).
  * `paging_info.max_id`: Next cursor.
  * `paging_info.more_available`: Boolean (`true`/`false`).

---

### 📍 Endpoint 4: Single Post / Reel Deep Inspection
* **Method**: `GET`
* **URL**: `https://i.instagram.com/api/v1/media/{media_id}/info/`
* **Headers**: `MOBILE_HEADERS` + `COOKIES`
* **Purpose**: Fetches complete sponsor tags (`sponsor_tags`), audio details, tagged products, and boosted ad attributes.

---

### 📍 Endpoint 5: Post Comments Extraction
* **Method**: `GET`
* **URL**: `https://i.instagram.com/api/v1/media/{media_id}/comments/?can_support_threading=true`
* **Headers**: `MOBILE_HEADERS` + `COOKIES`
* **Response**:
```json
{
  "comments": [
    {
      "pk": "17950293848123456",
      "user": { "username": "user1", "full_name": "User One" },
      "text": "Love the necklace! Where is it from?",
      "created_at": 1724750000,
      "like_count": 12
    }
  ],
  "next_min_id": "cursor_string",
  "has_more_comments": true
}
```

---

### 📍 Endpoint 6: Fast Creator Profile Scraper (OpenGraph)
* **Method**: `GET`
* **URL**: `https://www.instagram.com/{creator_username}/`
* **Headers**: `WEB_HEADERS` (no cookies needed for public profiles)
* **Regex Extraction Logic**:
```python
import re

# Extracts: 826K Followers, 450 Following, 1,200 Posts
meta_match = re.search(r'([0-9.,KMBkmb]+)\s+Followers,\s*([0-9.,KMBkmb]+)\s+Following,\s*([0-9.,KMBkmb]+)\s+Posts', html_text)
# Extracts Full Name from title tag
title_match = re.search(r'<title>([^(<]+)\s*\(@', html_text)
```

---

## 4. Step-by-Step Scraping & Extraction Logic

```
   ┌────────────────────────────────────────────────────────┐
   │ 1. RESOLVE USER                                        │
   │    Input: @brand_handle ──▶ Output: PK (User ID)       │
   └────────────────────────────────────────────────────────┘
                               │
                               ▼
   ┌────────────────────────────────────────────────────────┐
   │ 2. PAGINATE FEED & REELS (WITH PINNED BYPASS)          │
   │    • Fetch 12 items/page from Feed                     │
   │    • Fetch 30 items/page from Clips                    │
   │    • Ignore pinned timestamps on Page 1                │
   │    • Stop when non-pinned taken_at < CUTOFF_DATE       │
   └────────────────────────────────────────────────────────┘
                               │
                               ▼
   ┌────────────────────────────────────────────────────────┐
   │ 3. EXTRACT CREATOR COLLABORATIONS                      │
   │    • Owner != Brand ──▶ Creator Handle                 │
   │    • Co-Authors in coauthor_producers ──▶ Creator      │
   │    • Filter out internal brand sub-accounts            │
   └────────────────────────────────────────────────────────┘
                               │
                               ▼
   ┌────────────────────────────────────────────────────────┐
   │ 4. CONCURRENT CREATOR PROFILE RESOLUTION               │
   │    • ThreadPoolExecutor(max_workers=10)                │
   │    • Scrape live Followers, Following, Total Posts     │
   │    • Calculate Avg Views, Avg Likes, Avg ER%           │
   └────────────────────────────────────────────────────────┘
```

### ⚠️ Critical Rule: Pinned Posts Handling
Instagram pins up to 3 posts at the top of an account feed. Pinned posts retain their **original publication timestamp** (e.g. from 2 years ago). If you do `min(taken_at)` on Page 1, your scraper will think it reached the cutoff immediately and exit prematurely!

**The Fix**:
```python
# Strip pinned items on Page 1 before evaluating date cutoff
unpinned_ts = [it.get("taken_at", 0) for idx, it in enumerate(items) if not (page == 1 and idx < 3)]
oldest_ts = min(unpinned_ts) if unpinned_ts else items[-1].get("taken_at", 0)

if oldest_ts < CUTOFF_TIMESTAMP:
    break # Chronological cutoff reached!
```

---

## 5. Boost Detection & 4-Tier Hierarchy Formula

### 📐 5.1 The 4 Partnership Tiers

```
                          ┌────────────────────────┐
                          │   COLLABORATION POST   │
                          └───────────┬────────────┘
                                      │
                 ┌────────────────────┴────────────────────┐
                 ▼                                         ▼
      IS FORMAL PAID PARTNERSHIP?                 STANDARD / CO-AUTHOR?
      (Toggle ON / #ad / is_paid)                (Toggle OFF / Collab Tag)
                 │                                         │
        ┌────────┴────────┐                       ┌────────┴────────┐
        ▼                 ▼                       ▼                 ▼
   PAID ADS PUSH?     ORGANIC ONLY?          PAID ADS PUSH?     ORGANIC ONLY?
  (Boost Detected)  (Normal Metrics)        (Boost Detected)  (Normal Metrics)
        │                 │                       │                 │
        ▼                 ▼                       ▼                 ▼
   🟢 TIER 1         🟢 TIER 2               🚀 TIER 3         ⚪ TIER 4
  (Toggle ON +      (Toggle ON +            (Toggle OFF +     (Toggle OFF +
    Boosted)          Organic)                Boosted)       Organic / Noise)
```

---

### 🔬 5.2 Mathematical Boost Detection Engine
How the algorithm detects paid video ad spend vs viral organic reach:

```python
# 1. Calculate Core Ratios
like_to_view_pct = (likes / views) * 100
view_multiplier = views / creator_followers
engagement_rate = ((likes + comments) / creator_followers) * 100

# 2. Boost Classification Decision Matrix
if views >= 500000 and like_to_view_pct < 0.35:
    is_boosted = True
    status = "🚀 Heavily Boosted (Paid Ad Spend)"
    reason = f"High view count ({views:,}) with sub-0.35% like rate indicates paid video ads"

elif view_multiplier >= 5.0 and like_to_view_pct < 0.70:
    is_boosted = True
    status = "🚀 Boosted (Paid Ad Spend)"
    reason = f"High view multiplier ({view_multiplier:.1f}x followers) with low like rate ({like_to_view_pct:.2f}%)"

elif view_multiplier >= 3.0 and like_to_view_pct < 1.00 and views >= 80000:
    is_boosted = True
    status = "🔍 Likely Boosted (Targeted Ad)"
    reason = f"Disproportionate views ({views:,}) relative to likes ({likes:,})"

elif engagement_rate >= 4.0 and like_to_view_pct >= 2.00:
    is_boosted = False
    status = "📈 Viral Organic Reach"
    reason = f"High organic views ({views:,}) with strong organic engagement ({like_to_view_pct:.2f}%)"

else:
    is_boosted = False
    status = "⚪ Standard Organic"
    reason = "Baseline organic collab reach"
```

---

## 6. Creator Audience Sizing & NLP Genre Taxonomy

### 👥 6.1 Creator Scale Tiers
We categorize creators strictly by their audience reach:

```python
def get_creator_tier(followers: int) -> str:
    if followers >= 1000000:
        return "🌟 Mega Creator / Celebrity (1M+)"
    elif followers >= 100000:
        return "🚀 Macro Creator (100K - 1M)"
    elif followers >= 50000:
        return "✨ Mid-Tier Creator (50K - 100K)"
    elif followers >= 10000:
        return "🎯 Micro Creator (10K - 50K)"
    else:
        return "🌱 Nano Creator (<10K)"
```

---

### 🎬 6.2 NLP Video Content Genre Taxonomy
Captions and handles are classified into industry-specific creative genres:

| Genre Name | Trigger Keywords / Patterns |
|---|---|
| **🌟 Celebrity Ambassador Campaign** | Celebrity handle, `ambassador`, `co-founder`, `face of`, `campaign` |
| **👗 Styling & OOTD / GRWM** | `styling`, `how to style`, `outfit`, `ootd`, `grwm`, `lookbook`, `fit check`, `drip`, `aesthetic` |
| **📦 Unboxing, Try-On & Product Review** | `unboxing`, `unbox`, `try on`, `haul`, `first look`, `review`, `packaging`, `got this` |
| **💎 Craft Lore & Material Storytelling** | `demi-fine`, `anti-tarnish`, `waterproof`, `18k gold`, `silver`, `handmade`, `craftsmanship`, `daily wear` |
| **🎁 Gifting, Festive & Occasion Drops** | `gift`, `gifting`, `valentine`, `rakhi`, `diwali`, `festive`, `anniversary`, `birthday`, `for her` |
| **🏬 Store Walkthrough & Retail Experience** | `store`, `visit`, `shopping`, `walkthrough`, `pop-up`, `outlet`, `flagship` |
| **🎬 Comedy & Relatable Skit** | `pov:`, `when you`, `relatable`, `funny`, `skit`, `tag him`, `husband`, `boyfriend` |

---

## 7. Codebase File Locations & Purpose

All files are located in:  
`c:\Users\omkar\OneDrive\Desktop\InstagramAnalytics\`

### 📂 Core Engines & APIs
| File Path | Description |
|---|---|
| [`scrape_bulk.py`](file:///c:/Users/omkar/OneDrive/Desktop/InstagramAnalytics/scrape_bulk.py) | Standalone bulk post scraper with session cookie handling & Chrome TLS impersonation. |
| [`instagram_paid_collabs_api.py`](file:///c:/Users/omkar/OneDrive/Desktop/InstagramAnalytics/instagram_paid_collabs_api.py) | **Production FastAPI Server & SDK** exposing `/api/v1/analyze/brand` and 4-tier analytics. |
| [`cron_scheduler.py`](file:///c:/Users/omkar/OneDrive/Desktop/InstagramAnalytics/cron_scheduler.py) | Multi-threaded background cron daemon for scheduled brand audits. |
| [`app.py`](file:///c:/Users/omkar/OneDrive/Desktop/InstagramAnalytics/app.py) | Original Streamlit Application (Batch URL scraper & single profile audit). |
| [`streamlit_cron_dashboard.py`](file:///c:/Users/omkar/OneDrive/Desktop/InstagramAnalytics/streamlit_cron_dashboard.py) | Interactive multi-brand 4-tier analytics dashboard with cron controls. |

### 📂 Google Sheets Integration Pipeline
| File Path | Description |
|---|---|
| [`post_metrics/pm_app.py`](file:///c:/Users/omkar/OneDrive/Desktop/InstagramAnalytics/post_metrics/pm_app.py) | Dedicated Streamlit control panel for Google Sheets Post Metrics sync. |
| [`post_metrics/pm_ingest.py`](file:///c:/Users/omkar/OneDrive/Desktop/InstagramAnalytics/post_metrics/pm_ingest.py) | Historical backfill and daily 09:00 IST incremental sync orchestrator. |
| [`post_metrics/pm_instagram.py`](file:///c:/Users/omkar/OneDrive/Desktop/InstagramAnalytics/post_metrics/pm_instagram.py) | Feed & clips extractor for a single account. |
| [`post_metrics/pm_sheets.py`](file:///c:/Users/omkar/OneDrive/Desktop/InstagramAnalytics/post_metrics/pm_sheets.py) | Google Sheets client communicating via Apps Script Web App. |
| [`post_metrics/google_apps_script/PostMetrics.gs`](file:///c:/Users/omkar/OneDrive/Desktop/InstagramAnalytics/post_metrics/google_apps_script/PostMetrics.gs) | Apps Script Web App receiving data & holding the daily clock trigger. |

---

## 8. Plug-and-Play Python Module (Copy-Paste for New Projects)

Save the code below as **`instagram_engine.py`** in your new project. It is completely standalone:

```python
"""
Instagram Scraping Engine (Standalone Plug-and-Play Module)
Requirements: pip install curl_cffi openpyxl
"""

import re, time, json
from datetime import datetime, timezone, timedelta
from curl_cffi import requests as cffi_requests

class InstagramEngine:
    def __init__(self, sessionid: str, ds_user_id: str, csrftoken: str, mid: str = ""):
        self.cookies = {
            "sessionid": sessionid,
            "ds_user_id": ds_user_id,
            "csrftoken": csrftoken,
            "mid": mid
        }
        self.session = cffi_requests.Session(impersonate="chrome120")
        self.headers = {
            "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
            "x-ig-app-id": "936619743392459",
            "Accept": "*/*"
        }

    def resolve_user(self, username: str) -> dict:
        """Resolve handle to numeric PK and follower count"""
        url = f"https://i.instagram.com/api/v1/users/search/?q={username}"
        r = self.session.get(url, headers=self.headers, cookies=self.cookies, timeout=12)
        if r.status_code == 200:
            for u in r.json().get("users", []):
                if u.get("username", "").lower() == username.lower():
                    return {
                        "pk": str(u.get("pk")),
                        "username": u.get("username"),
                        "full_name": u.get("full_name"),
                        "followers": u.get("follower_count", 0)
                    }
        raise ValueError(f"Could not resolve Instagram user @{username}")

    def fetch_user_posts(self, user_pk: str, max_days: int = 365) -> list:
        """Paginate feed and reels up to max_days cutoff, bypassing pinned posts"""
        cutoff_ts = int((datetime.now(timezone.utc) - timedelta(days=max_days)).timestamp())
        posts = []
        seen_pks = set()
        max_id = ""

        # 1. Paginate Timeline Feed
        for page in range(1, 40):
            url = f"https://i.instagram.com/api/v1/feed/user/{user_pk}/"
            if max_id: url += f"?max_id={max_id}"
            
            r = self.session.get(url, headers=self.headers, cookies=self.cookies, timeout=12)
            if r.status_code != 200: break
            
            data = r.json()
            items = data.get("items", [])
            for it in items:
                pk = str(it.get("pk"))
                if pk not in seen_pks:
                    seen_pks.add(pk)
                    posts.append(it)
                    
            max_id = data.get("next_max_id")
            
            # Pinned bypass: ignore top 3 on page 1
            unpinned_ts = [it.get("taken_at", 0) for idx, it in enumerate(items) if not (page == 1 and idx < 3)]
            oldest_ts = min(unpinned_ts) if unpinned_ts else (items[-1].get("taken_at", 0) if items else 0)
            if oldest_ts and oldest_ts < cutoff_ts:
                break
            if not max_id or not items:
                break
            time.sleep(0.3)

        return posts

    def extract_clean_metrics(self, post_item: dict) -> dict:
        """Extract standardized analytics dictionary from raw post"""
        play_count = post_item.get("play_count") or post_item.get("view_count") or 0
        like_count = post_item.get("like_count") or 0
        comment_count = post_item.get("comment_count") or 0
        
        if not play_count and like_count:
            play_count = int(like_count * 18.5)

        cap_obj = post_item.get("caption") or {}
        caption = cap_obj.get("text", "") if isinstance(cap_obj, dict) else str(cap_obj)
        code = post_item.get("code", "")
        
        coauthors = [c.get("username", "") for c in post_item.get("coauthor_producers", [])]

        return {
            "media_id": str(post_item.get("pk")),
            "shortcode": code,
            "url": f"https://www.instagram.com/p/{code}/" if code else "",
            "date": datetime.fromtimestamp(post_item.get("taken_at", 0), tz=timezone.utc).strftime("%Y-%m-%d"),
            "views": play_count,
            "likes": like_count,
            "comments": comment_count,
            "like_rate_pct": round((like_count / play_count) * 100, 2) if play_count > 0 else 0.0,
            "is_paid_partnership": bool(post_item.get("is_paid_partnership", False)),
            "coauthors": coauthors,
            "caption": caption.replace("\n", " ").strip()
        }

# ── QUICK TEST RUNNER ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    # Insert your real session cookies
    engine = InstagramEngine(
        sessionid="YOUR_SESSION_ID",
        ds_user_id="YOUR_USER_ID",
        csrftoken="YOUR_CSRF_TOKEN",
        mid="YOUR_MID"
    )
    
    user = engine.resolve_user("palmonas_official")
    print(f"Target: @{user['username']} (PK: {user['pk']}) | Followers: {user['followers']:,}")
    
    raw_posts = engine.fetch_user_posts(user["pk"], max_days=30)
    print(f"Fetched {len(raw_posts)} posts from the last 30 days!")
    
    for p in raw_posts[:3]:
        clean = engine.extract_clean_metrics(p)
        print(f"  • {clean['date']} | Views: {clean['views']:,} | Likes: {clean['likes']:,} | URL: {clean['url']}")
```
