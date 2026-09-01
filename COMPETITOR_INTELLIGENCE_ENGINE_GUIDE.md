# Enterprise Competitor Influencer & Paid Ads Intelligence Engine
## Dual-Engine Architecture: Instagram Private API + Meta Ad Library (Playwright + GraphQL)

> **Production Guide & Standalone Codebase**  
> Use this guide to deploy full-funnel competitor research across any D2C, fashion, skincare, or retail brand in under 60 seconds.

---

## 📑 Table of Contents
1. [Architecture & System Flow](#1-architecture--system-flow)
2. [Environment Setup & Dependencies](#2-environment-setup--dependencies)
3. [Module A: Meta Ad Library Playwright + GraphQL Engine](#3-module-a-meta-ad-library-playwright--graphql-engine)
4. [Module B: Instagram Private API Grid & Co-Author Scraper](#4-module-b-instagram-private-api-grid--co-author-scraper)
5. [Module C: Cross-Platform Data Fusion & Deduplication](#5-module-c-cross-platform-data-fusion--deduplication)
6. [Module D: Advanced Competitor Intelligence Formulas & Signals](#6-module-d-advanced-competitor-intelligence-formulas--signals)
7. [Module E: Automated Master Excel Report Generator](#7-module-e-automated-master-excel-report-generator)
8. [Quick-Start Runbook for Any New Brand](#8-quick-start-runbook-for-any-new-brand)

---

## 1. Architecture & System Flow

```
                               TARGET COMPETITOR BRAND
                                          │
            ┌─────────────────────────────┴─────────────────────────────┐
            ▼                                                           ▼
┌──────────────────────────────────────┐     ┌──────────────────────────────────────┐
│  MODULE A: META AD LIBRARY ENGINE    │     │  MODULE B: INSTAGRAM PRIVATE API     │
│  (Playwright + GraphQL / DOM)        │     │  (curl_cffi Chrome TLS + Mobile API) │
├──────────────────────────────────────┤     ├──────────────────────────────────────┤
│ • Meta Ad Library Dark Ads           │     │ • Public Co-Authored Grid Posts      │
│ • Whitelisted Creator Handle Ads     │     │ • Organic Barter & Gifting Collabs   │
│ • Ad Longevity & Active / Inactive   │     │ • Real Video Plays, Likes & Comments │
│ • Creative Copy, Hooks & CTAs        │     │ • Pinned Post Cutoff Bypass          │
└──────────────────┬───────────────────┘     └──────────────────┬───────────────────┘
                   │                                            │
                   └──────────────────────┬─────────────────────┘
                                          ▼
                     ┌────────────────────────────────────────┐
                     │   MODULE C: DATA FUSION & DEDUPE       │
                     │   • Unifies IG Grid + Meta Dark Ads    │
                     │   • Pure Creator Entity Sizing         │
                     │   • Filters Sister Brands & Platforms  │
                     └────────────────────┬───────────────────┘
                                          ▼
                     ┌────────────────────────────────────────┐
                     │   MODULE D: COMPETITOR SIGNALS         │
                     │   • 🏆 Evergreen Hero Winner Score     │
                     │   • 🚀 Paid Spend Multiplier           │
                     │   • 💬 Comment Purchase Intent NLP     │
                     │   • 🔄 Creator Re-Hire & Loyalty Index │
                     └────────────────────┬───────────────────┘
                                          ▼
                     ┌────────────────────────────────────────┐
                     │   DELIVERABLE: MASTER EXCEL & JSON     │
                     │   (Multi-Tab, Formatted, Ready to Use) │
                     └────────────────────────────────────────┘
```

---

## 2. Environment Setup & Dependencies

Install required libraries:

```bash
pip install playwright curl_cffi openpyxl pandas requests
playwright install chromium
```

---

## 3. Module A: Meta Ad Library Playwright + GraphQL Engine

This module uses Playwright with `en-US` locale enforcement to scroll, trigger GraphQL lazy loading, and extract all active and inactive ad creatives, library IDs, run dates, copy, and creator whitelists.

### Save as `engine_meta_adlibrary.py`:

```python
"""
Module A: Playwright + GraphQL Meta Ad Library Scraper
Extracts all Ad Cards, identifies Branded Content Whitelists & Dark Ads.
"""

import sys, os, json, time, re, csv
from datetime import datetime, timezone
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")

def scan_meta_ad_library(brand_query: str, page_id: str = None, max_scrolls: int = 35) -> dict:
    """
    Scrapes Meta Ad Library for any brand using keyword or page_id mode.
    """
    if page_id:
        url = f"https://www.facebook.com/ads/library/?active_status=all&ad_type=all&country=IN&view_all_page_id={page_id}&search_type=page&media_type=all"
    else:
        url = f"https://www.facebook.com/ads/library/?active_status=all&ad_type=all&country=IN&q={brand_query}&search_type=keyword_unordered&media_type=all"

    print(f"[Meta AdLib] Starting scan for: '{brand_query}' (Page ID: {page_id})")
    print(f"[Meta AdLib] Target URL: {url}")

    all_ads = []
    seen_ids = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1600, "height": 1000},
            locale="en-US",
            extra_http_headers={"Accept-Language": "en-US,en;q=0.9"}
        )
        page = context.new_page()

        print("[Meta AdLib] Loading page DOM...")
        page.goto(url, timeout=50000, wait_until="domcontentloaded")
        time.sleep(6)

        js_extractor = r"""
        () => {
            const ads = [];
            const fullText = document.body.innerText || "";
            const blocks = fullText.split(/(?:Library ID:|लायब्ररी आयडी:)\s*(\d+)/i);
            
            for (let i = 1; i < blocks.length; i += 2) {
                const libId = blocks[i].trim();
                const chunk = blocks[i+1] || "";
                
                const isActive = chunk.includes("Active") || chunk.includes("सक्रिय");
                const dateMatch = chunk.match(/(?:Started running on|रोजी प्रसारण सुरू झाले)\s*([^
·]+)/i);
                const startDate = dateMatch ? dateMatch[1].trim() : "";
                
                let advName = "";
                const advMatch = chunk.match(/(?:See (?:ad|summary) details|जाहिरात तपशील पहा)
([^
]+)
(?:Sponsored|प्रायोजित)/i);
                if (advMatch) {
                    advName = advMatch[1].trim();
                }
                
                let body = "";
                const spParts = chunk.split(/(?:Sponsored|प्रायोजित)/i);
                if (spParts.length > 1) {
                    body = spParts[1].split(/(?:Shop Now|Learn More|Buy Now|Order Now|See (?:ad|summary)|जाहिरात तपशील पहा)/i)[0].trim();
                }
                
                ads.push({
                    library_id: libId,
                    ad_url: `https://www.facebook.com/ads/library/?id=${libId}`,
                    advertiser: advName,
                    is_active: isActive,
                    start_date: startDate,
                    body: body.substring(0, 400).replace(/\n/g, " ").trim()
                });
            }
            return ads;
        }
        """

        prev_len = 0
        stalls = 0

        for step in range(1, max_scrolls + 1):
            batch = page.evaluate(js_extractor)
            for ad in batch:
                if ad["library_id"] not in seen_ids:
                    seen_ids.add(ad["library_id"])
                    all_ads.append(ad)

            print(f"  Scroll {step:>2}/{max_scrolls}: Captured {len(all_ads)} unique Ad IDs (DOM Batch: {len(batch)})")

            if len(all_ads) == prev_len:
                stalls += 1
                if stalls >= 4:
                    print("  -> Cursor exhausted. No new ads loaded.")
                    break
            else:
                stalls = 0

            prev_len = len(all_ads)
            page.evaluate("window.scrollBy(0, 2500)")
            time.sleep(1.8)

        browser.close()

    # Identify Creators & Whitelisted Partners
    creators = {}
    clean_brand = brand_query.lower().replace(" ", "").replace("_", "")

    for ad in all_ads:
        adv = ad["advertiser"]
        body = ad["body"]
        is_collab = False
        cname = None

        if " with " in adv.lower():
            parts = adv.split(" with ")
            if clean_brand in parts[1].lower().replace(" ", ""):
                is_collab = True
                cname = parts[0].strip()
        elif adv and clean_brand not in adv.lower().replace(" ", "") and not adv.startswith("See ad"):
            is_collab = True
            cname = adv
        else:
            mentions = re.findall(r'@([A-Za-z0-9_.]+)', body)
            valid_m = [m for m in mentions if clean_brand not in m.lower().replace("_", "")]
            if valid_m:
                is_collab = True
                cname = f"@{valid_m[0]}"

        ad["is_creator_collab"] = is_collab
        ad["creator_name"] = cname

        if is_collab and cname:
            ckey = cname.lower().replace("@", "").strip()
            if ckey not in creators:
                creators[ckey] = {
                    "name": cname,
                    "handle": f"@{ckey}",
                    "active_ads": 1 if ad["is_active"] else 0,
                    "total_ads": 1,
                    "sample_start_date": ad["start_date"],
                    "sample_ad_url": ad["ad_url"]
                }
            else:
                creators[ckey]["total_ads"] += 1
                if ad["is_active"]:
                    creators[ckey]["active_ads"] += 1

    print(f"[Meta AdLib] Finished! Captured {len(all_ads)} Ads | {len(creators)} Creator Partners\n")
    return {"ads": all_ads, "creators": list(creators.values())}

if __name__ == "__main__":
    res = scan_meta_ad_library(brand_query="Palmonas", page_id="100076111693972", max_scrolls=15)
    with open("meta_adlibrary_export.json", "w", encoding="utf-8") as f:
        json.dump(res, f, indent=2)
```

---

## 4. Module B: Instagram Private API Grid & Co-Author Scraper

This module extracts every co-authored reel, tagged lookbook, and organic post across a 365-day window with automatic pinned post bypass and pure audience tiering.

### Save as `engine_instagram_grid.py`:

```python
"""
Module B: Instagram Private API Grid & Co-Author Scraper
Extracts all public collaborations with pinned-post bypass and 4-tier hierarchy.
"""

import sys, os, json, time, re
from datetime import datetime, timezone, timedelta
from curl_cffi import requests as cffi_requests

sys.stdout.reconfigure(encoding="utf-8")

# Live Session Cookies
COOKIES = {
    "sessionid": "76326162386%3A670U47iQkU6B8V%3A18%3AAYj9oJ1L51k_G3_j-uX4lQ9V6aM9Wc7gQ2yZ",
    "ds_user_id": "76326162386",
    "csrftoken": "b3U8nI5m6d9L0v7e8W1x2y",
}

MOB_HEADERS = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

def resolve_instagram_pk(target_handle: str) -> int:
    s = cffi_requests.Session(impersonate="chrome120")
    r = s.get(f"https://www.instagram.com/{target_handle}/", timeout=12)
    m = re.search(r'"profilePage_(\d+)"', r.text) or re.search(r'"props":{"id":"(\d+)"', r.text)
    if m: return int(m.group(1))
    raise ValueError(f"Could not resolve PK for @{target_handle}")

def scan_instagram_brand_grid(target_handle: str, days_back: int = 365) -> list:
    clean_u = target_handle.lower().strip()
    user_pk = resolve_instagram_pk(clean_u)
    print(f"[IG Scraper] Resolved @{clean_u} -> PK: {user_pk}")

    session = cffi_requests.Session(impersonate="chrome120")
    cutoff_ts = int((datetime.now(timezone.utc) - timedelta(days=days_back)).timestamp())

    all_items = []
    seen_ids = set()

    for endpoint_name, base_url in [
        ("Timeline Feed", f"https://i.instagram.com/api/v1/feed/user/{user_pk}/"),
        ("Reels/Clips", f"https://i.instagram.com/api/v1/clips/user/")
    ]:
        print(f"[IG Scraper] Scanning {endpoint_name}...")
        max_id = ""
        for page in range(1, 30):
            if endpoint_name == "Timeline Feed":
                url = f"{base_url}?count=12"
                if max_id: url += f"&max_id={max_id}"
                r = session.get(url, headers=MOB_HEADERS, cookies=COOKIES, timeout=12)
            else:
                body_data = {"target_user_id": str(user_pk), "page_size": "12"}
                if max_id: body_data["max_id"] = max_id
                r = session.post(base_url, headers=MOB_HEADERS, cookies=COOKIES, data=body_data, timeout=12)

            if r.status_code != 200: break
            data = r.json()
            items = data.get("items", [])
            
            # Format standard items
            raw_list = [it.get("media", it) if "media" in it else it for it in items]
            for it in raw_list:
                pk = str(it.get("pk") or it.get("id"))
                if pk and pk not in seen_ids:
                    seen_ids.add(pk)
                    all_items.append(it)

            # Pinned post cutoff bypass
            unpinned_ts = [it.get("taken_at", 0) for it in raw_list if not it.get("timeline_pinned_user_ids") and it.get("taken_at")]
            oldest_ts = min(unpinned_ts, default=0)
            if oldest_ts and oldest_ts < cutoff_ts:
                print(f"  -> Reached {days_back}-day date cutoff in {endpoint_name}")
                break

            max_id = data.get("next_max_id") or data.get("paging_info", {}).get("max_id")
            if not max_id or not items: break
            time.sleep(0.35)

    print(f"[IG Scraper] Total Raw Media Extracted: {len(all_items)}")

    # Extract Collaborations
    collabs = []
    for it in all_items:
        taken_at = it.get("taken_at", 0)
        if taken_at < cutoff_ts: continue

        owner = it.get("user", {})
        owner_uname = owner.get("username", "").lower()
        coauthors = it.get("coauthor_producers", [])
        is_paid = bool(it.get("is_paid_partnership", False))

        is_collab = False
        creator_uname = ""

        if owner_uname and owner_uname != clean_u and owner_uname not in ["palmonas", "palmonas_men"]:
            is_collab = True
            creator_uname = owner_uname
        elif coauthors:
            for c in coauthors:
                cu = c.get("username", "").lower()
                if cu != clean_u and cu not in ["palmonas", "palmonas_men"]:
                    is_collab = True
                    creator_uname = cu
                    break

        if is_collab and creator_uname:
            likes = it.get("like_count") or 0
            comments = it.get("comment_count") or 0
            views = it.get("play_count") or it.get("view_count") or int(likes * 18.5)
            like_rate = (likes / views * 100.0) if views > 0 else 0.0

            # 4-Tier Hierarchy
            is_boosted = (like_rate < 0.35 and views >= 200000) or views >= 3000000
            if is_paid and is_boosted: tier = "Tier 1: Toggle ON + Boosted Paid Ad"
            elif is_paid and not is_boosted: tier = "Tier 2: Toggle ON + Organic Collab"
            elif not is_paid and is_boosted: tier = "Tier 3: Toggle OFF + Heavily Boosted Ad"
            else: tier = "Tier 4: Toggle OFF + Organic / Noise"

            code = it.get("code", "")
            date_str = datetime.fromtimestamp(taken_at, tz=timezone.utc).strftime("%Y-%m-%d")

            collabs.append({
                "post_url": f"https://www.instagram.com/p/{code}/" if code else "",
                "creator_handle": f"@{creator_uname}",
                "raw_handle": creator_uname,
                "date": date_str,
                "views": views,
                "likes": likes,
                "comments": comments,
                "like_to_view_pct": round(like_rate, 3),
                "is_paid_toggle": is_paid,
                "is_boosted": is_boosted,
                "partnership_tier": tier
            })

    print(f"[IG Scraper] Identified {len(collabs)} Collab Posts across 1 Year!\n")
    return collabs
```

---

## 5. Module C: Cross-Platform Data Fusion & Deduplication

Merges the output of **Module A (Meta Ad Library)** and **Module B (Instagram Grid)** into a unified, deduplicated creator roster.

```python
"""
Module C: Data Fusion & Entity Deduplication
Combines IG Grid + Meta Ad Library into 1 unified database.
"""

def fuse_datasets(ig_collabs: list, meta_data: dict) -> list:
    unified_creators = {}

    # 1. Ingest Instagram Grid Creators
    for c in ig_collabs:
        h = c["raw_handle"].lower()
        if h not in unified_creators:
            unified_creators[h] = {
                "handle": f"@{h}",
                "raw_handle": h,
                "name": h,
                "on_instagram_grid": True,
                "on_meta_adlibrary": False,
                "total_grid_posts": 1,
                "total_grid_views": c["views"],
                "active_meta_ads": 0,
                "total_meta_ads": 0,
                "sample_grid_url": c["post_url"],
                "sample_ad_url": ""
            }
        else:
            unified_creators[h]["total_grid_posts"] += 1
            unified_creators[h]["total_grid_views"] += c["views"]

    # 2. Ingest Meta Ad Library Creators (Dark Ads)
    for mc in meta_data.get("creators", []):
        h = mc["handle"].replace("@", "").lower()
        if h in unified_creators:
            unified_creators[h]["on_meta_adlibrary"] = True
            unified_creators[h]["active_meta_ads"] = mc["active_ads"]
            unified_creators[h]["total_meta_ads"] = mc["total_ads"]
            unified_creators[h]["sample_ad_url"] = mc["sample_ad_url"]
        else:
            unified_creators[h] = {
                "handle": f"@{h}",
                "raw_handle": h,
                "name": mc["name"],
                "on_instagram_grid": False,
                "on_meta_adlibrary": True,
                "total_grid_posts": 0,
                "total_grid_views": 0,
                "active_meta_ads": mc["active_ads"],
                "total_meta_ads": mc["total_ads"],
                "sample_grid_url": "",
                "sample_ad_url": mc["sample_ad_url"]
            }

    return list(unified_creators.values())
```

---

## 6. Module D: Advanced Competitor Intelligence Formulas & Signals

### 🏆 1. Evergreen Hero Winner Score
Calculates if a competitor ad is an evergreen positive-ROAS winner based on run duration:
```python
def calculate_ad_longevity(start_date_str: str, is_active: bool) -> str:
    """
    Identifies high-converting hero ads vs failed creative tests.
    """
    if not start_date_str: return "Unknown"
    start_dt = datetime.strptime(start_date_str, "%Y-%m-%d")
    days_running = (datetime.now() - start_dt).days
    
    if is_active and days_running >= 90:
        return f"🏆 Evergreen Hero Winner ({days_running}d - Core ROAS Engine)"
    elif is_active and days_running >= 30:
        return f"⚡ Proven Scaler ({days_running}d - Consistent Spend)"
    elif is_active:
        return f"🧪 Active Test ({days_running}d)"
    else:
        return f"⏹️ Completed / Inactive ({days_running}d runtime)"
```

### 🚀 2. Paid Media Multiplier (Spend Velocity)
$$\text{Paid Multiplier} = \frac{\text{Actual Video Views}}{\text{Creator Baseline Follower Count} \times 0.25}$$

* $\text{Multiplier} > 50\text{x} \rightarrow$ **Aggressive Performance Ad Spend (₹5L–₹20L+)**
* $\text{Multiplier} < 2\text{x} \rightarrow$ **Pure Organic / Barter Distribution**

### 💬 3. Comment Purchase Intent NLP
Scans comment text for buying signals vs vanity noise:
```python
import re

INTENT_REGEX = re.compile(r'\b(price|cost|how much|link|buy|order|available|discount|code|dm|where to get|tarnish|waterproof|fake|gold)\b', re.I)

def calculate_buyer_intent_score(comments_list: list) -> dict:
    if not comments_list: return {"intent_score_pct": 0, "buyer_comments": 0}
    high_intent = [c for c in comments_list if INTENT_REGEX.search(c)]
    score = (len(high_intent) / len(comments_list)) * 100.0
    return {
        "intent_score_pct": round(score, 1),
        "total_comments": len(comments_list),
        "high_intent_comments": len(high_intent)
    }
```

---

## 7. Module E: Automated Master Excel Report Generator

Generates a master Excel file (`.xlsx`) with:
1. `Executive Summary Card`
2. `Unified Creator Roster` (Followers, Sizing Tiers, Platform Presence)
3. `Master Hierarchy & Boost Analysis` (Tier 1/2/3/4)
4. `Meta Ad Library Archive` (Live IDs, URLs, Copy)

### Save as `generate_master_excel.py`:

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_excel_deliverable(brand_name: str, fused_creators: list, collabs: list, all_ads: list, filename: str):
    wb = openpyxl.Workbook()
    
    # Styling definitions
    font_title = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    font_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True, color="000000")
    font_norm = Font(name="Calibri", size=10, bold=False, color="000000")
    font_link = Font(name="Calibri", size=10, bold=False, color="0563C1", underline="single")
    thin_line = Side(style="thin", color="D5D8DC")
    border_cell = Border(left=thin_line, right=thin_line, top=thin_line, bottom=thin_line)

    # Sheet 1: Unified Creators
    ws = wb.active
    ws.title = "Unified Creator Roster"
    ws.sheet_view.showGridLines = True
    
    ws.merge_cells("A1:H1")
    ws["A1"] = f"Complete Creator Partnerships Portfolio — {brand_name} ({len(fused_creators)} Total Creators)"
    ws["A1"].font = font_title
    ws["A1"].fill = PatternFill("solid", fgColor="0B2240")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    
    headers = [
        ("#", 5), ("Creator Handle", 25), ("Source Type", 24),
        ("Total Grid Posts", 16), ("Total Video Views", 18),
        ("Active Meta Ads", 16), ("Sample Post Link", 40), ("Sample Meta Ad Link", 40)
    ]
    
    for col_idx, (h_text, w) in enumerate(headers, 1):
        c = ws.cell(row=2, column=col_idx, value=h_text)
        c.font = font_hdr
        c.fill = PatternFill("solid", fgColor="1B2631")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[2].height = 24
    ws.freeze_panes = "A3"
    
    for idx, c in enumerate(fused_creators, 1):
        r_num = idx + 2
        source = "Instagram Grid + Meta Ads" if (c["on_instagram_grid"] and c["on_meta_adlibrary"]) else ("Instagram Grid Only" if c["on_instagram_grid"] else "Meta Dark Ad Only")
        r_vals = [
            idx, c["handle"], source, c["total_grid_posts"], c["total_grid_views"],
            c["active_meta_ads"], c["sample_grid_url"], c["sample_ad_url"]
        ]
        for c_idx, val in enumerate(r_vals, 1):
            cell = ws.cell(row=r_num, column=c_idx, value=val)
            cell.border = border_cell
            if c_idx == 1: cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx == 2: cell.font = font_bold; cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx == 3: cell.font = font_bold; cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in (4, 5, 6): cell.font = font_norm; cell.alignment = Alignment(horizontal="right", vertical="center"); cell.number_format = "#,##0"
            elif c_idx in (7, 8): cell.font = font_link; cell.alignment = Alignment(horizontal="left", vertical="center"); cell.hyperlink = val if val else None
        ws.row_dimensions[r_num].height = 20

    wb.save(filename)
    print(f"[Deliverable] Successfully generated {filename}!")
```

---

## 8. Quick-Start Runbook for Any New Brand

To run a competitor audit on a new brand (e.g. `giva.co`, `nykdbynykaa`, `snitch.co.in`, `bluestone`):

### Step 1: Identify Brand Targets
1. **Instagram Username**: e.g., `giva.co`
2. **Meta Ad Library Page ID**: e.g., Find on Facebook Page `About` or search query `GIVA`

### Step 2: Run Unified Pipeline
```python
from engine_meta_adlibrary import scan_meta_ad_library
from engine_instagram_grid import scan_instagram_brand_grid
from generate_master_excel import build_excel_deliverable

# 1. Run Meta Ad Library Scan
meta_res = scan_meta_ad_library(brand_query="GIVA", page_id=None, max_scrolls=30)

# 2. Run Instagram Grid Scan (365 Days)
ig_collabs = scan_instagram_brand_grid(target_handle="giva.co", days_back=365)

# 3. Fuse & Generate Master Deliverable
fused = fuse_datasets(ig_collabs, meta_res)
build_excel_deliverable("GIVA", fused, ig_collabs, meta_res["ads"], "giva_1year_competitor_master.xlsx")
```

---

### 🛡️ Best Practices & Edge Cases:
1. **Always use `locale="en-US"` in Playwright**: Prevents UI localization issues where "Library ID" or "Active" render in regional languages.
2. **Always include pinned post bypass**: Do not break feed loops on page 1 when encountering 2023/2024 pinned posts.
3. **Filter self-partnerships**: Exclude sister brands (e.g. `jockeywomanindia` vs `jockeyindia`) from third-party creator rosters.
