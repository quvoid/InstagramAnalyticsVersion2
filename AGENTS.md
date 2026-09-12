# 🤖 AGENTS.md — AI Agent & Developer Operating Manual

> **For AI Coding Assistants (Claude Code, Antigravity, Cursor, GitHub Copilot, Roo, Cline) & External Developers**  
> This file defines the operational framework, architectural guidelines, natural language interface, and strict rules for working with the **Instagram & Omnichannel Marketing Intelligence Engine**.

---

## 🎯 What This Repository Is

An enterprise-grade Python intelligence engine for:
1. **Instagram Collaboration Discovery**: Identifying co-authored reels, partner tags, and brand campaigns.
2. **Meta Paid Partnership Toggle Detection**: Verifying official `is_paid_partnership` disclosure labels.
3. **Paid Media Ad Spend Detection**: Classifying reels into 4 collaboration tiers based on view-to-like ratios and view-to-follower multipliers.
4. **Regional Creator Discovery**: Sourcing and live-auditing authentic regional creators (e.g. Kolkata/Bengal) with 10K+ followers.
5. **Cross-Platform Competitor Intelligence**: Analyzing Meta Ads Library, YouTube video transcripts/shorts, and Google Maps footfall/reviews across Retail, FinTech, Jewellery, Electronics, Footwear, and FMCG.

---

## 🚀 How to Run Commands Using Natural Language

Any user or agent can execute tasks directly from the terminal using natural language via `run.py`:

```bash
# 1. Profile Audit & Follower Verification
python run.py "audit profile @rjpraveen"

# 2. Regional Creator Discovery (Kolkata & Bengal)
python run.py "find kolkata creators 10k"

# 3. FinTech 53-Brand 4-Tier 2-Year Deep Scan
python run.py "scan fintech brands 2 years"
python run.py "scan fintech brands dry run"

# 4. Paid Media Ad Spend & Boost Detection
python run.py "detect boost 1250000 2500 45000"

# 5. List All Master Deliverable Workbooks
python run.py "list master workbooks"

# 6. Show Full Catalog & Directory
python run.py "show catalog"
```

---

## 🏛️ Core Project Directory Structure

```
InstagramAnalytics/
├── run.py                                    # Central Natural Language CLI & Dispatcher
├── AGENTS.md                                 # This file (AI agent operational instructions)
├── README.md                                 # Human-facing project overview
├── catalog/
│   ├── MASTER_DIRECTORY.md                   # Complete asset, deliverable, and domain index
│   └── REGISTRY.json                         # Intent-to-script mapping for natural language runner
├── core/
│   ├── kolkata_engine.py                     # Kolkata & Bengal Creator Discovery Engine
│   ├── fintech_engine.py                     # FinTech 53-Brand 4-Tier Collaboration Engine
│   ├── profile_auditor.py                    # Live Profile & Follower Verification Engine
│   └── ad_boost_engine.py                    # Paid Media Ad Spend & Boost Classifier
├── api_wrapper/
│   ├── client.py                             # Core Instagram Python SDK & cffi session client
│   ├── cli.py                                # Subcommand CLI for direct SDK calls
│   └── server.py                             # FastAPI REST endpoints
├── competitor/                               # Structured Competitor Intelligence Sub-Suite
│   ├── data/                                 # Raw JSON scrapes and intermediate dumps
│   ├── docs/                                 # Sector-specific documentation and research notes
│   ├── scripts/                              # Sector analysis scripts (Croma, Jewellery, Footwear)
│   └── workbooks/                            # Final sector-specific Excel workbooks
├── deliverables/                             # Every client workbook and CSV
├── data/                                     # Tracked JSON datasets and caches
├── docs/                                     # CREATOR_DB.md, PAGE_AUDIT.md, guides, prompts
├── .claude/skills/                           # ig-competitor-intelligence, ig-creator-intelligence, ig-partnership-timeline
└── scripts/legacy/                           # One-off scripts from past work (run from repo root)
```

---

## ⚖️ Non-Negotiable Operating Rules

When modifying code, creating scripts, or interacting with the user, you **MUST** follow these rules:

### 1. Accurate Follower Counts (No Fabrication)
* **Never guess, approximate, or fabricate follower counts.**
* Always resolve live follower metrics via `core/profile_auditor.py` (which uses `curl_cffi` with Chrome 120 impersonation) or Playwright with authenticated session cookies.
* If an account cannot be resolved, explicitly report it as `0` or `Unresolved`—do not invent numbers.

#### 1a. `og:description` is ROUNDED above ~10K — do not read it as exact (verified 2026-09-11)
Instagram's `og:description` meta tag prints exact counts only for small accounts
(`1,229 Followers`) and **abbreviates everything above roughly 10K** (`165K Followers`).
Parsing that token yields `165000` for an account that actually has `165,093` — which is
fabricated precision, and it lands on **every creator above the 10K gate**. Measured errors:
`whatkolkataeats` 126K vs 125,749 · `primark` 11,000,000 vs 11,435,534 (off by 435,534).
`og:description` is also cached and can disagree with the live profile.

**Resolution ladder — use `core/profile_auditor.resolve_profile()`, which implements it:**

| Rung | Source | Precision | Notes |
|---|---|---|---|
| 1 | `GET /api/v1/users/{pk}/info/` (Android UA + cookies) → `user.follower_count` | **EXACT** | also returns `public_email`, `category`, `is_verified`, `city_name`, branded-content flags |
| 2 | Rendered logged-in DOM: header node matching `^[\d.,km]+\s+followers$` → its `span[title]` | **EXACT** | `title` holds `125,749` while the text shows `125K`. There is **no** `a[href$="/followers/"]` any more — anchor-based selectors are dead |
| 3 | `og:description` | **ROUNDED** | fallback only, and it must be labelled `rounded` |

* Resolve the numeric `pk` from the profile HTML via `"profilePage_(\d+)"`.
* Every count carries `followers_precision` / `followers_source` / `resolved_at`.
  Counts drift minute to minute (`kolkatasutra` read 165,094 → 165,093 → 165,090 within
  one session), so a bare number with no timestamp is not a fact.
* A **rounded** count within 15% of the 10K gate is reported
  `BORDERLINE_ROUNDED_NEEDS_EXACT`, never counted as a pass.

#### 1b. Throttling is per-endpoint-family — a 429 is not "account not found"
Measured on the live session: `users/{pk}/info/`, `discover/chaining`, `web/search/topsearch`,
`fbsearch/places` and `locations/{pk}/sections/` all answered normally **while**
`web_profile_info` returned `429` and `feed/user/{pk}/` and `clips/user/` returned
`400 {"message":"feedback_required","is_spam":true}`. Both shapes mean *throttled*; the
`429` body is Instagram's "Page Not Found" HTML but the session is still logged in
(`class="no-js logged-in"`). Use `profile_auditor.is_throttled()` and never record a
throttled response as a missing account or a zero follower count.

> `api_wrapper/client.py:resolve_creator_profile()` still violates several of these rules
> (emoji tier strings, tier boundaries that disagree with the taxonomy below, no cookies on
> the request, and an unscoped `Followers` regex that can match the wrong number in the
> document). Six pipelines call it, so changing its tier boundaries would move existing
> deliverables — fix deliberately, not incidentally.

### 2. Zero Emojis in Excel Deliverables
* **Strictly no emojis** in final `.xlsx` files.
* Banners, tier strings, categories, headers, and sheet names must be plain text:
  - Use `Mega (1M+)`, `Macro (500K-1M)`, `Mid-Tier (100K-500K)`, `Micro (10K-100K)`
  - Do NOT use `🌟 Mega Creator`, `🚀 Boosted`, etc. in workbook cells unless explicitly requested.
* **Scraped text is the main leak, not your own strings.** Creator bios and display names
  are full of emojis and flag characters, and they reach cells untouched unless you strip
  them. Run every cell value through `core.kolkata_engine.clean_cell()`, which also removes
  lone UTF-16 surrogates (`openpyxl` raises `UnicodeEncodeError: surrogates not allowed` on
  save when it meets one). Keep sheet titles at **31 characters or fewer** — openpyxl warns
  and some spreadsheet apps then refuse the file.

### 3. Session Cookies Required
* All Instagram web requests must include the user's active session cookies.
* Cookie dict format (for `curl_cffi` or `requests`):
  ```python
# Never hardcode these. They live in the git-ignored .env and are loaded by
# core/session.py. See .env.example for the variable names.
from core.session import load_cookies, playwright_cookies
COOKIES = load_cookies()                      # dict for curl_cffi / requests
PLAYWRIGHT_COOKIES = playwright_cookies(COOKIES)   # list for context.add_cookies()
```
* Cookie list format (for Playwright `context.add_cookies()`):
  ```python
  PLAYWRIGHT_COOKIES = [
      {"name": k, "value": v, "domain": ".instagram.com", "path": "/"}
      for k, v in COOKIES.items()
  ]
  ```

### 3a. Creator Discovery — Source Precision Ranking (measured 2026-09-11)
All sources below are free. No paid discovery product is used. Implemented in
`core/discovery_sources.py`; orchestrated by `core/kolkata_engine.py`.

| Rank | Source | Endpoint / input | Measured yield & precision |
|---|---|---|---|
| **S1** | **Chaining graph** (best) | `GET /api/v1/discover/chaining/?target_id={pk}` | 77–80 accounts per seed, category-coherent. 13 Kolkata seeds → **630-handle pool**. Accounts suggested by many seeds are near-100% genuine |
| S2 | LLM answers | `core/llm_discovery.py` (live browser) or `quvoid/LLMxCitations` → `output.csv` | Strong on established names, blind to emerging ones. **Leads only — never a number.** Resolve names via topsearch; stated handles are often invented |
| S7 | Free directory pages | modash.io / socialveins / qoruz public ranking pages | modash yielded 49 handles for one Kolkata food page, plus audience-city %, engagement rate and fake-follower % |
| S3 | Topsearch grid | `GET /web/search/topsearch/?context=blended&query=` | ~5 users per query, matches username + display name. Pulls in shops named after the region |
| S4 | Geo sections | `fbsearch/places` → `POST /api/v1/locations/{pk}/sections/` | 63 authors from one Kolkata place. Low precision, but the **only physical-presence proof** |
| S5 | Hashtag sections | `POST /api/v1/tags/{tag}/sections/` | medium |
| S6 | Hub HTML crawl (legacy) | scrape `"username":"..."` out of hub pages | Lowest precision — picks up commenters and noise |

**Rules that follow from this:**
* **Seed lists rot — health-check before crawling.** 9 of the original 16 `HUB_ACCOUNTS`
  were gone, deleted, or under the gate. `kolkatabuzz` is dead while `thekolkatabuzz`
  (1,019,000 followers) is live — the chaining graph found the replacement itself.
* **Rank the audit queue by corroboration, not arbitrarily.** Weight each source family,
  score how many independently proposed a handle, audit highest-first. On the first run this
  gave **5 verified creators from 5 audits, zero skips**; the top candidate had been
  suggested by 10 different seeds.
* **Always seed `discover/chaining` with a pk resolved from the live profile.** A wrong pk
  returns confident, unrelated accounts from another country rather than an error.
* **Third-party numbers never enter the exact-count columns.** Vendor follower counts,
  engagement rates and audience splits are estimates; keep them in separate,
  source-attributed fields or leave them out.
* `@` in scraped prose is often an email or domain — reject handles ending in a TLD
  (`@gmail.com`, `@terareach.com`).

### 3b. The Permanent Creator Store — query it before you scrape

**`creator_intelligence.db` is the source of record for every creator, region, campaign and
brand collaborator. Read `docs/CREATOR_DB.md` and query the database before starting any new
discovery run.** The first backfill recovered **9,193 distinct handles** that were already
sitting in this repo's rosters and workbooks with no way to query them — which is why Kolkata
"only had 170 creators" and each Enamor state "barely 90".

Three steps, deliberately separated by cost:

```bash
python core/regional_engine.py backfill                                  # once
python core/regional_engine.py harvest --region punjab --pages 6         # cheap, wide
python core/regional_engine.py audit   --region punjab --limit 400       # expensive, ranked, resumable
python core/regional_engine.py deliver --region punjab --residence 2 --min 10000 --xlsx out.xlsx
python core/creator_db.py ask "kolkata food creators above 50k with email"
```

* **Residence is evidence, not a bio keyword.** `geo_evidence.COUNT(DISTINCT place_pk)` per
  region. Two or more distinct real Instagram locations means they operate there. Bio keyword
  matching stays a weak secondary signal — it was the single biggest false-negative source in
  the old engine, discarding every local creator who doesn't type their city in their bio.
* **Page the surfaces.** A single-page read of a location is leaving most of it behind: one
  Kolkata place gave 61 authors on page 0 and **176 by page 2**, still `more_available: true`.
  Use `location_sweep()` / `hashtag_sweep()`, never the single-page helpers, for volume work.
* **Harvest never writes a follower count.** Leads land as `PENDING_AUDIT` with
  `followers = 0, precision = 'unresolved'`. Only the audit phase writes a number, and only an
  exact one.

#### 3c. A PAST campaign cannot be swept from the `recent` tab (measured 2026-09-11)
Hashtag and location `recent` tabs return **newest-first**, so reaching a window a year back
would take thousands of pages. A 13-hashtag × 3-page sweep for Durga Puja 2025 produced exactly
**2 in-window posts** — everything recent under `#durgapuja` is from 2026.

For a past campaign:
1. sweep the **`top`** tab as well (ranked by engagement, so a past festival's best posts are
   still reachable) — `harvest --campaign X` now does both tabs and prints in-window vs outside;
2. the reliable route is **per-creator verification** during the audit —
   `audit --verify-campaign durga_puja_2025` reads each gated creator's own media and records a
   `campaign_evidence` row only for a post genuinely inside the window. It needs the feed
   endpoint, so it reports unavailability when throttled rather than guessing;
3. campaign venues (pandals) are useful because they only exist during the festival, but they
   still need deep paging to reach last year.

Never infer participation from a bio, a follower count, or a hashtag outside the date window.

#### 3d. Asking a chatbot: resolve NAMES, never trust stated handles
`core/llm_discovery.py` drives ChatGPT / duck.ai / Perplexity / Gemini in a real browser.
A plain HTTP client cannot do this — duck.ai answers `HTTP 418 ERR_CHALLENGE` because its
`x-vqd-hash-1` header is obfuscated JavaScript that has to be executed, and ChatGPT needs a
logged-in session. A persistent browser profile solves both, once.

Chatbots answer mostly with **names**, and the handles they state are frequently invented, so:
* extract handles *and* names, then resolve names through Instagram topsearch —
  `"Soham Sinha"` → `@kolkatadelites` (605,822, verified), `"Bong Eats"` → `@bongeats`;
* everything lands as `PENDING_AUDIT` with source `llm`; the audit decides. An invented handle
  simply returns `404_NOT_FOUND` (verified with a fabricated `@indrani_eats`);
* **an LLM never contributes a follower number.** It contributes a name to go and check.

The tool does **not** log in or accept terms of service by itself — `login <provider>` opens a
real window for the user to do it once, and `--accept-terms` is an explicit opt-in.

### 4. Incremental Progress & Caching
* For long-running scans (e.g. 50+ brands or 200+ creators), **always save progress after every batch/brand**.
* Check existing cache files before re-scraping to avoid unnecessary API calls and prevent rate-limiting:
  - FinTech: `fintech_collabs_cache.json`
  - Kolkata Creators: `discovery_verified_creators.json` (saved after every verified creator),
    `discovery_candidates_pool.json` (saved after every source), `discovery_run_report.json`
  - Write via a `.tmp` file plus `os.replace()` so an interrupted run cannot leave a
    half-written JSON behind.

### 5. Creator Discovery Commands

```bash
# Full multi-source discovery (all sources, 2 chaining hops)
python core/kolkata_engine.py 50

# One category, chaining only, single hop
python core/kolkata_engine.py 25 --categories=food --sources=chaining --hops=1

# Category-specific LLM prompts, then harvest what the models named
python core/discovery_sources.py prompts kolkata food fashion   # -> llm_discovery_prompts.csv
#   copy to the LLMxCitations checkout as prompts.csv, then:
#     python main.py --save-auth chatgpt
#     python main.py --platforms chatgpt,perplexity,gemini --input prompts.csv --output output.csv
python core/kolkata_engine.py 40 --llm-csv=../LLMxCitations/output.csv

# Re-resolve legacy roster rows to exact counts (flags, never deletes)
python core/kolkata_engine.py --reaudit

# Diagnostics
python core/discovery_sources.py health kolkatasutra kolkatabuzz
python core/discovery_sources.py grid kolkata food
python core/profile_auditor.py @whatkolkataeats --engagement
```

---

## 📊 The 4-Tier Collaboration Taxonomy

| Tier | Toggle Status | Ad Spend (Boost) | Strategic Meaning |
|---|---|---|---|
| **Tier 1** | Toggle ON | Boosted | Formal Paid Partnership + Paid Media Ad Spend |
| **Tier 2** | Toggle ON | Organic | Formal Paid Partnership + Organic Reach |
| **Tier 3** | Toggle OFF | Boosted | Co-Author Collab (Toggle OFF) + Heavy Paid Media Spend |
| **Tier 4** | Toggle OFF | Organic | Standard Organic Collab / Barter / Low Reach |

* **Boost Detection Logic**:
  - `Views >= 500,000` with `Like Rate < 0.35%` $\rightarrow$ **Heavily Boosted**
  - `Views >= 1,000,000` with `Like Rate < 0.60%` $\rightarrow$ **Heavily Boosted**
  - `Multiplier >= 10.0x` with `Like Rate < 0.50%` $\rightarrow$ **Likely Boosted**
  - `Multiplier >= 25.0x` with organic engagement $\rightarrow$ **Viral Organic**

---

## 🛠️ Adding New Workflows to the Natural Language Dispatcher

To register a new capability so that users can trigger it via natural language:
1. Place the underlying script in `core/` or appropriate folder.
2. Add an intent entry to `catalog/REGISTRY.json`:
   ```json
   {
     "id": "new_workflow_id",
     "name": "Human-Readable Name",
     "description": "What it accomplishes",
     "triggers": ["trigger phrase 1", "trigger phrase 2"],
     "script": "path/to/script.py"
   }
   ```
3. Update `run.py` to handle any custom parameter parsing if required.
4. Document the deliverable in `catalog/MASTER_DIRECTORY.md`.
