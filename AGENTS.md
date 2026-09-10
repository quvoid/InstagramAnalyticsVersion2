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
└── [Master Workbooks & Datasets in Root]     # 80+ Deliverable Excel files and JSON caches
```

---

## ⚖️ Non-Negotiable Operating Rules

When modifying code, creating scripts, or interacting with the user, you **MUST** follow these rules:

### 1. Accurate Follower Counts (No Fabrication)
* **Never guess, approximate, or fabricate follower counts.**
* Always resolve live follower metrics via `core/profile_auditor.py` (which uses `curl_cffi` with Chrome 120 impersonation) or Playwright with authenticated session cookies.
* If an account cannot be resolved, explicitly report it as `0` or `Unresolved`—do not invent numbers.

### 2. Zero Emojis in Excel Deliverables
* **Strictly no emojis** in final `.xlsx` files.
* Banners, tier strings, categories, headers, and sheet names must be plain text:
  - Use `Mega (1M+)`, `Macro (500K-1M)`, `Mid-Tier (100K-500K)`, `Micro (10K-100K)`
  - Do NOT use `🌟 Mega Creator`, `🚀 Boosted`, etc. in workbook cells unless explicitly requested.

### 3. Session Cookies Required
* All Instagram web requests must include the user's active session cookies.
* Cookie dict format (for `curl_cffi` or `requests`):
  ```python
  COOKIES = {
      "sessionid": "25113411270%3AyQFaao428g6Xb9%3A0%3AAYjwS_M5UIDcS-i41tvdVFu8WKSqfzfgEhlZLGMvFg",
      "csrftoken": "3gJbkGDZp99lA8QQ0brobyoHzOreuu8f",
      "mid": "afyCbwALAAFRStE-k17-dfO5_jfa",
      "ds_user_id": "25113411270",
  }
  ```
* Cookie list format (for Playwright `context.add_cookies()`):
  ```python
  PLAYWRIGHT_COOKIES = [
      {"name": k, "value": v, "domain": ".instagram.com", "path": "/"}
      for k, v in COOKIES.items()
  ]
  ```

### 4. Incremental Progress & Caching
* For long-running scans (e.g. 50+ brands or 200+ creators), **always save progress after every batch/brand**.
* Check existing cache files before re-scraping to avoid unnecessary API calls and prevent rate-limiting:
  - FinTech: `fintech_collabs_cache.json`
  - Kolkata Creators: `discovery_verified_creators.json`

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
