---
name: ig-competitor-intelligence
description: Find which creators an Instagram brand actually works with, tier every collaboration, and check what the brand is running as paid ads right now. Use this whenever the user names a brand (or a list of competitor brands) and wants its collaborators, partnership posts, creator tiers, paid-partnership toggles, boosted reels, Meta Ad Library ads, or a competitor benchmark workbook — phrasings like "who does X collab with", "get the partners of @brand", "scan these brands", "4-tier", "which creators is Y running ads with", "competitor analysis for Z", or handing over a list of brand handles. Also use it when the user asks how a past brand scan was done or where its Excel lives.
---

# Instagram competitor intelligence

Answers three questions about any brand, from live data, with no paid tools:

1. **Who does this brand collaborate with?** — scan the brand's own Instagram
   feed for creator partnerships (Method A).
2. **How serious is each collaboration?** — classify every partnership post
   into the 4-tier taxonomy and every creator into an audience tier.
3. **What is the brand paying to promote right now?** — pull its live Meta
   Ad Library ads and separate creator ads from plain brand ads (Method B).

The engines already exist. Do not rewrite them; call them.

## Before starting

- Read `PROJECT_CONTEXT.md` §3 and §4b. They hold the gotchas below in detail
  and the list of brands already processed — check it before re-scanning a
  brand that has a scratch file or a sheet already.
- Credentials come from `.env` via `core/session.py`. Never hardcode a
  `sessionid`. Confirm the session works first:
  `python core/profile_auditor.py @anyhandle` must print an EXACT count.
- Instagram throttles per endpoint family. `429` or
  `{"message":"feedback_required"}` is a rate limit, **not** a missing
  account. Stop and say so; never retry in a loop. `core/profile_auditor.is_throttled()`.

## Method A — organic partnership scan

**What it does:** paginates the brand's own feed and flags a post as a
partnership when any of these hold. Signal 1 is dominant.

| # | Signal | Field |
|---|---|---|
| 1 | Post is **owned by another account** (collab posts appear on the brand's timeline with the creator as owner) | `item.user.username != brand` |
| 2 | Brand-owned post **tags a co-author** | `item.coauthor_producers` |
| 3 | Meta's paid-partnership **toggle is on** | `item.is_paid_partnership == True` |
| 4 | Caption carries `#ad #paidpartnership #sponsored #collab` — pull `@mentions` as fallback | caption regex |
| 5 | Caption **credit lines** with no tag at all — `Director:`, `Talent:`, `Starring:`, `Featuring:`, `Cast:` (seen on brand-film posts). Resolve the named person to a handle separately | caption regex |

**Run it:**

```python
from api_wrapper.client import InstagramService
svc = InstagramService()
res = svc.get_partnerships("brandhandle", time_window="1y", max_pages=35)
# time_window accepts 7d / 30d / 90d / 6m / 1y / 2y or an int of days
```

or through the API (`api/`, port 8000): `GET /api/v1/partnerships/{brand}` and
`GET /api/v1/partnerships/{brand}/usernames`.

**Coverage is the most recent N posts** (`max_pages`), not full history.
Say so in any deliverable.

**Traps that produce false partners — check every list before handing it over:**

- **Sister accounts of the same company** show up as each other's partner
  (`jockeyindia` ↔ `jockeywomanindia`). Exclude self-referential pairs.
- **Multi-brand retailers / distributors** (`makobaindia`, `starmarkindia`,
  `luxor_india`) — most co-author hits are *other brands being cross-promoted*,
  not creators. Separate brand/distributor accounts from individuals.
- **Two related brand accounts asked for together** — keep the two partner
  lists separate; do not merge or dedupe unless told to.
- **Resolve `user_id` from profile HTML** (`"profilePage_(\d+)"`) when
  `web_profile_info` 400s with the `ig_business_category_subvertical` error —
  that is an Instagram bug on some business accounts, not a cookie problem.

## Tiering

**Creator audience tier** — exact live count via `core/profile_auditor.resolve_profile()`.
Plain text, no emojis:

`Mega (1M+)` · `Macro (500K-1M)` · `Mid-Tier (100K-500K)` · `Micro (10K-100K)` · `Nano (<10K)`

Never read the count off `og:description` — it is rounded above ~10K and
lands on every creator that matters. The auditor stamps
`followers_precision`; only ship `exact`.

**4-tier collaboration taxonomy** — toggle × boost:

| Tier | Toggle | Ad spend | Meaning |
|---|---|---|---|
| 1 | ON | Boosted | Formal paid partnership + paid media |
| 2 | ON | Organic | Formal paid partnership, organic reach |
| 3 | OFF | Boosted | Undisclosed collab with heavy paid push |
| 4 | OFF | Organic | Organic / barter / low reach |

Boost detection is `core/ad_boost_engine.evaluate_boost(views, likes, followers)`:
views ≥ 500K with like rate < 0.35%, or ≥ 1M with < 0.60%, or view-to-follower
multiplier ≥ 10x with like rate < 0.50% → boosted; multiplier ≥ 25x with
organic engagement → viral organic.

Hidden like counts: a page with "hide like count" returns `like_count = 3`
as a placeholder (`like_and_view_counts_disabled: true`). Treat the like
rate as unknown, never compute on the 3.

## Method B — live Meta Ad Library ads

**Use `fb_api/` (port 8001), never the Ad Library web UI.** Scrolling the UI
hits a false plateau — a manual scroll captured 108 Underneat ads when the
true count was 384. `fb_api` replays the same persisted GraphQL query the
UI uses and proves completeness by cursor exhaustion.

```
GET /api/v1/page/{handle}                -> ad_library_page_id (use page mode, it is exact)
GET /api/v1/adlibrary/{brand}?page_id=…  -> every live ad, with reported_total / captured / complete
GET /api/v1/adlibrary/{brand}/partners   -> creator partners only
GET /api/v1/session/check                -> is the FB cookie jar alive
```

- **Branded content is structural.** `snapshot.branded_content` non-null is
  the paid-partnership signal; then `snapshot.page_name` is the *creator*
  and top-level `page_name` is the *brand*. Do not string-match
  `"<Creator> with <Brand>"`.
- **Filter self-partnerships** — brands tag their own page as the creator
  (Underneat did on 63 ads). `fb_api` counts these separately.
- **Page mode over keyword mode.** Keyword `zivame` returns ~350 ads; Zivame's
  page has 69.
- Session lives in `fb_api/.env` as `FB_COOKIE`. Needs `c_user` + `xs`.
- **The official Ad Library API is useless for India** — it only returns
  political/social-issue ads outside the EU. Don't chase developer approval.
- Never say a scroll-based capture is "confirmed complete." State the count
  and the method.

For "when did this creator's collab go live" on an existing sheet, use the
`ig-partnership-timeline` skill instead of redoing the scan.

## Deliverables

Standard shape the user expects — one sheet per brand, rows of creators:

`Username · Full Name · Followers (exact) · Tier · Collab Post URL · Toggle (ON/OFF) ·
Boost verdict · 4-Tier · Views · Likes · Comments · Posted at · Source (organic / ad library)`

Plus a Summary tab: brand, posts scanned, window, partners found, tier counts,
live-ad count, confidence note.

Rules:
- **Zero emojis, zero colour fills.** Pass every cell through
  `core.kolkata_engine.clean_cell()` — scraped captions carry emojis and
  lone UTF-16 surrogates that crash `openpyxl` on save.
- Sheet titles ≤ 31 characters.
- If the output is a handful of rows, answer in chat; the user has said
  they don't want a file for small results.
- Append new brands to the running master (`live_partnership_ads_FINAL.xlsx`
  for ad-library work) rather than starting a new file, unless told otherwise.

Where past work lives: `deliverables/` (client workbooks and CSVs),
`competitor/workbooks/` (sector benchmarks), `scratch_*.json` files are raw
per-brand dumps and safe to ignore.

## Feeding the permanent store

Every partner found should also land in `creator_intelligence.db` so it is
queryable later without another scan:

```python
from core import creator_db as cdb
conn = cdb.connect()
cdb.add_brand_collab(conn, handle, brand, post_url, is_paid_partnership, tier, taken_at)
```

Then `python core/creator_db.py ask "collaborators of @brand"` answers from
the store. See the `ig-creator-intelligence` skill.
