# Instagram Analytics — Project Context

**Read this first in any new session on this project.** It captures everything built and learned so far, so you don't have to re-derive it from scratch or re-discover the same bugs.

---

## 1. What this project is

A collection of scripts + one-off analyses for:
1. **Scraping Instagram profiles/posts** (followers, engagement, captions) for a list of accounts.
2. **Finding "partnership" accounts** for a brand two different ways:
   - **Method A — Organic scan**: scan the brand's *own* Instagram feed for posts that are actually co-authored by / collab-tagged with another account (creator collab posts show up on the brand's own timeline).
   - **Method B — Paid scan**: search Meta's Ad Library for the brand's *currently active paid ads*, and identify which of those are boosted creator/branded-content ads (as opposed to plain brand-produced ads).
3. Merging that partner list with profile metrics (via Method A's scraper) and building Excel deliverables.

The user is doing competitive/vendor research on Indian D2C brands (lingerie, skincare, jewelry, hair/body care) — who they collab with, how many, what tier of creator, and what's currently live as a paid ad.

---

## 2. Key files in this repo

| File | What it does |
|---|---|
| `scrape_profiles.py` | Standalone script. Fill `USERNAMES` list + `COOKIES`, run it → `instagram_profile_analysis.xlsx` with follower/engagement metrics per account (3 sheets: Overview, Detailed Posts, Captions). Has a `_cache.json` safety net now (see §5) and a locked-file fallback so it never loses a scrape. |
| `scrape_bulk.py` | Standalone script for scraping metrics/comments off a specific list of (Name, Post URL) pairs rather than whole profiles. |
| `find_partnership_accounts.py` | Method A implementation as a standalone script — resolves a target username's partnership/collab accounts from its own feed. Target defaults to `officialzivame`; edit `TARGET_USERNAME` to reuse. |
| `PARTNERSHIP_INTEGRATION.md` | Write-up of Method A's logic + code, written for the user to drop into their own website. **Contains a real live Instagram session cookie in plain text — it's git-ignored on purpose, never remove it from `.gitignore`, never commit it.** |
| `live_partnership_ads_FINAL.xlsx` | The running Method B deliverable — one sheet per brand, each row = one live creator/partnership ad (Partner Name, Ad Library URL, Started Running, Caption). Summary tab tracks per-brand status/confidence. **Keep appending to this file for new brands rather than starting a new one, unless told otherwise.** |
| `api/` | The Instagram partnership + profile-metrics API (Method A as HTTP endpoints). Needs the live IG session cookie. |
| `fb_api/` | The Facebook API: Page metrics/posts + Meta Ad Library scans with completeness metadata. No credentials needed. See §4b. |
| `PROJECT_CONTEXT.md` | This file. |

Various `scratch_*.json` files in the repo root are working data dumps from individual scrape runs (per-brand partnership hits, Ad Library raw exports). Safe to ignore/delete; they're not deliverables, just intermediate caches from the sessions that produced them.

---

## 3. Method A — Organic partnership scan (Instagram's own feed API)

**Use when:** asked for "the partnership accounts for @brand" from raw Instagram data, or need it merged with profile metrics.

### How to resolve a user_id
`web_profile_info` REST endpoint sometimes 400s with a bizarre error (`"...ig_business_category_subvertical has been deleted..."`) — this is an **Instagram-side bug for certain business accounts**, not a cookie problem. Workaround: fetch the plain profile HTML and regex out `"profilePage_(\d+)"`:

```python
r = session.get(f"https://www.instagram.com/{username}/", headers=hdrs, cookies=COOKIES, timeout=15)
m = re.search(r'"profilePage_(\d+)"', r.text)
user_id = m.group(1) if m else None
```

### How to paginate posts and detect partnerships
Use the private feed endpoint (works even when `web_profile_info` is broken):
`https://www.instagram.com/api/v1/feed/user/{user_id}/?count=12&max_id={cursor}` — paginate via `next_max_id` / `more_available`.

For each item, a post counts as a **partnership signal** if any of:
1. `item['user']['username'] != target` — the post is *owned by the partner account*, not the brand. **This is the dominant signal** — Instagram's collab feature makes a creator's own post appear on the brand's timeline too, with `owner` = the creator.
2. `item['coauthor_producers']` contains someone other than the target — for posts the brand itself owns but tags a co-author (e.g. another brand, an event partner, a hospital, etc.).
3. `item['is_paid_partnership']` is `True`, or caption matches `#ad|#paidpartnership|#sponsored|#collab|#partnership` — pull `@mentions` out of the caption as a fallback.

**Watch out:** two sister/sub-brand accounts of the *same* company (e.g. `jockeyindia` and `jockeywomanindia`) will show up as each other's "partner" via signal #1/#2 — this is not a real third-party collab, just cross-posting. Check for and exclude self-referential brand pairs before handing over a partner list.

**Watch out #2 (multi-brand retailers):** for a brand whose Instagram is run by a retailer/distributor (e.g. Makoba India sells Monteverde, Diplomat, Conklin, Leonardo, Ranga, Penlux etc. as a multi-brand pen store), most "coauthor tag" hits will be *other brands being cross-promoted*, not creator partnerships. Same for pen-industry distributors like `starmarkindia`, `luxorparker`/`luxor_india` (Parker's Indian licensee, Luxor), `harmonydistributors`. Flag these as brand/distributor accounts, separate from genuine individual creators, rather than dumping them into one undifferentiated partner list.

**Signal #5 — caption credit lines (no IG tag at all):** some brands (seen on `montblanc`'s produced brand-film posts) don't use IG's collab/coauthor feature for talent at all — the post caption just has plain-text lines like `Director: Roman Coppola` / `Talent: Daniel Brühl`. Regex for `(Director|Talent|Starring|Featuring|Cast):\s*(.+)` in the caption to catch these. These names usually have **no @mention in the same post** — resolve their real IG handle separately (WebSearch + verify follower count/verified badge on the account before trusting it), or check if a *different* post on the same target account coauthor-tags that same person (happened for Daniel Brühl/Rupert Friend on montblanc — the credit and the IG tag were on two different posts for the same campaign).

**Coverage:** scans are capped (`MAX_POSTS`, usually 400) — this is the *most recent* N posts, not full account history, unless the account has fewer posts than the cap.

### Output format the user likes
```python
USERNAMES = [
    "handle1",
    "handle2",
]
# N unique accounts, from @brand's last M posts
```
Ready to paste into `scrape_profiles.py`'s `USERNAMES` list to get follower/engagement metrics on all of them.

**If asked to scan two related brand accounts, keep the two partner lists SEPARATE (don't merge/dedupe them) unless explicitly told to combine.**

---

## 4. Method B — Meta Ad Library live-ads scan

**Use when:** asked for "live"/"active" partnership *ads* (as opposed to organic posts).

### Search
`https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country=IN&q={brand}&search_type=keyword_unordered&media_type=all`

Generic brand names (Jockey, Clovia, Lyra, Amante...) pull in heavy noise from unrelated advertisers — always filter results down to advertiser names that actually contain the brand string, and sanity-check the filtered set before trusting a "0 results" conclusion.

### Loading ALL the ads — this is the part that's easy to get wrong
The page lazy-loads on scroll. **A single scroll-to-bottom pass is not enough** — you have to loop:
```js
async function step() {
  window.__adLibSeen = window.__adLibSeen || {};
  for (let i = 0; i < 14; i++) { window.scrollTo(0, document.body.scrollHeight); await new Promise(r => setTimeout(r, 500)); }
  const blocks = document.body.innerText.split(/\n​\nActive\n/).slice(1);
  blocks.forEach(b => {
    const idm = b.match(/Library ID: (\d+)/); if (!idm) return;
    if (window.__adLibSeen[idm[1]]) return;
    const advm = b.match(/See (?:ad|summary) details\n([^\n]+)\nSponsored/); if (!advm) return;
    const datem = b.match(/Started running on ([^\n·]+)/);
    const afterSponsored = b.slice(b.indexOf('Sponsored\n') + 10);
    let caption = afterSponsored.split(/\d:\d\d ?\/ ?\d+:?\d*/)[0].trim();
    window.__adLibSeen[idm[1]] = {library_id: idm[1], started: datem?.[1]?.trim(), advertiser: advm[1].trim(), caption};
  });
  return Object.keys(window.__adLibSeen).length;
}
```
Call `step()` repeatedly (as separate tool calls — one call's worth of scrolling risks a 30s timeout on big pages) until **two consecutive calls add zero new ads**. `window.__adLibSeen` persists across calls as long as you don't navigate away.

### ⚠️ CRITICAL CAVEAT — verified, not theoretical
**This scroll-plateau is a FALSE plateau.** Cross-checked against TrendTrack (a proper Ad Library-API-backed tool) for Underneat: we captured 108 creator ads / 75 partners after "reaching plateau"; TrendTrack showed 208 ads / 88 partners for the same brand at the same time. **We were missing roughly half the real ad volume**, and it wasn't random — the scroll-triggered lazy-load in the web UI has a soft ceiling that has nothing to do with the true total. A second cross-check (Nykd) came out much closer (74 vs 85 ads, 27/27 partners matched) — so the gap size is inconsistent and unpredictable per brand.

**Bottom line: never claim a scroll-based capture is "confirmed complete."** State the raw count, note it's scroll-based, and if the user has (or can get) numbers from a real Ad Library API tool, that's the number to trust over ours. Retract "confirmed complete" language from any brand sheet if this comes up again — it happened across the whole `live_partnership_ads_FINAL.xlsx` deliverable and needed walking back.

### Detecting creator vs. plain-brand ads
Meta's advertiser field literally spells this out: branded-content/collab ads show as `"<Creator Name> with <Brand>"`. Plain brand ads just show the brand name. Filter:
```js
const isCreator = a => / with /i.test(a.advertiser) && !new RegExp('^'+brandKey+' with','i').test(a.advertiser);
```

### Getting large datasets out of the browser efficiently
**Don't manually transcribe large JS tool outputs into Write calls** — it's slow and error-prone (did this for the first few brands before finding a better way). Instead, trigger a real browser download and grab it with Bash — **the browser pane and Bash share the same filesystem in this environment**, downloads land in `OneDrive/Desktop`:
```js
const blob = new Blob([JSON.stringify(data)], {type: 'application/json'});
const url = URL.createObjectURL(blob);
const el = document.createElement('a');
el.href = url; el.download = 'whatever.json';
document.body.appendChild(el); el.click(); document.body.removeChild(el);
```
Then `mv "/c/Users/omkar/OneDrive/Desktop/whatever.json" "<project>/scratch_whatever.json"`. **First `mv` attempt sometimes fails with "No such file" if the download hasn't flushed yet — just retry once, it's a timing issue, not a real failure.** Also: Chrome throttles/silently drops automatic downloads after several in a row without a user gesture in between — if a download goes missing, just re-trigger the same JS call once more.

### Excel cell surrogate-character bug
Captions copied out of the Ad Library sometimes contain lone UTF-16 surrogate characters (from JS `.slice()` cutting a 4-byte emoji in half) that crash `openpyxl` on save with `UnicodeEncodeError: surrogates not allowed`. Always sanitize before writing:
```python
import re
SURROGATE_RE = re.compile('[\ud800-\udfff]')
def clean(s): return SURROGATE_RE.sub('', s) if isinstance(s, str) else s
```

---

## 4a. `scrape_bulk.py` comment pagination — known-bad `has_more_comments`, fixed 2026-08-18

`fetch_comments()` in `scrape_bulk.py` paginates `GET /api/v1/media/{id}/comments/?can_support_threading=true`. Instagram's `has_more_comments` field on this endpoint is **unreliable — it reports `False` even when more pages genuinely exist** (verified empirically: manually continuing with `next_min_id` after a `has_more_comments: False` response returns a fresh page of different comments). Any version of this function that guards its pagination loop on `has_more_comments` will silently stop after page 1, capping every post at ~15 comments regardless of the configured `MAX_COMMENTS`. This was true from when the script was first written, not just after a later edit.

**Fix:** ignore `has_more_comments` entirely. Keep paginating as long as `next_min_id` is present and different from the previous `min_id`; stop on 2 consecutive empty pages (dedupe via seen comment `pk`/`id`) or when `next_min_id` repeats/disappears. This is what's in the repo now — don't reintroduce a `has_more_comments` check.

## 4b. `fb_api/` — Method B as a real API, 2026-08-22 (supersedes the scroll workflow in §4)

`fb_api/` is the Facebook counterpart to `api/` (same FastAPI shape, `X-API-Key`, live per request, background bulk jobs). Runs on port 8001 so it can sit alongside the Instagram API on 8000. **It needs no credentials** — Facebook Page data and the Ad Library both work logged-out, unlike anything on the Instagram side.

**Use `/api/v1/adlibrary/{brand}` instead of scrolling the Ad Library UI.** The scroll method's false plateau (§4) is solved: Meta server-renders the first page of results into the page HTML as a Relay payload under `"search_results_connection"`, and that payload carries `count` — **Meta's own total for the query**. Further pages come from the same persisted GraphQL query the UI fires on scroll (`AdLibrarySearchPaginationQuery`, cursor from `page_info.end_cursor`). So every response reports `reported_total` / `captured` / `complete` / `warning`, and you never have to guess whether a capture finished. Measured 2026-08-22: Underneat came back **382 of 384 ads in ~40s** (the manual scroll got 108 ads / 75 partners; TrendTrack showed 208 / 88).

Other things that changed with it:

- **Branded content is structural, not string-matched.** Forget parsing `"<Creator> with <Brand>"` advertiser labels. `result.snapshot.branded_content` non-null IS the paid-partnership signal, and in that case `snapshot.page_name` / `page_profile_uri` / `page_like_count` are the **creator's** page while top-level `page_id`/`page_name` are the **brand**. You get creator page ids and like counts for free.
- **Self-partnerships must be filtered.** Underneat tags its own page as the creator on 63 ads — same trap as the sister-brand cross-tagging in §3. `fb_api` excludes these from `partners` and counts them in `self_partnership_ads`.
- **Page mode >> keyword mode.** Keyword `zivame` in IN returns ~350 ads; Zivame's actual page has 69. `GET /api/v1/page/{handle}` returns `ad_library_page_id` (from the page HTML's `delegate_page`) — pass it as `?page_id=` for an exact scan.
- **Anti-bot challenge → the /ads/library/ HTML route is unusable, the GraphQL endpoint is not.** Meta serves HTTP 403 + its `__rd_verify_` challenge on `/ads/library/` **even for a fully logged-in session** (verified 2026-08-22 with a real FB cookie jar). It does **not** challenge `/api/graphql/`. So the default `AD_TRANSPORT=graphql` never loads that route: it pulls LSD/DTSG tokens off the plain facebook.com homepage (not challenged, cached 10 min) and queries `AdLibrarySearchPaginationQuery` directly with the session cookie, building the variables itself. Measured on the real session: Zivame page mode 69 ads in 9s, Underneat keyword 381 ads / 93 partners in 36s, both cursor-exhausted. **No Playwright needed** — `AD_TRANSPORT=browser` is an opt-in fallback that also auto-rediscovers the doc_id if Meta rotates it (else set `AD_DOC_ID` in .env).
- **Completeness on the GraphQL path is proven by cursor exhaustion, not by a count.** The HTML payload's `count` (Meta's own total) does not exist in the GraphQL response, so `reported_total` is null there and `completeness_basis` says `"cursor exhausted"` — Meta reporting `has_next_page: false`. That is stronger evidence than a number matching; don't "fix" it by reintroducing a count comparison.
- **Session lives in `fb_api/.env`** as `FB_COOKIE` (whole cookie header, must include `c_user` + `xs`) or `FB_COOKIE_FILE` (raw header / cookie-editor JSON / Netscape cookies.txt all parse). `GET /api/v1/session/check` reports cookie names (never values), whether the jar is logged in, and makes one live Ad Library call to prove it. A stale session = `503` with the reason spelled out. **Unlike the Instagram scripts, the FB session is never hardcoded in a .py file** — .env only, see §5.
- **The official Ad Library API is useless for this project.** Meta's `ads_archive` docs state that ads not reaching EU locations are only returned if categorised as social issues / elections / politics. For India that means **zero commercial brand ads**, no matter how good your token. Don't burn time on developer-app approval for it. `/api/v1/official/*` is wrapped anyway, for EU/UK ads (commercial ads are in the archive there post-DSA) and political ads.

**Facebook Page data caveats** (`/api/v1/page/{handle}`): likes and talking-about are exact; **followers are rounded only** ("794K") because that's all Facebook renders logged-out. There is **no verified/blue-badge flag** in the anonymous payload — the endpoint deliberately omits it rather than reading a lookalike field belonging to some other entity. `/posts` returns one timeline story plus ~6 videos (with play + reaction counts) — that's Facebook's ceiling logged-out, not a parser limit.

## 5. Other things worth knowing

- **Rate limiting**: this Instagram session's cookies get rate-limited (`429`) under heavy sequential use across a session — if a script is stuck retrying, kill it and check with a single lightweight request before relaunching a big batch, rather than assuming it's still making progress.
- **`scrape_profiles.py` won't lose a scrape anymore**: it now dumps `profiles_data` to a `*_cache.json` file right after scraping, before touching Excel, and falls back to a timestamped filename if the primary `.xlsx` is locked (open in Excel) instead of crashing. If a future run crashes anyway, check for the cache JSON before re-scraping from scratch.
- **⚠️ Live credential exposure, unresolved**: `scrape_profiles.py` and `scrape_bulk.py` have a real Instagram `sessionid`/`csrftoken`/etc. hardcoded, and that commit (`354a77c`) is already pushed to `origin/main` on GitHub (`quvoid/InstagramAnalyticsVersion2`). The user was told the fix is (a) rotate the session in Instagram's security settings — invalidates the leaked cookie regardless of git state — and (b) redact + history-rewrite the repo if they want it actually gone from GitHub. As of the last session this was **not yet done** — worth checking in on if it comes up, don't assume it's resolved.
- **Don't build Excel/files unless it's actually warranted.** The user has explicitly said "just give me the answer in chat" more than once when the output was small enough. Use judgment on volume — a handful of rows goes in chat; hundreds of rows goes in a file. When in doubt for a genuinely large ask, ask which they want rather than assuming.
- **Tone**: the user is blunt/informal (and occasionally hostile when frustrated with pace/accuracy) — don't take it personally, just adjust: stop chasing marginal completeness and deliver what exists when told to.

---

## 6. Brands processed so far (status snapshot — may be stale, check scratch files' dates)

**Method A (organic scan) done for:** officialzivame, moxiebeautyofficial, indewild, mamaearth.in, jockeywomanindia, jockeyindia, osmo.wellversed, ufcindia, chungreng_koren (this one's a partner account, not a brand — scanned for engagement/comments only), montblanc, parkerpenindia, crosspens.india, sheafferpen, lamy_india, makobaindia, submarine_pens (the last 7 = pen/writing-instrument brands, 2026-08-18/19 session; montblanc used the new caption-credit signal — see §3 signal #5).

**Meta Ad Library (Method B) — pen brands, 2026-08-19:** montblanc, parkerpenindia, crosspens.india, sheafferpen, makobaindia = **0 genuine creator-partnership ads** in India (keyword search returns only unrelated resellers/hotels/noise, or literally 0 results for Makoba). lamy_india and submarine_pens **do run active Meta ads directly** (12 and 28 respectively) but all are plain product/catalog ads — no `<creator> with <brand>` branded-content pattern found on any of the 7.

**Method B (Ad Library) done for**, with rough confidence:
| Brand | Creator ads found | Confidence |
|---|---|---|
| Underneat | 108 (+ 18 more known from TrendTrack) | Undercounted — proven |
| Zivame | 6 | Cross-checked, solid |
| Nykd By Nykaa | 74 (+1 more known) | Cross-checked, close |
| Palmonas | 86 | Scroll-plateau only, unverified |
| Hyphen | 458 | Scroll-plateau only, unverified |
| Giva | 150 | Scroll-plateau only, unverified |
| mCaffeine | 105 | Stopped early, not exhausted |
| Mama Earth | 31 | Old partial capture, not re-verified |
| Jockey India | 0 | Genuine — verified via advertiser breakdown |
| Clovia | 0 | Genuine — verified via advertiser breakdown |
| Lyra India | 0 active ads at all | Genuine — brand has no current Meta spend |
| Amante | 3 | Partial — tab hit perf limits mid-scroll |

All of this lives in `live_partnership_ads_FINAL.xlsx`. Also produced separately: `mamaearth_partnership_with_metrics.xlsx`, `partnership_ads_dates.xlsx`, `partnership_ads_batch2.xlsx`, `meta_ad_library_snapshot.xlsx` (earlier one-off deliverables, may be superseded by the FINAL file for overlapping brands).
