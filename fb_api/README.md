# Facebook Page + Ad Library API

The Facebook counterpart to [`../api/`](../api/README.md) (the Instagram
partnership + profile-metrics API), built the same way: FastAPI, `X-API-Key`
header, live per request, background jobs for anything bulk.

**Method B, properly.** The Ad Library endpoints replace the
scroll-the-UI-and-scrape-text workflow in `PROJECT_CONTEXT.md` §4. Instead of
scrolling until results stop appearing, they page through Meta's own cursor
until Meta says there are no more — so every response can state plainly
whether the scan finished, instead of guessing at a plateau.

## Where your session goes

`fb_api/.env`, same as any other secret here, and nothing else reads it:

```ini
FB_COOKIE=datr=...; sb=...; c_user=...; xs=...; fr=...
```

Getting the value: open facebook.com logged in → DevTools (F12) → **Network** →
reload → click any `facebook.com` request → **Request Headers** → right-click
the `cookie:` line → Copy value. Paste the whole thing — it needs `c_user` and
`xs` to count as a session, and a couple of hand-picked cookies won't do. A
leading `cookie:` label is stripped for you.

The header is 1–2KB, which is unpleasant on one line, so there's a file option:

```ini
FB_COOKIE_FILE=C:\Users\omkar\OneDrive\Desktop\InstagramAnalytics\fb_api\fb_cookie.txt
```

That file can hold the raw cookie header, a cookie-editor JSON export, or a
Netscape `cookies.txt` — all three parse. Keep it out of git (`fb_cookie*.txt`
is already gitignored).

**Check it before running a batch:**

```bash
curl -H "X-API-Key: your-key" "http://localhost:8001/api/v1/session/check"
```

```json
{"cookie_count":10,"cookie_names":["c_user","datr","fr","sb","xs",...],
 "logged_in":true,"missing_session_cookies":[],"source":"FB_COOKIE",
 "ad_transport":"graphql","ad_library_over_cookies":"ok",
 "probe":{"captured":27,"transport":"graphql"}}
```

It reports cookie **names** only, never values, then makes one real Ad Library
request to settle it. `"ad_library_over_cookies":"challenged"` means Meta
rejected the session as anonymous or stale — re-copy the header from a live
tab. Page endpoints keep working either way; they don't need a session at all.

Verified on a real session (2026-08-22): Zivame page mode returned 69/69 ads in
9s, and an Underneat keyword scan 381 ads / 93 partners in 36s — both over
cookies alone, no browser.

## Setup

```bash
cd fb_api
pip install -r requirements.txt
copy .env.example .env
```

No Playwright, no browser download — the default path is cookies over HTTP.

Set `API_KEY` (`python -c "import secrets; print(secrets.token_urlsafe(32))"`)
and `FB_COOKIE`. Then:

```bash
uvicorn main:app --reload --port 8001
```

Port 8001 keeps it clear of the Instagram API on 8000 — both can run at once.

## Endpoints

All require `X-API-Key`. `GET /health` doesn't, and reports which Ad Library
transport is active.

### Pages

**`GET /api/v1/page/{handle}`** — page ids, category, likes, talking-about,
follower estimate, bio.

```bash
curl -H "X-API-Key: your-key" "http://localhost:8001/api/v1/page/zivame"
```

```json
{"handle":"zivame","name":"Zivame","page_id":"100064316664778",
 "ad_library_page_id":"234603919914240","category":"Clothing (Brand)",
 "likes":794213,"talking_about":4694,"followers_display":"794K",
 "followers_approx":794000,"followers_is_approximate":true,"bio":"..."}
```

`ad_library_page_id` is the id to feed the Ad Library endpoints — see below,
it matters.

**`GET /api/v1/page/{handle}/posts?limit=12`** — recent content with reactions
and play counts, plus averages and an engagement rate.

**`POST /api/v1/page/bulk`** — `{"handles": [...], "include_posts": false}`,
returns a `job_id`.

### Ad Library

**`GET /api/v1/adlibrary/{brand}`** — every ad Meta will show for the brand.

```bash
# exact (preferred)
curl -H "X-API-Key: your-key" \
  "http://localhost:8001/api/v1/adlibrary/zivame?page_id=234603919914240"

# keyword search
curl -H "X-API-Key: your-key" "http://localhost:8001/api/v1/adlibrary/underneat"
```

Query params: `page_id`, `country` (default `IN`), `active_status`
(`active|inactive|all`), `media_type`, `max_ads`, `include_ads`.

**`GET /api/v1/adlibrary/{brand}/partners`** — just the creators, the
analogue of the Instagram API's `/partnerships/{brand}/usernames`:

```json
{"query":"page:234603919914240","reported_total_ads":null,"captured_ads":69,
 "complete":true,"completeness_basis":"cursor exhausted (Meta reported no further pages)",
 "branded_content_ads":5,"self_partnership_ads":0,"partner_count":5,
 "partners":[{"name":"Ishita Anand","profile_url":"https://www.facebook.com/anandishii/",
              "page_likes":41648,"ad_count":1,"first_seen":"2026-06-30",
              "last_seen":"2026-06-30","ad_library_urls":["..."]}]}
```

**`POST /api/v1/adlibrary/bulk`** — `{"brands": [...]}`; an all-digits entry is
treated as a Page id, anything else as a keyword.

**`GET /api/v1/jobs/{job_id}`** — poll any bulk job; `?include_results=false`
for progress only.

### Official Graph API (optional, needs a token)

**`GET /api/v1/official/ads`** and **`GET /api/v1/official/page/{page_id}`**
wrap Meta's sanctioned endpoints. Read the limitation below before bothering.

## Things worth knowing before you trust a number

**Page mode beats keyword mode, by a lot.** A keyword search for `zivame` in
India returns ~350 ads; Zivame's actual Page has 69. The rest are other
advertisers whose ad text happens to mention the brand — the same noise
`PROJECT_CONTEXT.md` §4 warns about for Jockey, Clovia, Lyra and Amante. Get
the page id from `/api/v1/page/{handle}` and pass it.

**Completeness is reported, not assumed.** Every Ad Library response carries
`captured`, `complete`, `completeness_basis` and a `warning`. This is the
direct fix for the false-plateau problem: the old scroll capture put Underneat
at 108 creator ads and looked done, while TrendTrack showed 208. This API pulls
**381 ads / 93 partners** for the same brand in ~36 seconds and reports
`"cursor exhausted"`, meaning Meta itself said there was no next page.

On the `http`/`browser` transports you also get `reported_total` — Meta's own
count for the query — and then `complete` can be checked the second way, by
comparison. A small gap there (a handful of ads) is normally Meta collating
duplicate versions of one ad; a large one means the scan is genuinely short,
and `warning` says which you're looking at. On the default `graphql` transport
`reported_total` is `null`, because that field only exists in the HTML
payload.

**Branded content is structural now.** The old method inferred creator ads
from an advertiser label reading `<Creator> with <Brand>`. Meta's payload
actually carries a `branded_content` object, so a creator ad is identified
directly and comes with the creator's page name, profile URL, page id and like
count. Ads where a brand tags *itself* as the creator are excluded from
`partners` and counted in `self_partnership_ads` (Underneat has 63 of them —
they would otherwise show up as the brand's biggest "partner").

**Followers are approximate; likes are exact.** Facebook renders `794K
followers` to logged-out callers and never the precise number. `likes` and
`talking_about` are exact. Note these are different things on Facebook — a
Page can be liked without being followed — so don't line `followers_approx` up
against an Instagram follower count as if they measure the same thing.

**There's no verified/blue-badge flag.** Facebook doesn't put one in the
anonymous page payload. The nearest lookalike fields in the same blob belong
to other entities, so the endpoint omits it rather than reporting a confident
guess. The Instagram API does return a real `verified` flag; these two are not
comparable.

**`/posts` is a sample, not a feed.** Logged-out, Facebook server-renders one
timeline story and about six videos. That's the ceiling, not a parser
limitation — there's no anonymous equivalent of Instagram's 12-per-request
feed pagination.

**The official Ad Library API is the wrong tool for this project.** Meta's
`ads_archive` reference states that ads not reaching EU locations are only
returned if they're categorised as social issues, elections or politics. For
India that means commercial D2C brand ads are simply not in the official
archive — a perfect token returns zero rows for Zivame, Underneat or
Mamaearth. `/api/v1/official/*` is there for EU/UK ads (commercial ads *are*
in the archive there, post-DSA, with reach data) and for political ads
anywhere. Everything else should use `/api/v1/adlibrary/*`.

## How the Ad Library endpoints work

Meta's anti-bot challenge sits on the `/ads/library/` **HTML route** — and it
fires there even for a fully logged-in session, so "just send cookies at the
page" doesn't work. What does work: that challenge is not on
`/api/graphql/`.

So `AD_TRANSPORT=graphql` (the default) never loads that route at all. It
takes the LSD and DTSG tokens off the ordinary facebook.com homepage — which
isn't challenged — and queries `AdLibrarySearchPaginationQuery` directly with
your cookies, building the query variables itself instead of lifting them out
of a rendered page. Tokens are cached for 10 minutes so a batch doesn't refetch
the homepage per brand.

The tradeoff: the HTML payload carries Meta's own `count` for a query, and the
GraphQL response doesn't. So on this transport `reported_total` is `null` and
completeness is proven the other way — by paginating until Meta reports
`has_next_page: false`. `completeness_basis` in every response says which test
was used (`"cursor exhausted"` vs `"captured >= Meta's reported total"`).
Running out of cursor is the stronger evidence of the two; it's Meta saying
there is nothing left, not a number happening to match.

The other two transports are still there, selectable per call with
`?transport=`:

- `http` — same cookies, but against the HTML route, so you get `reported_total`
  back. Expect `503`; Meta challenges it. Kept in case that changes.
- `browser` — headless Chromium via Playwright. Needs no cookies at all, and
  it's the one thing that can rediscover the GraphQL `doc_id` automatically if
  Meta rotates it. Install only if you want that: `pip install playwright &&
  python -m playwright install chromium`.

**If scans suddenly fail with "Meta most likely rotated the doc_id":** that's
the persisted-query id changing when Meta ships a new JS bundle. Either run one
scan with `?transport=browser` (it rediscovers and caches the new id), or open
the Ad Library in your browser, search the JS bundles for
`AdLibrarySearchPaginationQuery_facebookRelayOperation`, and put the number
next to it in `AD_DOC_ID` in `.env`.

Nothing is scraped off rendered pixels on any transport. Meta server-renders the
Rough timings on the cookie transport: 69 ads in ~9s, 381 ads in ~36s (13
pages). `MAX_CONCURRENT_AD_SCANS` only throttles browser scans; cookie scans
are plain HTTP and run freely.

## Before this is reachable by anything other than you

- Tighten `ALLOWED_ORIGINS` in `.env` to your actual domain instead of `*`.
- Don't commit `.env` — it holds your API key (already gitignored).
- `FB_COOKIE` is a real Facebook session — anyone holding it is logged in as
  you. It lives only in `.env` (gitignored), never in code, which is the one
  thing the Instagram side got wrong (`PROJECT_CONTEXT.md` §5: a live IG
  `sessionid` is hardcoded in `scrape_profiles.py` and already pushed to
  GitHub). Keep it that way, and if you ever paste a cookie into a `.py` file
  here, rotate it. `fb_cookie*.txt` / `fb_cookie*.json` are gitignored too.
- Facebook sessions expire. When scans start returning `503`, re-copy the
  header and restart the API — `/api/v1/session/check` tells you in one call
  whether that's what happened.
