"""
Instagram Metrics Scraper — Streamlit App
==========================================
Run with:  streamlit run app.py

Install:
    pip install streamlit curl_cffi requests openpyxl

Fill in your Instagram cookies in the COOKIES dict below.
"""

import streamlit as st
import time, random, json, io
from datetime import datetime, timezone

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from curl_cffi import requests as cffi_requests
    USE_CURL_CFFI = True
except ImportError:
    import requests as cffi_requests
    USE_CURL_CFFI = False

# ── PAGE CONFIG ───────────────────────────────────────────────
st.set_page_config(
    page_title="Instagram Metrics",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── CUSTOM CSS ────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');

*, html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
}

.stApp {
    background: #0A0A0A;
}

header[data-testid="stHeader"] { background: transparent; }

/* Inputs */
.stTextInput > div > div > input,
.stNumberInput > div > div > input {
    background: #141414 !important;
    border: 1px solid #282828 !important;
    border-radius: 10px !important;
    color: #E8E8E8 !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-size: 14px !important;
    padding: 12px 16px !important;
}
.stTextInput > div > div > input:focus,
.stNumberInput > div > div > input:focus {
    border-color: #E8405A !important;
    box-shadow: 0 0 0 2px rgba(232,64,90,0.15) !important;
}
.stTextInput label, .stNumberInput label {
    color: #888 !important;
    font-size: 11px !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 0.09em;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #E8405A, #FF6B35) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    padding: 13px 28px !important;
    width: 100% !important;
    letter-spacing: 0.03em;
    transition: all 0.2s ease !important;
}
.stButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 8px 24px rgba(232,64,90,0.3) !important;
}

.stDownloadButton > button {
    background: linear-gradient(135deg, #1DB954, #17A449) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    padding: 13px 28px !important;
    width: 100% !important;
    letter-spacing: 0.03em;
    transition: all 0.2s ease !important;
}
.stDownloadButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 8px 24px rgba(29,185,84,0.3) !important;
}

/* Expander */
.streamlit-expanderHeader {
    background: #141414 !important;
    border-radius: 10px !important;
    color: #CCC !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
}

/* Spinner */
.stSpinner > div { border-top-color: #E8405A !important; }

/* Scrollbar */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: #0A0A0A; }
::-webkit-scrollbar-thumb { background: #2A2A2A; border-radius: 3px; }

/* Metric grid */
.metric-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 10px;
    margin: 18px 0;
}
.metric-card {
    background: #141414;
    border: 1px solid #1E1E1E;
    border-radius: 12px;
    padding: 18px 14px;
    text-align: center;
}
.metric-value {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 24px;
    font-weight: 800;
    color: #E8E8E8;
    line-height: 1;
    letter-spacing: -0.02em;
}
.metric-label {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 10px;
    font-weight: 600;
    color: #555;
    margin-top: 6px;
    text-transform: uppercase;
    letter-spacing: 0.1em;
}
.metric-er .metric-value { color: #E8405A; }
.metric-followers .metric-value { color: #FF6B35; }

/* Profile card */
.profile-card {
    background: #141414;
    border: 1px solid #202020;
    border-radius: 14px;
    padding: 22px 20px;
    margin: 18px 0;
}
.profile-name {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 20px;
    font-weight: 800;
    color: #F0F0F0;
    letter-spacing: -0.01em;
}
.profile-username {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 13px;
    font-weight: 500;
    color: #E8405A;
    margin-top: 3px;
}
.profile-bio {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 13px;
    font-weight: 400;
    color: #777;
    margin-top: 10px;
    line-height: 1.55;
}
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 11px;
    font-weight: 600;
    margin-right: 6px;
    margin-top: 8px;
}
.badge-verified  { background: #1a2d40; color: #4fc3f7; }
.badge-business  { background: #2d1b00; color: #FF8C42; }
.badge-category  { background: #1e1228; color: #B07FE8; }

/* Section label */
.section-label {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 11px;
    font-weight: 700;
    color: #444;
    text-transform: uppercase;
    letter-spacing: 0.14em;
    margin: 26px 0 10px;
}

/* Post table */
.post-table {
    background: #0E0E0E;
    border: 1px solid #1A1A1A;
    border-radius: 12px;
    overflow: hidden;
}
.post-row {
    display: grid;
    grid-template-columns: 36px 90px 80px 1fr 80px 80px 68px;
    align-items: center;
    padding: 9px 14px;
    border-bottom: 1px solid #161616;
    gap: 4px;
}
.post-row:last-child { border-bottom: none; }
.post-row-header {
    background: #141414;
    border-bottom: 1px solid #202020 !important;
}
.post-row-header span {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 10px;
    font-weight: 700;
    color: #555;
    text-transform: uppercase;
    letter-spacing: 0.1em;
}
.post-cell {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 12px;
    font-weight: 400;
    color: #BDBDBD;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.post-cell-num {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 12px;
    font-weight: 600;
    color: #D0D0D0;
    text-align: center;
}
.post-cell-idx {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 11px;
    color: #444;
    text-align: center;
}
.post-cell-er {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 12px;
    font-weight: 700;
    color: #E8405A;
    text-align: center;
}
.type-image    { color: #4A9EFF; font-weight: 600; }
.type-video    { color: #A78BFA; font-weight: 600; }
.type-carousel { color: #FBBF24; font-weight: 600; }
.post-link {
    color: #4A9EFF;
    text-decoration: none;
    font-size: 12px;
    font-family: 'Plus Jakarta Sans', sans-serif;
}
.post-link:hover { text-decoration: underline; }
</style>
""", unsafe_allow_html=True)


# ── COOKIES ───────────────────────────────────────────────────
DEFAULT_COOKIES = {
    "sessionid":  "25113411270%3AyQFaao428g6Xb9%3A0%3AAYgcV3Qe5uiiJD1FyRkSD2vULzcMkEegPbBVeDmWYQ",
    "csrftoken":  "3gJbkGDZp99lA8QQ0brobyoHzOreuu8f",
    "mid":        "afyCbwALAAFRStE-k17-dfO5_jfa",
    "ds_user_id": "25113411270",
}

IG_APP_ID = "936619743392459"

BASE_HEADERS = {
    "accept":             "*/*",
    "accept-language":    "en-US,en;q=0.9",
    "accept-encoding":    "gzip, deflate, br",
    "user-agent":         (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "x-ig-app-id":        IG_APP_ID,
    "x-ig-www-claim":     "0",
    "x-asbd-id":          "129477",
    "x-requested-with":   "XMLHttpRequest",
    "x-csrftoken":        "",
    "origin":             "https://www.instagram.com",
    "referer":            "https://www.instagram.com/",
    "sec-ch-ua":          '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
    "sec-ch-ua-mobile":   "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest":     "empty",
    "sec-fetch-mode":     "cors",
    "sec-fetch-site":     "same-origin",
}


# ── HELPERS ───────────────────────────────────────────────────

def extract_username(raw: str) -> str:
    raw = raw.strip().lstrip("@")
    if "instagram.com" in raw:
        parts = [p for p in raw.replace("https://", "").replace("http://", "").split("/") if p and "instagram.com" not in p]
        raw = parts[0] if parts else raw
    return raw.split("?")[0].strip()

def human_delay(a=1.0, b=3.0):
    time.sleep(random.uniform(a, b))

def make_session(cookies: dict):
    headers = {**BASE_HEADERS, "x-csrftoken": cookies.get("csrftoken", "")}
    if USE_CURL_CFFI:
        s = cffi_requests.Session(impersonate="chrome120")
    else:
        import requests
        s = requests.Session()
    s.headers.update(headers)
    s.cookies.update(cookies)
    return s


# ── SCRAPER FUNCTIONS ─────────────────────────────────────────

def fetch_profile(username, session, cookies, status_fn=None):
    urls = [
        f"https://www.instagram.com/api/v1/users/web_profile_info/?username={username}",
        f"https://i.instagram.com/api/v1/users/web_profile_info/?username={username}",
    ]
    headers = {**BASE_HEADERS, "x-csrftoken": cookies.get("csrftoken", ""),
               "referer": f"https://www.instagram.com/{username}/"}

    for url in urls:
        for attempt in range(1, 4):
            try:
                if status_fn: status_fn(f"Fetching profile (attempt {attempt}/3)...")
                if USE_CURL_CFFI:
                    r = session.get(url, headers=headers, cookies=cookies,
                                    timeout=20, impersonate="chrome120")
                else:
                    r = session.get(url, headers=headers, cookies=cookies, timeout=20)

                if r.status_code == 200:
                    return r.json()
                elif r.status_code in (401, 403, 404):
                    return {"_error": r.status_code}
                elif r.status_code == 429:
                    if status_fn: status_fn(f"Rate limited. Waiting {30*attempt}s...")
                    time.sleep(30 * attempt)
                else:
                    human_delay(2, 5)
            except Exception:
                human_delay(2, 4)
    return None


def fetch_posts_v2(user_id, session, cookies, count=12, status_fn=None):
    url = f"https://www.instagram.com/api/v1/feed/user/{user_id}/?count={count}"
    headers = {**BASE_HEADERS, "x-csrftoken": cookies.get("csrftoken", "")}
    for attempt in range(1, 4):
        try:
            human_delay(1, 2.5)
            if status_fn: status_fn(f"Fetching posts via fallback feed (attempt {attempt}/3)...")
            if USE_CURL_CFFI:
                r = session.get(url, headers=headers, cookies=cookies,
                                timeout=20, impersonate="chrome120")
            else:
                r = session.get(url, headers=headers, cookies=cookies, timeout=20)
            if r.status_code == 200:
                return [parse_post_v2(i) for i in r.json().get("items", [])]
            elif r.status_code == 429:
                time.sleep(30 * attempt)
            else:
                human_delay(3, 5)
        except Exception:
            human_delay(2, 4)
    return []


def fetch_posts_v2_paginated(user_id, session, cookies, max_id, count=12):
    url = f"https://www.instagram.com/api/v1/feed/user/{user_id}/?count={count}&max_id={max_id}"
    headers = {**BASE_HEADERS, "x-csrftoken": cookies.get("csrftoken", "")}
    try:
        human_delay(2, 4)
        if USE_CURL_CFFI:
            r = session.get(url, headers=headers, cookies=cookies,
                            timeout=20, impersonate="chrome120")
        else:
            r = session.get(url, headers=headers, cookies=cookies, timeout=20)
        if r.status_code == 200:
            data = r.json()
            posts = [parse_post_v2(i) for i in data.get("items", [])]
            next_id = data.get("next_max_id") if data.get("more_available") else None
            return posts, next_id
    except Exception:
        pass
    return [], None


def fetch_more_posts_gql(user_id, end_cursor, session, cookies, count=12):
    QUERY_HASHES = [
        "e7e2f4da4b02303f74f0841279e52d76",
        "69cba40317214236af40e7efa9ca7448",
    ]
    variables = json.dumps({"id": user_id, "first": count, "after": end_cursor},
                           separators=(",", ":"))
    headers = {**BASE_HEADERS, "x-csrftoken": cookies.get("csrftoken", "")}
    for qhash in QUERY_HASHES:
        url = f"https://www.instagram.com/graphql/query/?query_hash={qhash}&variables={variables}"
        try:
            human_delay(2, 4)
            if USE_CURL_CFFI:
                r = session.get(url, headers=headers, cookies=cookies,
                                timeout=20, impersonate="chrome120")
            else:
                r = session.get(url, headers=headers, cookies=cookies, timeout=20)
            if r.status_code == 200:
                return r.json()
        except Exception:
            pass
    return None


def parse_post(node):
    likes    = node.get("edge_media_preview_like", {}).get("count", 0) or 0
    comments = node.get("edge_media_to_comment", {}).get("count", 0) or 0
    edges    = node.get("edge_media_to_caption", {}).get("edges", [])
    caption  = edges[0]["node"]["text"] if edges else ""
    mt_map   = {"GraphImage": "Image", "GraphVideo": "Video", "GraphSidecar": "Carousel"}
    ts       = node.get("taken_at_timestamp", 0)
    date_str = datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d") if ts else "N/A"
    return {
        "shortcode":   node.get("shortcode", ""),
        "url":         f"https://www.instagram.com/p/{node.get('shortcode', '')}/",
        "type":        mt_map.get(node.get("__typename", ""), "Image"),
        "date":        date_str,
        "likes":       likes,
        "comments":    comments,
        "video_views": node.get("video_view_count", 0) or 0,
        "caption":     caption[:150],
    }


def parse_post_v2(item):
    likes    = item.get("like_count", 0) or 0
    comments = item.get("comment_count", 0) or 0
    cap_obj  = item.get("caption")
    caption  = (cap_obj.get("text", "") if isinstance(cap_obj, dict) else "") or ""
    mt_map   = {1: "Image", 2: "Video", 8: "Carousel"}
    ts       = item.get("taken_at", 0)
    date_str = datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d") if ts else "N/A"
    code     = item.get("code", item.get("shortcode", ""))
    return {
        "shortcode":   code,
        "url":         f"https://www.instagram.com/p/{code}/",
        "type":        mt_map.get(item.get("media_type", 1), "Image"),
        "date":        date_str,
        "likes":       likes,
        "comments":    comments,
        "video_views": item.get("view_count", 0) or item.get("play_count", 0) or 0,
        "caption":     caption[:150],
    }


def run_scrape(username, num_posts, cookies, status_fn=None):
    session = make_session(cookies)

    raw = fetch_profile(username, session, cookies, status_fn)
    if not raw:
        return None, None, "Could not reach Instagram. Check your internet connection."
    if isinstance(raw, dict) and "_error" in raw:
        code = raw["_error"]
        if code == 401: return None, None, "401 — Cookies expired. Refresh from Chrome DevTools."
        if code == 403: return None, None, "403 — Rate limited or forbidden."
        if code == 404: return None, None, f"404 — @{username} not found. Check the username."
        return None, None, f"HTTP {code} error."

    user = raw.get("data", {}).get("user")
    if not user:
        return None, None, "User not found or account is private."

    profile = {
        "username":    username,
        "full_name":   user.get("full_name", "N/A"),
        "followers":   user.get("edge_followed_by", {}).get("count", 0),
        "following":   user.get("edge_follow", {}).get("count", 0),
        "total_posts": user.get("edge_owner_to_timeline_media", {}).get("count", 0),
        "bio":         user.get("biography", ""),
        "verified":    user.get("is_verified", False),
        "user_id":     user.get("id", ""),
        "category":    user.get("category_name", "N/A"),
        "is_business": user.get("is_business_account", False),
        "external_url":user.get("external_url", ""),
    }

    timeline   = user.get("edge_owner_to_timeline_media", {})
    edges      = timeline.get("edges", [])
    page_info  = timeline.get("page_info", {})
    has_next   = page_info.get("has_next_page", False)
    end_cursor = page_info.get("end_cursor", "")
    posts      = [parse_post(e["node"]) for e in edges]

    if status_fn: status_fn(f"Got {len(posts)} posts from profile. Collecting more...")

    if len(posts) == 0 and profile["user_id"]:
        if status_fn: status_fn("Business/verified account — using fallback feed endpoint...")
        posts = fetch_posts_v2(profile["user_id"], session, cookies,
                               count=min(num_posts, 12), status_fn=status_fn)
        if posts:
            next_max_id = posts[-1]["shortcode"]
            while len(posts) < num_posts:
                if status_fn: status_fn(f"Paginating... {len(posts)}/{num_posts} posts")
                new_posts, next_max_id = fetch_posts_v2_paginated(
                    profile["user_id"], session, cookies,
                    max_id=next_max_id, count=min(12, num_posts - len(posts))
                )
                if not new_posts: break
                posts.extend(new_posts)
                if not next_max_id: break
    elif len(posts) < num_posts:
        while len(posts) < num_posts and has_next and end_cursor:
            if status_fn: status_fn(f"GraphQL pagination... {len(posts)}/{num_posts} posts")
            page_data = fetch_more_posts_gql(
                profile["user_id"], end_cursor, session, cookies,
                count=min(12, num_posts - len(posts))
            )
            if not page_data: break
            media      = page_data.get("data", {}).get("user", {}).get("edge_owner_to_timeline_media", {})
            new_edges  = media.get("edges", [])
            page_info  = media.get("page_info", {})
            has_next   = page_info.get("has_next_page", False)
            end_cursor = page_info.get("end_cursor", "")
            posts.extend([parse_post(e["node"]) for e in new_edges])

    posts = posts[:num_posts]

    if not posts:
        return profile, [], None

    followers = profile["followers"]
    for p in posts:
        p["er_percent"] = round((p["likes"] + p["comments"]) / followers * 100, 4) if followers else 0

    return profile, posts, None


# ── EXCEL BUILDER ─────────────────────────────────────────────

def build_excel(profile, posts) -> bytes:
    wb   = openpyxl.Workbook()
    thin = Side(style="thin", color="222222")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    def fill(h): return PatternFill("solid", fgColor=h)
    def bf(sz=11, c="000000"): return Font(name="Calibri", bold=True,  size=sz, color=c)
    def nf(sz=11, c="000000"): return Font(name="Calibri", bold=False, size=sz, color=c)
    def ctr(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
    def lft(): return Alignment(horizontal="left",   vertical="center", wrap_text=True)

    followers      = profile["followers"]
    n              = len(posts)
    avg_er         = round(sum(p["er_percent"] for p in posts) / n, 4) if n else 0
    total_likes    = sum(p["likes"]       for p in posts)
    total_comments = sum(p["comments"]    for p in posts)
    total_views    = sum(p["video_views"] for p in posts)

    # ── Sheet 1: Post Metrics ──────────────────────────────
    ws = wb.active
    ws.title = "Post Metrics"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"

    ws.merge_cells("A1:K1")
    ws["A1"] = f"Instagram Metrics — @{profile['username']}"
    ws["A1"].font = bf(15, "FFFFFF"); ws["A1"].fill = fill("0D0D0D")
    ws["A1"].alignment = ctr(); ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:K2")
    ws["A2"] = (f"Followers: {followers:,}  |  Posts: {n}  |  "
                f"Avg ER%: {avg_er:.2f}%  |  ER = (Likes+Comments)/Followers×100  "
                f"[Saves & Shares: not available via any public API]")
    ws["A2"].font = nf(9, "CCCCCC"); ws["A2"].fill = fill("1A1A1A")
    ws["A2"].alignment = ctr(); ws.row_dimensions[2].height = 20

    headers = ["#", "Date", "Type", "Post URL", "Likes", "Comments",
               "Video Views", "Followers", "Saves", "Shares", "ER%"]
    widths  = [5,   13,     11,     50,          12,      12,
               13,          13,         12,       12,       10]
    for col, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font = bf(10, "FFFFFF"); c.fill = fill("E8405A")
        c.alignment = ctr(); c.border = bdr
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[3].height = 22

    type_colors = {"Image": "3B82F6", "Video": "8B5CF6", "Carousel": "F59E0B"}

    for i, p in enumerate(posts):
        row = i + 4
        bg  = "111111" if i % 2 == 0 else "161616"
        values = [i+1, p["date"], p["type"], p["url"],
                  p["likes"], p["comments"],
                  p["video_views"] or 0,
                  followers, "N/A — private", "N/A — private", None]
        for col, val in enumerate(values, 1):
            c = ws.cell(row=row, column=col)
            c.fill = fill(bg); c.border = bdr
            c.alignment = lft() if col == 4 else ctr()
            if col == 11:
                c.value = f"=(E{row}+F{row})/H{row}*100"
                c.number_format = '0.00"%"'; c.font = bf(10, "E8405A")
            elif col == 3:
                c.value = val; c.font = bf(10, type_colors.get(val, "AAAAAA"))
            elif col in (5, 6, 7, 8):
                c.value = val; c.number_format = "#,##0"; c.font = nf(10, "EEEEEE")
            elif col in (9, 10):
                c.value = val; c.font = nf(9, "555555")
            else:
                c.value = val; c.font = nf(10, "CCCCCC")
        ws.row_dimensions[row].height = 18

    tr = n + 4
    ws.merge_cells(f"A{tr}:D{tr}")
    ws[f"A{tr}"] = "TOTALS / AVERAGES"
    ws[f"A{tr}"].font = bf(10, "FFFFFF"); ws[f"A{tr}"].fill = fill("0D0D0D")
    ws[f"A{tr}"].alignment = ctr()
    for col, val in zip([5,6,7,8,9,10,11],
                        [total_likes, total_comments, total_views, followers, "", "", f"{avg_er:.2f}%"]):
        c = ws.cell(row=tr, column=col)
        c.fill = fill("0D0D0D"); c.border = bdr; c.alignment = ctr()
        c.font = bf(10, "F5C842") if col == 11 else bf(10, "FFFFFF")
        if col in (5,6,7,8) and isinstance(val, int):
            c.value = val; c.number_format = "#,##0"
        else:
            c.value = val
    ws.row_dimensions[tr].height = 22

    nr = tr + 1
    ws.merge_cells(f"A{nr}:K{nr}")
    ws[f"A{nr}"] = ("ℹ️  Instagram hides Saves and Shares from all public endpoints. "
                    "No scraper can return these.  ER% = (Likes + Comments) / Followers × 100.  "
                    f"Scraped: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    ws[f"A{nr}"].font = nf(9, "7C5500"); ws[f"A{nr}"].fill = fill("FFF8E1")
    ws[f"A{nr}"].alignment = lft(); ws.row_dimensions[nr].height = 28

    # ── Sheet 2: Profile Summary ───────────────────────────
    ws2 = wb.create_sheet("Profile Summary")
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions["A"].width = 26
    ws2.column_dimensions["B"].width = 44

    ws2.merge_cells("A1:B1")
    ws2["A1"] = f"Profile — @{profile['username']}"
    ws2["A1"].font = bf(13, "FFFFFF"); ws2["A1"].fill = fill("0D0D0D")
    ws2["A1"].alignment = ctr(); ws2.row_dimensions[1].height = 32

    rows2 = [
        ("Username",            f"@{profile['username']}"),
        ("Full Name",           profile["full_name"]),
        ("Category",            profile["category"]),
        ("Business Account",    "Yes" if profile["is_business"] else "No"),
        ("Verified",            "Yes" if profile["verified"] else "No"),
        ("Followers",           f"{profile['followers']:,}"),
        ("Following",           f"{profile['following']:,}"),
        ("Total Posts",         f"{profile['total_posts']:,}"),
        ("External URL",        profile["external_url"] or "—"),
        ("Bio",                 profile["bio"]),
        ("", ""),
        ("Posts Analysed",      n),
        ("Total Likes",         f"{total_likes:,}"),
        ("Total Comments",      f"{total_comments:,}"),
        ("Total Video Views",   f"{total_views:,}"),
        ("Avg Likes / Post",    f"{total_likes//n:,}" if n else "0"),
        ("Avg Comments / Post", f"{total_comments//n:,}" if n else "0"),
        ("Avg ER%",             f"{avg_er:.2f}%"),
        ("Saves Available",     "Not available — Instagram private"),
        ("Shares Available",    "Not available — Instagram private"),
        ("ER% Formula",         "(Likes + Comments) / Followers × 100"),
        ("Scrape Method",       "Instagram Internal API + Cookie Auth"),
        ("TLS Bypass",          "curl_cffi Chrome120" if USE_CURL_CFFI else "requests"),
        ("Scraped At",          datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    ]
    for i, (label, value) in enumerate(rows2):
        row = i + 2
        bg  = "111111" if i % 2 == 0 else "161616"
        if not label:
            ws2.row_dimensions[row].height = 8; continue
        lc = ws2.cell(row=row, column=1, value=label)
        lc.font = bf(10, "999999"); lc.fill = fill(bg)
        lc.alignment = lft(); lc.border = bdr
        vc = ws2.cell(row=row, column=2, value=value)
        vc.font = nf(10, "EEEEEE"); vc.fill = fill(bg)
        vc.alignment = lft(); vc.border = bdr
        ws2.row_dimensions[row].height = 20

    # ── Sheet 3: Captions ──────────────────────────────────
    ws3 = wb.create_sheet("Captions")
    ws3.sheet_view.showGridLines = False
    ws3.column_dimensions["A"].width = 5
    ws3.column_dimensions["B"].width = 14
    ws3.column_dimensions["C"].width = 55
    ws3.column_dimensions["D"].width = 12
    ws3.column_dimensions["E"].width = 12
    ws3.column_dimensions["F"].width = 12

    ws3.merge_cells("A1:F1")
    ws3["A1"] = f"Post Captions — @{profile['username']}"
    ws3["A1"].font = bf(13, "FFFFFF"); ws3["A1"].fill = fill("0D0D0D")
    ws3["A1"].alignment = ctr(); ws3.row_dimensions[1].height = 30

    for col, (h, _) in enumerate(zip(["#","Date","Caption","Likes","Comments","ER%"],
                                      [5, 14, 55, 12, 12, 12]), 1):
        c = ws3.cell(row=2, column=col, value=h)
        c.font = bf(10, "FFFFFF"); c.fill = fill("E8405A")
        c.alignment = ctr(); c.border = bdr

    for i, p in enumerate(posts):
        row = i + 3
        bg  = "111111" if i % 2 == 0 else "161616"
        for col, val in enumerate([i+1, p["date"], p["caption"] or "—",
                                    p["likes"], p["comments"], p["er_percent"]], 1):
            c = ws3.cell(row=row, column=col, value=val)
            c.fill = fill(bg); c.border = bdr
            c.alignment = lft() if col == 3 else ctr()
            if col == 6:
                c.number_format = '0.00"%"'; c.font = bf(10, "E8405A")
            elif col in (4,5):
                c.number_format = "#,##0"; c.font = nf(10, "EEEEEE")
            else:
                c.font = nf(10, "CCCCCC")
        ws3.row_dimensions[row].height = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── UTILS ─────────────────────────────────────────────────────

def fmt_num(n):
    if n >= 1_000_000: return f"{n/1_000_000:.1f}M"
    if n >= 1_000:     return f"{n/1_000:.1f}K"
    return str(n)


# ── STREAMLIT UI ──────────────────────────────────────────────

def main():
    # ── Session state init ──────────────────────────────────
    if "result_profile" not in st.session_state:
        st.session_state.result_profile = None
    if "result_posts" not in st.session_state:
        st.session_state.result_posts = None
    if "result_excel" not in st.session_state:
        st.session_state.result_excel = None
    if "result_fname" not in st.session_state:
        st.session_state.result_fname = None

    # ── Header ─────────────────────────────────────────────
    st.markdown("""
    <div style="padding: 36px 0 20px; text-align: center;">
      <div style="font-family:'Plus Jakarta Sans',sans-serif; font-size:11px; font-weight:700;
                  color:#E8405A; letter-spacing:0.22em; text-transform:uppercase; margin-bottom:10px;">
        Analytics Tool
      </div>
      <h1 style="font-family:'Plus Jakarta Sans',sans-serif; font-size:38px; font-weight:800;
                 color:#F0F0F0; margin:0; line-height:1.1; letter-spacing:-0.02em;">
        Instagram <span style="color:#E8405A;">Metrics</span>
      </h1>
      <p style="color:#555; font-size:13px; margin-top:10px; font-family:'Plus Jakarta Sans',sans-serif; font-weight:400;">
        Scrape post metrics for any public account &rarr; download as Excel
      </p>
    </div>
    """, unsafe_allow_html=True)

    # ── Cookies Sidebar ─────────────────────────────────────
    with st.sidebar:
        st.markdown("""
        <div style="font-family:'Plus Jakarta Sans',sans-serif; font-weight:800; font-size:17px;
                    color:#F0F0F0; margin-bottom:4px;">🍪 Session Cookies</div>
        <p style="color:#666; font-size:12px; margin-bottom:14px; font-family:'Plus Jakarta Sans',sans-serif;">
          Required to access Instagram's API.<br>
          Get from Chrome DevTools → Application → Cookies → instagram.com
        </p>
        """, unsafe_allow_html=True)

        sessionid  = st.text_input("sessionid",  value=DEFAULT_COOKIES.get("sessionid",""),  type="password")
        csrftoken  = st.text_input("csrftoken",  value=DEFAULT_COOKIES.get("csrftoken",""))
        ds_user_id = st.text_input("ds_user_id", value=DEFAULT_COOKIES.get("ds_user_id",""))
        mid        = st.text_input("mid",         value=DEFAULT_COOKIES.get("mid",""))

        st.markdown("""
        <div style="margin-top:14px; padding:12px; background:#141414; border-radius:10px;
                    font-size:11px; color:#666; line-height:1.7; font-family:'Plus Jakarta Sans',sans-serif;">
          <b style="color:#E8405A;">How to get cookies:</b><br>
          1. Open Instagram in Chrome<br>
          2. Press F12 → Application tab<br>
          3. Cookies → https://www.instagram.com<br>
          4. Copy the 4 values above
        </div>
        <div style="margin-top:10px; padding:12px; background:#141414; border-radius:10px;
                    font-size:11px; color:#555; line-height:1.7; font-family:'Plus Jakarta Sans',sans-serif;">
          <b style="color:#666;">Note:</b> Instagram hides Saves &amp; Shares from all public endpoints.
          ER% = (Likes + Comments) / Followers × 100.
        </div>
        """, unsafe_allow_html=True)

    cookies = {
        "sessionid":  sessionid,
        "csrftoken":  csrftoken,
        "ds_user_id": ds_user_id,
        "mid":        mid,
    }

    # ── Main Input ──────────────────────────────────────────
    col1, col2 = st.columns([3, 1])
    with col1:
        raw_input = st.text_input(
            "Instagram URL or Username",
            placeholder="username  or  https://www.instagram.com/username/",
        )
    with col2:
        num_posts = st.number_input("Posts", min_value=1, max_value=50, value=15, step=5)

    go = st.button("▶  Fetch Metrics", use_container_width=True)

    # ── Run Scrape ──────────────────────────────────────────
    if go:
        if not raw_input.strip():
            st.error("Please enter an Instagram username or URL.")
            return
        if not all([sessionid, csrftoken, ds_user_id, mid]):
            st.error("⚠️ Fill in all 4 cookies in the sidebar first.")
            return

        # Clear previous results so we don't flash stale data
        st.session_state.result_profile = None
        st.session_state.result_posts   = None
        st.session_state.result_excel   = None

        username    = extract_username(raw_input)
        status_box  = st.empty()
        progress    = st.progress(0)

        def update_status(msg):
            status_box.info(f"⏳ {msg}")

        update_status("Connecting to Instagram...")
        progress.progress(10)

        profile, posts, err = run_scrape(
            username, int(num_posts), cookies,
            status_fn=update_status
        )

        if err:
            status_box.empty(); progress.empty()
            st.error(f"❌ {err}")
            return

        progress.progress(80)
        update_status("Building Excel file...")

        if not posts:
            status_box.empty(); progress.empty()
            st.warning("⚠️ Profile found but no posts could be retrieved. "
                       "Cookies may be expired or the account is private.")
            return

        excel_bytes = build_excel(profile, posts)
        fname       = f"ig_metrics_{username}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        progress.progress(100)
        status_box.empty()
        progress.empty()

        # Store in session state so download button doesn't cause re-scrape
        st.session_state.result_profile = profile
        st.session_state.result_posts   = posts
        st.session_state.result_excel   = excel_bytes
        st.session_state.result_fname   = fname

    # ── Render Results (from session state) ─────────────────
    profile = st.session_state.result_profile
    posts   = st.session_state.result_posts

    if not profile or not posts:
        return

    # Success banner
    st.markdown(f"""
    <div style="background:#0d1f16; border:1px solid #1a4d30; border-radius:12px;
                padding:18px 22px; margin:14px 0;">
      <div style="font-family:'Plus Jakarta Sans',sans-serif; font-size:17px; font-weight:800;
                  color:#1DB954; letter-spacing:-0.01em;">✅ Done — @{profile['username']}</div>
      <div style="font-family:'Plus Jakarta Sans',sans-serif; font-size:13px; color:#3d7a55; margin-top:3px;">
        {len(posts)} posts fetched successfully
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Profile card
    badges = ""
    if profile["verified"]:    badges += '<span class="badge badge-verified">✓ Verified</span>'
    if profile["is_business"]: badges += '<span class="badge badge-business">Business</span>'
    if profile["category"] and profile["category"] != "N/A":
        badges += f'<span class="badge badge-category">{profile["category"]}</span>'

    st.markdown(f"""
    <div class="profile-card">
      <div class="profile-name">{profile['full_name']}</div>
      <div class="profile-username">@{profile['username']}</div>
      <div>{badges}</div>
      <div class="profile-bio">{profile['bio'] or ''}</div>
    </div>
    """, unsafe_allow_html=True)

    # Metric cards
    followers = profile["followers"]
    n         = len(posts)
    avg_er    = round(sum(p["er_percent"] for p in posts) / n, 4) if n else 0
    total_l   = sum(p["likes"]       for p in posts)
    total_c   = sum(p["comments"]    for p in posts)
    total_v   = sum(p["video_views"] for p in posts)

    st.markdown(f"""
    <div class="metric-grid">
      <div class="metric-card metric-followers">
        <div class="metric-value">{fmt_num(followers)}</div>
        <div class="metric-label">Followers</div>
      </div>
      <div class="metric-card">
        <div class="metric-value">{fmt_num(total_l)}</div>
        <div class="metric-label">Total Likes</div>
      </div>
      <div class="metric-card">
        <div class="metric-value">{fmt_num(total_c)}</div>
        <div class="metric-label">Total Comments</div>
      </div>
      <div class="metric-card">
        <div class="metric-value">{fmt_num(total_l // n if n else 0)}</div>
        <div class="metric-label">Avg Likes / Post</div>
      </div>
      <div class="metric-card">
        <div class="metric-value">{fmt_num(total_v)}</div>
        <div class="metric-label">Video Views</div>
      </div>
      <div class="metric-card metric-er">
        <div class="metric-value">{avg_er:.2f}%</div>
        <div class="metric-label">Avg ER%</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Post Table — FIX: build entire HTML block in one shot ──
    st.markdown('<div class="section-label">Recent Posts</div>', unsafe_allow_html=True)

    type_cls = {"Image": "type-image", "Video": "type-video", "Carousel": "type-carousel"}

    rows_html = ""
    for i, p in enumerate(posts):
        tc        = type_cls.get(p["type"], "")
        shortcode = p.get("shortcode", "")
        link_text = shortcode[:14] + "↗" if shortcode else "—"
        rows_html += (
            f'<div class="post-row">'
            f'  <div class="post-cell-idx">{i+1}</div>'
            f'  <div class="post-cell">{p["date"]}</div>'
            f'  <div class="post-cell {tc}">{p["type"]}</div>'
            f'  <div class="post-cell"><a class="post-link" href="{p["url"]}" target="_blank">{link_text}</a></div>'
            f'  <div class="post-cell-num">{p["likes"]:,}</div>'
            f'  <div class="post-cell-num">{p["comments"]:,}</div>'
            f'  <div class="post-cell-er">{p["er_percent"]:.2f}%</div>'
            f'</div>'
        )

    st.markdown(
        f'<div class="post-table">'
        f'  <div class="post-row post-row-header">'
        f'    <span>#</span><span>Date</span><span>Type</span><span>URL</span>'
        f'    <span>Likes</span><span>Comments</span><span>ER%</span>'
        f'  </div>'
        f'  {rows_html}'
        f'</div>',
        unsafe_allow_html=True
    )

    # ── Download — reads from session_state, no re-scrape ──
    st.markdown('<div class="section-label">Export</div>', unsafe_allow_html=True)
    st.download_button(
        label="⬇  Download Excel Report",
        data=st.session_state.result_excel,
        file_name=st.session_state.result_fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.markdown("""
    <div style="text-align:center; font-family:'Plus Jakarta Sans',sans-serif;
                font-size:11px; color:#383838; margin-top:6px;">
      Includes 3 sheets: Post Metrics · Profile Summary · Captions
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()