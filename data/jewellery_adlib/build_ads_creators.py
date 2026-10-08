# -*- coding: utf-8 -*-
"""Jewellery brands - Ad Library creatives + creators, two tabs per brand.

Tab 1  <Brand> - Ads       one row per distinct creative: hook, problem, solution,
                           offer, tone/format, campaign, partner creator, links, transcript.
Tab 2  <Brand> - Creators  top 10 creators; Ad Library partnership ads; Ad Library
                           branded content; every creator collab reel on the brand's own
                           page since 1 Apr 2026 with metrics.
Source for ads: the client's Ad Library Google Sheet (dumped to sheet_dump.json).
Source for our creators: brand_scan_cache.json (own page: Posts grid + Reels tab).
No spend estimates, no winner scores, no tiers.
"""
import json, re, sys, collections, importlib.util
from datetime import datetime, timezone
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = "C:/Users/omkar/OneDrive/Desktop/InstagramAnalytics"
sys.path.insert(0, ROOT)
from core.kolkata_engine import clean_cell
from core.brand_collab_scan import account_type
spec = importlib.util.spec_from_file_location("ad_themes", ROOT + "/data/jewellery_adlib/ad_themes.py")
TH = importlib.util.module_from_spec(spec); spec.loader.exec_module(TH)

OUT = ROOT + "/deliverables/Jewellery_Ads_and_Creators_Oct2026.xlsx"
D = json.load(open(ROOT + "/data/jewellery_adlib/sheet_dump.json", encoding="utf-8"))
CACHE = json.load(open(ROOT + "/brand_scan_cache.json", encoding="utf-8"))
APR_KEY = "2026-04-01_2026-10-07_own"


def _load(path):
    try:
        return json.load(open(path, encoding="utf-8"))
    except Exception:
        return {}


TRANSCRIPTS = {k: v for k, v in _load(ROOT + "/jewellery_audio_transcripts.json").items()
               if v.get("model") == "large-v3-turbo"}
import glob as _glob
for _f in sorted(_glob.glob(ROOT + "/jewellery_audio_transcripts_groq*.json")):   # Groq wins where both exist
    TRANSCRIPTS.update(_load(_f))
LANG = {"en": "English", "hi": "Hindi", "ta": "Tamil", "ml": "Malayalam", "bn": "Bengali", "te": "Telugu",
        "kn": "Kannada", "mr": "Marathi", "gu": "Gujarati", "ur": "Urdu", "pa": "Punjabi"}


REEL_AN = {}
for _f in sorted(_glob.glob(ROOT + "/data/jewellery_adlib/reel_analysis/*.json")):
    REEL_AN.update(_load(_f))


def transcript_of(url):
    code = ([x for x in str(url).split("/") if x] or [""])[-1]
    t = TRANSCRIPTS.get(code)
    if not t:
        return "not transcribed yet", ""
    txt = t.get("text") or ""
    lang = t.get("lang") or ""
    lang = LANG.get(lang, lang.title() if lang else "")
    if txt.startswith("no speech"):
        return "Music only - no voiceover", ""
    if junk(txt):                                       # lyric loops: 'Ba na na na...'
        return "Music / lyrics only - no usable voiceover", ""
    if lang and lang not in LANG.values():            # e.g. 'Chinese' on a music clip = misdetection
        return "Unclear - likely music (language misdetected as " + lang + ")", ""
    return txt, lang

BRANDS = [  # display, ads tab, transcript tab, creators tab, IG handle (None = not scanned)
 ("Tanishq", "Tanishq", "Tanishq transcript", "Tanishq Creators", "tanishqjewellery"),
 ("Malabar Gold", "Malabar Gold", "Malabar Gold Transcript", "Malabar Creators", "malabargoldanddiamonds"),
 ("Joyalukkas", "Joyalukkas", "Joyalukkas Transcript", None, "joyalukkas"),
 ("Senco", "Senco Gold & Diamonds", "Senco Gold & Diamonds Transcript", "Senco Creators", "sencogoldanddiamonds"),
 ("Jos Alukkas", "Jos Alukkas", "Jos Alukkas Transcript", None, "josalukkas"),
 ("BlueStone", "Bluestone", "Bluestone Transcript", "Bluestone Creators", "bluestone_jewellery"),
 ("Kalyan Jewellers", "Kalyan Jewellers", "Kalyan Transcript", None, "kalyanjewellers_official"),
 ("Thangamayil", "Thangamayil Jewellery Ltd", "Thangamayil Jewellery Ltd Transcript", None, "thangamayiljewellery"),
 ("PC Jeweller", "PC Jeweller", "PC Jeweller Transcript", None, "pc_jeweller"),
 ("CaratLane", "Carat Lane", "Carat Lane Transcript", None, "caratlane"),
 ("Jewellery Khazana (Bengali)", "Jewellery khazana ( Bengali )", "Jewellery khazana ( Bengali ) Transcript", None, None),
]
THEME_KEY = {"Malabar Gold": "Malabar Gold", "Senco": "Senco", "Thangamayil": "Thangamayil"}
SHORT = {"Jewellery Khazana (Bengali)": "Khazana Bengali", "Kalyan Jewellers": "Kalyan"}
BRAND_FIX = [(r"\bTun(i|ee)sh[kq]?\b|\bTanishk\b", "Tanishq"), (r"\bJ(w|oo|yo)l(l)?ers\b|\bJyolars\b", "Jewellers"),
             (r"\bSenko\b", "Senco"), (r"\bThangamayal\b", "Thangamayil"), (r"\bMara Bargordan\b", "Malabar Gold"),
             (r"\bMalabar Gold and Diamonds\b", "Malabar Gold & Diamonds")]

norm = lambda s: re.sub(r"\W+", " ", (s or "").lower()).strip()[:400]
H1 = Font(bold=True, size=13); B = Font(bold=True); W = Font(bold=True, color="FFFFFF")
HDR = PatternFill("solid", fgColor="1F3864"); SEC = PatternFill("solid", fgColor="D9E1F2")
WRAP = Alignment(wrap_text=True, vertical="top")
LINK = Font(color="0563C1", underline="single")


def fix_brand(t):
    for p, r in BRAND_FIX:
        t = re.sub(p, r, t, flags=re.I)
    return t


def junk(t):
    """Whisper-base on music: repetition, or only 'Music / Thank you / Subscribe'."""
    w = re.findall(r"\w+", (t or "").lower())
    if not w:
        return True
    if len(w) <= 6 and set(w) <= {"music", "thank", "you", "thanks", "for", "watching", "subscribe", "to", "the", "channel", "our"}:
        return True
    if len(w) >= 12:
        c = collections.Counter(w)
        if c.most_common(1)[0][1] / len(w) > 0.22 or len(c) / len(w) < 0.3:
            return True
    return False


def first_sentence(t, n=170):
    t = re.sub(r"\s+", " ", t or "").strip()
    m = re.match(r"(.{12,}?[.!?])(\s|$)", t)
    s = m.group(1) if m else t
    return (s[:n] + "...") if len(s) > n else s


def tab_rows(name):
    return D.get(name.strip(), []) if name else []


def extract_offer(copy, offers_col, said=""):
    """Plain-language offer pulled from the ad's own copy; overlapping matches collapsed."""
    text = re.sub(r"\s+", " ", (copy or "") + " " + (said or ""))
    pats = [r"flat \d+% (?:off|discount)(?: on [a-z ]{3,30})?", r"up ?to \d+% off(?: on [a-z ]{3,40})?",
            r"\d+% off(?: on [a-z ]{3,40})?", r"\d+% deduction\*?(?: on [a-z ]{3,25})?",
            r"₹\s?[\d,]+(?:\s?[–-]\s?₹?[\d,]+)?\s?off(?: per gram)?", r"(?:just |only )?\d+% advance",
            r"cash[^.]{0,30}15 minutes", r"free [a-z]{3,15}(?: [a-z]{3,15})?(?: week)?", r"under ₹\s?[\d,]+k?",
            r"lower (?:gold )?rate[^.]{0,30}", r"\d+-hours?[^.]{0,15}", r"lock[^.]{0,25}rate"]
    found = []
    for pat in pats:
        for m in re.finditer(pat, text, re.I):
            sx = re.sub(r"[*\s]+$", "", m.group(0).strip(" .,"))
            if not any(sx.lower() in f.lower() or f.lower() in sx.lower() for f in found):
                found.append(sx)
    if not found and re.search(r"(\d+)\s?%", text):          # regional-language copy (Malayalam, Tamil...)
        m = re.search(r"(\d+)\s?%", text)
        tag = "advance" if re.search(r"advance|അഡ്വാൻ|அட்வான்ஸ்", text, re.I) else "off"
        found.append(f"{m.group(1)}% {tag} (regional-language copy)")
    if not found and offers_col:
        found.append(offers_col)
    return "; ".join(found[:3]) or "No offer stated in this ad"


CTA = {"LEARN_MORE": "Learn more", "SHOP_NOW": "Shop now", "BOOK_TRAVEL": "Book now", "CONTACT_US": "Contact us",
       "GET_OFFER": "Get offer", "SEND_WHATSAPP_MESSAGE": "WhatsApp us", "WHATSAPP_MESSAGE": "WhatsApp us",
       "SIGN_UP": "Sign up", "MESSAGE_PAGE": "Message us", "GET_DIRECTIONS": "Get directions", "CALL_NOW": "Call now",
       "APPLY_NOW": "Apply now", "ORDER_NOW": "Order now", "SEE_MORE": "See more", "WATCH_MORE": "Watch more",
       "BOOK_NOW": "Book now", "SUBSCRIBE": "Subscribe", "DOWNLOAD": "Download", "NO_BUTTON": "No button"}


def creatives(at, tt):
    A = tab_rows(at); T = tab_rows(tt)
    if not A:
        return []
    ix = {h: i for i, h in enumerate(A[0])}
    tix = {h: i for i, h in enumerate(T[0])} if T else {}
    g = lambda r, k: (r[ix[k]] if k in ix and len(r) > ix[k] else "") or ""
    tr = {}
    for r in T[1:]:
        if tix and len(r) > tix.get("Full transcript", 999):
            tr[r[tix["Library ID"]]] = r
    tg = lambda r, k: (r[tix[k]] if r and k in tix and len(r) > tix[k] else "") or ""
    out = collections.OrderedDict()
    for r in A[1:]:
        if not g(r, "Library ID"):
            continue
        trow = tr.get(g(r, "Library ID"))
        t = tg(trow, "Full transcript") or g(r, "Transcript")
        copy, head = g(r, "Ad copy"), g(r, "Headline")
        cat = "{{" in (copy + head) and not t
        k = ("CAT:" + norm(g(r, "Offers") + g(r, "Landing group"))) if cat else (norm(t) if t else norm(head + " " + copy))
        e = out.setdefault(k, {"catalogue": cat, "ads": [], "t": t, "copy": copy, "head": head,
                               "format": g(r, "Format"), "lang": tg(trow, "Voiceover language") or g(r, "Language"),
                               "offers": g(r, "Offers"), "said": tg(trow, "Offers said") or g(r, "Offers said"),
                               "cta": g(r, "CTA"), "landing": g(r, "Landing URL"), "len": g(r, "Video length (s)")})
        e["ads"].append({"id": g(r, "Library ID"), "link": g(r, "Library link"), "status": g(r, "Status"),
                         "started": g(r, "Started"), "last": g(r, "Last seen / ended"), "days": g(r, "Days running"),
                         "creator": g(r, "Creator (partnership)"), "creator_ig": g(r, "Creator IG handle"),
                         "platforms": g(r, "Platforms")})
    return list(out.values())


def theme_for(brand, x):
    """Match on what the ad leads with (headline + opening of copy/transcript) first, so a
    passing mention lower in the copy (e.g. '0% deduction' in an earring ad) cannot decide it."""
    if x["catalogue"]:
        return ("Catalogue / dynamic product ad",) + TH.CATALOGUE[1:]
    lead = " ".join([x["head"], (x["copy"] or "")[:180], (x["t"] or "")[:160]])
    full = " ".join([x["head"], x["copy"], x["t"] or ""])
    for text in (lead, full):
        for th in TH.T.get(brand, []):
            if th[1] != "." and re.search(th[1], text, re.I):
                return (th[0], th[2], th[3], th[4])
    for th in TH.T.get(brand, []):
        if th[1] == ".":
            return (th[0], th[2], th[3], th[4])
    return ("Product / brand showcase", "Shoppers browsing for jewellery.", "Showcases designs; no specific claim.",
            "Product creative.")


HALLUCINATION = re.compile(r"click the link|description box|subscribe|thanks for watching|in the name of god|"
                           r"like and share|see you in the next|next episode", re.I)
GENERIC_HEAD = {"tanishq", "pc jeweller", "jos alukkas", "malabar gold and diamonds", "thangamayil jewellery ltd",
                "jewellery khazana", "kalyan jewellers", "caratlane", "bluestone"}


def hook_of(x, brand):
    """Text hook first (headline / copy opener - exact), spoken line second (auto-transcribed,
    shown only when it reads as real speech)."""
    parts = []
    head = x["head"] if x["head"] and "{{" not in x["head"] and not re.match(r"^(www\.|api\.|instagram\.com)", x["head"]) else ""
    if head and norm(head) not in {norm(g) for g in GENERIC_HEAD} | {norm(brand)}:
        parts.append('Headline: "' + head[:150] + '"')
    elif x["copy"] and "{{" not in x["copy"]:
        parts.append('Copy opens: "' + first_sentence(x["copy"], 150) + '"')
    t = x["t"] or ""
    if t and not junk(t):
        line = fix_brand(first_sentence(t))
        if HALLUCINATION.search(line):
            parts.append("Spoken: unclear in auto-transcript")
        else:
            parts.append('Spoken (auto-transcribed): "' + line + '"')
    elif t:
        parts.append("Spoken: music only, no voiceover")
    return "\n".join(parts) or "No text - visual only"


# ---------------- creators -------------------------------------------------------
def sheet_creator_sections(ct):
    """Client Creators tabs vary in titles and order, so sections are recognised by their
    header columns: 'Library IDs' = partnership ads; 'Date'+'Post link' = branded content
    per post; 'Paid-partnership posts' = branded content summarised per creator."""
    out = {"partnership": [], "bc_posts": [], "bc_creators": []}
    cur, hdr = None, None
    for r in tab_rows(ct):
        cells = [c for c in r if c]
        if not cells:
            continue
        if r[0] == "Brand" and len(cells) > 3:
            hdr = r
            cur = ("partnership" if "Library IDs" in r else "bc_posts" if "Post link" in r
                   else "bc_creators" if "Paid-partnership posts" in r else None)
            continue
        if len(cells) == 1:          # a section title - the next header row decides the type
            cur, hdr = None, None; continue
        if cur and hdr:
            out[cur].append(dict(zip(hdr, r + [""] * (len(hdr) - len(r)))))
    return out


def our_reels(handle):
    res = (CACHE.get(APR_KEY) or {}).get(handle) if handle else None
    if not res:
        return None, None
    reels = [p for p in res.get("posts", []) if p.get("kind") == "Reel"]
    reels.sort(key=lambda p: -(p.get("taken_at") or 0))
    return reels, res


# ---------------- workbook -------------------------------------------------------
wb = openpyxl.Workbook()
ov = wb.active; ov.title = "Overview"
overview = []


def header(ws, row, cols, widths=None):
    for c, h in enumerate(cols, 1):
        x = ws.cell(row, c, h); x.font = W; x.fill = HDR; x.alignment = Alignment(wrap_text=True, vertical="center")
    if widths:
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w


def section(ws, row, text, ncol):
    x = ws.cell(row, 1, text); x.font = Font(bold=True, size=12)
    for c in range(1, ncol + 1):
        ws.cell(row, c).fill = SEC
    return row + 1


def put(ws, row, vals, link_cols=()):
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row, c, clean_cell(v.replace("�", "") if isinstance(v, str) else v))
        cell.alignment = WRAP
        if c in link_cols and isinstance(v, str) and v.startswith("http"):
            cell.hyperlink = v; cell.font = LINK


for disp, at, tt, ct, handle in BRANDS:
    short = SHORT.get(disp, disp)
    G = creatives(at, tt)
    themes_used = collections.Counter()
    # ---------------- Ads tab
    ws = wb.create_sheet(f"{short} - Ads"[:31])
    nads = sum(len(x["ads"]) for x in G); live = sum(a["status"] == "Active" for x in G for a in x["ads"])
    COLS = ["#", "Status", "Started", "Days Running", "Ads Using It", "Format", "Language", "Campaign",
            "Hook (opening line)", "Problem It Targets", "Solution / Claim", "Offer", "CTA", "Tone & Format",
            "Partner Creator", "Headline", "Ad Copy", "Landing Page", "Ad Library Link", "Full Transcript (auto)"]
    WID = [5, 9, 11, 9, 8, 11, 10, 28, 46, 40, 46, 34, 11, 36, 20, 30, 60, 34, 34, 70]
    rows = []
    for x in G:
        th = theme_for(disp if disp in TH.T else THEME_KEY.get(disp, disp), x)
        themes_used[th[0]] += len(x["ads"])
        ads = sorted(x["ads"], key=lambda a: (a["status"] != "Active", a["started"]), reverse=False)
        a0 = ads[0]
        days = max((int(a["days"]) for a in ads if str(a["days"]).isdigit()), default="")
        started = min((a["started"] for a in ads if a["started"]), default="")
        creators = sorted({a["creator"] for a in ads if a["creator"]})
        status = "Live" if any(a["status"] == "Active" for a in ads) else "Ended"
        t = x["t"] or ""
        rows.append([status, started, days, len(ads), x["format"], x["lang"], th[0], hook_of(x, disp), th[1], th[2],
                     extract_offer(x["copy"], x["offers"], x["said"]), CTA.get(x["cta"], x["cta"].replace("_", " ").title() if x["cta"] else "No button"),
                     th[3], ", ".join(creators) or "Brand ad", x["head"] if "{{" not in x["head"] else "catalogue - filled per product",
                     x["copy"] if "{{" not in x["copy"] else "catalogue - filled per product", x["landing"] or "No link",
                     a0["link"], fix_brand(t) if t and not junk(t) else ("Music / lyrics only" if t else "No transcript - not a video or no speech")])
    rows.sort(key=lambda r: (r[0] != "Live", -(r[3] or 0), str(r[1])), reverse=False)
    ws.cell(1, 1, f"{disp} - every Meta Ad Library creative ({len(G)} distinct creatives across {nads} ads, {live} live). "
                  f"One row per creative; 'Ads Using It' counts the ads that run it. Hook is the ad's own opening line. "
                  f"Problem / Solution / Tone describe the campaign it belongs to. Transcripts are auto-generated (Whisper) "
                  f"and can mis-hear words.").font = B
    ws.row_dimensions[1].height = 30
    ws.cell(2, 1, "Campaigns: " + " | ".join(f"{k} ({v})" for k, v in themes_used.most_common()))
    header(ws, 4, COLS, WID)
    for i, r in enumerate(rows, 1):
        put(ws, 4 + i, [i] + r, link_cols=(18, 19))
    ws.freeze_panes = "C5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(COLS))}{4 + len(rows)}"

    # ---------------- Ads tab, part 2: creator partnership reels, transcribed and analysed
    r_reels, _ = our_reels(handle)
    row = 4 + len(rows) + 3
    row = section(ws, row, "CREATOR PARTNERSHIP REELS ON THE BRAND PAGE SINCE 1 APR 2026 - transcribed and analysed "
                           "(same breakdown as the Ad Library creatives above)", len(COLS))
    if not r_reels:
        put(ws, row, ["", "Not scanned (excluded on request)." if not handle else "No collab reels found."])
    else:
        ws.cell(row, 1, "Hook is the reel's spoken opening line (English rendering for regional-language reels); "
                        "where a reel has no voiceover, its caption line is used and marked 'Caption'. "
                        "Boosted is an estimate from views vs likes and followers.").font = B
        row += 1
        header(ws, row, ["#", "Post Date", "Creator", "Account Type", "Followers (Exact)", "Views", "Boosted? (est.)",
                         "Paid Label", "Language", "Hook (opening line)", "Problem It Targets", "Solution / Claim",
                         "Offer & CTA", "Tone & Format", "Reel Link", "Full Transcript (auto)"]); row += 1
        for i, p in enumerate(r_reels, 1):
            an = REEL_AN.get(p["code"]) or {}
            txt, lang = transcript_of(p["post_url"])
            at_ = account_type(handle, p["creator"], p.get("creator_name") or "", p.get("creator_category") or "")
            put(ws, row, [i, datetime.fromtimestamp(p["taken_at"], tz=timezone.utc).strftime("%Y-%m-%d"), "@" + p["creator"],
                          "Creator" if at_ == "Creator / individual" else ("Brand's own sub-account" if at_ == "Own sub-brand" else "Brand / company"),
                          p.get("followers") if p.get("followers") is not None else "not resolved",
                          p.get("views") if p.get("views") is not None else "not shown",
                          "Boosted" if p.get("is_boosted") else "Not boosted",
                          "ON" if p.get("is_paid_partnership") else "OFF", lang or "-",
                          an.get("hook") or ("Caption: '" + first_sentence(p.get("caption") or "") + "'"),
                          an.get("problem", ""), an.get("solution", ""), an.get("offer", ""), an.get("tone", ""),
                          p["post_url"], txt], link_cols=(15,)); row += 1

    # ---------------- Creators tab
    wc = wb.create_sheet(f"{short} - Creators"[:31])
    NC = 12
    for c, w in enumerate([5, 24, 22, 14, 12, 12, 11, 11, 14, 16, 40, 40], 1):
        wc.column_dimensions[get_column_letter(c)].width = w
    wc.cell(1, 1, f"{disp} - creators: top 10, Meta Ad Library partnership ads and branded content, "
                  f"and every creator collab reel on @{handle or 'n/a'}'s own page since 1 Apr 2026").font = B
    row = 13          # rows 3-11 hold the summary block, written once the sections are counted
    reels, res = our_reels(handle)
    creators_reels = [p for p in (reels or []) if account_type(handle, p["creator"], p.get("creator_name") or "", p.get("creator_category") or "") == "Creator / individual"] if reels else []

    # 1. top 10
    row = section(wc, row, "1. TOP 10 CREATORS - collab reels on the brand's own page since 1 Apr 2026, ranked by best reel views", NC)
    if handle is None:
        put(wc, row, ["", "Not scanned - Jewellery Khazana (Bengali) was excluded from the Instagram scan on request."]); row += 2
    elif reels is None:
        put(wc, row, ["", "1 Apr scan still running for this brand - rerun the build when it finishes."]); row += 2
    else:
        header(wc, row, ["#", "Creator", "Name", "Followers (Exact)", "Collab Reels", "Best Views", "Total Views",
                         "Total Likes", "Latest Collab", "Boosted? (est.)", "Best Reel", "Profile"]); row += 1
        agg = collections.OrderedDict()
        for p in creators_reels:
            a = agg.setdefault(p["creator"], {"name": p.get("creator_name") or "", "f": p.get("followers"), "n": 0,
                                               "best": 0, "bestu": "", "views": 0, "likes": 0, "last": 0, "boost": 0})
            v = p.get("views") or 0
            a["n"] += 1; a["views"] += v; a["likes"] += p.get("likes") or 0; a["last"] = max(a["last"], p.get("taken_at") or 0)
            a["boost"] += 1 if p.get("is_boosted") else 0
            if v >= a["best"]:
                a["best"], a["bestu"] = v, p["post_url"]
        top = sorted(agg.items(), key=lambda kv: -kv[1]["best"])[:10]
        if not top:
            put(wc, row, ["", "No creator collab reels on the brand's page since 1 Apr 2026."]); row += 1
        for i, (h, a) in enumerate(top, 1):
            put(wc, row, [i, "@" + h, a["name"], a["f"] if a["f"] is not None else "not resolved", a["n"], a["best"], a["views"],
                          a["likes"], datetime.fromtimestamp(a["last"], tz=timezone.utc).strftime("%Y-%m-%d") if a["last"] else "",
                          f"{a['boost']} of {a['n']} boosted" if a["boost"] else "Not boosted", a["bestu"],
                          f"https://www.instagram.com/{h}/"], link_cols=(11, 12)); row += 1
        row += 1

    # 2. Ad Library partnership ads (from the ads tab + the sheet's live-creators section)
    row = section(wc, row, "2. META AD LIBRARY - PARTNERSHIP ADS (creator runs as partner in the brand's paid ads)", NC)
    header(wc, row, ["#", "Creator", "IG Handle", "Ads", "Live Ads", "First Seen", "Latest Seen", "FB Page Likes",
                     "Status", "Campaign", "Ad Library Link", "Creator Profile"]); row += 1
    part = collections.OrderedDict()
    for x in G:
        th = theme_for(disp if disp in TH.T else THEME_KEY.get(disp, disp), x)[0]
        for a in x["ads"]:
            if a["creator"]:
                p = part.setdefault(a["creator"], {"ig": a["creator_ig"], "ads": [], "camp": th, "likes": "", "profile": ""})
                p["ads"].append(a)
    secs = sheet_creator_sections(ct)
    if True:
            for r in secs["partnership"]:
                nm = r.get("Creator") or ""
                p = part.setdefault(nm, {"ig": r.get("IG handle") or "", "ads": [], "camp": "", "likes": "", "profile": ""})
                p["likes"] = r.get("Facebook page likes") or p["likes"]; p["profile"] = r.get("Profile link") or p["profile"]
                for lid in (r.get("Library IDs") or "").split():
                    if not any(a["id"] == lid for a in p["ads"]):
                        p["ads"].append({"id": lid, "link": f"https://www.facebook.com/ads/library/?id={lid}",
                                         "status": "Active", "started": r.get("First collab seen") or "",
                                         "last": r.get("Latest collab seen") or ""})
    if not part:
        put(wc, row, ["", "No creator partnership ads in this brand's Ad Library data."]); row += 1
    for i, (nm, p) in enumerate(sorted(part.items(), key=lambda kv: -len(kv[1]["ads"])), 1):
        ads = p["ads"]; lv = sum(a.get("status") == "Active" for a in ads)
        put(wc, row, [i, nm, ("@" + p["ig"]) if p["ig"] else "not given", len(ads), lv,
                      min((a.get("started") or "" for a in ads), default=""), max((a.get("last") or "" for a in ads), default=""),
                      p["likes"] or "not given", "Live" if lv else "Ended", p["camp"] or "Creator partnership",
                      ads[0]["link"] if ads else "", p["profile"] or (f"https://www.instagram.com/{p['ig']}/" if p["ig"] else "")],
            link_cols=(11, 12)); row += 1
    row += 1

    # 3. Ad Library branded content
    row = section(wc, row, "3. META AD LIBRARY - BRANDED CONTENT (creator posts tagged with the brand as paid partner)", NC)
    header(wc, row, ["#", "Creator", "Handle", "Latest Post", "First Post", "Paid-Partner Posts", "Reels", "Posts",
                     "Stories", "Type / Platform", "Post Link(s)", "Profile"]); row += 1
    bc = []
    for r in secs["bc_creators"]:
        links = (r.get("Post links") or "").split()
        bc.append([r.get("Creator"), "@" + (r.get("Handle") or ""), r.get("Latest collab"), r.get("First collab"),
                   r.get("Paid-partnership posts"), r.get("Reels"), r.get("Posts"), r.get("Stories"), r.get("Platforms"),
                   "\n".join(links[:6]) + (f"\n(+{len(links)-6} more)" if len(links) > 6 else ""), r.get("Profile")])
    for r in secs["bc_posts"]:
        bc.append([r.get("Creator"), "@" + (r.get("Handle") or ""), r.get("Date"), r.get("Date"), 1,
                   1 if r.get("Type") == "Reel" else 0, 1 if r.get("Type") == "Post" else 0, 1 if r.get("Type") == "Story" else 0,
                   f"{r.get('Type')} / {r.get('Platform')}", r.get("Post link"), r.get("Profile")])
    if not bc:
        put(wc, row, ["", "No branded-content entries for this brand in the Ad Library data."]); row += 1
    for i, r in enumerate(sorted(bc, key=lambda r: r[2] or "", reverse=True), 1):
        put(wc, row, [i] + r, link_cols=(11, 12)); row += 1
    row += 1

    # 4. our full list
    row = section(wc, row, "4. ALL CREATOR COLLAB REELS ON THE BRAND'S OWN PAGE SINCE 1 APR 2026 (Posts grid + Reels tab)", NC)
    if not reels:
        put(wc, row, ["", "Not available yet." if handle else "Not scanned (excluded on request)."]); row += 1
    else:
        header(wc, row, ["#", "Creator", "Account Type", "Post Date", "Followers (Exact)", "Views", "Likes", "Comments",
                         "Paid Label", "Boosted? (est.)", "Reel Link", "Why Boosted (estimate)",
                         "Audio Language", "Audio Transcript"]); row += 1
        wc.column_dimensions["M"].width = 12; wc.column_dimensions["N"].width = 90
        for i, p in enumerate(reels, 1):
            at_ = account_type(handle, p["creator"], p.get("creator_name") or "", p.get("creator_category") or "")
            put(wc, row, [i, "@" + p["creator"],
                          "Creator" if at_ == "Creator / individual" else ("Brand's own sub-account" if at_ == "Own sub-brand" else "Brand / company"),
                          datetime.fromtimestamp(p["taken_at"], tz=timezone.utc).strftime("%Y-%m-%d"),
                          p.get("followers") if p.get("followers") is not None else "not resolved",
                          p.get("views") if p.get("views") is not None else "not shown",
                          "hidden" if p.get("likes_hidden") else p.get("likes"), p.get("comments"),
                          "ON" if p.get("is_paid_partnership") else "OFF", "Boosted" if p.get("is_boosted") else "Not boosted",
                          p["post_url"], p.get("boost_reason") or ""] + list(transcript_of(p["post_url"]))[::-1],
                link_cols=(11,)); row += 1
    # ---- summary block (top of tab) ---------------------------------------------
    norm_h = lambda h: re.sub(r"[^a-z0-9]", "", str(h or "").lower())
    org = {norm_h(p["creator"]) for p in creators_reels}
    boosted = {norm_h(p["creator"]) for p in creators_reels if p.get("is_boosted")}
    paid = {norm_h(p["creator"]) for p in creators_reels if p.get("is_paid_partnership")}
    live_part = {nm for nm, p in part.items() if any(a.get("status") == "Active" for a in p["ads"])}
    part_keys = set()
    for nm, p in part.items():
        part_keys |= {norm_h(nm), norm_h(p.get("ig"))} - {""}
    bc_keys = {norm_h(r[1]) for r in bc} | {norm_h(r[0]) for r in bc}
    bc_creators = {norm_h(r[1]) for r in bc}
    overlap = {h for h in org if h in part_keys or h in bc_keys}
    na = "not scanned" if handle is None else ("pending" if reels is None else None)
    kpis = [
        ("Total creators (collab reels on the brand's page since 1 Apr)", na or len(org)),
        ("Live creators - partnership ads running now in the Ad Library", len(live_part)),
        ("Total Ad Library partnership creators (live + ended)", len(part)),
        ("Branded content creators (Ad Library)", len(bc_creators)),
        ("Boosted creators - at least one reel looks boosted (ESTIMATE)", na or len(boosted)),
        ("Paid-partnership-label creators (Instagram's own tag)", na or len(paid)),
        ("Creators in both the organic collabs and the Ad Library", na or len(overlap)),
    ]
    x = wc.cell(3, 1, "SUMMARY"); x.font = Font(bold=True, size=12)
    for c in range(1, 4):
        wc.cell(3, c).fill = SEC
    for i, (k, v) in enumerate(kpis, 4):
        a = wc.cell(i, 2, k); a.font = B
        wc.merge_cells(start_row=i, start_column=2, end_row=i, end_column=6)
        b = wc.cell(i, 7, v); b.font = Font(bold=True, size=12); b.alignment = Alignment(horizontal="center")
    wc.freeze_panes = "B3"

    overview.append([disp, "@" + handle if handle else "not scanned", len(G), nads, live,
                     ", ".join(k for k, _ in themes_used.most_common(3)), len(part),
                     len(bc), len(reels) if reels else ("n/a" if handle is None else "pending"),
                     len({p['creator'] for p in creators_reels}) if reels else ""] + [v for _, v in kpis[1:6]])

# ---------------- overview
ov.cell(1, 1, "Jewellery brands - Meta Ad Library creatives and creators (as of 7 Oct 2026). Two tabs per brand: "
              "'- Ads' (hook, problem, solution, offer per creative) and '- Creators'.").font = H1
header(ov, 3, ["Brand", "Instagram", "Distinct Creatives", "Ads", "Live Ads", "Top Campaigns", "Partnership-Ad Creators",
               "Branded-Content Posts", "Our Collab Reels Since 1 Apr", "Distinct Creators",
               "Live Ad-Library Creators", "All Ad-Library Partnership Creators", "Branded-Content Creators",
               "Boosted Creators (est.)", "Paid-Label Creators"],
       [28, 28, 11, 8, 8, 60, 13, 13, 14, 11, 12, 13, 13, 12, 11])
for r, row in enumerate(overview, 4):
    put(ov, r, row)
notes = ["How to read it:",
         "- Hook is the ad's own first line: what is said in the first seconds (auto-transcribed) and/or the headline / copy opener.",
         "- Problem, Solution and Tone & Format are written per campaign after reading every creative; each ad row inherits its campaign's read.",
         "- Spend estimates, winner scores and other model fields from the source sheet are removed on purpose.",
         "- Transcripts come from the source sheet (Whisper base) and mis-hear names; brand names were corrected, nothing else.",
         "- 'Boosted' on our creator reels is an ESTIMATE from views vs engagement; Instagram does not publish it. 'Paid Label' is Instagram's own tag.",
         "- Our creator reels include both the brand's Posts grid and its Reels tab - creator-owned collab reels only appear on the Reels tab."]
for i, t in enumerate(notes, len(overview) + 6):
    ov.cell(i, 1, t).font = B if i == len(overview) + 6 else Font()
wb.save(OUT)
print(OUT)
for r in overview:
    print("  ", r)
