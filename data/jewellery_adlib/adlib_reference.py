# -*- coding: utf-8 -*-
"""'Ads library reference only' tab - the client's raw Meta Ad Library export (sheet 1tK_Hi...),
re-transcribed with Groq whisper-large-v3. Spend, impressions and winner-score columns are left
out on purpose. Called from build_diwali.py: add_tab(wb, helpers)."""
import csv, glob, json, os, re, collections

HERE = os.path.dirname(os.path.abspath(__file__))
NOT_JEWELLERY = {"Urbanscape Properties", "Muggo.Shoes", "CM Yuva", "One District One Product", "Sarees By Aari",
                 "Tooletries", "HKJC Entertainment", "VAN HAM Kunstauktionen", "Makeupbykavya", "Nevaeh Store"}
IMITATION = re.compile(r"imitation|artificial|fashion jewel|rental|for rent|one gram|1 gram|anti.?tarnish|demi.?fine|plated", re.I)
OFFER_RX = re.compile(r"(up ?to \d+ ?% ?off[^.,;!\n]{0,40}|flat \d+ ?% ?off[^.,;!\n]{0,30}|\d+ ?% ?off[^.,;!\n]{0,40}|"
                      r"buy \d+ get \d+[^.,;!\n]{0,20}|free shipping|cod available|cash on delivery|free gift[^.,;!\n]{0,20}|"
                      r"(?:rs\.?|inr|₹) ?[\d,]+(?:/-)?[^.,;!\n]{0,25})", re.I)


def _rows(name):
    return list(csv.DictReader(open(HERE + "/" + name, encoding="utf-8")))


def _lid(r):
    return re.sub(r"\D", "", (r.get("Library link") or "").split("id=")[-1]) or r.get("Library ID", "")


def _junk(t):
    w = t.split()
    return not w or (len(w) > 12 and len(set(w)) < len(w) / 5)


def _first(t, n=180):
    t = re.sub(r"\s+", " ", t or "").strip()
    return (re.split(r"(?<=[.!?])\s", t)[0] if t else "")[:n]


def category(adv, copy):
    if adv in NOT_JEWELLERY:
        return "Not jewellery (keyword match)"
    if IMITATION.search(adv + " " + copy):
        return "Imitation / fashion / rental jewellery"
    if re.search(r"silver|925", adv + " " + copy, re.I):
        return "Silver jewellery"
    if re.search(r"diamond|solitaire|lab.?grown", adv + " " + copy, re.I):
        return "Diamond / fine jewellery"
    if re.search(r"sabyasachi|abhinav mishra|designer|couture|atelier", adv + " " + copy, re.I):
        return "Designer / couture"
    return "Gold / fine jewellery"


def theme(copy, angles):
    c = copy.lower(); out = []
    for k, rx in [("Bridal / wedding", r"bride|bridal|wedding|shaadi|vivaah"), ("Festive", r"diwali|dhanteras|navratri|festive|onam|karwa|durga|pujo|akshaya"),
                  ("Gifting", r"gift|anniversary|valentine|rakhi"), ("Offer / sale", r"% off|sale|discount|offer|buy 1|free"),
                  ("New launch", r"new collection|launch|introducing|unveil"), ("Everyday / lightweight", r"everyday|daily wear|lightweight|office"),
                  ("Trust / certified", r"hallmark|certified|bis|igi|gia|purity"), ("Rental", r"rent")]:
        if re.search(rx, c):
            out.append(k)
    if "Creator" in angles:
        out.append("Creator / UGC")
    return ", ".join(dict.fromkeys(out)) or "Product showcase"


def add_tab(wb, put, header, B, SEC, HDR, W, WRAP, LINK, Font):
    raw, tr_sheet = _rows("grt_adslibraryraw.csv"), {_lid(r): r for r in _rows("grt_adslibrarytranscript.csv")}
    ours = {}
    for f in glob.glob(HERE + "/adlib_ref_transcripts_*.json"):
        ours.update(json.load(open(f, encoding="utf-8")))
    seen, rows = set(), []
    for r in raw:
        lid = _lid(r)
        if lid in seen:
            continue
        seen.add(lid)
        t2 = tr_sheet.get(lid, {})
        copy = r.get("Ad copy") or t2.get("Ad copy") or ""
        g = ours.get(lid) or {}
        if g.get("text"):
            tr = g["text"]; hook_sp = g.get("hook3s") or ""; lang = g.get("lang", "").title(); src = "Groq Whisper large-v3"
        else:
            tr = r.get("Transcript") or t2.get("Full transcript") or ""; hook_sp = r.get("Spoken hook (first 3 s)") or t2.get("Spoken hook (first 3 s)") or ""
            lang = r.get("Voiceover language") or t2.get("Voiceover language") or ""; src = "sheet (whisper-base)" if tr else ""
        is_video = r.get("Format") in ("VIDEO", "DCO") and (tr or g)
        if tr.startswith("no speech") or tr == "no audio track" or (not tr and is_video):
            spoken = "Music only - no voiceover"; tr_out = "Music only - no voiceover"
        elif _junk(tr):
            spoken = ""; tr_out = "Unusable (speech model looped) - see ad copy"
        else:
            spoken = hook_sp if hook_sp and not _junk(hook_sp) else _first(tr, 140); tr_out = tr
        head = r.get("Headline") or ""
        hook = ("Spoken: '" + spoken + "'") if spoken and not spoken.startswith("Music") else \
               ("Headline: '" + head + "'" if head and "{{" not in head else "Copy: '" + _first(copy) + "'")
        offers = sorted({m.group(0).strip(" .") for m in OFFER_RX.finditer(copy + " " + tr_out)}, key=str.lower)
        offers = "; ".join(dict.fromkeys([r.get("Offers") or ""] + offers[:5])).strip("; ") or "No offer stated"
        creator = r.get("Creator (partnership)") or ""
        rows.append([r["Advertiser"], category(r["Advertiser"], copy), "Live" if r["Status"] == "Active" else "Ended",
                     r.get("Started", ""), r.get("Last seen / ended", ""), r.get("Days running", ""), r.get("Format", ""),
                     r.get("Video length (s)", "") or t2.get("Video length (s)", ""), lang or r.get("Language", ""),
                     hook, theme(copy, r.get("Angles") or ""), offers, r.get("CTA", "").replace("_", " ").title(),
                     ("@" + (r.get("Creator IG handle") or creator)) if creator else "-",
                     head if "{{" not in head else "catalogue - filled per product",
                     copy if "{{" not in copy else "catalogue - filled per product",
                     r.get("Landing URL", ""), r["Library link"], tr_out, src])
    rows.sort(key=lambda x: (x[1].startswith("Not"), x[2] != "Live", x[0].lower(), -(int(x[5]) if str(x[5]).isdigit() else 0)))

    ws = wb.create_sheet("Ads library reference only")
    ws.sheet_view.showGridLines = False
    live = sum(1 for x in rows if x[2] == "Live"); vids = sum(1 for x in rows if x[6] in ("VIDEO", "DCO"))
    cats = collections.Counter(x[1] for x in rows)
    ws.cell(1, 1, "ADS LIBRARY - REFERENCE ONLY (client's raw Meta Ad Library export, India, jewellery keyword search)").font = Font(bold=True, size=14, color="1F3864")
    ws.cell(2, 1, f"{len(rows)} ads from {len({x[0] for x in rows})} advertisers - {live} live, {vids} video. Spend, impressions and "
                  f"winner scores are deliberately left out. Videos were re-transcribed with Whisper large-v3 because the sheet's "
                  f"own transcripts (whisper-base) were often looped or garbled. Hook = the ad's spoken opening line; where "
                  f"there is no voiceover, its headline or first line of copy.")
    ws.cell(2, 1).alignment = WRAP; ws.merge_cells("A2:L2"); ws.row_dimensions[2].height = 45
    ws.cell(3, 1, "Mix: " + " | ".join(f"{k} {v}" for k, v in cats.most_common())).font = B
    ws.cell(4, 1, "Note: these advertisers are mostly small, imitation and designer labels - NOT GRT's direct competitors. "
                  "Rows marked 'Not jewellery' came in through the keyword search and are listed last.").font = Font(italic=True, color="9C0006")
    cols = ["Advertiser", "Category", "Status", "Started", "Last Seen", "Days Running", "Format", "Video (s)", "Language",
            "Hook (opening line)", "Theme", "Offer", "CTA", "Creator Partner", "Headline", "Ad Copy", "Landing Page",
            "Ad Library Link", "Full Transcript", "Transcript Source"]
    header(ws, 6, cols, [22, 22, 8, 11, 11, 8, 9, 7, 10, 50, 26, 34, 12, 16, 30, 60, 30, 30, 70, 16])
    for i, x in enumerate(rows, 7):
        put(ws, i, x, link_cols=(17, 18))
    ws.freeze_panes = "B7"
    from openpyxl.utils import get_column_letter
    ws.auto_filter.ref = f"A6:{get_column_letter(len(cols))}{6 + len(rows)}"
    return len(rows)
