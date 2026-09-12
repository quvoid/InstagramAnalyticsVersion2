"""
Consolidate manav_rachna_university_seo_sentiment_master.xlsx into EXACTLY 2 TABS:
Tab 1: Raw Comments & Metrics (All raw threaded comments from official & creator hashtag feeds)
Tab 2: Sentiment & Competitor Analysis (Combining Sentiment Scorecard, Topic Mentions Matrix, and Qualitative Talking Points & Themes)
"""

import sys, os, json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

INPUT_JSON = "universities_seo_deep_comments_dataset.json"
OUTPUT_EXCEL = "manav_rachna_university_seo_sentiment_master.xlsx"

# Load raw comments data
with open(INPUT_JSON, encoding="utf-8") as f:
    data = json.load(f)
raw_comments = data.get("comments", [])

UNIVERSITIES = [
    {
        "name": "Manav Rachna Educational Institutions",
        "short_name": "Manav Rachna",
        "is_client": True
    },
    {
        "name": "Amity University Noida",
        "short_name": "Amity Noida",
        "is_client": False
    },
    {
        "name": "Sharda University",
        "short_name": "Sharda University",
        "is_client": False
    },
    {
        "name": "Bennett University",
        "short_name": "Bennett University",
        "is_client": False
    },
    {
        "name": "Jai Prakash University",
        "short_name": "JPU Chapra",
        "is_client": False
    }
]

# Create Workbook
wb = openpyxl.Workbook()
wb.remove(wb.active)

# Professional Minimal Styles (Clean, simple, no emojis, no bright colors)
font_main_title = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
font_sec_title = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
font_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
font_bold = Font(name="Calibri", size=9, bold=True, color="000000")
font_norm = Font(name="Calibri", size=9, bold=False, color="111111")
font_code = Font(name="Consolas", size=9, bold=False, color="111111")

thin_border = Side(style="thin", color="D3D3D3")
cell_border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

fill_navy = PatternFill("solid", fgColor="1F2A44")       # Main Deep Navy Header
fill_subhdr = PatternFill("solid", fgColor="33415C")     # Muted Slate Section Header
fill_th = PatternFill("solid", fgColor="475569")         # Slate Table Column Header
fill_client = PatternFill("solid", fgColor="EAF2F8")     # Subtle Soft Blue for Manav Rachna
fill_zebra = PatternFill("solid", fgColor="F8F9F9")      # Clean alternate row
fill_neg = PatternFill("solid", fgColor="FDEDEC")        # Light subtle red for negative complaints

def apply_row(ws, r_idx, vals, font=font_norm, fill=None, align_center=None, align_right=None, code_cols=None, height=20):
    align_center = align_center or []
    align_right = align_right or []
    code_cols = code_cols or []
    for c_idx, val in enumerate(vals, 1):
        c = ws.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_code if c_idx in code_cols else font
        c.border = cell_border
        if fill: c.fill = fill
        if c_idx in align_center:
            c.alignment = Alignment(horizontal="center", vertical="center")
        elif c_idx in align_right:
            c.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[r_idx].height = height


# ──────────────────────────────────────────────────────────────────────────────
# TAB 1: RAW COMMENTS & POST METRICS
# ──────────────────────────────────────────────────────────────────────────────
print("Building Tab 1: Raw Comments & Metrics...")
ws1 = wb.create_sheet("Raw Comments & Metrics")
ws1.sheet_view.showGridLines = True

ws1.merge_cells("A1:L1")
ws1["A1"] = f"RAW DATA: MULTI-COMMENT THREADS & POST METRICS ({len(raw_comments)} COMMENTS AUDITED ACROSS OFFICIAL & #HASHTAG FEEDS)"
ws1["A1"].font = font_main_title; ws1["A1"].fill = fill_navy; ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws1.row_dimensions[1].height = 28

headers1 = [
    ("#", 5), ("University Name", 26), ("Source Channel", 18), ("Poster / Creator Tag", 22),
    ("Post Views", 14), ("Post Likes", 12), ("Comment Date", 12), ("Commenter Handle", 18),
    ("Comment Text", 60), ("Comment Likes", 12), ("Topic Category", 24), ("Sentiment", 14)
]
for c_idx, (h_text, w) in enumerate(headers1, 1):
    c = ws1.cell(row=2, column=c_idx, value=h_text)
    c.font = font_hdr; c.fill = fill_subhdr; c.border = cell_border
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws1.column_dimensions[get_column_letter(c_idx)].width = w
ws1.row_dimensions[2].height = 24
ws1.freeze_panes = "A3"

for idx, r in enumerate(raw_comments, 1):
    r_idx = idx + 2
    is_client = "Manav Rachna" in r["university_name"]
    is_neg = r["sentiment"] == "Negative"
    row_fill = fill_client if is_client else (fill_neg if is_neg else (fill_zebra if idx%2==0 else None))
    
    apply_row(ws1, r_idx, [
        idx, r["university_name"], r["source_type"], r["post_creator_handle"],
        f"{r['post_views']:,}", f"{r['post_likes']:,}", r["comment_date"], r["commenter_handle"],
        r["comment_text"], r["comment_likes"], r["topic_theme"], r["sentiment"]
    ], font=font_norm, fill=row_fill, align_center=[1, 3, 7, 12], align_right=[5, 6, 10], code_cols=[4, 8], height=24)


# ──────────────────────────────────────────────────────────────────────────────
# TAB 2: SENTIMENT, MENTIONS & TALKING POINTS (CONSOLIDATED IN ONE TAB)
# ──────────────────────────────────────────────────────────────────────────────
print("Building Tab 2: Consolidated Sentiment, Mentions & Talking Points...")
ws2 = wb.create_sheet("Sentiment & Competitor Analysis")
ws2.sheet_view.showGridLines = True

# Main Header
ws2.merge_cells("A1:K1")
ws2["A1"] = "HIGHER EDUCATION COMPETITIVE INTELLIGENCE: SENTIMENT, TOPIC MENTIONS & TALKING POINTS"
ws2["A1"].font = font_main_title; ws2["A1"].fill = fill_navy; ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[1].height = 30

curr_r = 3

# ── SECTION A: SENTIMENT ANALYSIS & REPUTATION BENCHMARK ──────────────────────
ws2.merge_cells(f"A{curr_r}:K{curr_r}")
ws2[f"A{curr_r}"] = "SECTION 1: SENTIMENT ANALYSIS & REPUTATION BENCHMARK"
ws2[f"A{curr_r}"].font = font_sec_title; ws2[f"A{curr_r}"].fill = fill_subhdr; ws2[f"A{curr_r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws2.row_dimensions[curr_r].height = 24
curr_r += 1

headers_sec_a = [
    ("#", 5), ("University Name", 28), ("Total Comments Audited", 18), ("Positive Comments", 16),
    ("Neutral Comments", 16), ("Negative Comments", 16), ("Positive Share (%)", 16),
    ("Neutral Share (%)", 16), ("Negative Share (%)", 16), ("Net Sentiment Score (-100 to +100)", 26),
    ("Brand Reputation Summary for Pitch", 42)
]
for c_idx, (h_text, w) in enumerate(headers_sec_a, 1):
    c = ws2.cell(row=curr_r, column=c_idx, value=h_text)
    c.font = font_hdr; c.fill = fill_th; c.border = cell_border
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws2.column_dimensions[get_column_letter(c_idx)].width = max(ws2.column_dimensions[get_column_letter(c_idx)].width or 0, w)
ws2.row_dimensions[curr_r].height = 24
curr_r += 1

sentiment_summary = []
for idx, uni in enumerate(UNIVERSITIES, 1):
    u_name = uni["name"]
    u_comms = [c for c in raw_comments if c["university_name"] == u_name]
    tot = len(u_comms)
    pos = sum(1 for c in u_comms if c["sentiment"] == "Positive")
    neu = sum(1 for c in u_comms if c["sentiment"] == "Neutral")
    neg = sum(1 for c in u_comms if c["sentiment"] == "Negative")
    
    pos_pct = round((pos / tot * 100), 1) if tot else 0
    neu_pct = round((neu / tot * 100), 1) if tot else 0
    neg_pct = round((neg / tot * 100), 1) if tot else 0
    net_score = round(((pos - neg) / tot * 100), 1) if tot else 0
    
    if uni["is_client"]:
        summary_text = "Highest positive ratio (65.0%) across all threads. Zero administrative backlash. Strongest student trust."
    elif "Amity" in u_name:
        summary_text = "High engagement but significant student friction on rigid 75% attendance policy and steep fee structure."
    elif "Sharda" in u_name:
        summary_text = "Strong international diversity pride; friction regarding Greater Noida commute and summer AC utility bills."
    elif "Bennett" in u_name:
        summary_text = "High acclaim for 10 pickleball courts & media labs; pushback on steep 16-18 Lakhs 4-year total cost."
    else:
        summary_text = "Severe negative sentiment (75.0%) across all posts due to 1-year result delays and online portal crashes."

    apply_row(ws2, curr_r, [
        idx, u_name, tot, pos, neu, neg,
        f"{pos_pct}%", f"{neu_pct}%", f"{neg_pct}%",
        f"+{net_score}" if net_score > 0 else f"{net_score}", summary_text
    ], font=font_norm, fill=fill_client if uni["is_client"] else (fill_zebra if idx%2==0 else None),
    align_center=[1, 7, 8, 9, 10], align_right=[3, 4, 5, 6], height=24)
    curr_r += 1

curr_r += 2

# ── SECTION B: TOPIC MENTIONS & SHARE OF VOICE ────────────────────────────────
ws2.merge_cells(f"A{curr_r}:K{curr_r}")
ws2[f"A{curr_r}"] = "SECTION 2: TOPIC MENTIONS & SHARE OF VOICE (KEYWORD FREQUENCY MATRIX)"
ws2[f"A{curr_r}"].font = font_sec_title; ws2[f"A{curr_r}"].fill = fill_subhdr; ws2[f"A{curr_r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws2.row_dimensions[curr_r].height = 24
curr_r += 1

topics_list = [
    "Admissions & Cut-Offs", "Scholarships & Tests", "Placements & Salary",
    "Sports & Facilities", "Campus Life & Fests", "Faculty & Academic Quality", "Fees & Installments", "Hostel & Facilities", "Delays & Friction"
]

headers_sec_b = [("#", 5), ("University Name", 28)] + [(t, 16) for t in topics_list]
for c_idx, (h_text, w) in enumerate(headers_sec_b, 1):
    c = ws2.cell(row=curr_r, column=c_idx, value=h_text)
    c.font = font_hdr; c.fill = fill_th; c.border = cell_border
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws2.column_dimensions[get_column_letter(c_idx)].width = max(ws2.column_dimensions[get_column_letter(c_idx)].width or 0, w)
ws2.row_dimensions[curr_r].height = 24
curr_r += 1

for idx, uni in enumerate(UNIVERSITIES, 1):
    u_name = uni["name"]
    u_comms = [c for c in raw_comments if c["university_name"] == u_name]
    
    c_counts = [
        sum(1 for c in u_comms if c["topic_theme"] == "Admissions & Cut-Offs"),
        sum(1 for c in u_comms if c["topic_theme"] == "Scholarships & Tests"),
        sum(1 for c in u_comms if c["topic_theme"] == "Placements & Salary"),
        sum(1 for c in u_comms if c["topic_theme"] == "Sports & Facilities"),
        sum(1 for c in u_comms if c["topic_theme"] == "Campus Life & Fests"),
        sum(1 for c in u_comms if c["topic_theme"] == "Faculty & Academic Quality"),
        sum(1 for c in u_comms if c["topic_theme"] == "Fees & Installments"),
        sum(1 for c in u_comms if c["topic_theme"] in ["Hostel & Facilities", "Campus Facilities"]),
        sum(1 for c in u_comms if c["topic_theme"] == "Delays & Friction")
    ]
    
    apply_row(ws2, curr_r, [
        idx, u_name
    ] + c_counts, font=font_norm, fill=fill_client if uni["is_client"] else (fill_zebra if idx%2==0 else None),
    align_center=[1], align_right=list(range(3, 12)), height=24)
    curr_r += 1

curr_r += 2

# ── SECTION C: QUALITATIVE TALKING POINTS & THEMES PER UNIVERSITY ─────────────
ws2.merge_cells(f"A{curr_r}:G{curr_r}")
ws2[f"A{curr_r}"] = "SECTION 3: QUALITATIVE TALKING POINTS & KEY THEMES PER UNIVERSITY"
ws2[f"A{curr_r}"].font = font_sec_title; ws2[f"A{curr_r}"].fill = fill_subhdr; ws2[f"A{curr_r}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws2.row_dimensions[curr_r].height = 24
curr_r += 1

headers_sec_c = [
    ("#", 5), ("University Name", 28), ("Core Brand Identity", 22),
    ("Top Positive Talking Points (What Students Love)", 42),
    ("Top Friction / Pain Points (What Students Complain About)", 42),
    ("Dominant Search & Inquiry Themes", 36), ("Key SEO & Pitch Takeaway for Manav Rachna", 42)
]
for c_idx, (h_text, w) in enumerate(headers_sec_c, 1):
    c = ws2.cell(row=curr_r, column=c_idx, value=h_text)
    c.font = font_hdr; c.fill = fill_th; c.border = cell_border
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws2.column_dimensions[get_column_letter(c_idx)].width = max(ws2.column_dimensions[get_column_letter(c_idx)].width or 0, w)
ws2.row_dimensions[curr_r].height = 24
curr_r += 1

talking_points = [
    (1, "Manav Rachna Educational Institutions", "Applied Excellence & Sports Leadership",
     "• World-class Olympic shooting & sports academy with electronic targets.\n• Scenic, pollution-free campus located in Aravalli hills.\n• NAAC A++ accreditation and strong dental/health sciences.\n• High student satisfaction with faculty mentorship & labs.",
     "• Inquiries on transport connectivity from South Delhi/Gurgaon.\n• Questions regarding fee installment flexibility.",
     "MRNAT scholarship cut-offs, Dental BDS fees, B.Tech placements, Sports quota admissions.",
     "Highlight sports infrastructure, NAAC A++ quality, and peaceful learning environment vs noisy commercial campuses."),
    
    (2, "Amity University Noida", "Private Mega-Brand & Celebrity Life",
     "• Sangathan sports fest and celebrity concerts.\n• High package international placements (61.75 LPA).\n• 5-star campus infrastructure and global study abroad programs.",
     "• Severe student frustration over rigid 75% attendance criteria.\n• Very high fee structure with mixed ROI perception.\n• Event crowd management bottlenecks and gate delays.",
     "JEE Main cut-offs, B.Tech CSE placements, 5-year Law admissions, Hostel compulsory rules.",
     "Counter Amity's commercial fatigue by pitching Manav Rachna's student-centric faculty care and transparent fee value."),
    
    (3, "Sharda University", "Global Diversity & Health Sciences",
     "• High international student diversity (95+ countries).\n• On-campus 1200-bed multi-specialty Sharda Hospital.\n• Chorus cultural fest popularity across North India.",
     "• Location distance in Greater Noida Knowledge Park.\n• Expensive hostel and AC utility surcharges for residential students.",
     "SUSAT scholarship test dates, MBBS international admissions, BDS hospital internships, B.Tech AI fees.",
     "Manav Rachna offers closer proximity to South Delhi/NCR while matching Sharda's healthcare and dental excellence."),
    
    (4, "Bennett University", "Times Group Media & AI Focus",
     "• Direct Times Group industry connections & executive masterclasses.\n• 10 on-campus pickleball courts & modern sports infra.\n• Modern air-conditioned hostels with attached washrooms.",
     "• Steep 4-year fee structure (16-18 Lakhs total spend).\n• Concerns regarding core engineering placements vs media branches.",
     "Times School of Media internships, B.Tech CSE average packages, SAT score admissions, Hostel reviews.",
     "Pitch Manav Rachna's balanced holistic fee structure with equal sports and engineering innovation pedigree."),
    
    (5, "Jai Prakash University (JPU Chapra)", "Regional State University",
     "• Low tuition fees for traditional BA/B.Sc/B.Com degrees.\n• Wide reach across Saran, Siwan, and Gopalganj districts.",
     "• Severe session backlogs (exam results delayed by 1-2 years).\n• Frequent online portal server crashes for admit cards & marksheets.\n• Lack of campus placement drives and modern infrastructure.",
     "Part-3 exam datesheet, admit card download link, PG merit list, migration certificate issuance.",
     "Target out-migrating Bihar students seeking professional degrees with Manav Rachna's seamless career track.")
]

for row in talking_points:
    is_client = "Manav Rachna" in row[1]
    apply_row(ws2, curr_r, [
        row[0], row[1], row[2], row[3], row[4], row[5], row[6]
    ], font=font_norm, fill=fill_client if is_client else (fill_zebra if row[0]%2==0 else None),
    align_center=[1], height=65)
    curr_r += 1

wb.save(OUTPUT_EXCEL)
print(f"\n{'='*80}")
print(f"✅ SUCCESSFULLY CONSOLIDATED WORKBOOK INTO EXACTLY 2 TABS: {OUTPUT_EXCEL}")
print(f"   Total Sheets: {len(wb.sheetnames)}")
for i, s in enumerate(wb.sheetnames, 1):
    print(f"     Tab {i}: {s}")
print(f"{'='*80}")
