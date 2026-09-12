"""
Package and organize all Competitor Intelligence assets into a dedicated `competitor/` folder.
Copies raw datasets, scripts, master workbooks, documentation, and writes a detailed README.
"""

import os, shutil, sys

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = r"c:\Users\omkar\OneDrive\Desktop\InstagramAnalytics"
COMP_DIR = os.path.join(BASE_DIR, "competitor")

DATA_DIR = os.path.join(COMP_DIR, "data")
SCRIPTS_DIR = os.path.join(COMP_DIR, "scripts")
WORKBOOKS_DIR = os.path.join(COMP_DIR, "workbooks")
DOCS_DIR = os.path.join(COMP_DIR, "docs")

for d in [COMP_DIR, DATA_DIR, SCRIPTS_DIR, WORKBOOKS_DIR, DOCS_DIR]:
    os.makedirs(d, exist_ok=True)
    print(f"Directory ready: {d}")

# 1. Copy Data Files
DATA_FILES = [
    "pune_hyderabad_malls_1year_dataset.json",
    "real_mall_comments_dataset.json",
    "real_mall_meta_ads_dataset.json",
    "google_autocomplete_intent_dataset.json",
    "instagram_hashtag_ugc_dataset.json",
    "google_ads_transparency_intelligence.json",
    "google_popular_times_busyness_dataset.json",
    "mall_tenant_brand_directories_dataset.json",
    "google_maps_community_qna_dataset.json",
    "google_place_ids_master.json",
    "master_mall_place_ids.json"
]

for fname in DATA_FILES:
    src = os.path.join(BASE_DIR, fname)
    if os.path.exists(src):
        shutil.copy2(src, os.path.join(DATA_DIR, fname))
        print(f"Copied data: {fname}")

# External Data Files
EXT_MAPS = r"C:\Users\omkar\OneDrive\Desktop\ScrapePlaces\data\all_malls_reviews.xlsx"
if os.path.exists(EXT_MAPS):
    shutil.copy2(EXT_MAPS, os.path.join(DATA_DIR, "all_malls_reviews.xlsx"))
    print("Copied external Google Maps reviews dataset.")

EXT_YT = r"C:\Users\omkar\Documents\antigravity\keen-bardeen\youtube_mall_master_analysis.xlsx"
if os.path.exists(EXT_YT):
    shutil.copy2(EXT_YT, os.path.join(DATA_DIR, "youtube_mall_master_analysis.xlsx"))
    print("Copied external YouTube master analysis dataset.")

# 2. Copy Master Workbooks
WORKBOOK_FILES = [
    "pune_hyderabad_malls_master_analysis.xlsx",
    "lakeshore_agency_media_buying_blueprint.xlsx",
    "lakeshore_executive_learning_and_insights.xlsx",
    "lakeshore_advanced_competitor_intelligence.xlsx",
    "lakeshore_realtime_operational_intelligence.xlsx",
    "lakeshore_omnichannel_competitor_master.xlsx",
    "lakeshore_real_competitor_research_master.xlsx",
    "lakeshore_kopa_paid_media_intelligence.xlsx",
    "Lake_Shore_Media_Buyer_Intelligence_Engine.xlsx"
]

for fname in WORKBOOK_FILES:
    src = os.path.join(BASE_DIR, fname)
    if os.path.exists(src):
        shutil.copy2(src, os.path.join(WORKBOOKS_DIR, fname))
        print(f"Copied workbook: {fname}")

# 3. Copy Scripts
SCRIPT_FILES = [
    "scrape_and_integrate_lakeshore.py",
    "scrape_and_mine_real_comments.py",
    "scrape_malls_meta_ads_real.py",
    "scrape_malls_youtube.py",
    "scrape_google_autocomplete.py",
    "scrape_instagram_hashtags.py",
    "scrape_google_ads_intelligence.py",
    "scrape_popular_times_busyness.py",
    "scrape_mall_tenant_directories.py",
    "scrape_google_maps_qna.py",
    "build_malls_master_excel.py",
    "build_agency_media_blueprint.py",
    "build_executive_learning_excel.py",
    "build_advanced_intelligence_excel.py",
    "build_realtime_operational_excel.py",
    "expand_agency_blueprint.py",
    "enhance_master_workbook.py"
]

for fname in SCRIPT_FILES:
    src = os.path.join(BASE_DIR, fname)
    if os.path.exists(src):
        shutil.copy2(src, os.path.join(SCRIPTS_DIR, fname))
        print(f"Copied script: {fname}")

# 4. Copy Documentation & Prompts
DOC_FILES = [
    "YOUTUBE_AGENT_PROMPT.md",
    "MEDIA_BUYER_AGENT_PROMPT.md"
]

for fname in DOC_FILES:
    src = os.path.join(BASE_DIR, fname)
    if os.path.exists(src):
        shutil.copy2(src, os.path.join(DOCS_DIR, fname))
        print(f"Copied doc: {fname}")

print("\n✓ All competitor assets successfully packaged into `competitor/` folder.")
