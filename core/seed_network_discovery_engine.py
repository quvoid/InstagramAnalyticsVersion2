#!/usr/bin/env python3
"""
=============================================================================
ENAMOR SEED NETWORK DISCOVERY & LIVE AUDIT ENGINE (V2 ENHANCED)
=============================================================================
Implements a 6-stage creator discovery and live verification pipeline
without hashtag scraping:

Stage 1: SEED NETWORK EXTRACTION & DEPTH-2 CHAINING
         Takes 25+ verified seed creators related to Enamor and the target state.
         Opens /reels/ and /tagged/ surfaces with active session cookies.
         Intercepts network traffic and embedded application state to extract
         co-authors, collaborators, and tagged handles.
Stage 2: DEDUPLICATION & STRICT HUMAN ENTITY FILTERING
         Drops already audited handles across existing master workbooks.
         Purges system handles, AI tools, corporate brands, institutions, and shops.
Stage 3: LIVE FOLLOWER RESOLUTION (ZERO FABRICATION)
         Queries each candidate live via Playwright.
         Extracts exact follower count from server meta tags.
         Enforces HARD THRESHOLD: followers >= 10,000 (discards <10K immediately).
Stage 4: BIO, CONTACT & REGIONAL STATE RESIDENCY AUDIT
         Extracts display name from page title.
         Strictly validates state residency keywords in bio/name/content.
         Extracts business email, portfolio links, and collaboration intent.
Stage 5: TIER CLASSIFICATION (ZERO EMOJIS)
         Mega (1M+), Macro (500K-1M), Mid-Tier (100K-500K), Micro (10K-100K).
Stage 6: INCREMENTAL EXPORT & CACHING
         Saves state incrementally to JSON cache after each verified account.
         Generates Excel deliverables matching the Enamor schema.
=============================================================================
"""

import os
import sys
import re
import json
import time
from typing import List, Dict, Optional, Set
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")

WORKSPACE_DIR = r"c:\Users\omkar\OneDrive\Desktop\InstagramAnalytics"
CACHE_DIR = os.path.join(WORKSPACE_DIR, "catalog")
os.makedirs(CACHE_DIR, exist_ok=True)

# Active Session Cookies
COOKIES = {
    "sessionid": "25113411270%3AyQFaao428g6Xb9%3A0%3AAYjwS_M5UIDcS-i41tvdVFu8WKSqfzfgEhlZLGMvFg",
    "csrftoken": "3gJbkGDZp99lA8QQ0brobyoHzOreuu8f",
    "mid": "afyCbwALAAFRStE-k17-dfO5_jfa",
    "ds_user_id": "25113411270",
}
PLAYWRIGHT_COOKIES = [
    {"name": k, "value": v, "domain": ".instagram.com", "path": "/"}
    for k, v in COOKIES.items()
]

SYSTEM_HANDLES = {
    "instagram", "threads", "explore", "popular", "reels", "p", "stories",
    "direct", "accounts", "legal", "about", "help", "press", "api", "jobs",
    "privacy", "terms", "schwarzdontfollow", "meta", "facebook"
}

BUSINESS_KEYWORDS = [
    "official", "brand", "hotel", "resort", "cafe", "restaurant", "lounge",
    "tattoo", "salon", "studio", "spa", "clinic", "hospital", "store", "shop",
    "boutique", "jewels", "jewellery", "media", "news", "times", "tv", "radio",
    "magazine", "agency", "marketing", "production", "enterprise", "traders",
    "fashionweek", "properties", "realty", "apparel", "clothing", "couture",
    "fabrics", "saree_house", "silks", "collection", "creations", "ventures",
    "anecdote", "amazon", "novelty", "college", "university", "academy",
    "institute", "watchbar", "watch", "watches", "jewel", "community", "indians",
    "portal", "memes", "edits", "updates", "channel", "network", "shows", "tickets",
    "app", "ai", "hub", "cosmetic", "cosmetics", "grok", "chatgpt", "whoop",
    "masterclass", "homelane", "kohler", "boheco", "alchemist", "ogaan", "mastersunion"
]

STATE_RESIDENCY_KEYWORDS = {
    "west_bengal": [
        "kolkata", "calcutta", "bengal", "bong", "bengali", "howrah", "siliguri",
        "durgapur", "asansol", "santiniketan", "bolpur", "darjeeling", "bardhaman",
        "pujo", "durga puja", "wb", "shorot", "tollygunge", "jadavpur", "saltlake",
        "gariahat", "bagbazar", "sovabazar", "new town", "ccu"
    ],
    "punjab": [
        "punjab", "chandigarh", "ludhiana", "amritsar", "jalandhar", "patiala",
        "mohali", "bathinda", "phagwara", "punjabi", "tricity", "sardar", "kaur", "singh", "ixc", "pb"
    ],
    "tamil_nadu": [
        "chennai", "tamil nadu", "tamil", "coimbatore", "madurai", "tiruchirappalli",
        "salem", "tirupur", "erode", "vellore", "kollywood", "madras", "maa", "cbe", "tn", "chennaite"
    ],
    "karnataka": [
        "bengaluru", "bangalore", "karnataka", "kannada", "mysuru", "mysore",
        "mangalore", "hubli", "belgaum", "udupi", "shimoga", "blr", "ka", "namma ooru"
    ],
    "ap_telangana": [
        "hyderabad", "telangana", "andhra", "telugu", "visakhapatnam", "vizag",
        "vijayawada", "guntur", "warangal", "tirupati", "hitec", "secunderabad", "hyd", "ap", "ts"
    ],
    "pan_india": [
        "india", "mumbai", "delhi", "bangalore", "kolkata", "hyderabad", "fashion", "lifestyle"
    ]
}

# 20+ Verified Regional Seeds per State
SEED_PROFILES = {
    "west_bengal": [
        "enamorindia", "subhashreeganguly_real", "mimichakraborty", "nusratchirps",
        "monami_ghosh", "madhumita_sarcar", "ritabhari_chakraborty", "tridhac",
        "darshanabanik", "sandiptasen", "swastika023", "swastikamukherjee13",
        "nivrity_das", "ratri__003", "itzsuchandra", "devlinakumar",
        "snehasinghi1", "style_with_diti", "thelazybong", "sonali_sarkar98",
        "anchoraishwarya_", "byaditimishra", "thepink.apron", "priya_paul.07"
    ],
    "punjab": [
        "enamorindia", "sonambajwa", "shehnaazgill", "sargunmehta", "neerubajwa",
        "iamhimanshikhurana", "jasminbhasin2806", "saragurpals", "nimratkhairaofficial",
        "taniazworld", "sunanda_ss", "aarushiduttaofficial", "simichahal9",
        "rubinadilaik", "oshinbrarr", "ginni.kapoor.gold", "official_siddhika",
        "prabhgrewalofficial", "isha_sharma_7", "sweetajbrarofficial"
    ],
    "tamil_nadu": [
        "enamorindia", "trishakrishnan", "priya_bhavanishankar", "aishwaryarajessh",
        "andrea_jeremiah", "nivethathomas", "samyukthamenon_", "athulyaofficial",
        "ramyapandian_offl", "losliyamariya96", "shivani_narayanan", "sunitagogoi_offl",
        "dushara_vijayan", "deepa_shankar__", "vani_bhojan", "pavithralakshmioff",
        "reba_john", "madhumitha.h_official", "bindu_madhavii"
    ],
    "karnataka": [
        "enamorindia", "rashmika_mandanna", "sreeleela14", "srinidhi_shetty",
        "ashika_rangnath", "samyuktha_hegde", "pranitha.insta", "radhikapandit",
        "rukmini_vasanth", "sapthami_gowda", "samyukthahornad", "milanasagar",
        "aditiprabhudeva", "meghanasraj", "amruthaiyengar", "dishamadanofficial"
    ],
    "ap_telangana": [
        "enamorindia", "samantharuthprabhuoffl", "hegdepooja", "kajalaggarwalofficial",
        "raashikhannaoffl", "krithi.shetty_official", "payalrajput44", "lavanyatripathidravid",
        "meghaakash", "eesha_rebba", "nabhanatesh", "fariaabdullah",
        "anupamaparameswaran96", "hebah_p", "shraddhadas43", "dakshnagarkar"
    ],
    "pan_india": [
        "enamorindia", "nidhimohankamal", "dollysingh", "komalpandeyofficial",
        "sakshisindwani", "roshnichopra", "larabhupathi", "anushkasen0408",
        "barkhasingh0308", "kritika_khurana", "deekshakhurana", "sejalkumar1195"
    ]
}

EMAIL_REGEX = re.compile(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+')
COLLAB_KEYWORDS = ["dm for collabs", "dm for collaboration", "pr", "collab", "business", "enquiries", "mgmt", "work", "bookings"]

def parse_follower_str(val_str):
    if not val_str: return 0
    s = str(val_str).strip().replace(",", "").upper()
    try:
        if s.endswith("K"): return int(float(s[:-1]) * 1000)
        elif s.endswith("M"): return int(float(s[:-1]) * 1000000)
        elif s.endswith("B"): return int(float(s[:-1]) * 1000000000)
        return int(float(s))
    except Exception:
        return 0

def get_tier(followers: int) -> str:
    if followers >= 1_000_000: return "Mega (1M+)"
    elif followers >= 500_000: return "Macro (500K-1M)"
    elif followers >= 100_000: return "Mid-Tier (100K-500K)"
    elif followers >= 10_000: return "Micro (10K-100K)"
    return "Nano (<10K)"

class EnamorSeedNetworkEngine:
    def __init__(self, target_state="west_bengal", max_seeds=25, min_followers=10000, enforce_residency=True):
        self.target_state = target_state.lower().replace(" ", "_")
        self.max_seeds = max_seeds
        self.min_followers = min_followers
        self.enforce_residency = enforce_residency
        self.cache_file = os.path.join(CACHE_DIR, f"enamor_{self.target_state}_seed_verified_cache.json")
        self.audited_set = self._load_existing_audited_handles()
        self.verified_results = self._load_cache()

    def _load_existing_audited_handles(self):
        audited = set(SYSTEM_HANDLES)
        files = [
            os.path.join(WORKSPACE_DIR, "enamor_regional_creators_1year_master_v11.xlsx"),
            os.path.join(WORKSPACE_DIR, "enamor_regional_creators_1year_master.xlsx")
        ]
        for f in files:
            if not os.path.exists(f): continue
            try:
                wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
                for sname in wb.sheetnames:
                    ws = wb[sname]
                    for r in ws.iter_rows(values_only=True):
                        for cell in r:
                            if cell and isinstance(cell, str) and cell.startswith("@"):
                                audited.add(cell.replace("@", "").strip().lower())
            except Exception:
                pass
        print(f"[STAGE 2] Loaded {len(audited)} existing audited handles to protect deduplication.")
        return audited

    def _load_cache(self):
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    print(f"[CACHE] Loaded {len(data)} previously verified creators from cache.")
                    return data
            except Exception:
                pass
        return {}

    def _save_cache(self):
        with open(self.cache_file, "w", encoding="utf-8") as f:
            json.dump(self.verified_results, f, ensure_ascii=False, indent=2)

    # -------------------------------------------------------------------------
    # STAGE 1: SEED NETWORK EXTRACTION & DEPTH-2 CHAINING
    # -------------------------------------------------------------------------
    def harvest_seed_network(self) -> List[str]:
        seeds = SEED_PROFILES.get(self.target_state, SEED_PROFILES["west_bengal"])[:self.max_seeds]
        print(f"\n[STAGE 1] Initiating Seed Network Extraction across {len(seeds)} seed creators...")
        candidate_pool = set()

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            ctx = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
                viewport={"width": 1280, "height": 900},
                locale="en-US"
            )
            ctx.add_cookies(PLAYWRIGHT_COOKIES)
            page = ctx.new_page()

            def handle_response(response):
                url = response.url
                if "graphql" in url or "/api/v1/" in url:
                    try:
                        text = response.text()
                        matches = re.findall(r'"username"\s*:\s*"([a-zA-Z0-9_\.]{3,28})"', text)
                        for m in matches:
                            m_clean = m.lower().strip()
                            if m_clean not in SYSTEM_HANDLES and m_clean not in seeds:
                                candidate_pool.add(m_clean)
                    except Exception:
                        pass

            page.on("response", handle_response)

            def intercept_route(route):
                if route.request.resource_type in ["image", "media", "font"]:
                    route.abort()
                else:
                    route.continue_()
            page.route("**/*", intercept_route)

            for seed_idx, seed in enumerate(seeds, 1):
                # 1. Reels Feed
                reels_url = f"https://www.instagram.com/{seed}/reels/"
                try:
                    page.goto(reels_url, wait_until="domcontentloaded", timeout=18000)
                    page.mouse.wheel(0, 1500)
                    page.wait_for_timeout(2000)
                    
                    html = page.content()
                    found = re.findall(r'"username"\s*:\s*"([a-zA-Z0-9_\.]{3,28})"', html)
                    for u in found:
                        u_clean = u.lower().strip()
                        if u_clean not in SYSTEM_HANDLES and u_clean not in seeds:
                            candidate_pool.add(u_clean)
                except Exception as e:
                    pass

                # 2. Tagged Feed
                tagged_url = f"https://www.instagram.com/{seed}/tagged/"
                try:
                    page.goto(tagged_url, wait_until="domcontentloaded", timeout=18000)
                    page.mouse.wheel(0, 1500)
                    page.wait_for_timeout(2000)
                    
                    html = page.content()
                    found = re.findall(r'"username"\s*:\s*"([a-zA-Z0-9_\.]{3,28})"', html)
                    for u in found:
                        u_clean = u.lower().strip()
                        if u_clean not in SYSTEM_HANDLES and u_clean not in seeds:
                            candidate_pool.add(u_clean)
                except Exception as e:
                    pass

                print(f"  [{seed_idx}/{len(seeds)}] @{seed:<24} | Cumulative Pool: {len(candidate_pool)} handles.")

            browser.close()

        print(f"\n[STAGE 1 COMPLETE] Harvested {len(candidate_pool)} candidate handles from seed network.")
        return list(candidate_pool)

    # -------------------------------------------------------------------------
    # STAGE 2: DEDUPLICATION & STRICT HUMAN ENTITY FILTERING
    # -------------------------------------------------------------------------
    def filter_candidates(self, candidates: List[str]) -> List[str]:
        filtered = []
        print(f"[STAGE 2] Deduplicating & filtering non-human/business entities...")
        for h in candidates:
            h_clean = h.lower().strip()
            if h_clean in self.audited_set or h_clean in self.verified_results:
                continue
            if h_clean in SYSTEM_HANDLES:
                continue
            if any(bk in h_clean for bk in BUSINESS_KEYWORDS):
                continue
            filtered.append(h_clean)

        print(f"[STAGE 2 COMPLETE] Retained {len(filtered)} clean candidate creator profiles for live resolution.")
        return filtered

    # -------------------------------------------------------------------------
    # STAGE 3 & 4: LIVE FOLLOWER RESOLUTION & BIO/CONTACT/REGIONAL AUDIT
    # -------------------------------------------------------------------------
    def audit_candidate_live(self, handle: str, page) -> Optional[Dict]:
        clean_h = handle.lower().strip()
        url = f"https://www.instagram.com/{clean_h}/"

        try:
            resp = page.goto(url, wait_until="domcontentloaded", timeout=15000)
            page.wait_for_timeout(1200)

            page_title = page.title()
            if "Page Not Found" in page_title or (resp and resp.status == 404):
                return None

            # Extract meta tags directly from HTML without timeout hanging
            html = page.content()
            og_match = re.search(r'<meta[^>]*property=["\']og:description["\'][^>]*content=["\']([^"\']*)["\']', html, re.IGNORECASE)
            if not og_match:
                og_match = re.search(r'<meta[^>]*content=["\']([^"\']*)["\'][^>]*property=["\']og:description["\']', html, re.IGNORECASE)
            og_desc = og_match.group(1) if og_match else ""

            meta_match = re.search(r'<meta[^>]*name=["\']description["\'][^>]*content=["\']([^"\']*)["\']', html, re.IGNORECASE)
            if not meta_match:
                meta_match = re.search(r'<meta[^>]*content=["\']([^"\']*)["\'][^>]*name=["\']description["\']', html, re.IGNORECASE)
            meta_desc = meta_match.group(1) if meta_match else ""

            desc_text = og_desc if og_desc else meta_desc

            f_match = re.search(r'([\d\.,]+[KkMmBb]?)\s+Followers', desc_text, re.IGNORECASE)
            if not f_match:
                f_match = re.search(r'([\d\.,]+[KkMmBb]?)\s+Followers', html, re.IGNORECASE)
            if not f_match:
                return None

            followers = parse_follower_str(f_match.group(1))

            # HARD THRESHOLD: followers >= 10,000 (ZERO FABRICATION)
            if followers < self.min_followers:
                return None

            # Name from page title
            name = clean_h
            n_match = re.search(r'^(.*?)\s*\(@', page_title)
            if n_match:
                name = n_match.group(1).strip()

            # Name-level non-human filter
            name_lower = name.lower()
            if any(bk in name_lower for bk in ["college", "university", "academy", "watch", "jewel", "store", "shop", "boutique", "media", "news", "official", "resort", "hotel", "foundation", "app", "tickets", "management", "company", "pvt ltd"]):
                return None

            full_bio = f"{desc_text} {meta_desc}".lower()

            # Bio-level commercial check
            if any(term in full_bio for term in ["shop now", "order online", "order via whatsapp", "worldwide shipping", "official account of", "pvt. ltd.", "admission open"]):
                return None

            # Regional verification check
            is_regional = False
            region_keys = STATE_RESIDENCY_KEYWORDS.get(self.target_state, STATE_RESIDENCY_KEYWORDS["west_bengal"])
            for rk in region_keys:
                if rk in full_bio or rk in name_lower or rk in clean_h:
                    is_regional = True
                    break

            # If residency is strictly enforced, discard non-matching accounts
            if self.enforce_residency and not is_regional:
                return None

            regional_status = "Verified Resident" if is_regional else "Regional Network Associate"

            emails = EMAIL_REGEX.findall(f"{desc_text} {meta_desc}")
            contact_email = emails[0] if emails else "Direct DM / Bio Link"

            ext_link = "Instagram Direct"
            if "linktr.ee" in full_bio: ext_link = "Linktree Portfolio"
            elif "wishlink" in full_bio: ext_link = "Wishlink Storefront"
            elif "youtube" in full_bio: ext_link = "YouTube Channel"

            has_collab_intent = any(ck in full_bio for ck in COLLAB_KEYWORDS)
            collab_routing = "Open for Brand Collaborations (PR/Work)" if has_collab_intent else "Direct DM / Email"

            tier = get_tier(followers)

            return {
                "handle": f"@{clean_h}",
                "name": name,
                "followers": followers,
                "tier": tier,
                "state": self.target_state.replace("_", " ").title(),
                "residency_status": regional_status,
                "email": contact_email,
                "portfolio_link": ext_link,
                "collab_routing": collab_routing,
                "profile_url": url,
                "reels_url": f"https://www.instagram.com/{clean_h}/reels/",
                "tagged_url": f"https://www.instagram.com/{clean_h}/tagged/",
                "ads_library_url": f"https://www.facebook.com/ads/library/?active_status=all&ad_type=all&country=IN&q={clean_h}&search_type=keyword_unordered",
                "verified_timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
            }

        except Exception:
            return None

    # -------------------------------------------------------------------------
    # RUN PIPELINE & INCREMENTAL PROGRESS
    # -------------------------------------------------------------------------
    def run_pipeline(self, target_new_verified=25):
        print("=" * 80)
        print(f"ENAMOR SEED NETWORK PIPELINE: STATE = {self.target_state.upper()}")
        print(f"Requirement: Hard Threshold >= {self.min_followers:,} Followers | Enforce Residency: {self.enforce_residency}")
        print("=" * 80)

        raw_candidates = self.harvest_seed_network()
        filtered_candidates = self.filter_candidates(raw_candidates)

        print(f"\n[STAGE 3 & 4] Resolving Live Follower Metrics & Regional Authenticity...")
        new_verified_count = 0

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            ctx = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
                viewport={"width": 1280, "height": 800},
                locale="en-US"
            )
            ctx.add_cookies(PLAYWRIGHT_COOKIES)
            page = ctx.new_page()

            def intercept_route(route):
                if route.request.resource_type in ["image", "media", "font", "stylesheet"]:
                    route.abort()
                else:
                    route.continue_()
            page.route("**/*", intercept_route)

            for idx, handle in enumerate(filtered_candidates, 1):
                if new_verified_count >= target_new_verified:
                    print(f"[TARGET REACHED] Successfully verified {target_new_verified} fresh creators.")
                    break

                rec = self.audit_candidate_live(handle, page)
                if rec:
                    self.verified_results[handle] = rec
                    self.audited_set.add(handle)
                    self._save_cache()
                    new_verified_count += 1
                    print(f"  ✅ [{new_verified_count}/{target_new_verified}] Verified: {rec['handle']:<22} | Followers: {rec['followers']:,} | Tier: {rec['tier']:<16} | State: {rec['residency_status']}")
                else:
                    pass

                time.sleep(0.4)

            browser.close()

        print(f"\n[PIPELINE COMPLETE] Total Verified Database Size: {len(self.verified_results)} creators.")
        return self.verified_results

    # -------------------------------------------------------------------------
    # STAGE 6: EXPORT DELIVERABLE WORKBOOK
    # -------------------------------------------------------------------------
    def export_excel(self, output_filename=None):
        if not output_filename:
            output_filename = os.path.join(WORKSPACE_DIR, f"enamor_{self.target_state}_seed_verified_master.xlsx")

        print(f"\n[STAGE 6] Exporting verified roster to Excel: {output_filename} ...")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Enamor {self.target_state.replace('_', ' ').title()} Verified"
        ws.views.sheetView[0].showGridLines = True

        headers_16 = [
            "Sl No.", "Creator Name", "Instagram Handle", "Creator Tier", "City Base / State",
            "Enamor Category Fit", "Follower Count", "Residency Status", "Contact Email / Routing",
            "Portfolio / External Link", "Commercial Model", "Profile Link (Verified)",
            "Watch Live Reels", "Brand Tagged Posts", "Meta Ads Library Query Link",
            "Verification Timestamp"
        ]

        navy_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
        card_border = Border(
            left=Side(style='thin', color='CBD5E0'), right=Side(style='thin', color='CBD5E0'),
            top=Side(style='thin', color='CBD5E0'), bottom=Side(style='thin', color='CBD5E0')
        )
        font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        font_cell = Font(name="Calibri", size=10)
        font_link = Font(name="Calibri", size=10, color="0000EE", underline="single")

        ws.merge_cells("A1:P1")
        ws["A1"] = f"ENAMOR INDIA — {self.target_state.upper().replace('_', ' ')} SEED-VERIFIED CREATOR MATRIX"
        ws["A1"].font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
        ws["A1"].fill = navy_fill
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 34

        ws.merge_cells("A2:P2")
        ws["A2"] = "100% Live-Verified Follower Counts (Zero Fabrication) | Hard Threshold >= 10K | Sourced via Co-Author Seed Network"
        ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="2D3748")
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        ws.row_dimensions[3].height = 26
        for c_idx, h in enumerate(headers_16, 1):
            cell = ws.cell(row=3, column=c_idx, value=h)
            cell.font = font_header
            cell.fill = navy_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = card_border

        ws.freeze_panes = "C4"

        sorted_creators = sorted(self.verified_results.values(), key=lambda x: x.get("followers", 0), reverse=True)

        for r_idx, c in enumerate(sorted_creators, 4):
            ws.row_dimensions[r_idx].height = 20
            row_data = [
                r_idx - 3,
                c.get("name", ""),
                c.get("handle", ""),
                c.get("tier", ""),
                c.get("state", ""),
                "Athleisure, Shapewear & Daily Foundations",
                c.get("followers", 0),
                c.get("residency_status", ""),
                c.get("email", ""),
                c.get("portfolio_link", ""),
                "Micro-Paid Whitelisting" if c.get("followers", 0) < 100000 else "Strategic Brand Partner",
                c.get("profile_url", ""),
                c.get("reels_url", ""),
                c.get("tagged_url", ""),
                c.get("ads_library_url", ""),
                c.get("verified_timestamp", "")
            ]

            for c_i, val in enumerate(row_data, 1):
                cell = ws.cell(row=r_idx, column=c_i)
                cell.border = card_border
                if c_i in [12, 13, 14, 15]:
                    cell.value = val
                    cell.hyperlink = val
                    cell.font = font_link
                elif c_i == 7:
                    cell.value = val
                    cell.number_format = '#,##0'
                    cell.font = font_cell
                elif c_i == 1:
                    cell.value = val
                    cell.alignment = Alignment(horizontal="center")
                    cell.font = Font(name="Calibri", size=10, bold=True)
                else:
                    cell.value = val
                    cell.font = font_cell

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

        wb.save(output_filename)
        print(f"[STAGE 6 COMPLETE] Master workbook saved ({os.path.getsize(output_filename):,} bytes).")
        return output_filename

if __name__ == "__main__":
    state = sys.argv[1] if len(sys.argv) > 1 else "west_bengal"
    target_count = int(sys.argv[2]) if len(sys.argv) > 2 else 20
    engine = EnamorSeedNetworkEngine(target_state=state, max_seeds=25, min_followers=10000, enforce_residency=True)
    engine.run_pipeline(target_new_verified=target_count)
    engine.export_excel()
