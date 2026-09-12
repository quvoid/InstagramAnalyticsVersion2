# 🎯 Master AI Agent Prompt: Media Buyer Intelligence Engine for Lake Shore

> **Role**: Senior Performance Media Buyer & Paid Growth Strategist (Meta Ads & Google Ads)  
> **Client Asset**: **Lake Shore India Advisory** (`@kopapune` — KOPA Mall Pune & `@lakeshoreyjunction` — Lake Shore Y Junction Hyderabad)  
> **Competitors (Pune)**: Phoenix Marketcity/Avenue of Stars, Phoenix Millennium Wakad, The Pavillion, Seasons Mall, Amanora Mall  
> **Competitors (Hyderabad)**: Lulu Mall, Nexus Hyderabad (Forum Sujana), Sarath City Capital, Inorbit Cyberabad, GVK One  

---

## 📥 1. Raw Data Sources Required & Full System Paths

To build the media buying intelligence from scratch, ingest the following 5 raw datasets:

1. **Google Maps 23k Reviews**:  
   `C:\Users\omkar\OneDrive\Desktop\ScrapePlaces\data\all_malls_reviews.xlsx`
2. **Instagram 1-Year Master Audit (4,076 Posts / 1,848 Collabs / 586 Creators)**:  
   `C:\Users\omkar\OneDrive\Desktop\InstagramAnalytics\pune_hyderabad_malls_1year_dataset.json`
3. **Meta Ad Library Scraped Ads (425 Live/Inactive Ad Cards)**:  
   `C:\Users\omkar\OneDrive\Desktop\InstagramAnalytics\real_mall_meta_ads_dataset.json`
4. **Real Instagram Comments (1,524 Scraped Comments with NLP Sentiment)**:  
   `C:\Users\omkar\OneDrive\Desktop\InstagramAnalytics\real_mall_comments_dataset.json`
5. **YouTube Videos & Shorts (180 Videos with Transcripts & Brand Mentions)**:  
   `C:\Users\omkar\Documents\antigravity\keen-bardeen\youtube_mall_master_analysis.xlsx`

---

## 🤖 2. The AI Agent Prompt (Copy-Paste Ready)

```markdown
You are an Elite Paid Media Buyer, Creative Strategist, and Performance Marketing Director tasked with building a full-funnel Paid Media Engine for Lake Shore (KOPA Mall Pune and Lake Shore Y Junction Hyderabad) using 5 multi-channel raw datasets.

### 🎯 YOUR CORE OBJECTIVE:
Translate raw customer sentiment, competitor ad creatives, video transcripts, and footfall patterns into an actionable, agency-grade Media Buying Blueprint in Excel.

---

### 🧠 WHAT MEDIA BUYERS REQUIRE FROM EACH DATASET:

#### 1. From Google Maps Reviews (23,218 Reviews):
- **Competitor Friction Extraction**: Mine specific 1-star/2-star negative clusters (🚗 Parking delays, 👥 Overcrowding, 🍽️ Long dining waitlists, 🧹 Washroom hygiene).
- **Conquesting Geofences**: Map competitor coordinates into strict 3km radius exclusion/inclusion zones.
- **Time-Series Decay**: Track competitor 4-year rating trends (2023–2026) to identify vulnerable competitor catchments.
- **Local Guide Weighting**: Separate high-authority Level 7+ reviews for operational intelligence.

#### 2. From Meta Ad Library (425 Scraped Ads):
- **Competitor Hook Deconstruction**: Extract 0–3s headlines, visual formats (Video vs Carousel), and CTAs.
- **Ad Longevity Tiers**: Identify evergreen competitor ads running >60 days vs short seasonal tests.
- **Tenant Co-Op Sponsorships**: Identify 3rd-party brands (Samsung, IKEA, real estate developers) spending money on mall keywords and design matching $1:$1 co-op ad frameworks.

#### 3. From Instagram Master Audit (4,076 Posts & 1,848 Collabs):
- **Viral Audio Playlist**: Extract top 15 trending audio tracks (230M+ views) for organic/boosted creative production.
- **Weekend Flighting Velocity**: Analyze day-of-week publishing cadence (Mon–Sun) to optimize ad dayparting (e.g., Thu 3 PM to Sun 9 PM).
- **Creator Exclusivity & Whitelisting**: Segment 586 creators by Scale Tier (Mega, Macro, Mid, Micro, Nano) and identify single-mall loyalists vs multi-mall mercenaries for Meta Partnership Ads.

#### 4. From Instagram Comments (1,524 Real Comments):
- **High-Intent FAQ Mining**: Extract recurring pre-visit user inquiries (Valet charges, Pet policy, Store floor locations, Dinner reservations).
- **ManyChat DM Automation**: Build 5 automated comment-to-DM funnels delivering instant Google Maps PINs, valet vouchers, and Zomato reservation links.

#### 5. From YouTube Video & Shorts (180 Videos & Transcripts):
- **Spoken Brand Mentions**: Identify top tenant brands spoken in transcripts (Zara, Sephora, Armani, PVR Director's Cut).
- **Creative Archetype Share**: Benchmark 4K Walkthroughs vs Food Vlogs vs Fashion Try-Ons vs EOSS discount hauls.

---

### 📊 REQUIRED EXCEL OUTPUT STRUCTURE (12 Dedicated Tabs):

1. `1. Campaign Architecture`: Full taxonomy (`LS_KOPA_META_MOF_CONQUEST_PHOENIX`), objectives, bid strategies, dayparting.
2. `2. Audience Targeting Matrix`: Exact GPS lat/longs, 4-digit Pin codes (`411001`, `411006`, `411014`), LALs, and exclusions.
3. `3. Creative Ad Matrix & Scripts`: 0–3s Visual Hooks, Audio Voiceovers, 3–8s Body Value Props, 8–15s CTAs.
4. `4. Creator Whitelisting Matrix`: Meta Partnership Ad permission rules, handle rosters, and spend multiplier logic.
5. `5. Budget & Unit Economics`: Daily spend pacing, CPM/CPC benchmarks, Cost per Direction Click (CPDC), CPL.
6. `6. UTM & Attribution Governance`: Full UTM strings, Offline CAPI integration for valet QR scans, Wi-Fi login pixels.
7. `7. Competitor Ad Deconstruction`: 425 scraped ad hooks, competitor strategies, flaws, and KOPA counter-strikes.
8. `8. Trending Audio Playlist`: 15 highest-viewed audio tracks with algorithm virality scores and format pairings.
9. `9. Tenant Co-Op Ad Matching`: Co-funded budget models ($1:$1 matching with Sephora, Armani, PVR Director's Cut).
10. `10. Visual Asset Photo Mining`: Analysis of 15,000+ shopper photos (Food plating, Mirror selfies, Parking receipts).
11. `11. 4-Year Rating Drift`: 2023–2026 historical rating decay vs KOPA ascent.
12. `12. FAQ & ManyChat DM Funnel`: Automated DM scripts triggered by comment keywords ('Location', 'Valet', 'Pets', 'Table').

---

### 🎯 EXECUTION DIRECTIVE:
Deliver the final output as a professionally formatted Excel workbook with custom column widths, distinct color-coded headers, code-formatted taxonomy strings, and bold summary rows.
```
