"""
Streamlit Cron Job & Instagram Creator Intelligence Dashboard
Interactive UI with Background Scheduled Automation, 4-Tier Analytics, and Excel Exports
"""

import sys, os, json, time
import pandas as pd
import streamlit as st
from datetime import datetime

from cron_scheduler import (
    load_cron_config,
    save_cron_config,
    get_or_start_cron_daemon,
    execute_scheduled_brand_audit,
    CRON_STATUS_FILE
)

st.set_page_config(
    page_title="Instagram Creator & Paid Collabs Intelligence",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Start background cron daemon silently
daemon = get_or_start_cron_daemon()

# ── CUSTOM CSS STYLING ────────────────────────────────────────────────────────
st.markdown("""
<style>
    .metric-card {
        background-color: #F8F9FA;
        border-radius: 10px;
        padding: 18px;
        border-left: 5px solid #1B4F72;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .tier-badge-t1 { background-color: #D4EFDF; color: #145A32; font-weight: bold; padding: 4px 8px; border-radius: 4px; }
    .tier-badge-t2 { background-color: #EAFAF1; color: #1E8449; font-weight: bold; padding: 4px 8px; border-radius: 4px; }
    .tier-badge-t3 { background-color: #FEF9E7; color: #B7950B; font-weight: bold; padding: 4px 8px; border-radius: 4px; }
    .tier-badge-t4 { background-color: #F2F4F4; color: #566573; font-weight: normal; padding: 4px 8px; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR CONTROLS ──────────────────────────────────────────────────────────
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/e/e7/Instagram_logo_2016.svg", width=50)
st.sidebar.title("Brand Intelligence")
st.sidebar.caption("4-Tier Paid Hierarchy & Cron Automation Engine")

active_dataset = st.sidebar.selectbox(
    "Select Industry Dataset:",
    [
        "Footwear & Sneakers (Skechers, Gully Labs, Comet)",
        "Spiritual & Rudraksha (6 Brands - 2 Years)",
        "GRT Oriana (Oriana by GRT Jewellers - 2 Years)",
        "Consumer Electronics Retail (Croma & 10 Competitors)"
    ]
)

# ── LOAD SELECTED DATASET ─────────────────────────────────────────────────────
@st.cache_data
def load_dataset_by_selection(selection):
    if "Footwear" in selection:
        file_path = "footwear_1year_4tier_dataset.json"
        prof_path = "footwear_creators_profile_metrics.json"
        excel_path = "footwear_sneaker_brands_master_analysis.xlsx"
    elif "Spiritual" in selection:
        file_path = "spiritual_2year_4tier_dataset.json"
        prof_path = "spiritual_creators_profile_metrics.json"
        excel_path = "spiritual_rudraksha_brands_master_analysis.xlsx"
    elif "GRT Oriana" in selection:
        file_path = "grt_oriana_2year_4tier_dataset.json"
        prof_path = "grt_oriana_creators_profile_metrics.json"
        excel_path = "grt_oriana_2year_master_analysis.xlsx"
    else:
        file_path = "footwear_1year_4tier_dataset.json"
        prof_path = "footwear_creators_profile_metrics.json"
        excel_path = "footwear_sneaker_brands_master_analysis.xlsx"
        
    posts = []
    profiles = []
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            posts = json.load(f)
    if os.path.exists(prof_path):
        with open(prof_path, "r", encoding="utf-8") as f:
            profiles = json.load(f)
            
    return pd.DataFrame(posts), pd.DataFrame(profiles), excel_path

df_posts, df_profiles, master_excel_path = load_dataset_by_selection(active_dataset)

# ── HEADER & NAVIGATION TABS ──────────────────────────────────────────────────
st.title("🚀 Instagram Creator & Paid Collabs Dashboard")

tabs = st.tabs([
    "📊 Executive Summary",
    "⏰ Automated Cron Scheduler",
    "👥 Creator Sizing & Directory",
    "🏷️ 4-Tier Collaborations Explorer",
    "📥 Master Exports"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1: EXECUTIVE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.subheader(f"Executive Overview: {active_dataset}")
    
    if not df_posts.empty:
        tot_posts = len(df_posts)
        tot_creators = len(df_profiles)
        
        t1_cnt = len(df_posts[df_posts["tier"] == 1])
        t2_cnt = len(df_posts[df_posts["tier"] == 2])
        t3_cnt = len(df_posts[df_posts["tier"] == 3])
        t4_cnt = len(df_posts[df_posts["tier"] == 4])
        high_intent = t1_cnt + t2_cnt + t3_cnt
        high_intent_rate = (high_intent / tot_posts * 100) if tot_posts > 0 else 0
        
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Collab Posts", f"{tot_posts:,}")
        c2.metric("Unique Creators", f"{tot_creators:,}")
        c3.metric("💎 High-Intent Paid Ads", f"{high_intent:,}")
        c4.metric("High-Intent Adoption Rate", f"{high_intent_rate:.1f}%")
        c5.metric("⚪ Tier 4 Noise / Barter", f"{t4_cnt:,}")
        
        st.divider()
        
        col_left, col_right = st.columns([1, 1])
        
        with col_left:
            st.markdown("#### 🏛️ 4-Tier Partnership Breakdown")
            tier_df = pd.DataFrame([
                {"Tier": "🟢 Tier 1: Toggle ON + Boosted", "Count": t1_cnt},
                {"Tier": "🟢 Tier 2: Toggle ON + Organic", "Count": t2_cnt},
                {"Tier": "🚀 Tier 3: Toggle OFF + Boosted", "Count": t3_cnt},
                {"Tier": "⚪ Tier 4: Toggle OFF + Organic (Noise)", "Count": t4_cnt}
            ])
            st.bar_chart(tier_df.set_index("Tier"), color="#1B4F72", height=280)
            
        with col_right:
            st.markdown("#### 🎬 Top Video Content Genres")
            if "video_genre" in df_posts.columns:
                genre_counts = df_posts["video_genre"].value_counts().reset_index()
                genre_counts.columns = ["Video Genre", "Posts"]
                st.dataframe(genre_counts, use_container_width=True, height=280, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2: AUTOMATED CRON SCHEDULER
# ══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.subheader("⏰ Automated Cron Job & Scheduling Controls")
    st.info("Configure how often the system automatically scrapes your target brand roster, calculates 4-tiers, and updates the database.")
    
    cfg = load_cron_config()
    
    col_c1, col_c2 = st.columns([1, 1])
    
    with col_c1:
        st.markdown("### ⚙️ Schedule Configuration")
        cron_enabled = st.toggle("Enable Background Scheduled Cron Job", value=cfg.get("enabled", True))
        
        schedule_mode = st.selectbox(
            "Schedule Frequency:",
            ["Daily (Specified Time)", "Weekly (Specified Day & Time)", "Every X Hours", "Hourly"],
            index=0 if cfg.get("schedule_type") == "daily" else 1
        )
        
        if "Daily" in schedule_mode:
            sched_time = st.time_input("Run Every Day At:", datetime.strptime(cfg.get("scheduled_time", "09:00"), "%H:%M").time())
            cfg["schedule_type"] = "daily"
            cfg["scheduled_time"] = sched_time.strftime("%H:%M")
        elif "Weekly" in schedule_mode:
            wday = st.selectbox("Run Every Week On:", ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"], index=0)
            sched_time = st.time_input("At Time:", datetime.strptime(cfg.get("scheduled_time", "09:00"), "%H:%M").time())
            cfg["schedule_type"] = "weekly"
            cfg["weekday"] = wday.lower()
            cfg["scheduled_time"] = sched_time.strftime("%H:%M")
        elif "Every X Hours" in schedule_mode:
            hrs = st.number_input("Interval (Hours):", min_value=1, max_value=168, value=cfg.get("interval_hours", 24))
            cfg["schedule_type"] = "interval_hours"
            cfg["interval_hours"] = hrs
            
        cfg["enabled"] = cron_enabled
        
        if st.button("💾 Save & Apply Cron Schedule", type="primary"):
            save_cron_config(cfg)
            st.success("✅ Cron schedule configuration updated and applied successfully!")
            
    with col_c2:
        st.markdown("### 📡 Live Cron Status & Manual Trigger")
        
        # Load status
        status_msg = "Idle (Waiting for next scheduled run)"
        if os.path.exists(CRON_STATUS_FILE):
            try:
                with open(CRON_STATUS_FILE, "r") as f:
                    sdata = json.load(f)
                    status_msg = sdata.get("status_message", status_msg)
            except Exception:
                pass
                
        st.markdown(f"**Status**: `{status_msg}`")
        st.markdown(f"**Last Run**: `{cfg.get('last_run') or 'Never'}`")
        st.markdown(f"**Total Automatic Runs**: `{cfg.get('run_count', 0)}`")
        
        st.divider()
        st.markdown("#### ▶ Manual Immediate Run")
        if st.button("🚀 Trigger Full Audit Now (All Configured Brands)", type="secondary"):
            with st.spinner("Executing live audit across all configured brands..."):
                execute_scheduled_brand_audit()
            st.success("✅ Audit completed! Refreshing metrics...")
            time.sleep(1)
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3: CREATOR DIRECTORY & SIZING
# ══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.subheader("👥 Creator Roster & Audience Sizing Taxonomy")
    
    if not df_profiles.empty:
        col_f1, col_f2 = st.columns([1, 2])
        with col_f1:
            tier_filter = st.multiselect(
                "Filter by Creator Tier / Scale:",
                options=df_profiles["creator_tier"].unique() if "creator_tier" in df_profiles.columns else [],
                default=df_profiles["creator_tier"].unique() if "creator_tier" in df_profiles.columns else []
            )
        with col_f2:
            search_query = st.text_input("Search Creator Handle or Name:", "")
            
        filtered_profiles = df_profiles.copy()
        if tier_filter and "creator_tier" in filtered_profiles.columns:
            filtered_profiles = filtered_profiles[filtered_profiles["creator_tier"].isin(tier_filter)]
        if search_query:
            filtered_profiles = filtered_profiles[
                filtered_profiles["handle"].str.contains(search_query, case=False, na=False) |
                filtered_profiles["full_name"].str.contains(search_query, case=False, na=False)
            ]
            
        st.markdown(f"Showing **{len(filtered_profiles)}** of **{len(df_profiles)}** creators")
        
        cols_to_show = ["handle", "creator_tier", "brands", "full_name", "followers", "total_posts", "avg_likes", "avg_er"]
        avail_cols = [c for c in cols_to_show if c in filtered_profiles.columns]
        
        st.dataframe(
            filtered_profiles[avail_cols].style.format({"followers": "{:,.0f}", "total_posts": "{:,.0f}", "avg_likes": "{:,.0f}", "avg_er": "{:.2f}%"}),
            use_container_width=True,
            height=450,
            hide_index=True
        )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4: 4-TIER COLLABORATIONS EXPLORER
# ══════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.subheader("🏷️ All Collaboration Posts (4-Tier Hierarchy)")
    
    if not df_posts.empty:
        col_t1, col_t2 = st.columns([1, 2])
        with col_t1:
            t_sel = st.multiselect("Filter by Tier:", options=df_posts["tier_name"].unique(), default=df_posts["tier_name"].unique())
        with col_t2:
            b_sel = st.multiselect("Filter by Brand:", options=df_posts["brand"].unique(), default=df_posts["brand"].unique())
            
        filtered_posts = df_posts[df_posts["tier_name"].isin(t_sel) & df_posts["brand"].isin(b_sel)]
        st.markdown(f"Showing **{len(filtered_posts)}** collaboration posts")
        
        p_cols = ["tier_name", "brand", "handle", "followers", "video_genre", "views", "likes", "er_pct", "date", "url", "boost_status"]
        p_avail = [c for c in p_cols if c in filtered_posts.columns]
        
        st.dataframe(
            filtered_posts[p_avail].style.format({"followers": "{:,.0f}", "views": "{:,.0f}", "likes": "{:,.0f}", "er_pct": "{:.2f}%"}),
            use_container_width=True,
            height=450,
            hide_index=True
        )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5: MASTER EXPORTS
# ══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.subheader("📥 Master Reports & Deliverables")
    st.info("Download formatted Excel workbooks with colored tier styling or raw CSV exports.")
    
    if os.path.exists(master_excel_path):
        with open(master_excel_path, "rb") as f:
            excel_bytes = f.read()
        st.download_button(
            label=f"📗 Download Master Excel Workbook ({os.path.basename(master_excel_path)})",
            data=excel_bytes,
            file_name=os.path.basename(master_excel_path),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
    st.divider()
    st.markdown("#### 📄 Flat CSV Downloads:")
    c_csv1, c_csv2 = st.columns(2)
    with c_csv1:
        if not df_posts.empty:
            st.download_button("Download All Collabs CSV", df_posts.to_csv(index=False).encode("utf-8"), "all_collabs_4tier.csv", "text/csv")
    with c_csv2:
        if not df_profiles.empty:
            st.download_button("Download Creator Profiles CSV", df_profiles.to_csv(index=False).encode("utf-8"), "creator_profiles.csv", "text/csv")
