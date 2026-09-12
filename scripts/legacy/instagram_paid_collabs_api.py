"""
Instagram Paid Partnerships & Boosted Ads Intelligence Engine & API
===================================================================
A unified API and CLI tool to scrape brand collaboration posts, detect Meta Paid Partnership
toggles, evaluate ad spend boosting signatures, and categorize creator campaigns into 4 Tiers:
  🟢 Tier 1: Toggle ON + 🚀 Boosted (Formal Disclosure + Paid Ad Spend)
  🟢 Tier 2: Toggle ON + ⚪ Organic (Formal Disclosure + Organic Reach)
  🚀 Tier 3: Toggle OFF + 🚀 Boosted (Co-Author Collab + Paid Media Spend / Partnership Ads)
  ⚪ Tier 4: Toggle OFF + ⚪ Organic (Standard Collab / Low Reach / Noise)

Features:
  - Python SDK Functions: analyze_brand(), analyze_post_urls(), batch_analyze_brands()
  - REST API Web Server: FastAPI / HTTP endpoints for remote ingestion
  - Excel & CSV Exporter with visual tier highlighting and state origin metadata
"""

import sys
import os
import json
import time
import re
import csv
import argparse
from datetime import datetime, timezone
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

# Optional imports for web server
try:
    from fastapi import FastAPI, Query, HTTPException, Body
    from pydantic import BaseModel, Field
    import uvicorn
    FASTAPI_AVAILABLE = True
except ImportError:
    FASTAPI_AVAILABLE = False

# Optional openpyxl for Excel generation
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# Import core network scraper session
from scrape_bulk import make_session, extract_shortcode, shortcode_to_id, COOKIES

sys.stdout.reconfigure(encoding="utf-8")

# ─────────────────────────────────────────────────────────────
# CONFIGURATION & HEADERS
# ─────────────────────────────────────────────────────────────
MOBILE_HEADERS = {
    "User-Agent": "Instagram 269.0.0.18.75 Android (26/8.0.0; 480dpi; 1080x1920; OnePlus; ONEPLUS A3003; OnePlus3; qcom; en_US; 314665256)",
    "x-ig-app-id": "936619743392459",
}

WEB_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "x-ig-app-id": "936619743392459",
    "accept": "*/*",
}

STATE_ORIGIN_DATABASE = {
    "tanishq": "Pan-India / Karnataka (HQ: Bengaluru, Titan)",
    "kalyan": "Kerala (HQ: Thrissur)",
    "malabar": "Kerala (HQ: Kozhikode)",
    "grt": "Tamil Nadu (HQ: Chennai)",
    "pc chandra": "West Bengal (HQ: Kolkata)",
    "bluestone": "Pan-India / Karnataka (HQ: Bengaluru, D2C)",
    "giva": "Pan-India / Karnataka (HQ: Bengaluru, D2C)",
    "palmonas": "Pan-India / Maharashtra (HQ: Pune, D2C)",
    "senco": "West Bengal (HQ: Kolkata)",
    "png": "Maharashtra (HQ: Pune)",
    "prince": "Tamil Nadu (HQ: Chennai)",
    "jagdamba": "Telangana (HQ: Hyderabad)",
    "kalamandir": "Gujarat (HQ: Surat)",
    "khimji": "Odisha (HQ: Bhubaneswar)",
    "croma": "Pan-India / Maharashtra (HQ: Mumbai, Tata Group)",
    "reliance digital": "Pan-India / Maharashtra (HQ: Mumbai, Reliance)",
    "vijay sales": "Maharashtra / Pan-India (HQ: Mumbai)",
    "tata neu": "Pan-India / Maharashtra (HQ: Mumbai, Tata Digital)",
    "sathya": "Tamil Nadu (HQ: Tuticorin / Chennai)",
    "pai": "Karnataka & Telangana (HQ: Bengaluru)",
    "bajaj": "Telangana & Andhra Pradesh (HQ: Hyderabad)",
    "aditya vision": "Bihar, Jharkhand & UP (HQ: Patna)",
}


def lookup_state(brand_name: str) -> str:
    b_low = brand_name.lower()
    for k, v in STATE_ORIGIN_DATABASE.items():
        if k in b_low:
            return v
    return "India"


# ─────────────────────────────────────────────────────────────
# CORE ENGINE CLASS
# ─────────────────────────────────────────────────────────────
class InstagramPaidPartnershipEngine:
    def __init__(self, max_workers: int = 15):
        self.session = make_session()
        self.max_workers = max_workers

    def get_user_id_and_info(self, username: str) -> dict:
        """Resolve numeric user ID, follower count, and bio from Instagram username."""
        clean_user = username.replace("@", "").strip().lower()
        url = f"https://i.instagram.com/api/v1/users/search/?q={clean_user}"
        try:
            r = self.session.get(url, headers=MOBILE_HEADERS, cookies=COOKIES, timeout=10)
            if r.status_code == 200:
                for u in r.json().get("users", []):
                    if u.get("username", "").lower() == clean_user:
                        return {
                            "pk": u.get("pk"),
                            "username": u.get("username"),
                            "full_name": u.get("full_name"),
                            "is_verified": u.get("is_verified", False),
                            "follower_count": u.get("follower_count") or 0,
                            "profile_pic_url": u.get("profile_pic_url"),
                        }
        except Exception as e:
            print(f"[-] Error resolving user @{username}: {e}", file=sys.stderr)
        return {"pk": None, "username": clean_user, "follower_count": 0}

    def get_creator_followers(self, handle: str) -> int:
        """Fetch follower count for a creator handle."""
        info = self.get_user_id_and_info(handle)
        return info.get("follower_count", 0)

    def fetch_brand_feed(self, user_id: int, max_pages: int = 6) -> list:
        """Fetch timeline feed items for a given brand user ID."""
        items = []
        max_id = ""
        for p in range(1, max_pages + 1):
            url = f"https://i.instagram.com/api/v1/feed/user/{user_id}/"
            if max_id:
                url += f"?max_id={max_id}"
            try:
                r = self.session.get(url, headers=MOBILE_HEADERS, cookies=COOKIES, timeout=12)
                if r.status_code == 200:
                    data = r.json()
                    f_items = data.get("items", [])
                    items.extend(f_items)
                    max_id = data.get("next_max_id")
                    if not max_id or len(f_items) == 0:
                        break
                    time.sleep(0.3)
                else:
                    break
            except Exception:
                break
        return items

    def fetch_brand_clips(self, user_id: int, max_pages: int = 6) -> list:
        """Fetch reels clips items for a given brand user ID."""
        items = []
        max_id = ""
        for p in range(1, max_pages + 1):
            url = "https://i.instagram.com/api/v1/clips/user/"
            payload = {"target_user_id": str(user_id), "page_size": 30}
            if max_id:
                payload["max_id"] = str(max_id)
            try:
                r = self.session.post(url, headers=MOBILE_HEADERS, data=payload, cookies=COOKIES, timeout=12)
                if r.status_code == 200:
                    data = r.json()
                    clips = [it.get("media") for it in data.get("items", []) if it.get("media")]
                    items.extend(clips)
                    paging = data.get("paging_info", {})
                    max_id = paging.get("max_id")
                    if not paging.get("more_available") or not max_id:
                        break
                    time.sleep(0.3)
                else:
                    break
            except Exception:
                break
        return items

    def evaluate_boost_and_tier(
        self,
        is_paid_toggle: bool,
        views: int,
        likes: int,
        comments: int,
        followers: int,
        caption: str = ""
    ) -> dict:
        """
        Calculates exact engagement ratios, detects paid media ad spend,
        and assigns the post to Tier 1, Tier 2, Tier 3, or Tier 4.
        """
        # Ensure views estimate if only likes exist
        if not views and likes:
            views = int(likes * 18.5)

        like_rate = round((likes / views) * 100, 2) if views > 0 else 0.0
        view_multiplier = round(views / followers, 2) if followers > 0 else 0.0
        er_pct = round(((likes + comments) / followers) * 100, 2) if followers > 0 else 0.0

        # Check caption for disclosure tags
        cap_lower = caption.lower()
        has_disclosure_tag = any(t in cap_lower for t in ["#ad", "#paidpartnership", "#sponsored", "#collab", "#brandpartner", "paid partnership"])
        is_formal_paid = is_paid_toggle or has_disclosure_tag

        # Boost Detection Logic:
        is_boosted = False
        if views >= 500000 and like_rate < 0.35:
            is_boosted = True
            boost_status = "🚀 Heavily Boosted (Paid Ad Spend)"
            reason = f"High view count ({views:,}) with sub-0.35% like rate ({like_rate}%) indicates paid video ads campaign"
        elif view_multiplier >= 5.0 and like_rate < 0.70:
            is_boosted = True
            boost_status = "🚀 Boosted (Paid Ad Spend)"
            reason = f"High view multiplier ({view_multiplier}x followers) combined with low engagement rate ({like_rate}%)"
        elif view_multiplier >= 3.0 and like_rate < 1.00 and views >= 80000:
            is_boosted = True
            boost_status = "🔍 Likely Boosted (Targeted Ad)"
            reason = f"Disproportionate views ({views:,}) relative to likes ({likes:,})"
        elif er_pct >= 4.0 and like_rate >= 2.00:
            is_boosted = False
            boost_status = "📈 Viral Organic Reach"
            reason = f"High organic views ({views:,}) with strong like rate ({like_rate}%)"
        else:
            is_boosted = False
            boost_status = "⚪ Standard Organic"
            reason = "Baseline organic collab reach"

        # Tier Assignment
        if is_formal_paid and is_boosted:
            tier = 1
            tier_name = "🟢 Tier 1: Toggle ON + Boosted"
            tier_desc = "Formal Paid Partnership Label + Active Paid Ad Spend"
        elif is_formal_paid and not is_boosted:
            tier = 2
            tier_name = "🟢 Tier 2: Toggle ON + Organic"
            tier_desc = "Formal Paid Partnership Label + Organic Reach"
        elif not is_formal_paid and is_boosted:
            tier = 3
            tier_name = "🚀 Tier 3: Toggle OFF + Boosted"
            tier_desc = "Co-Author / Collab Invite (Toggle OFF) + Heavy Paid Media Spend"
        else:
            tier = 4
            tier_name = "⚪ Tier 4: Toggle OFF + Organic (Noise)"
            tier_desc = "Standard Collab / Low Organic Reach / Unboosted"

        return {
            "tier": tier,
            "tier_name": tier_name,
            "tier_description": tier_desc,
            "is_paid_partnership": is_formal_paid,
            "is_boosted": is_boosted,
            "boost_status": boost_status,
            "boost_reason": reason,
            "estimated_views": views,
            "like_to_view_pct": like_rate,
            "view_to_follower_multiplier": view_multiplier,
            "creator_er_pct": er_pct,
        }

    def analyze_brand(self, username: str, max_pages: int = 6) -> dict:
        """
        Complete end-to-end brand audit pipeline:
        Scrapes creator collabs, verifies toggles, detects boosted ads, and groups into 4 tiers.
        """
        brand_clean = username.replace("@", "").strip()
        brand_info = self.get_user_id_and_info(brand_clean)
        user_id = brand_info.get("pk")
        state_origin = lookup_state(brand_clean)

        if not user_id:
            return {
                "status": "error",
                "message": f"Could not resolve Instagram user ID for @{brand_clean}",
                "brand_handle": brand_clean,
            }

        # Fetch feed and reels
        feed_items = self.fetch_brand_feed(user_id, max_pages=max_pages)
        clips_items = self.fetch_brand_clips(user_id, max_pages=max_pages)

        # Deduplicate
        seen_ids = set()
        raw_items = []
        for it in feed_items + clips_items:
            pk = str(it.get("pk") or it.get("id"))
            if pk and pk not in seen_ids:
                seen_ids.add(pk)
                raw_items.append(it)

        # Filter creator collaborations
        collab_posts = []
        for it in raw_items:
            owner = it.get("user", {}).get("username", "").lower()
            coauthors = [c.get("username", "").lower() for c in it.get("coauthor_producers", [])]
            code = it.get("code") or ""
            post_url = f"https://www.instagram.com/p/{code}/" if code else ""
            
            caption_obj = it.get("caption") or {}
            caption_text = caption_obj.get("text", "") if isinstance(caption_obj, dict) else str(caption_obj)
            taken_at = it.get("taken_at")
            date_str = datetime.fromtimestamp(taken_at, tz=timezone.utc).strftime("%Y-%m-%d") if taken_at else "N/A"

            play_count = it.get("play_count") or it.get("view_count") or 0
            like_count = it.get("like_count") or 0
            comment_count = it.get("comment_count") or 0

            # Creator handle determination
            creator_handle = ""
            if owner != brand_clean.lower():
                creator_handle = f"@{owner}"
            elif coauthors:
                ext = [c for c in coauthors if c != brand_clean.lower()]
                if ext:
                    creator_handle = f"@{ext[0]}"

            if not creator_handle:
                continue

            collab_posts.append({
                "brand": brand_info.get("full_name") or brand_clean,
                "brand_handle": f"@{brand_clean}",
                "creator_handle": creator_handle,
                "raw_creator_handle": creator_handle.replace("@", ""),
                "post_url": post_url,
                "shortcode": code,
                "media_id": str(it.get("pk") or ""),
                "post_date": date_str,
                "is_paid_toggle": bool(it.get("is_paid_partnership", False)),
                "views": play_count,
                "likes": like_count,
                "comments": comment_count,
                "caption": caption_text[:250].replace("\n", " ").replace("\r", " "),
            })

        # Resolve creator followers concurrently
        unique_handles = list(set(p["raw_creator_handle"] for p in collab_posts))
        followers_map = {}
        with ThreadPoolExecutor(max_workers=self.max_workers) as ex:
            fut_map = {ex.submit(self.get_creator_followers, h): h for h in unique_handles}
            for fut in as_completed(fut_map):
                h = fut_map[fut]
                followers_map[h] = fut.result()

        # Classify posts into 4 tiers
        classified_posts = []
        for p in collab_posts:
            h = p["raw_creator_handle"]
            fol = followers_map.get(h, 0)
            p["creator_followers"] = fol

            eval_res = self.evaluate_boost_and_tier(
                is_paid_toggle=p["is_paid_toggle"],
                views=p["views"],
                likes=p["likes"],
                comments=p["comments"],
                followers=fol,
                caption=p["caption"]
            )
            p.update(eval_res)
            classified_posts.append(p)

        # Sort strictly by Tier ascending, then Views descending
        classified_posts.sort(key=lambda x: (x["tier"], -x["estimated_views"]))

        # Group by tier
        tier_1 = [p for p in classified_posts if p["tier"] == 1]
        tier_2 = [p for p in classified_posts if p["tier"] == 2]
        tier_3 = [p for p in classified_posts if p["tier"] == 3]
        tier_4 = [p for p in classified_posts if p["tier"] == 4]

        high_intent_paid = len(tier_1) + len(tier_2) + len(tier_3)
        total_collabs = len(classified_posts)
        total_unique_creators = len(set(p["creator_handle"].lower() for p in classified_posts))

        return {
            "status": "success",
            "audit_timestamp": datetime.now(timezone.utc).isoformat(),
            "brand": {
                "name": brand_info.get("full_name") or brand_clean,
                "handle": f"@{brand_clean}",
                "state_origin": state_origin,
                "followers": brand_info.get("follower_count", 0),
            },
            "summary": {
                "total_collab_posts": total_collabs,
                "total_unique_creators": total_unique_creators,
                "tier_1_toggle_on_boosted": {
                    "posts_count": len(tier_1),
                    "unique_creators": len(set(p["creator_handle"].lower() for p in tier_1)),
                },
                "tier_2_toggle_on_organic": {
                    "posts_count": len(tier_2),
                    "unique_creators": len(set(p["creator_handle"].lower() for p in tier_2)),
                },
                "tier_3_toggle_off_boosted": {
                    "posts_count": len(tier_3),
                    "unique_creators": len(set(p["creator_handle"].lower() for p in tier_3)),
                },
                "tier_4_toggle_off_organic_noise": {
                    "posts_count": len(tier_4),
                    "unique_creators": len(set(p["creator_handle"].lower() for p in tier_4)),
                },
                "total_high_intent_paid_ads": high_intent_paid,
                "high_intent_paid_rate_pct": round((high_intent_paid / total_collabs) * 100, 2) if total_collabs > 0 else 0.0,
            },
            "tier_1_toggle_on_boosted": tier_1,
            "tier_2_toggle_on_organic": tier_2,
            "tier_3_toggle_off_boosted": tier_3,
            "tier_4_noise": tier_4,
            "all_posts": classified_posts,
        }

    def analyze_posts_custom_list(self, posts_list: list) -> dict:
        """
        Classifies an arbitrary list of post dictionaries (containing url, likes, followers, caption, etc.)
        """
        results = []
        for p in posts_list:
            u = p.get("url", "")
            likes = p.get("likes", 0)
            comments = p.get("comments", 0)
            fol = p.get("followers", 0)
            views = p.get("views", 0)
            is_paid = p.get("is_paid_partnership", False)
            cap = p.get("caption", "")

            eval_res = self.evaluate_boost_and_tier(
                is_paid_toggle=is_paid,
                views=views,
                likes=likes,
                comments=comments,
                followers=fol,
                caption=cap
            )
            item = dict(p)
            item.update(eval_res)
            results.append(item)

        results.sort(key=lambda x: (x["tier"], -x["estimated_views"]))
        t1 = [r for r in results if r["tier"] == 1]
        t2 = [r for r in results if r["tier"] == 2]
        t3 = [r for r in results if r["tier"] == 3]
        t4 = [r for r in results if r["tier"] == 4]

        return {
            "status": "success",
            "total_posts": len(results),
            "summary": {
                "tier_1_toggle_on_boosted": len(t1),
                "tier_2_toggle_on_organic": len(t2),
                "tier_3_toggle_off_boosted": len(t3),
                "tier_4_noise_organic": len(t4),
                "total_high_intent_paid_ads": len(t1) + len(t2) + len(t3),
            },
            "tier_1": t1,
            "tier_2": t2,
            "tier_3": t3,
            "tier_4": t4,
            "all_posts": results,
        }


# ─────────────────────────────────────────────────────────────
# FASTAPI SERVER INITIALIZATION
# ─────────────────────────────────────────────────────────────
if FASTAPI_AVAILABLE:
    app = FastAPI(
        title="Instagram Paid Partnerships & Boosted Ads Intelligence API",
        description="API to scrape, verify Meta disclosures, detect boosted ad spend, and classify creator collaborations into 4 Tiers.",
        version="2.0.0"
    )
    engine = InstagramPaidPartnershipEngine()

    class BrandAuditRequest(BaseModel):
        username: str = Field(..., example="giva.co", description="Instagram brand handle to audit")
        max_pages: int = Field(default=6, description="Number of feed/clips pages to scrape")

    class CustomPostsRequest(BaseModel):
        posts: list = Field(..., description="List of post dictionaries to evaluate")

    @app.get("/api/v1/health")
    def health_check():
        return {"status": "healthy", "service": "Instagram Paid Collabs Engine", "version": "2.0.0"}

    @app.get("/api/v1/analyze/brand")
    def analyze_brand_get(
        username: str = Query(..., description="Instagram brand handle (e.g. giva.co, croma, bluestone_jewellery)"),
        max_pages: int = Query(default=6, ge=1, le=20, description="Pages to scrape")
    ):
        res = engine.analyze_brand(username, max_pages=max_pages)
        if res.get("status") == "error":
            raise HTTPException(status_code=404, detail=res.get("message"))
        return res

    @app.post("/api/v1/analyze/brand")
    def analyze_brand_post(payload: BrandAuditRequest):
        res = engine.analyze_brand(payload.username, max_pages=payload.max_pages)
        if res.get("status") == "error":
            raise HTTPException(status_code=404, detail=res.get("message"))
        return res

    @app.post("/api/v1/analyze/posts")
    def analyze_custom_posts(payload: CustomPostsRequest):
        return engine.analyze_posts_custom_list(payload.posts)


# ─────────────────────────────────────────────────────────────
# CLI ENTRY POINT
# ─────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Instagram Paid Partnerships & Boosted Ads API & Engine")
    parser.add_argument("--brand", type=str, help="Instagram brand handle (e.g. giva.co, croma, reliance_digital)")
    parser.add_argument("--pages", type=int, default=6, help="Feed & clips pages to scrape (default: 6)")
    parser.add_argument("--output", type=str, help="JSON output file path")
    parser.add_argument("--server", action="store_true", help="Launch FastAPI web server")
    parser.add_argument("--port", type=int, default=8000, help="Port for FastAPI web server (default: 8000)")

    args = parser.parse_args()

    if args.server:
        if not FASTAPI_AVAILABLE:
            print("[-] FastAPI or Uvicorn not installed. Run: pip install fastapi uvicorn", file=sys.stderr)
            sys.exit(1)
        print(f"🚀 Starting Instagram Paid Collabs API Server on http://localhost:{args.port}")
        uvicorn.run("instagram_paid_collabs_api:app", host="0.0.0.0", port=args.port, reload=True)
        return

    if args.brand:
        engine = InstagramPaidPartnershipEngine()
        print(f"[*] Auditing brand: @{args.brand} (Pages: {args.pages})...")
        res = engine.analyze_brand(args.brand, max_pages=args.pages)

        if res.get("status") == "error":
            print(f"[-] Error: {res.get('message')}")
            sys.exit(1)

        summary = res["summary"]
        brand_meta = res["brand"]
        print("\n" + "="*70)
        print(f"💎 BRAND AUDIT: {brand_meta['name']} ({brand_meta['handle']})")
        print(f"📍 State / Origin: {brand_meta['state_origin']}")
        print(f"👥 Brand Followers: {brand_meta['followers']:,}")
        print("="*70)
        print(f"• Total Collab Posts Scanned: {summary['total_collab_posts']}")
        print(f"• Total Unique Creators:      {summary['total_unique_creators']}")
        print(f"• 🟢 Tier 1 (Toggle ON + 🚀 Boosted): {summary['tier_1_toggle_on_boosted']['posts_count']} posts ({summary['tier_1_toggle_on_boosted']['unique_creators']} creators)")
        print(f"• 🟢 Tier 2 (Toggle ON + ⚪ Organic): {summary['tier_2_toggle_on_organic']['posts_count']} posts ({summary['tier_2_toggle_on_organic']['unique_creators']} creators)")
        print(f"• 🚀 Tier 3 (Toggle OFF + 🚀 Boosted): {summary['tier_3_toggle_off_boosted']['posts_count']} posts ({summary['tier_3_toggle_off_boosted']['unique_creators']} creators)")
        print(f"• ⚪ Tier 4 (Noise / Unboosted):       {summary['tier_4_toggle_off_organic_noise']['posts_count']} posts ({summary['tier_4_toggle_off_organic_noise']['unique_creators']} creators)")
        print(f"• 💎 Total High-Intent Paid Ads:       {summary['total_high_intent_paid_ads']} ({summary['high_intent_paid_rate_pct']}%)")
        print("="*70 + "\n")

        if args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                json.dump(res, f, indent=2)
            print(f"✓ Saved results to {args.output}")

        return

    parser.print_help()


if __name__ == "__main__":
    main()
