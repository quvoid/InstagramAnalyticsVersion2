"""
Unified FastAPI REST Server for Instagram, Facebook, and Meta Ad Library Intelligence.
Run:
    uvicorn api_wrapper.server:app --reload --port 8080
"""

import os, sys
from typing import Optional, Dict, Any, List
from fastapi import FastAPI, Query, HTTPException, Header, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from .client import CompetitorIntelligenceClient

app = FastAPI(
    title="Unified Competitor & Paid Media Intelligence API",
    description="Enterprise REST API for Instagram Creator Collabs, Facebook Pages, and Meta Ad Library Dark Ads.",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

client = CompetitorIntelligenceClient()

@app.get("/")
def root():
    return {
        "status": "online",
        "service": "Unified Competitor Intelligence API",
        "endpoints": {
            "docs": "/docs",
            "instagram_profile": "/api/v1/instagram/profile/{username}",
            "instagram_partnerships": "/api/v1/instagram/partnerships/{brand_username}",
            "facebook_page": "/api/v1/facebook/page/{page_handle}",
            "meta_ad_library": "/api/v1/meta/ads?query={brand}",
            "full_audit": "/api/v1/intelligence/audit/{brand_username}"
        }
    }

# 1. Instagram Profile Metrics
@app.get("/api/v1/instagram/profile/{username}")
def get_instagram_profile(username: str):
    try:
        return client.instagram.get_profile_metrics(username)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# 2. Instagram 1-Year Partnerships & Boost Hierarchy
@app.get("/api/v1/instagram/partnerships/{brand_username}")
def get_instagram_partnerships(
    brand_username: str,
    days_back: int = Query(default=365, ge=1, le=730),
    max_pages: int = Query(default=25, ge=1, le=50)
):
    try:
        return client.instagram.get_partnerships(brand_username, days_back=days_back, max_pages=max_pages)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# 3. Facebook Page Metrics
@app.get("/api/v1/facebook/page/{page_handle}")
def get_facebook_page(page_handle: str):
    try:
        return client.facebook.get_page_info(page_handle)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# 4. Meta Ad Library Search
@app.get("/api/v1/meta/ads")
def search_meta_ads(
    query: str = Query(..., description="Brand name or keyword"),
    page_id: Optional[str] = Query(default=None, description="Optional Meta Ad Library Page ID"),
    active_only: bool = Query(default=False),
    max_scrolls: int = Query(default=30, ge=1, le=50)
):
    try:
        return client.ad_library.search_ads(query=query, page_id=page_id, active_only=active_only, max_scrolls=max_scrolls)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# 5. Full-Funnel Unified Competitor Audit
@app.get("/api/v1/intelligence/audit/{brand_username}")
def audit_competitor_brand(
    brand_username: str,
    page_id: Optional[str] = Query(default=None),
    days_back: int = Query(default=365, ge=1, le=730),
    export_excel: bool = Query(default=True)
):
    try:
        return client.audit_brand(
            target_brand=brand_username,
            fb_page_id=page_id,
            days_back=days_back,
            export_excel=export_excel
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
