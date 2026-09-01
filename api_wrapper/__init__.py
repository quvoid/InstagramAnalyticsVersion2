"""
Unified Competitor Analytics & Paid Media API Wrapper
Combines Instagram API, Facebook Page API, and Meta Ad Library (GraphQL + Playwright)
"""

from .client import CompetitorIntelligenceClient, BrandAnalyticsClient

__version__ = "2.0.0"
__all__ = ["CompetitorIntelligenceClient", "BrandAnalyticsClient"]
