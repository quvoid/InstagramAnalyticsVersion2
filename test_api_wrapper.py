"""
Test script for api_wrapper
"""

import sys, json
from api_wrapper import CompetitorIntelligenceClient

sys.stdout.reconfigure(encoding="utf-8")

client = CompetitorIntelligenceClient()

print("1. Testing Instagram Profile Metrics...")
ig_prof = client.instagram.get_profile_metrics("palmonas_official")
print("   Instagram Profile:", ig_prof)

print("\n2. Testing Facebook Page Info...")
fb_page = client.facebook.get_page_info("Palmonas")
print("   Facebook Page:", fb_page)

print("\n✓ api_wrapper tests passed successfully!")
