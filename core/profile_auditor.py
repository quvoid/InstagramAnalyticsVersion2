"""
================================================================================
CORE MODULE: LIVE PROFILE AUDITOR & METRICS RESOLVER
================================================================================
Fast and live verification of Instagram profiles:
  - Follower count (parsed to exact integer and human-readable string)
  - Full display name
  - Verification badge status
  - Contact email and phone from bio
  - External bio link
  - Creator Tier (Mega, Macro, Mid, Micro, Nano)
================================================================================
"""

import sys, os, re, json
from typing import Dict, Any

# Add project root to sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from api_wrapper.client import resolve_creator_profile

def parse_count(c_str):
    if not c_str: return 0
    s = str(c_str).strip().upper().replace(',', '')
    try:
        if s.endswith('M'): return int(float(s[:-1]) * 1_000_000)
        elif s.endswith('K'): return int(float(s[:-1]) * 1_000)
        elif s.endswith('B'): return int(float(s[:-1]) * 1_000_000_000)
        return int(float(s))
    except:
        return 0

EMAIL_REGEX = re.compile(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+')
PHONE_REGEX = re.compile(r'(?:\+91[\-\s]?)?[6789]\d{9}')


def audit_profile(handle: str) -> Dict[str, Any]:
    """Resolves an Instagram handle and returns a clean audit dict."""
    clean_h = handle.replace("@", "").strip().lower()
    
    # 1. First fast attempt using curl_cffi with chrome120
    data = resolve_creator_profile(clean_h)
    
    # Enrich tier string to remove emojis for clean compliance
    followers = data.get("followers", 0)
    if followers >= 1_000_000:
        tier = "Mega (1M+)"
    elif followers >= 500_000:
        tier = "Macro (500K-1M)"
    elif followers >= 100_000:
        tier = "Mid-Tier (100K-500K)"
    elif followers >= 10_000:
        tier = "Micro (10K-100K)"
    else:
        tier = "Nano (<10K)"

    data["tier"] = tier
    return data


def format_followers(num: int) -> str:
    if num >= 1_000_000:
        return f"{num/1_000_000:.1f}M".replace('.0M', 'M')
    elif num >= 1_000:
        return f"{num/1_000:.1f}K".replace('.0K', 'K')
    return str(int(num))


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "@rjpraveen"
    result = audit_profile(target)
    print("=" * 60)
    print(f"PROFILE AUDIT: {result['handle']}")
    print("=" * 60)
    print(f"Name:       {result.get('full_name', '')}")
    print(f"Followers:  {result.get('followers', 0):,} ({format_followers(result.get('followers', 0))})")
    print(f"Tier:       {result.get('tier', '')}")
    print(f"URL:        {result.get('profile_url', '')}")
    print("=" * 60)
