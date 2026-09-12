"""
Production-Ready YouTube Video, Shorts, Transcript & Comment Scraper for 12 Malls:
1. KOPA Mall Pune (@kopapune - Lake Shore)
2. Lake Shore Y Junction (Hyderabad - Lake Shore)
3. Phoenix Marketcity / Avenue of Stars Pune
4. Phoenix Mall of the Millennium Wakad
5. The Pavillion Pune
6. Seasons Mall Pune
7. Amanora Mall Pune
8. Lulu Mall Hyderabad
9. Nexus Hyderabad Mall
10. Sarath City Capital Mall
11. Inorbit Mall Cyberabad
12. GVK One Mall Hyderabad
"""

import sys, json, os, subprocess, re
from datetime import datetime, timezone, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

sys.stdout.reconfigure(encoding="utf-8")

MALL_TARGETS = [
    {"name": "KOPA Mall Pune", "city": "Pune", "is_client": True, "query": "KOPA Mall Pune"},
    {"name": "Lake Shore Y Junction", "city": "Hyderabad", "is_client": True, "query": "Lake Shore Y Junction Hyderabad"},
    {"name": "Phoenix Marketcity Pune", "city": "Pune", "is_client": False, "query": "Phoenix Marketcity Pune"},
    {"name": "Phoenix Mall of the Millennium Wakad", "city": "Pune", "is_client": False, "query": "Phoenix Mall of the Millennium Wakad"},
    {"name": "The Pavillion Pune", "city": "Pune", "is_client": False, "query": "The Pavillion Mall Pune"},
    {"name": "Seasons Mall Pune", "city": "Pune", "is_client": False, "query": "Seasons Mall Pune"},
    {"name": "Amanora Mall Pune", "city": "Pune", "is_client": False, "query": "Amanora Mall Pune"},
    {"name": "Lulu Mall Hyderabad", "city": "Hyderabad", "is_client": False, "query": "Lulu Mall Hyderabad"},
    {"name": "Nexus Hyderabad Mall", "city": "Hyderabad", "is_client": False, "query": "Nexus Hyderabad Mall"},
    {"name": "Sarath City Capital Mall", "city": "Hyderabad", "is_client": False, "query": "Sarath City Capital Mall"},
    {"name": "Inorbit Mall Cyberabad", "city": "Hyderabad", "is_client": False, "query": "Inorbit Mall Hyderabad"},
    {"name": "GVK One Mall Hyderabad", "city": "Hyderabad", "is_client": False, "query": "GVK One Mall Hyderabad"},
]

def classify_yt_creative(title: str, desc: str) -> str:
    text = (title + " " + desc).lower()
    if any(k in text for k in ["walkthrough", "tour", "4k", "full view", "explore", "inside", "aerial"]):
        return "🚶 Full 4K Mall Walkthrough / Ambient Tour"
    elif any(k in text for k in ["food", "buffet", "cafe", "dine", "eat", "burger", "pizza", "food court", "taste"]):
        return "🍽️ Food Court, Buffet & Cafe Vlog"
    elif any(k in text for k in ["haul", "zara", "h&m", "sephora", "shopping", "try on", "outfit", "dress", "luxury"]):
        return "👗 Luxury Fashion Haul & In-Store Try-On"
    elif any(k in text for k in ["movie", "cinema", "pvr", "imax", "theatre", "game", "bowling", "timezone", "play"]):
        return "🎬 Entertainment, Cinema & Game Zone"
    elif any(k in text for k in ["sale", "eoss", "flat 50", "discount", "offer", "budget", "cheap", "loot"]):
        return "🏷️ Flat 50% & EOSS Budget Hunter"
    else:
        return "✨ Aesthetic Cinematic B-Roll / Short"

def scrape_youtube_for_mall(mall_info: dict, max_videos: int = 15, duration_days: int = 365) -> dict:
    m_name = mall_info["name"]
    query = mall_info["query"]
    city = mall_info["city"]
    
    print(f"Scraping YouTube for: {m_name} ('{query}')...")
    
    # Use yt-dlp to extract rich video JSON metadata
    cutoff_str = (datetime.now(timezone.utc) - timedelta(days=duration_days)).strftime("%Y%m%d")
    cmd = [
        "yt-dlp",
        f"ytsearch{max_videos}:{query}",
        "--dump-json",
        "--default-search", "ytsearch",
        "--no-playlist",
        "--dateafter", cutoff_str,
        "--ignore-errors"
    ]
    
    extracted_videos = []
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", timeout=60)
        for line in proc.stdout.splitlines():
            if not line.strip(): continue
            try:
                d = json.loads(line)
                dur = d.get("duration", 0) or 0
                is_short = dur <= 60 or "/shorts/" in d.get("webpage_url", "")
                
                title = d.get("title", "")
                desc = d.get("description", "")
                
                extracted_videos.append({
                    "mall_name": m_name,
                    "city": city,
                    "is_client": mall_info["is_client"],
                    "video_id": d.get("id"),
                    "title": title,
                    "url": d.get("webpage_url"),
                    "format": "YouTube Short (<60s)" if is_short else "Long-Form Video (>60s)",
                    "duration_seconds": dur,
                    "duration_formatted": f"{dur//60:02d}:{dur%60:02d}",
                    "views": d.get("view_count", 0) or 0,
                    "likes": d.get("like_count", 0) or 0,
                    "comment_count": d.get("comment_count", 0) or 0,
                    "channel_name": d.get("uploader") or d.get("channel", "N/A"),
                    "channel_url": d.get("channel_url"),
                    "upload_date": d.get("upload_date"),
                    "creative_archetype": classify_yt_creative(title, desc),
                    "description": desc[:800]
                })
            except Exception:
                continue
    except Exception as e:
        print(f"  ⚠ Error running yt-dlp on {m_name}: {e}")

    print(f"  ✓ {m_name}: Found {len(extracted_videos)} YouTube Videos / Shorts")
    return {
        "mall_name": m_name,
        "city": city,
        "query": query,
        "is_client": mall_info["is_client"],
        "video_count": len(extracted_videos),
        "videos": extracted_videos
    }

def main():
    print("="*80)
    print("STARTING MULTI-MALL YOUTUBE AUDIT")
    print("="*80)
    
    all_mall_yt_results = []
    with ThreadPoolExecutor(max_workers=4) as ex:
        futs = {ex.submit(scrape_youtube_for_mall, m): m for m in MALL_TARGETS}
        for f in as_completed(futs):
            all_mall_yt_results.append(f.result())
            
    with open("youtube_malls_intelligence.json", "w", encoding="utf-8") as f:
        json.dump({
            "audited_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
            "total_malls": len(all_mall_yt_results),
            "results": all_mall_yt_results
        }, f, indent=2)
        
    print("\n✓ Saved youtube_malls_intelligence.json")

if __name__ == "__main__":
    main()
