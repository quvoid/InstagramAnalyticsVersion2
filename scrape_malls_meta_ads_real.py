"""
Scrape REAL Meta Ad Library Ads for Lake Shore and Competitor Malls
"""

import sys, json, time, re
from datetime import datetime, timezone
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")

mall_searches = [
    {"mall": "Phoenix Marketcity Pune", "query": "Phoenix Marketcity Pune", "city": "Pune"},
    {"mall": "Phoenix Mall of the Millennium Wakad", "query": "Phoenix Mall of the Millennium", "city": "Pune"},
    {"mall": "The Pavillion Pune", "query": "The Pavillion Pune", "city": "Pune"},
    {"mall": "Seasons Mall Pune", "query": "Seasons Mall Pune", "city": "Pune"},
    {"mall": "Amanora Mall Pune", "query": "Amanora Mall", "city": "Pune"},
    {"mall": "KOPA Mall Pune", "query": "Kopa Pune", "city": "Pune"},
    {"mall": "Lulu Mall Hyderabad", "query": "Lulu Mall Hyderabad", "city": "Hyderabad"},
    {"mall": "Nexus Hyderabad Mall", "query": "Nexus Hyderabad", "city": "Hyderabad"},
    {"mall": "Sarath City Capital Mall", "query": "Sarath City Capital Mall", "city": "Hyderabad"},
    {"mall": "Inorbit Mall Cyberabad", "query": "Inorbit Mall Hyderabad", "city": "Hyderabad"},
]

print("="*80)
print("SCRAPING REAL META AD LIBRARY ADS FOR MALL COMPETITORS")
print("="*80)

all_scraped_ads = []
seen_ad_ids = set()

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        viewport={"width": 1600, "height": 1000},
        locale="en-US",
        extra_http_headers={"Accept-Language": "en-US,en;q=0.9"}
    )
    page = context.new_page()

    js_extractor = r"""
    () => {
        const ads = [];
        const fullText = document.body.innerText || "";
        const blocks = fullText.split(/(?:Library ID:|लायब्ररी आयडी:)\s*(\d+)/i);
        
        for (let i = 1; i < blocks.length; i += 2) {
            const libId = blocks[i].trim();
            const chunk = blocks[i+1] || "";
            
            const isActive = chunk.includes("Active") || chunk.includes("सक्रिय");
            const dateMatch = chunk.match(/(?:Started running on|रोजी प्रसारण सुरू झाले)\s*([^\n·]+)/i);
            const startDate = dateMatch ? dateMatch[1].trim() : "";
            
            let advName = "";
            const advMatch = chunk.match(/(?:See (?:ad|summary) details|जाहिरात तपशील पहा)\n([^\n]+)\n(?:Sponsored|प्रायोजित)/i);
            if (advMatch) advName = advMatch[1].trim();
            
            let body = "";
            const spParts = chunk.split(/(?:Sponsored|प्रायोजित)/i);
            if (spParts.length > 1) {
                body = spParts[1].split(/(?:Shop Now|Learn More|Buy Now|Order Now|Get Offer|Book Now|See (?:ad|summary)|जाहिरात तपशील पहा)/i)[0].trim();
            }
            
            let cta = "Learn More";
            if (chunk.includes("Shop Now")) cta = "Shop Now";
            else if (chunk.includes("Book Now")) cta = "Book Now";
            else if (chunk.includes("Get Offer")) cta = "Get Offer";
            else if (chunk.includes("Sign Up")) cta = "Sign Up";
            
            ads.push({
                library_id: libId,
                ad_url: `https://www.facebook.com/ads/library/?id=${libId}`,
                advertiser: advName,
                is_active: isActive,
                start_date: startDate,
                cta: cta,
                body: body.substring(0, 500).replace(/\n/g, " ").trim()
            });
        }
        return ads;
    }
    """

    for m in mall_searches:
        mname = m["mall"]
        q = m["query"]
        print(f"\nScanning Meta Ad Library for: {mname} ('{q}')...")
        
        url = f"https://www.facebook.com/ads/library/?active_status=all&ad_type=all&country=IN&q={q.replace(' ', '%20')}&search_type=keyword_unordered&media_type=all"
        page.goto(url, timeout=40000, wait_until="domcontentloaded")
        time.sleep(5)
        
        # Scroll 6-8 times to collect ads
        for _ in range(6):
            page.evaluate("window.scrollBy(0, 2500)")
            time.sleep(1.5)
            
        ads_batch = page.evaluate(js_extractor)
        mall_ads_count = 0
        for ad in ads_batch:
            if ad["library_id"] not in seen_ad_ids:
                seen_ad_ids.add(ad["library_id"])
                ad["target_mall"] = mname
                ad["city"] = m["city"]
                all_scraped_ads.append(ad)
                mall_ads_count += 1
                
        print(f"  ✓ Captured {mall_ads_count} real Meta Ads for {mname}")

    browser.close()

with open("real_mall_meta_ads_dataset.json", "w", encoding="utf-8") as f:
    json.dump({
        "total_ads_captured": len(all_scraped_ads),
        "scraped_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
        "ads": all_scraped_ads
    }, f, indent=2)

print(f"\n✓ Saved {len(all_scraped_ads)} Real Meta Ads to real_mall_meta_ads_dataset.json")
