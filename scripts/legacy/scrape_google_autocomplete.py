"""
Module 1: Google Autocomplete & Search Intent Miner
Extracts Google Autocomplete suggestions across all 12 malls with intent classification,
keyword difficulty/intent score, and negative keyword mapping for Google Ads.
"""

import sys, json, time, urllib.parse, urllib.request
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8")

MALLS = [
    {"name": "KOPA Mall Pune", "short": "KOPA Pune", "city": "Pune", "brand": "KOPA", "is_client": True},
    {"name": "Phoenix Marketcity Pune", "short": "Phoenix Marketcity", "city": "Pune", "brand": "Phoenix Avenue of Stars", "is_client": False},
    {"name": "Phoenix Mall of the Millennium", "short": "Phoenix Millennium", "city": "Pune", "brand": "Phoenix Millennium Wakad", "is_client": False},
    {"name": "The Pavillion Mall Pune", "short": "The Pavillion", "city": "Pune", "brand": "The Pavillion", "is_client": False},
    {"name": "Seasons Mall Pune", "short": "Seasons Mall", "city": "Pune", "brand": "Seasons Mall", "is_client": False},
    {"name": "Amanora Mall Pune", "short": "Amanora Mall", "city": "Pune", "brand": "Amanora Mall", "is_client": False},
    {"name": "Lake Shore Y Junction", "short": "LS Y Junction", "city": "Hyderabad", "brand": "Lake Shore Y Junction", "is_client": True},
    {"name": "Lulu Mall Hyderabad", "short": "Lulu Mall", "city": "Hyderabad", "brand": "Lulu Mall", "is_client": False},
    {"name": "Nexus Hyderabad Mall", "short": "Nexus Mall", "city": "Hyderabad", "brand": "Nexus Hyderabad", "is_client": False},
    {"name": "Sarath City Capital Mall", "short": "Sarath City", "city": "Hyderabad", "brand": "Sarath City Capital", "is_client": False},
    {"name": "Inorbit Mall Cyberabad", "short": "Inorbit Mall", "city": "Hyderabad", "brand": "Inorbit Cyberabad", "is_client": False},
    {"name": "GVK One Mall Hyderabad", "short": "GVK One", "city": "Hyderabad", "brand": "GVK One", "is_client": False}
]

# Query prefixes / modifiers to capture all buyer intents
INTENT_MODIFIERS = [
    "", # Direct brand search
    "brands",
    "stores",
    "parking",
    "parking charges",
    "restaurants",
    "food court",
    "movie",
    "cinema",
    "pvr",
    "timings",
    "entry fee",
    "offers",
    "sales",
    "zara",
    "sephora",
    "valet",
    "reviews",
    "owner",
    "how to reach",
    "metro station"
]

def fetch_google_autocomplete(query):
    encoded = urllib.parse.quote(query)
    url = f"https://suggestqueries.google.com/complete/search?client=chrome&q={encoded}&hl=en&gl=in"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode('utf-8'))
            suggestions = data[1] if len(data) > 1 else []
            return suggestions
    except Exception as e:
        # Fallback to firefox endpoint
        url_ff = f"https://suggestqueries.google.com/complete/search?client=firefox&q={encoded}&hl=en&gl=in"
        try:
            req_ff = urllib.request.Request(url_ff, headers=headers)
            with urllib.request.urlopen(req_ff, timeout=5) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                return data[1] if len(data) > 1 else []
        except Exception:
            return []

def classify_intent(query_str):
    q = query_str.lower()
    if any(k in q for k in ["parking", "valet", "car park", "two wheeler", "bike parking", "parking charges", "parking fee"]):
        return "🚗 Parking & Accessibility (Conquest Hook)"
    elif any(k in q for k in ["brand", "store", "zara", "h&m", "sephora", "mango", "armani", "tira", "uniqlo", "starbucks", "apple", "shopping"]):
        return "🛍️ Tenant Brand Search (High Purchase Intent)"
    elif any(k in q for k in ["food", "restaurant", "cafe", "dining", "rooftop", "bar", "buffet", "court", "eat", "zomato"]):
        return "🍽️ F&B & Dining Intent (High AOV)"
    elif any(k in q for k in ["movie", "cinema", "pvr", "imax", "theatre", "ticket", "show", "screen", "director"]):
        return "🎬 Entertainment & Cinema Intent"
    elif any(k in q for k in ["timing", "open", "close", "time", "hours", "sunday", "holiday"]):
        return "🕒 Operational / Visit Planning"
    elif any(k in q for k in ["reach", "metro", "bus", "location", "address", "route", "direction", "where"]):
        return "📍 Navigation & Commute Intent"
    elif any(k in q for k in ["offer", "sale", "discount", "eoss", "deal", "cheap", "price"]):
        return "🏷️ Bargain & Promotional Search"
    elif any(k in q for k in ["owner", "developer", "area", "sq ft", "jobs", "rent", "leasing"]):
        return "🏢 B2B & Commercial / Leasing"
    else:
        return "🔍 General Brand Awareness"

print("="*80)
print("SCRAPING GOOGLE AUTOCOMPLETE & SEARCH INTENT ACROSS 12 MALLS")
print("="*80)

results = []
all_suggestions_unique = set()
intent_breakdown = defaultdict(int)
mall_intent_map = defaultdict(lambda: defaultdict(list))

for mall in MALLS:
    m_name = mall["name"]
    print(f"Scraping search queries for {m_name}...")
    mall_queries_set = set()
    
    for mod in INTENT_MODIFIERS:
        search_term = f"{m_name} {mod}".strip()
        suggestions = fetch_google_autocomplete(search_term)
        time.sleep(0.15) # respectful delay
        
        for sugg in suggestions:
            if sugg not in mall_queries_set:
                mall_queries_set.add(sugg)
                intent = classify_intent(sugg)
                intent_breakdown[intent] += 1
                mall_intent_map[m_name][intent].append(sugg)
                
                # Assign bidding strategy recommendation
                if "Parking" in intent:
                    rec_action = "🔴 Bid on Phrase Match: Highlight Zero-Wait Valet at KOPA"
                    bid_prio = "HIGH CONQUEST"
                elif "Tenant" in intent:
                    rec_action = "🟢 Exact Match: Dynamic Tenant Directory Ad with Store Logos"
                    bid_prio = "HIGH INTENT"
                elif "Dining" in intent:
                    rec_action = "🟡 Dayparted Match (Thu-Sat): Rooftop & Table Booking CTA"
                    bid_prio = "HIGH CONVERSION"
                elif "B2B" in intent or "Bargain" in intent:
                    rec_action = "⛔ Add as Negative Keyword (Avoid wasted ad spend)"
                    bid_prio = "NEGATIVE LIST"
                else:
                    rec_action = "🔵 Broad Modified: General Catchment Brand Traffic"
                    bid_prio = "MEDIUM"

                results.append({
                    "mall_name": m_name,
                    "city": mall["city"],
                    "is_client": mall["is_client"],
                    "seed_search": search_term,
                    "autocomplete_query": sugg,
                    "intent_category": intent,
                    "bid_priority": bid_prio,
                    "google_ads_recommended_action": rec_action
                })

print(f"\n✓ Captured {len(results)} distinct Google Autocomplete search queries across 12 malls.")

output_file = "google_autocomplete_intent_dataset.json"
with open(output_file, "w", encoding="utf-8") as f:
    json.dump({
        "total_queries_captured": len(results),
        "intent_distribution": dict(intent_breakdown),
        "queries": results
    }, f, ensure_ascii=False, indent=2)

print(f"✓ Saved to {output_file}")
