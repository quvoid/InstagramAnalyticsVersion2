"""
Classify Video Genre & Content Format for all 147 Footwear Collaboration Posts
"""

import sys, json, re
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

with open("footwear_1year_4tier_dataset.json", encoding="utf-8") as f:
    posts = json.load(f)

# Genre classification rules based on creator niche, caption NLP, and keywords
def detect_video_genre(handle, caption, brand):
    h = handle.lower()
    c = caption.lower()
    
    # 1. Celebrity & Ambassador Campaign Films
    if any(k in h for k in ["kartikaaryan", "ananyapanday", "surya_14kumar", "alayaf", "dianapenty"]):
        if any(w in c for w in ["run", "match", "game", "cricket", "train", "sport"]):
            return "⚡ Athlete & Sports Performance"
        return "🌟 Celebrity Ambassador Campaign"
        
    # 2. Sports & Athletic Performance
    if any(w in c for w in ["cricket", "run", "marathon", "workout", "fitness", "training", "athlete", "game day", "pitch", "performance"]):
        return "⚡ Athlete & Sports Performance"
    if any(k in h for k in ["skecherscricket", "niviasports", "akshay_freestyle", "capt_dvs", "parmar_shailesh"]):
        return "⚡ Athlete & Sports Performance"
        
    # 3. Comedy & Relatable Entertainment Skits
    if any(k in h for k in ["thevishnukaushal", "itskivitime", "bhaiyajiismile", "cartoon_snacks", "harsh_zii"]):
        return "🎬 Comedy & Relatable Skit"
    if any(w in c for w in ["pov:", "when you", "relatable", "funny", "skit", "tag someone", "me trying to", "comedy"]):
        return "🎬 Comedy & Relatable Skit"
        
    # 4. Sneaker Unboxing & Product Review
    if any(w in c for w in ["unboxing", "unbox", "review", "first look", "first impression", "materials", "packaging", "detailed look", "hands on"]):
        return "📦 Unboxing & Sneaker Review"
    if any(k in h for k in ["smartkicksofficial", "winonakicks", "sneakergeekin", "thesneakerfreak", "kicks"]):
        return "📦 Unboxing & Sneaker Review"
        
    # 5. Styling, OOTD & GRWM (Fashion)
    if any(w in c for w in ["grwm", "styling", "style", "outfit", "ootd", "lookbook", "fit check", "how to style", "wear it with", "drip", "wardrobe"]):
        return "👗 Styling & OOTD / GRWM"
    if any(k in h for k in ["harshiitarora", "yashvi.bhaia", "sameoldseth", "amra.ldn", "kanchhiiii", "evaxfried", "anuj.mp4", "shreyasharma"]):
        return "👗 Styling & OOTD / GRWM"
        
    # 6. Brand-to-Brand Collab & Special Edition Drops
    if any(k in h for k in ["unoindiaofficial", "farakwear", "myntra", "amazonfashionin", "bombaysweetshop", "casabacardiin", "cmf.tech", "royalenfield"]):
        return "🤝 Brand Collab Drop / Co-Creation"
    if any(w in c for w in ["collab", "collaboration", "exclusive drop", "limited edition", "collection drop", "capsule collection", "co-created"]):
        return "🤝 Brand Collab Drop / Co-Creation"
        
    # 7. Event, Pop-Up & Festival Activations
    if any(w in c for w in ["event", "festival", "pop up", "popup", "launch party", "live at", "exhibition", "weekend", "booth", "store visit"]):
        return "🎉 Event & Pop-Up Activation"
    if any(k in h for k in ["indiansneakerfestival", "delhiartweekend", "districtupdates", "thethirdspacedelhi", "kommunedelhincr"]):
        return "🎉 Event & Pop-Up Activation"
        
    # 8. Design Craft, Streetwear Art & Storytelling
    if any(w in c for w in ["crafted", "design", "handmade", "story behind", "details", "heritage", "concept", "art", "aesthetic"]):
        return "🎨 Streetwear Design & Craft Storytelling"
    if any(k in h for k in ["anikjaindesign", "art", "design", "creative"]):
        return "🎨 Streetwear Design & Craft Storytelling"
        
    # Default to Lifestyle / Culture
    return "👟 Streetwear Lifestyle & Culture"

print("Sample Video Genre Classification on Top 20 Footwear Posts:\n")

for p in posts[:20]:
    g = detect_video_genre(p["handle"], p["caption"], p["brand"])
    p["video_genre"] = g
    print(f"[{g:<35}] {p['brand']:<16} | {p['handle']:<22} | Caption: {p['caption'][:60]}...")

for p in posts:
    p["video_genre"] = detect_video_genre(p["handle"], p["caption"], p["brand"])

gc = Counter(p["video_genre"] for p in posts)
print("\n" + "="*70)
print("FOOTWEAR 1-YEAR VIDEO GENRE BREAKDOWN (147 Posts):")
print("="*70)
for k, v in gc.most_common():
    print(f"  • {k:<40}: {v:>3} videos ({v/len(posts)*100:>4.1f}%)")
print("="*70)
