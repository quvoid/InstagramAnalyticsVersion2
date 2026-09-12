"""
Update CSVs with State/Region Information
"""

import sys, json, csv

sys.stdout.reconfigure(encoding="utf-8")

STATE_MAP = {
    "Malabar Gold & Diamonds": "Kerala (HQ: Kozhikode)",
    "GIVA Jewellery": "Pan-India / Karnataka (HQ: Bengaluru, D2C)",
    "Palmonas": "Pan-India / Maharashtra (HQ: Pune, D2C)",
    "GRT Jewellers": "Tamil Nadu (HQ: Chennai)",
    "P. C. Chandra Jewellers": "West Bengal (HQ: Kolkata)",
    "BlueStone": "Pan-India / Karnataka (HQ: Bengaluru, D2C)",
    "Tanishq": "Pan-India / Karnataka (HQ: Bengaluru, Titan)",
    "Sri Jagdamba Pearls": "Telangana (HQ: Hyderabad)",
    "Kalamandir Jewellers": "Gujarat (HQ: Surat)",
    "Senco Gold & Diamonds": "West Bengal (HQ: Kolkata)",
    "Khimji Jewellers": "Odisha (HQ: Bhubaneswar)",
    "C.H. Jewellers": "Gujarat (HQ: Vadodara)",
    "Anopchand Tilokchand (AT Jewellers)": "Chhattisgarh (HQ: Raipur)",
    "Anopchand Tilokchand (AT Jewell": "Chhattisgarh (HQ: Raipur)",
    "P. N. Gadgil Jewellers (PNG)": "Maharashtra (HQ: Pune)",
    "Lalithaa Jewellery": "Tamil Nadu (HQ: Chennai)",
    "Aisshpra Gems & Jewels": "Uttar Pradesh (HQ: Gorakhpur)",
    "Amrapali Jewels": "Rajasthan (HQ: Jaipur)",
    "Anjali Jewellers": "West Bengal (HQ: Kolkata)",
    "Bhima Jewellers": "Karnataka / Kerala (HQ: Bengaluru)",
    "Birdhichand Ghanshyamdas": "Rajasthan (HQ: Jaipur)",
    "C. Krishniah Chetty & Co (CKC)": "Karnataka (HQ: Bengaluru)",
    "DP Abhushan (DP Jewellers)": "Madhya Pradesh (HQ: Ratlam)",
    "Hazoorilal Legacy": "Delhi NCR (HQ: New Delhi)",
    "Joyalukkas": "Kerala (HQ: Thrissur)",
    "Kalyan Jewellers": "Kerala (HQ: Thrissur)",
    "Kashi Jewellers": "Uttar Pradesh (HQ: Kanpur)",
    "Khanna Jewellers": "Delhi NCR (HQ: New Delhi)",
    "Manik Chand Jewellers": "Assam / Meghalaya (HQ: Guwahati)",
    "Mangatrai Pearls & Jewellers": "Telangana (HQ: Hyderabad)",
    "Motisons Jewellers": "Rajasthan (HQ: Jaipur)",
    "Navrathan Jewellers": "Karnataka (HQ: Bengaluru)",
    "Nikka Mal Pyare Lal": "Punjab (HQ: Ludhiana)",
    "Prince Jewellery": "Tamil Nadu (HQ: Chennai)",
    "TBZ - The Original": "Maharashtra (HQ: Mumbai)",
    "Vaibhav Jewellers": "Andhra Pradesh (HQ: Visakhapatnam)",
    "Waman Hari Pethe (WHP)": "Maharashtra (HQ: Mumbai)",
}

def get_state(brand_name):
    for k, v in STATE_MAP.items():
        if k.lower() in brand_name.lower() or brand_name.lower() in k.lower():
            return v
    return "India"

with open("unified_master_dataset.json", encoding="utf-8") as f:
    master_records = json.load(f)

master_records.sort(key=lambda x: (x["tier"], -x["views"]))

# Write CSV with State column
with open("All_Brands_4Tier_Master.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow([
        "#", "Hierarchy Tier", "Brand Name", "State / Origin (HQ)", "Creator Handle", "Followers",
        "Views / Plays", "Likes", "Comments", "Like-to-View %", "Creator ER%",
        "Post Date", "Direct Instagram URL", "Boost Status & Classification", "Caption Preview"
    ])
    for idx, p in enumerate(master_records, 1):
        st = get_state(p["brand"])
        w.writerow([
            idx,
            p["tier_name"],
            p["brand"],
            st,
            p["handle"],
            p["followers"],
            p["views"],
            p["likes"],
            p["comments"],
            f"{p['like_rate_pct']:.2f}%",
            f"{p['er_pct']:.2f}%",
            p["post_date"],
            p["url"],
            f"{p['boost_status']}: {p['reason']}" if p['reason'] else p['boost_status'],
            p["caption"]
        ])

print("✓ Updated All_Brands_4Tier_Master.csv with State column")
