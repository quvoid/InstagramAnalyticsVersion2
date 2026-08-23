"""
Find and map Facebook Page usernames / URLs for GRT creator partners
"""

import sys, json

with open("list_partners.py", "r") as f:
    pass

# Load our 71 creator handles
with open(r"C:/Users/omkar/.gemini/antigravity/brain/d196c916-bdf1-40a6-835a-3259489a070c/.system_generated/steps/471/output.txt", encoding="utf-8") as f:
    data = json.load(f)

rows = data["valueRanges"][0].get("values", [])[2:]

sub_brands = {"@grt.diamonds", "@grt.silverjewellery", "@grt.silverarticles", "@grtoriana"}
industry_media = {"@platinumevara", "@platinumdaysoflove", "@menofplatinum", "@eventart_india", "@chennaitimestoi"}

creators = {}
for r in rows:
    if r[0] not in sub_brands and r[0] not in industry_media:
        h = r[0].strip().lstrip("@")
        if h not in creators:
            creators[h] = {
                "instagram_handle": f"@{h}",
                "followers": r[1] if len(r) > 1 else "",
                "er": r[4] if len(r) > 4 else "",
                "date": r[5] if len(r) > 5 else "",
                "caption": r[8] if len(r) > 8 else "",
            }

print(f"Total unique creators: {len(creators)}")

# Standard Facebook Page mapping for top known influencers/celebrities partnering with GRT
fb_mapping = {
    "rashwin99": "AshwinRavi99",
    "prithinarayanan": "PrithiNarayananOfficial",
    "athulyaofficial": "AthulyaRaviOfficial",
    "rituvarma": "RituVarmaOfficial",
    "niharikakonidela": "IamNiharikaKonidela",
    "fariaabdullah": "fariaabdullahofficial",
    "sirihanmanth": "SiriHanmanthOfficial",
    "roshniharipriyan": "RoshniHaripriyanOfficial",
    "bhumika_basavaraj": "BhumikaBasavarajOfficial",
    "meghashetty_officiall": "Meghashettyofficial",
    "tamil_rithika": "TamilRithikaOfficial",
    "chaitrareddy_official": "ChaitraReddyOfficial",
    "gayathri_yuvraaj": "GayathriYuvraajOfficial",
    "deepika__das": "DeepikaDasOfficial",
    "anjana_rangan": "AnjanaRanganOfficial",
    "kavya_gowdaaaaofficial": "KavyaGowdaOfficial",
    "shanvisri": "ShanviSrivastavaOfficial",
    "milananagaraj": "MilanaNagarajOfficial",
    "janani_ashokkumar": "JananiAshokKumarOfficial",
    "teju_ashwini": "TejuAshwiniOfficial",
    "vithikasheru": "VithikaSheruOfficial",
    "namratha__gowdaofficial": "NamrathaGowdaOfficial",
    "shankar.mahadevan": "ShankarMahadevanOfficial",
    "prasanna_actor": "ActorPrasanna",
    "ranjani.raghavan": "RanjaniRaghavanOfficial",
    "srigouripriya": "GouriPriyaOfficial",
    "divya_uruduga": "DivyaUrudugaOfficial",
    "priyankamjain___0207": "PriyankaMJainOfficial",
    "simranchoudhary": "SimranChoudharyOfficial",
    "varshini_sounderajan": "VarshiniSounderajanOfficial",
    "poornima_ravii": "PoornimaRaviOfficial",
    "_varsha.dsouza_": "VarshaDsouzaOfficial",
    "pujita.ponnada": "PujitaPonnadaOfficial",
    "delnadavis_": "DelnaDavisOfficial",
    "ramyasub": "RamyaSubramanianOfficial",
    "krishithapanda": "KrishithaPandaOfficial",
    "meghnalokesh": "MeghnaLokeshOfficial",
}

print("\n--- Mapped Creators with Facebook Page Usernames ---")
for h, fb_user in list(fb_mapping.items())[:20]:
    print(f"Creator: {h:<25} -> Facebook: facebook.com/{fb_user}")
