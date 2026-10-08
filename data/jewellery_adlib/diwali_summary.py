# -*- coding: utf-8 -*-
"""Content for the 'Diwali Summary' sheet - written from a read of every caption and transcript
in the Diwali 2025 window (30 Sep - 1 Nov 2025). Plain language, theme by theme, GRT vs competitors."""

R = "https://www.instagram.com/reel/"
P = "https://www.instagram.com/p/"

HEADLINE = [
    "GRT posted 33 creatives this Diwali. Most were festival greetings, product photos, its doctor-awards event and the "
    "MS Dhoni platinum launch. There was only ONE clear Diwali offer and almost no Diwali story.",
    "Competitors won on three things: a big, simple Diwali OFFER, a STORY that makes people feel something, and FACES - "
    "celebrities or popular creators carrying that offer or story.",
    "Thangamayil, GRT's closest Tamil Nadu rival, was the loudest brand of all: 107 creatives, several reels above 10 million "
    "views, all built around offers told by TV faces in Tamil.",
]

# (brand, creatives, main Diwali idea in plain words, biggest reel views, faces used)
GLANCE = [
    ("GRT Jewellers", 33, "Greetings, product photos, 'Golden Doctor' awards, MS Dhoni platinum edition (Men of Platinum). One creator Deepavali get-ready reel.", "Views not shown on GRT reels", "3 actresses in product photos; MS Dhoni (platinum)"),
    ("Thangamayil", 107, "'Deepa Velli' offer - buy gold, get silver free; Mega Chain Festival; Mega Savings Festival; Kanda Sashti devotion; Naya collection.", "22.6M", "Alya Manasa, Nadiya, Archana Chandhoke (TV faces)"),
    ("Malabar Gold", 60, "Lock today's gold rate with 10% advance; Diwali Collection 2025 product reels; store openings abroad.", "21.4M (New Zealand store)", "Anil Kapoor (store launch)"),
    ("Joyalukkas", 40, "Collection-by-collection product posts (Apurva, Pride, Eleganza uncut); one big Diwali brand film.", "18.5M (Diwali film)", "Samantha Ruth Prabhu"),
    ("Tanishq", 26, "One big story world - 'Mriganka' fantasy collection (palaces in the sky, mystical birds); celebrity styling; Bhai Dooj sibling reel.", "12.4M", "Manushi Chhillar, Jennifer Winget, Saba Pataudi"),
    ("Khazana", 30, "'Togetherness' theme - short family films in Telugu, Kannada and Tamil; Rs 505/gram off gold, 20% off diamonds & making.", "5.5M", "None - actors in short films"),
    ("Regal Jewellers", 34, "Store opening thank-you; Polki bridal; Tvisa milestones; Diwali wishes by stars.", "2.6M", "Radhika Pandit, Manju Warrier"),
    ("Lalithaa Jewellery", 16, "Festival greetings and product photo/carousels only - no reels, no offer, no faces.", "-", "None"),
    ("Bhima", 0, "Nothing posted on Instagram in this window.", "-", "-"),
    ("Akshaya Thangamaligai (ATM)", 0, "Nothing posted on Instagram in this window.", "-", "-"),
]

# (theme, what competitors did in simple words, best examples [(label, url)], what GRT did, what GRT can do)
THEMES = [
    ("1. A big, easy-to-remember Diwali offer",
     "Each strong brand had ONE offer that everyone could repeat. Thangamayil: 'Hands full of gold, bags full of silver' - "
     "silver free with every 10 g of gold (Deepa Velli). Khazana: Rs 505 off per gram + 20% off diamonds and making. "
     "Malabar: pay just 10% to lock today's gold rate. They repeated the same offer in many reels, photos and languages.",
     [("Thangamayil Deepa Velli - 21.7M views", R + "DPc3qqVDx3K/"), ("Thangamayil Deepa Velli - 21.8M", R + "DPlpRBYD9ps/"),
      ("Thangamayil Mega Savings Festival - 22.6M", R + "DQRlywaj2Gn/"), ("Malabar 10% advance rate-lock", R + "DPOoo42AHwp/"),
      ("Khazana Rs 505/gram offer", P + "DPanbNMkku-/")],
     "No headline Diwali offer in its Instagram posts - the only promotion was a 'sparkle and win' contest.",
     "Pick ONE Diwali offer with a catchy Tamil line and repeat it across 10-15 creatives, from Navaratri until Diwali day."),
    ("2. Famous faces carrying the offer",
     "Competitors put TV actresses and celebrities in front of the offer, not just the jewellery. Thangamayil used Tamil TV star "
     "Alya Manasa in skits (21M+ views). Joyalukkas used Samantha (4.9M). Regal used Radhika Pandit and Manju Warrier for Diwali wishes. "
     "Tanishq had Manushi Chhillar and Jennifer Winget style the collection (12.4M and 5M).",
     [("Thangamayil x Alya Manasa", R + "DPc3qqVDx3K/"), ("Joyalukkas x Samantha", R + "DP-4a6nkccu/"),
      ("Tanishq x Manushi Chhillar - 12.4M", R + "DP6U_UykryC/"), ("Tanishq x Jennifer Winget - 5M", R + "DP1asPukiZJ/"),
      ("Regal x Radhika Pandit", R + "DP_Mhrtk2W2/")],
     "Three actresses appeared only in still product photos (4 Oct). The one creator reel (Deepavali get-ready) was not paired with an offer.",
     "Use a well-known Tamil TV face or creator in short, funny or warm reels that SAY the offer out loud."),
    ("3. A Diwali story people feel",
     "Khazana made short family films - 'The finest decoration is togetherness', 'Her first Diwali' (a new bride) - and dubbed "
     "them into Telugu, Kannada and Tamil (up to 5.5M views each). Joyalukkas made one Diwali film ('What are you hiding? Give light "
     "to a new thought') that reached 18.5M. Tanishq made a sibling reel for Bhai Dooj.",
     [("Khazana - Togetherness (Telugu) 5.5M", R + "DPOAIU8jiY5/"), ("Khazana - Her First Diwali (Telugu) 5.2M", R + "DPWNItxDhxT/"),
      ("Khazana - Togetherness (Tamil)", R + "DPOAwwXDIf9/"), ("Joyalukkas Diwali film - 18.5M", R + "DP8cEunES0B/"),
      ("Tanishq Bhai Dooj sibling reel", R + "DQJ3-pYksCc/")],
     "Greeting reels (Lakshmi, Vijayadashami, Dhanteras, Deepavali wishes) - nice, but no story and no characters.",
     "Make 2-3 short Tamil family films (new bride's first Deepavali, mother passing gold to daughter) and end them with GRT."),
    ("4. One big 'world' for the collection",
     "Tanishq built its whole Diwali around one idea - 'Mriganka', a fantasy world of palaces in the sky and mystical birds. Every post, "
     "event and celebrity reel belonged to it, so the campaign felt big and premium.",
     [("Mriganka with Manushi - 12.4M", R + "DP6U_UykryC/"), ("Mriganka event", P + "DQBPX1LjtWq/"),
      ("Mriganka 3D pendant", R + "DQWmQ3zEkoY/")],
     "GRT's Diwali posts were a mix of separate topics (Dhoni, doctors, greetings, temple necklaces) with no single Diwali idea joining them.",
     "Give the Diwali season one name and one look, and make every post part of it."),
    ("5. Product showcases",
     "Everyone posted product reels and photos. Malabar posted almost daily (Diwali Collection 2025). Thangamayil showed collections by "
     "type (chains, studs under 4 g, bangles, antique malai, 'one jewel - five styles'). Joyalukkas named each collection (Apurva, Pride, Eleganza).",
     [("Thangamayil bangles - 6.3M", R + "DPSrVf8DLfo/"), ("Thangamayil men's chain - 14M", R + "DPOkRSjElvn/"),
      ("Thangamayil 'one jewel, five styles'", R + "DQO4SWsjK2i/"), ("Malabar Diwali Collection 2025", R + "DPQeL7iEZXW/")],
     "Temple and heritage product photos (Lord Shiva necklace, choker + haram) - good pieces, but mostly still photos, not reels.",
     "Turn product photos into short reels and add a useful angle: 'under 4 grams', 'one piece, many looks', price-from."),
    ("6. Light, everyday and men's jewellery",
     "Thangamayil pushed lightweight studs under 4 g, men's gold chains (14M) and a 'Family Chain' collection. Regal used Tvisa for "
     "everyday milestones.",
     [("Thangamayil studs under 4 g", R + "DPyfRE-jziO/"), ("Thangamayil men's chain - 14M", R + "DPOkRSjElvn/")],
     "MS Dhoni Signature Edition (platinum, via Men of Platinum) - a strong men's story, but platinum, not Diwali gold.",
     "Use the Dhoni partnership for a men's Diwali gifting angle, and add a lightweight-gold line for young buyers."),
    ("7. Devotion and local festivals",
     "Thangamayil rode Kanda Sashti / Murugan devotion right after Diwali (6M-7.7M views per reel) and sold Lord Shiva / deity rings. "
     "Regal greeted Kerala Piravi and Kannada Rajyotsava.",
     [("Thangamayil Murugan devotion - 7.7M", R + "DQLa2pLj771/"), ("Thangamayil Divine Ring Collection", R + "DP3QuHoj7jr/")],
     "Greetings for Lakshmi, Vijayadashami, Gandhi Jayanti, Dhanteras and Deepavali; a Lord Shiva necklace photo.",
     "Link devotional pieces to the festival day itself (Dhanteras buying, Kanda Sashti) with a reason to buy now."),
    ("8. Trust and gold-price help",
     "Thangamayil posted a 'gold awareness' explainer as prices rose. Malabar's rate-lock answered the same worry: 'what if gold gets dearer?'",
     [("Thangamayil gold awareness", R + "DPNysYMDwB_/"), ("Malabar rate-lock", R + "DPOpBKbk0lH/")],
     "'Golden Doctor' awards with IMA Coimbatore - good for trust and goodwill, but not tied to buying for Diwali.",
     "Add a simple gold-price explainer or rate-lock / scheme reel before Dhanteras."),
]

GRT_DID = [
    ("Festival greetings", "Lakshmi/Saraswati/Durga, Vijayadashami, Gandhi Jayanti, Dhanteras, Deepavali wishes", R + "DP81-zfiWqP/"),
    ("Actress product photos", "Pujita Ponnada, Hema Rajkumar, Sanjanaa Anand in gold sets (photos only)", P + "DPYzD4VCVGJ/"),
    ("Creator Deepavali reel", "Creator 'getting ready for Deepavali' with a GRT antique set (Tamil)", R + "DP3sxQokjqM/"),
    ("Contest", "'This Deepavali, sparkle and win' - post your festive look", R + "DP1RApDidPY/"),
    ("MS Dhoni Signature Edition", "Platinum men's line with Men of Platinum - 6 reels + 3 photos", R + "DQHIoW7CXYw/"),
    ("Golden Doctor awards", "Event with IMA Coimbatore honouring doctors - 6 posts", R + "DQdvi0QiSXE/"),
    ("Heritage product photos", "Temple jewellery, Lord Shiva necklace, choker + haram", P + "DQJzT3xieHV/"),
]

GAPS = [
    "No single Diwali offer that people could remember and repeat.",
    "No celebrity or creator saying the offer out loud in a reel.",
    "No Diwali story or short film - competitors' films reached 5M to 18M views.",
    "No one campaign idea joining the Diwali posts together.",
    "Few reels - many of GRT's Diwali creatives were still photos.",
    "No regional-language versions of the same film (Khazana dubbed into 3 languages).",
    "No lightweight / under-X-grams or price-led product reels for young buyers.",
]
