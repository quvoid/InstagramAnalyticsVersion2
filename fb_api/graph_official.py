"""
graph_official.py -- thin wrapper over Meta's OFFICIAL Graph API, for the
cases where the sanctioned endpoint is the right tool.

Read this before reaching for it, because the official Ad Library API almost
certainly does NOT do what this project needs:

    Meta's own ads_archive reference says ad_reached_countries is required and
    that "ads not reaching EU locations only return if categorized as social
    issues, elections or politics".

    In plain terms: for India -- every brand in this project -- ads_archive
    returns POLITICAL ADS ONLY. Searching it for Zivame, Underneat, Mamaearth
    or any other D2C brand returns nothing, no matter how good your token is.
    That is not a bug or a permissions problem, it is the documented scope.

So: use scan_adlibrary.py (the Ad Library the public actually browses) for
Indian commercial ads. Use this module when you want
  * EU/UK-delivered ads, where commercial ads ARE in the archive post-DSA,
    with eu_total_reach and targeting fields attached, or
  * political/issue ads anywhere, with spend and impression ranges, or
  * Page fields via a token you already hold for a Page you administer.

Access, per Meta's docs: a Meta developer app, identity+location confirmation
at facebook.com/ID, and a user access token. Tokens are long-lived (~60 days),
not permanent.

Field availability, straight from the ArchivedAd reference:
  all ads          id, ad_creation_time, ad_creative_bodies,
                   ad_creative_link_captions/descriptions/titles,
                   ad_delivery_start_time, ad_delivery_stop_time,
                   ad_snapshot_url, languages, page_id, page_name,
                   publisher_platforms, total_reach_by_location
  EU only          eu_total_reach, beneficiary_payers
  EU/UK + BR pol.  age_country_gender_reach_breakdown, target_ages,
                   target_gender, target_locations
  political only   spend, impressions, currency, bylines,
                   demographic_distribution, delivery_by_region,
                   estimated_audience_size
"""
import os
import json

from fb_session import FacebookError, new_session

GRAPH_VERSION = os.environ.get("GRAPH_API_VERSION", "v23.0")
GRAPH_BASE = f"https://graph.facebook.com/{GRAPH_VERSION}"

DEFAULT_AD_FIELDS = [
    "id",
    "ad_creation_time",
    "ad_creative_bodies",
    "ad_creative_link_captions",
    "ad_creative_link_descriptions",
    "ad_creative_link_titles",
    "ad_delivery_start_time",
    "ad_delivery_stop_time",
    "ad_snapshot_url",
    "languages",
    "page_id",
    "page_name",
    "publisher_platforms",
]

DEFAULT_PAGE_FIELDS = [
    "id",
    "name",
    "username",
    "link",
    "category",
    "fan_count",
    "followers_count",
    "talking_about_count",
    "verification_status",
    "about",
    "website",
    "is_published",
]


class GraphError(FacebookError):
    pass


def _require_token(access_token):
    token = access_token or os.environ.get("FB_GRAPH_TOKEN", "")
    if not token:
        raise GraphError(
            "no Graph API token. Set FB_GRAPH_TOKEN in .env (Meta developer app + identity "
            "confirmation at facebook.com/ID), or use the Ad Library endpoints instead -- "
            "those need no token."
        )
    return token


def _get(path, params, access_token):
    session = new_session()
    params = dict(params)
    params["access_token"] = _require_token(access_token)
    r = session.get(f"{GRAPH_BASE}/{path}", params=params, timeout=45)
    try:
        payload = r.json()
    except json.JSONDecodeError:
        raise GraphError(f"non-JSON response from Graph API (HTTP {r.status_code}): {r.text[:200]}")
    if isinstance(payload, dict) and payload.get("error"):
        err = payload["error"]
        raise GraphError(f"Graph API error {err.get('code')}: {err.get('message')}")
    return payload


def ads_archive(search_terms=None, search_page_ids=None, countries=("IN",),
                ad_type="ALL", ad_active_status="ACTIVE", ad_delivery_date_min=None,
                ad_delivery_date_max=None, media_type=None, publisher_platforms=None,
                search_type=None, languages=None, fields=None, limit=100,
                max_pages=10, access_token=None) -> dict:
    """Query the official ads_archive endpoint, following paging cursors."""
    params = {
        "ad_reached_countries": json.dumps(list(countries)),
        "ad_type": ad_type,
        "ad_active_status": ad_active_status,
        "fields": ",".join(fields or DEFAULT_AD_FIELDS),
        "limit": limit,
    }
    if search_terms:
        params["search_terms"] = search_terms
    if search_page_ids:
        params["search_page_ids"] = json.dumps([str(p) for p in search_page_ids][:10])
    if search_type:
        params["search_type"] = search_type
    if ad_delivery_date_min:
        params["ad_delivery_date_min"] = ad_delivery_date_min
    if ad_delivery_date_max:
        params["ad_delivery_date_max"] = ad_delivery_date_max
    if media_type:
        params["media_type"] = media_type
    if publisher_platforms:
        params["publisher_platforms"] = json.dumps(list(publisher_platforms))
    if languages:
        params["languages"] = json.dumps(list(languages))

    ads = []
    pages = 0
    payload = _get("ads_archive", params, access_token)
    while True:
        ads.extend(payload.get("data") or [])
        pages += 1
        after = ((payload.get("paging") or {}).get("cursors") or {}).get("after")
        if not after or pages >= max_pages or not (payload.get("paging") or {}).get("next"):
            break
        payload = _get("ads_archive", {**params, "after": after}, access_token)

    non_eu = [c for c in countries if c.upper() not in EU_EEA_UK]
    note = None
    if non_eu and ad_type == "ALL" and not ads:
        note = (f"0 results for {', '.join(non_eu)}. Expected: outside the EU/UK the official "
                f"archive only holds political and social-issue ads. Use "
                f"GET /api/v1/adlibrary/{{brand}} for commercial ads in these countries.")

    return {
        "source": "graph_ads_archive",
        "graph_version": GRAPH_VERSION,
        "countries": list(countries),
        "ad_type": ad_type,
        "count": len(ads),
        "pages_fetched": pages,
        "ads": ads,
        "note": note,
    }


EU_EEA_UK = {
    "AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE",
    "IT", "LV", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE",
    "IS", "LI", "NO", "GB",
}


def page_fields(page_id_or_username, fields=None, access_token=None) -> dict:
    """Read a Page node.

    Most useful fields (fan_count, followers_count, about, website) need either
    a Page access token for a Page you administer, or the Page Public Content
    Access / Page Public Metadata Access features on your app -- both require
    app review. If you just want public brand numbers, GET /api/v1/page/{handle}
    scrapes them with no token and no review.
    """
    payload = _get(str(page_id_or_username),
                   {"fields": ",".join(fields or DEFAULT_PAGE_FIELDS)},
                   access_token)
    return {"source": "graph_page", "graph_version": GRAPH_VERSION, "page": payload}
