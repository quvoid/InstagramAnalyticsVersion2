"""
CLI tool for the Unified Competitor Intelligence API Wrapper.
Usage:
    python -m api_wrapper.cli audit --brand palmonas_official --range 3m
    python -m api_wrapper.cli insta --brand giva.co --range 1w
    python -m api_wrapper.cli adlib --query Palmonas
    python -m api_wrapper.cli fb --page zivame
"""

import sys, os, argparse, json
from .client import CompetitorIntelligenceClient

def main():
    parser = argparse.ArgumentParser(description="Competitor Intelligence CLI")
    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # Audit command
    audit_p = subparsers.add_parser("audit", help="Run 360-degree competitor audit")
    audit_p.add_argument("--brand", required=True, help="Instagram brand handle")
    audit_p.add_argument("--page-id", default=None, help="Meta Ad Library Page ID")
    audit_p.add_argument("--range", default="1y", help="Date range: '1w', '1m', '3m', '6m', '1y', '2y'")
    audit_p.add_argument("--no-excel", action="store_true", help="Disable Excel export")

    # Ad Library command
    adlib_p = subparsers.add_parser("adlib", help="Scan Meta Ad Library")
    adlib_p.add_argument("--query", required=True, help="Brand name or keyword")
    adlib_p.add_argument("--page-id", default=None, help="Meta Ad Library Page ID")
    adlib_p.add_argument("--active-only", action="store_true", help="Active ads only")

    # Instagram command
    insta_p = subparsers.add_parser("insta", help="Scan Instagram Partnerships")
    insta_p.add_argument("--brand", required=True, help="Instagram brand handle")
    insta_p.add_argument("--range", default="1y", help="Date range: '1w', '1m', '3m', '6m', '1y', '2y'")

    # Facebook command
    fb_p = subparsers.add_parser("fb", help="Scan Facebook Page info")
    fb_p.add_argument("--page", required=True, help="Facebook page handle")

    args = parser.parse_args()
    client = CompetitorIntelligenceClient()

    if args.command == "audit":
        res = client.audit_brand(
            target_brand=args.brand,
            fb_page_id=args.page_id,
            time_window=args.range,
            export_excel=not args.no_excel
        )
        print(f"\nAudit complete! Found {res['total_unique_creators']} unique creators across {res['time_window']} window.")
        if res.get("excel_file"):
            print(f"Master Excel Deliverable: {res['excel_file']}")

    elif args.command == "adlib":
        res = client.ad_library.search_ads(
            query=args.query,
            page_id=args.page_id,
            active_only=args.active_only
        )
        print(f"\nCaptured {res['total_ads_captured']} ads and {res['unique_creators_count']} creator partners.")

    elif args.command == "insta":
        res = client.instagram.get_partnerships(
            target_brand=args.brand,
            time_window=args.range
        )
        print(f"\nFound {res['total_collab_posts']} collab posts across {res['unique_creators_count']} creators ({res['time_window_requested']} window).")
        print(f"Cohorts: {res['date_range_cohorts']}")

    elif args.command == "fb":
        res = client.facebook.get_page_info(page_handle=args.page)
        print(json.dumps(res, indent=2))

    else:
        parser.print_help()

if __name__ == "__main__":
    main()
