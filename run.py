"""
================================================================================
INSTAGRAM & OMNICHANNEL MARKETING INTELLIGENCE: NATURAL LANGUAGE DISPATCHER
================================================================================
Allows any user, AI agent, or external developer to execute repository tools,
scrapers, audits, and workbook exporters using natural language queries.

Usage:
  python run.py "<natural language prompt>"
  python run.py "audit profile @rjpraveen"
  python run.py "find kolkata creators 10k"
  python run.py "scan fintech brands 2 years"
  python run.py "list master workbooks"
  python run.py "show catalog"
================================================================================
"""

import sys, os, re, json, glob
from datetime import datetime
from typing import Dict, List, Any, Optional

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

REGISTRY_FILE = os.path.join(BASE_DIR, "catalog", "REGISTRY.json")


def load_registry() -> Dict[str, Any]:
    if os.path.exists(REGISTRY_FILE):
        with open(REGISTRY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"intents": []}


def list_master_workbooks():
    """Prints a beautiful table of all deliverable Excel workbooks in the repo."""
    print("=" * 80)
    print("MASTER EXCEL WORKBOOKS & DELIVERABLES")
    print("=" * 80)
    xlsx_files = glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
    if not xlsx_files:
        print("No .xlsx files found in root.")
        return

    # Categorize and sort by size
    items = []
    for fp in xlsx_files:
        fn = os.path.basename(fp)
        size_kb = os.path.getsize(fp) / 1024
        mtime = datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%Y-%m-%d")
        items.append((fn, size_kb, mtime))

    items.sort(key=lambda x: -x[1])

    print(f"{'File Name':<55} | {'Size (KB)':<10} | {'Last Modified':<12}")
    print("-" * 80)
    for fn, sz, mt in items[:25]:
        print(f"{fn:<55} | {sz:>9.1f} | {mt:<12}")
    print("=" * 80)
    print(f"Total Workbooks: {len(items)}")


def show_catalog():
    cat_path = os.path.join(BASE_DIR, "catalog", "MASTER_DIRECTORY.md")
    if os.path.exists(cat_path):
        with open(cat_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
            for line in lines[:60]:
                print(line, end="")
        print(f"\n... Read full catalog at: {cat_path}")
    else:
        print("Catalog not found.")


def parse_intent(query: str, registry: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    q = query.lower().strip()

    # Direct keyword matches
    if any(k in q for k in ["list workbook", "list excel", "show excel", "what datasets", "list master"]):
        return {"id": "list_workbooks"}
    if any(k in q for k in ["show catalog", "view catalog", "what is this repo", "help", "menu"]):
        return {"id": "show_catalog"}

    # Match against registry triggers
    best_intent = None
    max_matches = 0

    for intent in registry.get("intents", []):
        matches = 0
        for trigger in intent.get("triggers", []):
            t_words = trigger.lower().split()
            if all(w in q for w in t_words):
                matches += len(t_words)
        if matches > max_matches:
            max_matches = matches
            best_intent = intent

    return best_intent


def execute_query(query: str):
    print("=" * 80)
    print(f"OMNICHANNEL INTELLIGENCE DISPATCHER: \"{query}\"")
    print("=" * 80)

    registry = load_registry()
    intent = parse_intent(query, registry)

    if not intent:
        print("[!] No exact matching intent found. Available options:")
        print("  1. 'find kolkata creators 10k'")
        print("  2. 'audit profile @handle'")
        print("  3. 'scan fintech brands 2 years'")
        print("  4. 'detect boost <views> <likes> <followers>'")
        print("  5. 'list master workbooks'")
        print("  6. 'show catalog'")
        return

    intent_id = intent.get("id")

    if intent_id == "list_workbooks":
        list_master_workbooks()
        return

    if intent_id == "show_catalog":
        show_catalog()
        return

    print(f"[+] Routing to Engine: {intent.get('name')}")
    print(f"    Description:      {intent.get('description')}")
    print(f"    Target Script:    {intent.get('script')}")
    print("-" * 80)

    # 1. Profile Audit
    if intent_id == "audit_profile":
        from core.profile_auditor import audit_profile, format_followers
        # Extract handle
        handle_match = re.search(r'@[a-zA-Z0-9_\.]+', query)
        if not handle_match:
            words = query.split()
            handle = words[-1].replace("@", "")
        else:
            handle = handle_match.group(0)

        print(f"[*] Auditing profile: {handle} ...")
        res = audit_profile(handle)
        print("\n" + "=" * 50)
        print(f"Handle:    {res['handle']}")
        print(f"Name:      {res.get('full_name', '')}")
        print(f"Followers: {res.get('followers', 0):,} ({format_followers(res.get('followers', 0))})")
        print(f"Tier:      {res.get('tier', '')}")
        print(f"URL:       {res.get('profile_url', '')}")
        print("=" * 50)
        return

    # 2. Boost Detection
    if intent_id == "detect_boost_ad_spend":
        from core.ad_boost_engine import evaluate_boost
        nums = [int(s) for s in re.findall(r'\b\d+\b', query)]
        views = nums[0] if len(nums) > 0 else 1_250_000
        likes = nums[1] if len(nums) > 1 else 2_500
        followers = nums[2] if len(nums) > 2 else 45_000
        res = evaluate_boost(views, likes, followers)
        print("\n" + "=" * 50)
        print(f"Views:       {views:,}")
        print(f"Likes:       {likes:,}")
        print(f"Followers:   {followers:,}")
        print(f"Like Rate:   {res['like_to_view_pct']}")
        print(f"Multiplier:  {res['view_to_follower_multiplier']}")
        print(f"Status:      {res['boost_status']}")
        print(f"Reason:      {res['boost_reason']}")
        print("=" * 50)
        return

    # 3. Kolkata Creator Discovery
    if intent_id == "kolkata_creator_discovery":
        from core.kolkata_engine import run_discovery
        nums = [int(s) for s in re.findall(r'\b\d+\b', query) if int(s) < 1000]
        target_count = nums[0] if nums else 20
        print(f"[*] Launching Kolkata creator discovery (Target: {target_count} new)...")
        run_discovery(target_new=target_count)
        return

    # 4. FinTech 2-Year Scan
    if intent_id == "fintech_deep_scan_2years":
        from core.fintech_engine import run_fintech_scan
        is_dry = "dry" in query.lower()
        print(f"[*] Launching FinTech 4-Tier 2-Year Scan (dry_run={is_dry})...")
        run_fintech_scan(dry_run=is_dry)
        return

    # Fallback to executing the script directly
    script = intent.get("script")
    if script and os.path.exists(os.path.join(BASE_DIR, script)):
        print(f"[*] Executing {script} ...")
        os.system(f"python {script}")
    else:
        print(f"[!] Target script {script} not found.")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    user_query = " ".join(sys.argv[1:])
    execute_query(user_query)
