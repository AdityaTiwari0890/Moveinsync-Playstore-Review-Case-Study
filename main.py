"""
App Review Intelligence Scraper - Main Application
Orchestrates multi-app review collection across Google Play and Apple App Store,
performs business intelligence analysis, and exports into an executive Excel workbook.
"""

from typing import List, Dict, Any, Optional
import argparse
import datetime
import os
import re
import sys
from dateutil.relativedelta import relativedelta

# Configure UTF-8 encoding for Windows terminal safety
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass

from play_store_scraper import search_play_store, scrape_play_store_reviews
from app_store_scraper import search_app_store, scrape_app_store_reviews
from metadata import get_play_store_metadata, get_app_store_metadata
from analysis import (
    calculate_sample_summary,
    calculate_monthly_trends,
    calculate_version_trends,
    calculate_segment_comparisons,
    extract_issue_themes,
)
from excel_exporter import export_intelligence_workbook


def parse_store_id(input_str: str, platform: str) -> str:
    """Extract store identifier from URL or return raw identifier."""
    cleaned = input_str.strip()
    if platform.lower() == "android":
        # Google play URL: https://play.google.com/store/apps/details?id=com.moveinsync.ets
        match = re.search(r"id=([a-zA-Z0-9._]+)", cleaned)
        if match:
            return match.group(1)
        return cleaned
    else:
        # Apple app store URL: https://apps.apple.com/in/app/moveinsync/id962949521
        match = re.search(r"id(\d+)", cleaned)
        if match:
            return match.group(1)
        # Direct numeric ID
        digits = re.findall(r"\d+", cleaned)
        if digits:
            return digits[0]
        return cleaned


def interactive_select_candidate(candidates: List[Dict[str, Any]], platform: str) -> Optional[Dict[str, Any]]:
    """Prompt the user to explicitly select from matching candidates or enter manual ID."""
    if not candidates:
        print(f"\n[!] No matching {platform} apps found automatically.")
        manual_id = input(f"Enter {platform} Store URL or App ID manually (or press Enter to skip): ").strip()
        if manual_id:
            clean_id = parse_store_id(manual_id, platform)
            return {"title": "Manual Entry", "app_id": clean_id, "developer": "N/A", "platform": platform}
        return None

    print(f"\n--- Matching {platform} Candidates Found ---")
    for idx, c in enumerate(candidates, 1):
        score_str = f" | Rating: {c.get('score')} stars" if c.get("score") else ""
        print(f"  [{idx}] {c.get('title')} ({c.get('app_id')}) by {c.get('developer')}{score_str}")
    print("  [M] Enter manual ID / Store URL")
    print("  [S] Skip this platform")

    while True:
        choice = input(f"Select {platform} app [1-{len(candidates)}, M, S]: ").strip().upper()
        if choice == "S":
            return None
        if choice == "M":
            manual_id = input(f"Enter {platform} Store URL or App ID: ").strip()
            if manual_id:
                clean_id = parse_store_id(manual_id, platform)
                return {"title": "Manual Entry", "app_id": clean_id, "developer": "N/A", "platform": platform}
            return None
        if choice.isdigit() and 1 <= int(choice) <= len(candidates):
            selected = candidates[int(choice) - 1]
            print(f"--> Selected: {selected.get('title')} ({selected.get('app_id')})")
            return selected
        print("Invalid option. Please try again.")


def collect_app_target_interactive() -> Optional[Dict[str, Any]]:
    """Gather details for an app interactively."""
    print("\n" + "=" * 60)
    print("ADD APPLICATION TO SCRAPING QUEUE")
    print("=" * 60)
    app_name = input("App Name (e.g. MoveInSync, Routematic): ").strip()
    if not app_name:
        return None

    # Role
    print("\nSelect App Category / Role:")
    print("  [1] Employee-side app (e.g. MoveInSync Employee)")
    print("  [2] Driver-side app (e.g. MoveInSync Driver)")
    print("  [3] Competitor app (e.g. Routematic, Safetrax, Whistle)")
    role_choice = input("Enter choice [1-3] (default 1): ").strip()
    role_map = {"1": "employee", "2": "driver", "3": "competitor"}
    app_role = role_map.get(role_choice, "employee")

    # Platform
    print("\nSelect Platforms to Scrape:")
    print("  [1] Both Android & iOS")
    print("  [2] Android only (Google Play)")
    print("  [3] iOS only (Apple App Store)")
    plat_choice = input("Enter choice [1-3] (default 1): ").strip()
    plat_map = {"1": "both", "2": "android", "3": "ios"}
    platform = plat_map.get(plat_choice, "both")

    android_id = None
    ios_id = None

    # Android Resolution
    if platform in ("both", "android"):
        play_url = input("\nGoogle Play URL or Package ID (leave empty to search): ").strip()
        if play_url:
            android_id = parse_store_id(play_url, "android")
        else:
            print(f"Searching Google Play Store for '{app_name}'...")
            candidates = search_play_store(app_name, limit=5)
            selected = interactive_select_candidate(candidates, "Android")
            if selected:
                android_id = selected.get("app_id")

    # iOS Resolution
    if platform in ("both", "ios"):
        ios_url = input("\nApple App Store URL or App ID (leave empty to search): ").strip()
        if ios_url:
            ios_id = parse_store_id(ios_url, "ios")
        else:
            print(f"Searching Apple App Store for '{app_name}'...")
            candidates = search_app_store(app_name, limit=5)
            selected = interactive_select_candidate(candidates, "iOS")
            if selected:
                ios_id = selected.get("app_id")

    if not android_id and not ios_id:
        print("[!] No valid store identifiers configured for this app.")
        return None

    # Review Count
    count_str = input("\nTarget number of reviews per platform (min 100, default 500, max 10000): ").strip()
    count = 500
    if count_str.isdigit():
        count = max(100, min(10000, int(count_str)))

    return {
        "app_name": app_name,
        "app_role": app_role,
        "platform": platform,
        "android_id": android_id,
        "ios_id": ios_id,
        "count": count,
    }


def main():
    parser = argparse.ArgumentParser(
        description="App Review Intelligence Scraper - Multi-App Play Store & App Store Review Intelligence"
    )
    # CLI Mode Arguments
    parser.add_argument("--app", type=str, help="Application display name")
    parser.add_argument("--role", type=str, choices=["employee", "driver", "competitor"], default="employee", help="App role category")
    parser.add_argument("--play-url", type=str, help="Google Play Store URL or package ID")
    parser.add_argument("--app-store-url", type=str, help="Apple App Store URL or numeric app ID")
    parser.add_argument("--platform", type=str, choices=["android", "ios", "both"], default="both", help="Platforms to collect")
    parser.add_argument("--count", type=int, default=500, help="Target review count per platform (min 100, max 10000)")
    parser.add_argument("--months", type=int, choices=[6, 12], help="Filter reviews to the last 6 or 12 months")
    parser.add_argument("--start-date", type=str, help="Earliest review date (YYYY-MM-DD)")
    parser.add_argument("--end-date", type=str, help="Latest review date (YYYY-MM-DD)")
    parser.add_argument("--country", type=str, default="in", help="Two-letter store country code (default 'in')")
    parser.add_argument("--output", type=str, help="Output Excel filename (.xlsx)")
    parser.add_argument("--non-interactive", action="store_true", help="Run without terminal prompts (CLI arguments only)")

    args = parser.parse_args()

    print("=" * 70)
    print("      APP REVIEW INTELLIGENCE SCRAPER - ENTERPRISE EDITION      ")
    print("=" * 70)

    # Date Range Calculation
    start_date = None
    end_date = None
    if args.months:
        start_date = datetime.date.today() - relativedelta(months=args.months)
        print(f"[*] Date Range Filter: Last {args.months} months (since {start_date})")
    elif args.start_date:
        try:
            start_date = datetime.date.fromisoformat(args.start_date)
            print(f"[*] Start Date Filter: {start_date}")
        except Exception:
            print(f"[!] Invalid start-date format: {args.start_date}")

    if args.end_date:
        try:
            end_date = datetime.date.fromisoformat(args.end_date)
            print(f"[*] End Date Filter: {end_date}")
        except Exception:
            print(f"[!] Invalid end-date format: {args.end_date}")

    apps_queue: List[Dict[str, Any]] = []

    # Check CLI input
    if args.app or args.play_url or args.app_store_url:
        app_name = args.app or "Target Application"
        android_id = parse_store_id(args.play_url, "android") if args.play_url else None
        ios_id = parse_store_id(args.app_store_url, "ios") if args.app_store_url else None

        # If IDs missing in CLI mode and not non-interactive, prompt
        if not android_id and args.platform in ("both", "android"):
            print(f"Searching Google Play Store for '{app_name}'...")
            cands = search_play_store(app_name, limit=5, country=args.country)
            if args.non_interactive and cands:
                android_id = cands[0]["app_id"]
            else:
                sel = interactive_select_candidate(cands, "Android")
                if sel:
                    android_id = sel.get("app_id")

        if not ios_id and args.platform in ("both", "ios"):
            print(f"Searching Apple App Store for '{app_name}'...")
            cands = search_app_store(app_name, limit=5, country=args.country)
            if args.non_interactive and cands:
                ios_id = cands[0]["app_id"]
            else:
                sel = interactive_select_candidate(cands, "iOS")
                if sel:
                    ios_id = sel.get("app_id")

        target_count = max(100, min(10000, args.count))
        apps_queue.append({
            "app_name": app_name,
            "app_role": args.role,
            "platform": args.platform,
            "android_id": android_id,
            "ios_id": ios_id,
            "count": target_count,
        })
    else:
        # Full Interactive Terminal Wizard
        print("\nEntering Interactive Configuration Wizard...")
        while True:
            target = collect_app_target_interactive()
            if target:
                apps_queue.append(target)
                print(f"[+] Added '{target['app_name']}' ({target['app_role']}) to queue.")

            add_more = input("\nWould you like to add another app (e.g. driver-side or competitor)? [y/N]: ").strip().lower()
            if add_more not in ("y", "yes"):
                break

    if not apps_queue:
        print("[!] No apps queued for scraping. Exiting.")
        sys.exit(0)

    print("\n" + "=" * 70)
    print(f"CONFIRMED APPS TO SCRAPE ({len(apps_queue)} total):")
    for idx, a in enumerate(apps_queue, 1):
        print(f"  [{idx}] {a['app_name']} ({a['app_role']}) | Android: {a.get('android_id') or 'None'} | iOS: {a.get('ios_id') or 'None'} | Target: {a['count']} reviews")
    print("=" * 70)

    if not args.non_interactive:
        confirm = input("\nProceed with data collection? [Y/n]: ").strip().lower()
        if confirm in ("n", "no"):
            print("Operation aborted by user.")
            sys.exit(0)

    # Output Filename
    timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    if args.output:
        output_file = args.output
        if not output_file.endswith(".xlsx"):
            output_file += ".xlsx"
    else:
        first_name = apps_queue[0]["app_name"].replace(" ", "_")
        output_file = f"App_Review_Intelligence_{first_name}_{timestamp_str}.xlsx"

    all_android_reviews: List[Dict[str, Any]] = []
    all_ios_reviews: List[Dict[str, Any]] = []
    all_metadata: List[Dict[str, Any]] = []
    collection_logs: List[Dict[str, Any]] = []
    summary_metrics_list: List[Dict[str, Any]] = []

    def log_collector(level: str, msg: str, app_name: str, platform: str, store_id: str, req_cnt: int = 0, ret_cnt: int = 0, source: str = "Store API"):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{ts}] [{level}] [{app_name} - {platform}] {msg}")
        collection_logs.append({
            "timestamp": ts,
            "app_name": app_name,
            "platform": platform,
            "store_id": store_id,
            "status": level,
            "requested_count": req_cnt,
            "retrieved_count": ret_cnt,
            "data_source": source,
            "message": msg,
            "limitation_note": "" if level in ("INFO", "SUCCESS") else msg,
        })

    # Execute Collection Loop
    for target in apps_queue:
        name = target["app_name"]
        role = target["app_role"]
        req_count = target["count"]

        # 1. Android Collection
        if target.get("android_id"):
            pkg = target["android_id"]
            print(f"\n>>> Collecting Android Metadata & Reviews for '{name}' ({pkg})...")
            # Metadata
            meta_and = get_play_store_metadata(pkg, country=args.country)
            all_metadata.append(meta_and)
            log_collector(meta_and["status"], f"Fetched Android store metadata (Rating: {meta_and.get('store_rating')} stars)", name, "Android", pkg, 0, 0, "Google Play API")

            # Reviews
            def and_cb(lvl, m):
                log_collector(lvl, m, name, "Android", pkg, req_count, len(all_android_reviews), "google-play-scraper")

            and_revs = scrape_play_store_reviews(
                package_id=pkg,
                app_name=name,
                app_role=role,
                count=req_count,
                start_date=start_date,
                end_date=end_date,
                country=args.country,
                log_callback=and_cb,
            )
            all_android_reviews.extend(and_revs)

            # Analytics for this app's Android sample
            sm_and = calculate_sample_summary(and_revs, requested_count=req_count)
            sm_and["app_name"] = name
            sm_and["platform"] = "Android"
            sm_and["app_role"] = role
            summary_metrics_list.append(sm_and)

        # 2. iOS Collection
        if target.get("ios_id"):
            app_id = target["ios_id"]
            print(f"\n>>> Collecting iOS Metadata & Reviews for '{name}' ({app_id})...")
            # Metadata
            meta_ios = get_app_store_metadata(app_id, country=args.country)
            all_metadata.append(meta_ios)
            log_collector(meta_ios["status"], f"Fetched iOS store metadata (Rating: {meta_ios.get('store_rating')} stars)", name, "iOS", app_id, 0, 0, "iTunes Lookup API")

            # Reviews
            def ios_cb(lvl, m):
                log_collector(lvl, m, name, "iOS", app_id, req_count, len(all_ios_reviews), "Apple iTunes RSS Feed")

            ios_revs = scrape_app_store_reviews(
                app_id=app_id,
                app_name=name,
                app_role=role,
                count=req_count,
                start_date=start_date,
                end_date=end_date,
                country=args.country,
                log_callback=ios_cb,
            )
            all_ios_reviews.extend(ios_revs)

            # Analytics for this app's iOS sample
            sm_ios = calculate_sample_summary(ios_revs, requested_count=req_count)
            sm_ios["app_name"] = name
            sm_ios["platform"] = "iOS"
            sm_ios["app_role"] = role
            summary_metrics_list.append(sm_ios)

    # 3. Combined Dataset & Global Analytics
    all_combined_reviews = all_android_reviews + all_ios_reviews
    all_combined_reviews.sort(key=lambda r: r.get("review_date") or "", reverse=True)

    print("\n" + "=" * 70)
    print("RUNNING BUSINESS INTELLIGENCE & THEME EXTRACTION ANALYSIS...")
    print("=" * 70)

    monthly_trends = calculate_monthly_trends(all_combined_reviews)
    segment_comparisons = calculate_segment_comparisons(all_combined_reviews)
    issue_themes = extract_issue_themes(all_combined_reviews)

    print(f"[*] Total Android Reviews: {len(all_android_reviews)}")
    print(f"[*] Total iOS Reviews:     {len(all_ios_reviews)}")
    print(f"[*] Total Combined:        {len(all_combined_reviews)}")
    print(f"[*] Issue Themes Identified: {len(issue_themes)}")

    # 4. Generate Excel Workbook
    print(f"\n[*] Exporting comprehensive 6-sheet intelligence workbook to: '{output_file}'...")
    export_intelligence_workbook(
        android_reviews=all_android_reviews,
        ios_reviews=all_ios_reviews,
        combined_reviews=all_combined_reviews,
        metadata_list=all_metadata,
        summary_metrics=summary_metrics_list,
        monthly_trends=monthly_trends,
        segment_comparisons=segment_comparisons,
        issue_themes=issue_themes,
        collection_logs=collection_logs,
        output_filepath=output_file,
    )

    print("\n" + "=" * 70)
    print("SUCCESS: COLLECTION AND BUSINESS ANALYSIS COMPLETE!")
    print(f"Excel Workbook Saved: {os.path.abspath(output_file)}")
    print("Sheets Generated:")
    print("  1. Android_Reviews   (Raw Android data with all fields)")
    print("  2. iOS_Reviews       (Raw iOS data with all fields)")
    print("  3. Combined_Reviews  (Unified cross-platform dataset, newest first)")
    print("  4. App_Summary       (Store listing metadata & sample metrics & segment comparisons)")
    print("  5. Issue_Themes      (Thematic complaint & positive analysis)")
    print("  6. Collection_Log    (Complete audit trail of requests, results & API limits)")
    print("=" * 70)


if __name__ == "__main__":
    main()
