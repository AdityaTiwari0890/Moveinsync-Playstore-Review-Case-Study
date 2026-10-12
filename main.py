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
