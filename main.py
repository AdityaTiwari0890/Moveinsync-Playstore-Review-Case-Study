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
