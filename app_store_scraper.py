"""
Apple App Store Scraper Module
Extracts customer reviews and searches apps using Apple iTunes Search & Customer Reviews RSS feeds.
Optionally supports SerpApi if SERPAPI_API_KEY is provided in the environment.
"""

from typing import List, Dict, Any, Optional, Callable
import datetime
import hashlib
import time
import os
import xml.etree.ElementTree as ET
import requests
from dotenv import load_dotenv

load_dotenv()


def search_app_store(query: str, country: str = "in", limit: int = 5) -> List[Dict[str, Any]]:
    """
    Search Apple App Store for candidate iOS applications matching a query.
    """
    candidates: List[Dict[str, Any]] = []
    url = f"https://itunes.apple.com/search?term={requests.utils.quote(query)}&entity=software&country={country}&limit={limit + 2}"
    try:
        resp = requests.get(url, headers={"User-Agent": "AppReviewIntelligence/1.0"}, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            for item in data.get("results", []):
                app_id = str(item.get("trackId", ""))
                if not app_id:
                    continue
                candidates.append({
                    "title": item.get("trackName", "Unknown App"),
                    "app_id": app_id,
                    "developer": item.get("artistName", "Unknown Developer"),
                    "score": round(float(item.get("averageUserRating") or 0.0), 2) if item.get("averageUserRating") is not None else None,
                    "installs": "N/A (iOS)",
                    "icon": item.get("artworkUrl60", ""),
                    "store_url": item.get("trackViewUrl") or f"https://apps.apple.com/app/id{app_id}",
                    "platform": "iOS"
                })
                if len(candidates) >= limit:
                    break
    except Exception:
        pass
    return candidates


def _parse_itunes_json_page(json_data: dict) -> List[Dict[str, Any]]:
    """Parse raw reviews from an Apple iTunes JSON customer reviews response."""
    raw_entries = json_data.get("feed", {}).get("entry", [])
    reviews = []
    # In Apple's feed, entry[0] is often the app metadata object if multiple entries exist
