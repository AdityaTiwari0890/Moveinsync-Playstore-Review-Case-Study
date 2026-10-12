"""
Google Play Store Scraper Module
Extracts customer reviews and searches apps using google-play-scraper.
"""

from typing import List, Dict, Any, Optional, Callable
import datetime
import hashlib
import time
import re
from urllib.parse import quote
from google_play_scraper import reviews, Sort, search
from google_play_scraper.features.search import Formats, Regex, get, json


def search_play_store(query: str, country: str = "in", lang: str = "en", limit: int = 5) -> List[Dict[str, Any]]:
    """
    Search Google Play Store for candidate applications matching a query.
    Handles extraction edge cases to ensure valid package IDs.
    """
    candidates: List[Dict[str, Any]] = []
    seen_ids = set()

    try:
        # First attempt using standard search
        raw_results = search(query, n_hits=limit + 3, lang=lang, country=country)
        for item in raw_results:
            app_id = item.get("appId")
            # If appId is None on top card, try regex extraction from raw dom
            if not app_id and item.get("title"):
                try:
                    q_enc = quote(query)
                    url = Formats.Searchresults.build(query=q_enc, lang=lang, country=country)
                    dom = get(url)
                    # Find package id matching store url
                    matches = re.findall(r"details\?id=([a-zA-Z0-9._]+)", dom)
                    for m in matches:
                        if m not in seen_ids and not m.startswith("play."):
                            app_id = m
                            break
                except Exception:
                    pass

            if app_id and app_id not in seen_ids:
                seen_ids.add(app_id)
                candidates.append({
                    "title": item.get("title", "Unknown App"),
                    "app_id": app_id,
                    "developer": item.get("developer", "Unknown Developer"),
                    "score": round(float(item.get("score") or 0.0), 2) if item.get("score") else None,
                    "installs": item.get("installs", ""),
                    "icon": item.get("icon", ""),
                    "store_url": f"https://play.google.com/store/apps/details?id={app_id}",
                    "platform": "Android"
                })
                if len(candidates) >= limit:
                    break
    except Exception as exc:
        pass

    return candidates


def scrape_play_store_reviews(
    package_id: str,
    app_name: str,
