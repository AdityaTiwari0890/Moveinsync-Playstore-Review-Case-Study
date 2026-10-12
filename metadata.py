"""
App Metadata Collector
Fetches public store listing metadata for Google Play Store and Apple App Store apps.
"""

from typing import Dict, Any, Optional
import datetime
import urllib.parse
import requests
from google_play_scraper import app as play_store_app
from google_play_scraper.exceptions import NotFoundError


def get_play_store_metadata(package_id: str, country: str = "in", lang: str = "en") -> Dict[str, Any]:
    """
    Fetch public metadata for an Android app on Google Play Store.

    Args:
        package_id: The Android application package name (e.g., 'com.moveinsync.ets')
        country: Two-letter ISO country code (default 'in')
        lang: Language code (default 'en')

    Returns:
        Dict containing public metadata fields and documented limitations.
    """
    try:
        data = play_store_app(package_id, lang=lang, country=country)
        updated_dt = None
        if data.get("updated"):
            try:
                # updated can be timestamp (int) or formatted string
                if isinstance(data["updated"], (int, float)):
                    updated_dt = datetime.datetime.fromtimestamp(data["updated"]).strftime("%Y-%m-%d %H:%M:%S")
                else:
                    updated_dt = str(data["updated"])
            except Exception:
                updated_dt = str(data.get("updated"))

        return {
            "platform": "Android",
            "store_id": package_id,
            "title": data.get("title", ""),
            "developer": data.get("developer", ""),
            "store_url": data.get("url") or f"https://play.google.com/store/apps/details?id={package_id}",
            "current_version": data.get("version") or "Varies with device",
            "last_updated": updated_dt,
            "store_rating": round(float(data.get("score") or 0.0), 2) if data.get("score") else None,
