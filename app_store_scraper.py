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
    if isinstance(raw_entries, list):
        for entry in raw_entries:
            if not isinstance(entry, dict) or "im:rating" not in entry:
                continue

            r_id = entry.get("id", {}).get("label")
            rating_str = entry.get("im:rating", {}).get("label", "0")
            title = entry.get("title", {}).get("label", "")
            content = entry.get("content", {}).get("label", "")
            user_name = entry.get("author", {}).get("name", {}).get("label", "Anonymous")
            version = entry.get("im:version", {}).get("label")
            updated = entry.get("updated", {}).get("label")
            votes = entry.get("im:voteCount", {}).get("label") or entry.get("im:voteSum", {}).get("label") or 0

            reviews.append({
                "id": str(r_id) if r_id else None,
                "rating": int(rating_str) if rating_str.isdigit() else 0,
                "title": title,
                "content": content,
                "user_name": user_name,
                "version": version,
                "updated": updated,
                "votes": int(votes) if str(votes).isdigit() else 0,
            })
    return reviews


def _parse_itunes_xml_page(xml_bytes: bytes) -> List[Dict[str, Any]]:
    """Parse reviews from an Apple iTunes XML customer reviews response."""
    reviews = []
    try:
        root = ET.fromstring(xml_bytes)
        ns = {
            "atom": "http://www.w3.org/2005/Atom",
            "im": "http://itunes.apple.com/rss",
        }
        entries = root.findall("atom:entry", ns)
        for entry in entries:
            # Skip app info entry if missing rating
            rating_elem = entry.find("im:rating", ns)
            if rating_elem is None or not rating_elem.text:
                continue

            r_id = entry.findtext("atom:id", "", ns)
            title = entry.findtext("atom:title", "", ns)
            content = entry.findtext("atom:content", "", ns)
            author = entry.findtext("atom:author/atom:name", "Anonymous", ns)
            version = entry.findtext("im:version", None, ns)
            updated = entry.findtext("atom:updated", None, ns)
            votes = entry.findtext("im:voteCount", "0", ns)

            reviews.append({
                "id": r_id,
                "rating": int(rating_elem.text) if rating_elem.text.isdigit() else 0,
