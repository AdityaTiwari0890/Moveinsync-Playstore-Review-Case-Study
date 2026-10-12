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
            "store_ratings_count": data.get("ratings") or data.get("reviews"),
            "category": data.get("genre", ""),
            "installs": data.get("installs", ""),
            "description": data.get("description", ""),
            "limitation_note": (
                "Google Play public store pages do not expose historical version release dates "
                "or full changelog history without developer console credentials."
            ),
            "status": "SUCCESS",
        }
    except NotFoundError:
        return {
            "platform": "Android",
            "store_id": package_id,
            "title": "",
            "developer": "",
            "store_url": f"https://play.google.com/store/apps/details?id={package_id}",
            "current_version": None,
            "last_updated": None,
            "store_rating": None,
            "store_ratings_count": None,
            "category": "",
            "installs": "",
            "description": "",
            "limitation_note": f"Package ID '{package_id}' not found on Google Play Store.",
            "status": "NOT_FOUND",
        }
    except Exception as exc:
        return {
            "platform": "Android",
            "store_id": package_id,
            "title": "",
            "developer": "",
            "store_url": f"https://play.google.com/store/apps/details?id={package_id}",
            "current_version": None,
            "last_updated": None,
            "store_rating": None,
            "store_ratings_count": None,
            "category": "",
            "installs": "",
            "description": "",
            "limitation_note": f"Error fetching Google Play metadata: {str(exc)}",
            "status": "ERROR",
        }


def get_app_store_metadata(app_id: str, country: str = "in") -> Dict[str, Any]:
    """
    Fetch public metadata for an iOS app via Apple iTunes Lookup API.

    Args:
        app_id: The numeric Apple App Store app ID (e.g., '962949521') or bundle ID
        country: Two-letter ISO country code (default 'in')

    Returns:
        Dict containing public metadata fields and documented limitations.
    """
    clean_id = str(app_id).strip()
    if clean_id.lower().startswith("id"):
        clean_id = clean_id[2:]

    # Lookup by numeric ID or bundleId
    query_param = "id" if clean_id.isdigit() else "bundleId"
    url = f"https://itunes.apple.com/lookup?{query_param}={clean_id}&country={country}"

    try:
        resp = requests.get(url, headers={"User-Agent": "AppReviewIntelligence/1.0"}, timeout=10)
        resp.raise_for_status()
        payload = resp.json()
        results = payload.get("results", [])

        if not results:
            # Fallback retry without country if regional lookup returned empty
            url_fallback = f"https://itunes.apple.com/lookup?{query_param}={clean_id}"
            resp_fb = requests.get(url_fallback, headers={"User-Agent": "AppReviewIntelligence/1.0"}, timeout=10)
            if resp_fb.status_code == 200:
                results = resp_fb.json().get("results", [])

        if not results:
            return {
                "platform": "iOS",
                "store_id": clean_id,
                "title": "",
                "developer": "",
                "store_url": f"https://apps.apple.com/app/id{clean_id}",
                "current_version": None,
                "last_updated": None,
                "store_rating": None,
                "store_ratings_count": None,
                "category": "",
                "installs": "N/A (iOS)",
                "description": "",
                "limitation_note": f"App ID '{clean_id}' not found on Apple App Store.",
                "status": "NOT_FOUND",
            }

        app_info = results[0]
        release_date = app_info.get("currentVersionReleaseDate")
        if release_date:
            try:
                # ISO date parsing
                dt = datetime.datetime.fromisoformat(release_date.replace("Z", "+00:00"))
                release_date = dt.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                pass

        return {
            "platform": "iOS",
            "store_id": str(app_info.get("trackId", clean_id)),
            "title": app_info.get("trackName", ""),
            "developer": app_info.get("artistName", ""),
            "store_url": app_info.get("trackViewUrl") or f"https://apps.apple.com/app/id{clean_id}",
            "current_version": app_info.get("version"),
            "last_updated": release_date,
            "store_rating": round(float(app_info.get("averageUserRating") or 0.0), 2) if app_info.get("averageUserRating") is not None else None,
            "store_ratings_count": app_info.get("userRatingCount"),
            "category": app_info.get("primaryGenreName", ""),
            "installs": "N/A (iOS)",
            "description": app_info.get("description", ""),
            "limitation_note": (
                "Apple iTunes Lookup API exposes current version release date but does not "
                "expose historical release dates for prior versions."
            ),
            "status": "SUCCESS",
        }
    except Exception as exc:
        return {
            "platform": "iOS",
            "store_id": clean_id,
            "title": "",
            "developer": "",
            "store_url": f"https://apps.apple.com/app/id{clean_id}",
            "current_version": None,
            "last_updated": None,
            "store_rating": None,
            "store_ratings_count": None,
            "category": "",
            "installs": "N/A (iOS)",
            "description": "",
            "limitation_note": f"Error fetching App Store metadata: {str(exc)}",
            "status": "ERROR",
        }
