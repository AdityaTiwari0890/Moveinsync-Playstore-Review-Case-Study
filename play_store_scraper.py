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
    app_role: str = "employee",
    count: int = 500,
    start_date: Optional[datetime.date] = None,
    end_date: Optional[datetime.date] = None,
    country: str = "in",
    lang: str = "en",
    log_callback: Optional[Callable[[str, str], None]] = None,
) -> List[Dict[str, Any]]:
    """
    Collect the most recent reviews available for an Android app from Google Play Store.

    Args:
        package_id: Package name (e.g. 'com.moveinsync.ets')
        app_name: Display name of the application
        app_role: Application role ('employee', 'driver', 'competitor')
        count: Maximum number of reviews to collect (min 100, default 500)
        start_date: Optional earliest review date to collect
        end_date: Optional latest review date to collect
        country: Two-letter country code (default 'in')
        lang: Two-letter language code (default 'en')
        log_callback: Optional callback func(level, message)

    Returns:
        List of standardized review dictionaries sorted newest first.
    """
    collected_reviews: List[Dict[str, Any]] = []
    seen_ids = set()
    continuation_token = None
    batch_size = min(200, count)
    collection_ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    source_url = f"https://play.google.com/store/apps/details?id={package_id}"

    def log(level: str, msg: str):
        if log_callback:
            log_callback(level, msg)

    log("INFO", f"Starting Google Play scrape for '{app_name}' ({package_id}), target: {count} reviews.")

    reached_date_boundary = False
    retry_count = 0
    max_retries = 3

    while len(collected_reviews) < count and not reached_date_boundary:
        current_request_count = min(batch_size, count - len(collected_reviews))
        try:
            batch_result, next_token = reviews(
                package_id,
                lang=lang,
                country=country,
                sort=Sort.NEWEST,
                count=current_request_count,
                continuation_token=continuation_token,
            )
            retry_count = 0  # reset retries on success
        except Exception as exc:
            retry_count += 1
            if retry_count <= max_retries:
                wait_time = 2 ** retry_count
                log("WARNING", f"Transient error fetching Play Store reviews ({exc}). Retrying in {wait_time}s...")
                time.sleep(wait_time)
                continue
            else:
                log("ERROR", f"Failed fetching Play Store reviews after {max_retries} retries: {exc}")
                break

        if not batch_result:
            log("INFO", f"No further reviews returned by Google Play for '{package_id}'.")
            break

        new_items_in_batch = 0
        for item in batch_result:
            review_dt = item.get("at")
            if isinstance(review_dt, datetime.datetime):
                review_date_val = review_dt.date()
                review_dt_str = review_dt.strftime("%Y-%m-%d %H:%M:%S")
            elif review_dt:
                review_dt_str = str(review_dt)
                try:
                    review_date_val = datetime.date.fromisoformat(review_dt_str[:10])
                except Exception:
                    review_date_val = None
            else:
                review_dt_str = None
                review_date_val = None

            # Date Range Filters
            if review_date_val and end_date and review_date_val > end_date:
                continue  # Skip reviews newer than end_date

            if review_date_val and start_date and review_date_val < start_date:
                # Because Sort.NEWEST returns newest first, reaching an older review signals boundary
                reached_date_boundary = True
                break

            # Deduplication
            r_id = item.get("reviewId")
            if not r_id:
                # Fallback deterministic hash
                hash_seed = f"{item.get('userName')}_{review_dt_str}_{item.get('content', '')[:50]}"
                r_id = f"fallback_{hashlib.md5(hash_seed.encode('utf-8', errors='ignore')).hexdigest()}"

            if r_id in seen_ids:
                continue

            seen_ids.add(r_id)
            new_items_in_batch += 1

            # Format Developer Reply Date
            rep_dt = item.get("repliedAt")
            rep_dt_str = None
            if isinstance(rep_dt, datetime.datetime):
                rep_dt_str = rep_dt.strftime("%Y-%m-%d %H:%M:%S")
            elif rep_dt:
                rep_dt_str = str(rep_dt)

            standardized_review: Dict[str, Any] = {
                "app_name": app_name,
                "store_id": package_id,
                "platform": "Android",
                "app_role": app_role,
                "review_id": r_id,
                "reviewer_name": item.get("userName") or "Anonymous",
                "rating": int(item.get("score") or 0),
                "review_title": None,  # Not present in Play Store
                "review_text": item.get("content") or "",
                "review_date": review_dt_str,
                "app_version": item.get("reviewCreatedVersion") or item.get("appVersion") or None,
                "developer_reply": item.get("replyContent") or None,
                "developer_reply_date": rep_dt_str,
                "helpful_votes": int(item.get("thumbsUpCount") or 0),
                "collection_timestamp": collection_ts,
                "source_url": source_url,
            }
