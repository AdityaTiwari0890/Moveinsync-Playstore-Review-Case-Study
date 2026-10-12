"""
ratings_watch/normalization.py
Map incoming review dicts (from scrapers or Excel import) to DB schema.
"""

from __future__ import annotations
import json
import re
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional


def _clean_str(v: Any) -> Optional[str]:
    """Return stripped string or None."""
    s = str(v).strip() if v is not None else None
    if s in (None, "", "nan", "None", "NaT"):
        return None
    return s


def _safe_int(v: Any, default: int = 0) -> int:
    try:
        return int(float(v))
    except (TypeError, ValueError, OverflowError):
        return default


def _safe_rating(v: Any) -> Optional[int]:
    try:
        numeric = float(v)
    except (TypeError, ValueError):
        return None
    if not numeric.is_integer():
        return None
    r = int(numeric)
    return r if 1 <= r <= 5 else None


def _parse_date(v: Any) -> Optional[str]:
    """Normalize various date formats to ISO-8601 string."""
    if v is None:
        return None
    s = str(v).strip()
    if s in ("", "nan", "None", "NaT"):
        return None
    # Already looks ISO-like
    if re.match(r"\d{4}-\d{2}-\d{2}", s):
        return s[:19]
    for fmt in ("%d-%m-%Y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s[:19]


def _source_review_id(review: Dict[str, Any]) -> Optional[str]:
    """Return only a source-provided review identifier."""
    return _clean_str(review.get("review_id") or review.get("Review ID") or review.get("reviewId") or review.get("id"))


def normalize_review(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Convert any incoming review dict (from scraper or Excel) to the canonical DB schema.
    Returns None if the review has no valid source platform, ID, or rating (1–5).
    """
    # Column-name aliasing (Excel uses 'App Name', scrapers use 'app_name', etc.)
    def _get(*keys):
        for k in keys:
            if raw.get(k) is not None:
                return raw[k]
        return None

    platform = _clean_str(_get("platform", "Platform"))
    if not platform:
        return None
    rating = _safe_rating(_get("rating", "Rating", "score", "Score"))
    if rating is None:
        return None

    review_id = _source_review_id(raw)
    if not review_id or review_id.startswith("fallback_"):
        return None
    themes_raw = _get("assigned_themes", "Assigned Themes")
    if isinstance(themes_raw, list):
        themes = themes_raw
    elif isinstance(themes_raw, str) and themes_raw.startswith("["):
        try:
            themes = json.loads(themes_raw)
        except Exception:
            themes = [themes_raw] if themes_raw else []
    else:
        themes = [themes_raw] if themes_raw else []
    unclassified_raw = _get("is_unclassified", "Is Unclassified")
    is_unclassified = (
        unclassified_raw is True
        or unclassified_raw == 1
        or str(unclassified_raw).strip().lower() in {"true", "1", "yes"}
    )

    return {
        "review_id": review_id,
        "platform": platform,
        "store_id": _clean_str(_get("store_id", "Store ID", "package_id")),
        "app_name": _clean_str(_get("app_name", "App Name")),
        "app_role": _clean_str(_get("app_role", "App Role")),
        "reviewer_name": _clean_str(_get("reviewer_name", "Reviewer Name")),
        "rating": rating,
        "review_title": _clean_str(_get("review_title", "Review Title")),
        "review_text": _clean_str(_get("review_text", "Review Text", "content")) or "",
        "review_date": _parse_date(_get("review_date", "Review Date", "at", "updated")),
        "app_version": _clean_str(_get("app_version", "App Version")),
        "developer_reply": _clean_str(_get("developer_reply", "Developer Reply")),
        "developer_reply_date": _parse_date(_get("developer_reply_date", "Developer Reply Date")),
        "helpful_votes": _safe_int(_get("helpful_votes", "Helpful Votes")),
        "assigned_themes": themes,
        "classification_method": _clean_str(_get("classification_method")),
        "heuristic_confidence": _clean_str(_get("heuristic_confidence")),
        "is_unclassified": is_unclassified,
        "source_url": _clean_str(_get("source_url", "Source URL")),
        "source_collection_timestamp": _parse_date(_get("source_collection_timestamp", "Collection Timestamp", "collection_timestamp")),
    }


def normalize_reviews(raws: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Normalize a batch, dropping records without valid source identifiers and ratings."""
    results = []
    for r in raws:
        normalized = normalize_review(r)
        if normalized:
            results.append(normalized)
    return results
