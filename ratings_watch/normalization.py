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

