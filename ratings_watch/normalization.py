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
