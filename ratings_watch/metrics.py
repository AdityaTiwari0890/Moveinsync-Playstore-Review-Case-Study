from __future__ import annotations

from collections import Counter
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
import sqlite3

from ratings_watch.database import get_connection


def _safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def get_rating_summary(db_path: str, platform: str = "all", start_date: Optional[str] = None, end_date: Optional[str] = None) -> Dict[str, Any]:
    conn = sqlite3.connect(db_path)
    query = "SELECT rating, review_date, review_id, platform FROM reviews WHERE 1=1"
    params: List[Any] = []
    if platform != "all":
        query += " AND platform = ?"
        params.append(platform)
    if start_date:
        query += " AND review_date >= ?"
        params.append(start_date)
    if end_date:
        query += " AND review_date <= ?"
        params.append(end_date)
    rows = conn.execute(query, params).fetchall()
    if not rows:
        return {"current_rating": None, "average_rating": None, "total_reviews": 0, "negative_share": 0.0, "one_star_count": 0, "most_common_theme": None}

    ratings = [r[0] for r in rows]
    total_reviews = len(ratings)
    negative_count = sum(1 for r in ratings if r <= 2)
    avg = sum(ratings) / total_reviews if total_reviews else 0.0
    current_rating = sum(ratings) / total_reviews if total_reviews else None
    most_common = Counter(r[0] for r in rows)
    # theme summary is expected to be computed from theme table when present
    return {
        "current_rating": round(current_rating, 2) if current_rating is not None else None,
        "average_rating": round(avg, 2),
        "total_reviews": total_reviews,
        "negative_share": round((negative_count / total_reviews) * 100, 2) if total_reviews else 0.0,
