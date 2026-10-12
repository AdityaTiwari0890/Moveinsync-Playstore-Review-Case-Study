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
        "one_star_count": sum(1 for r in ratings if r == 1),
        "most_common_theme": most_common.most_common(1)[0][0] if most_common else None,
    }


def get_rating_trend(db_path: str, platform: str = "all", days: int = 30) -> List[Dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    query = "SELECT substr(review_date,1,10) as d, AVG(rating) as avg_rating, COUNT(*) as review_count FROM reviews WHERE review_date IS NOT NULL"
    params: List[Any] = []
    if platform != "all":
        query += " AND platform = ?"
        params.append(platform)
    query += " GROUP BY substr(review_date,1,10) ORDER BY d DESC LIMIT ?"
    params.append(days)
    rows = conn.execute(query, params).fetchall()
    return [
        {"date": row[0], "avg_rating": round(_safe_float(row[1], 0), 2), "review_count": row[2]}
        for row in rows
    ][::-1]


def get_review_counts(db_path: str, platform: str = "all") -> Dict[str, int]:
    conn = sqlite3.connect(db_path)
    query = "SELECT platform, COUNT(*) FROM reviews"
    if platform != "all":
        query = "SELECT platform, COUNT(*) FROM reviews WHERE platform = ?"
        params = [platform]
    else:
        params = []
    rows = conn.execute(query, params).fetchall()
    total = sum(int(r[1]) for r in rows)
    return {"total": total, "by_platform": {r[0]: int(r[1]) for r in rows}}


def get_negative_review_share(db_path: str, platform: str = "all", start_date: Optional[str] = None, end_date: Optional[str] = None) -> float:
    conn = sqlite3.connect(db_path)
    query = "SELECT COUNT(*) FROM reviews WHERE 1=1"
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
    total_rows = conn.execute(query, params).fetchone()[0]
    neg_rows = conn.execute(query.replace("SELECT COUNT(*)", "SELECT COUNT(*)") + " AND rating <= 2", params).fetchone()[0]
    return round((neg_rows / total_rows) * 100, 2) if total_rows else 0.0


def get_top_themes(db_path: str, platform: str = "all", limit: int = 5) -> List[Dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    rows = conn.execute(
        "SELECT assigned_themes FROM reviews WHERE assigned_themes IS NOT NULL AND assigned_themes != '[]'" + (" AND platform = ?" if platform != "all" else ""),
        ([platform] if platform != "all" else [])
    ).fetchall()
    counts = Counter()
    for row in rows:
        try:
            themes = __import__('json').loads(row[0])
        except Exception:
            themes = []
        if isinstance(themes, list):
            for theme in themes:
                counts[theme] += 1
    return [{"theme": theme, "count": count} for theme, count in counts.most_common(limit)]


def get_theme_evidence(db_path: str, theme_name: str, limit: int = 5) -> List[Dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    rows = conn.execute(
        "SELECT review_id, platform, rating, review_text, review_date FROM reviews WHERE assigned_themes LIKE ? LIMIT ?",
        (f"%{theme_name}%", limit),
    ).fetchall()
    return [
        {"review_id": row[0], "platform": row[1], "rating": row[2], "review_text": row[3], "review_date": row[4]}
        for row in rows
    ]
