from __future__ import annotations

import json
import re
import sqlite3
from datetime import date, timedelta
from typing import Any, Dict, List, Optional

from ratings_watch.config import DB_PATH, VECTOR_STORE_PATH
from ratings_watch.llm_assistant import LLMAssistant
from ratings_watch.vector_store import get_index_status, search_documents


def _question_filters(question: str) -> Dict[str, Any]:
    lowered = question.lower()
    filters: Dict[str, Any] = {}
    mentions_android = bool(re.search(r"\bandroid\b", lowered))
    mentions_ios = bool(re.search(r"\bios\b|iphone|ipad", lowered))
    if mentions_android != mentions_ios:
        filters["platform"] = "Android" if mentions_android else "iOS"
    if "driver" in lowered:
        filters["app_role"] = "driver"
    elif "employee" in lowered:
        filters["app_role"] = "employee"
    iso_dates = re.findall(r"\b\d{4}-\d{2}-\d{2}\b", question)
    if len(iso_dates) == 1:
        filters["date_after"] = iso_dates[0]
    elif len(iso_dates) >= 2:
        filters["date_after"], filters["date_before"] = sorted(iso_dates[:2])
    elif "last 30 days" in lowered or "past 30 days" in lowered:
        filters["date_after"] = (date.today() - timedelta(days=30)).isoformat()
    return filters


def _structured_evidence(db_path: str, filters: Dict[str, Any]) -> Dict[str, Any]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        def build_where(include_platform: bool) -> tuple[str, List[Any]]:
            clauses: List[str] = []
            parameters: List[Any] = []
            if include_platform and filters.get("platform"):
                clauses.append("platform = ?")
                parameters.append(filters["platform"])
            if filters.get("app_role"):
                clauses.append("app_role = ?")
                parameters.append(filters["app_role"])
            if filters.get("app_name"):
                clauses.append("app_name = ?")
                parameters.append(filters["app_name"])
            if filters.get("store_id"):
                clauses.append("store_id = ?")
                parameters.append(filters["store_id"])
            if filters.get("date_after"):
                clauses.append("review_date >= ?")
                parameters.append(filters["date_after"])
            if filters.get("date_before"):
                clauses.append("review_date <= ?")
                parameters.append(filters["date_before"])
            if filters.get("theme"):
                clauses.append("EXISTS (SELECT 1 FROM json_each(reviews.assigned_themes) WHERE value = ?)")
                parameters.append(filters["theme"])
            return (" WHERE " + " AND ".join(clauses) if clauses else "", parameters)

        where, parameters = build_where(include_platform=True)
        row = conn.execute(
            f"""
            SELECT COUNT(*) AS total, AVG(rating) AS avg_rating,
                   SUM(CASE WHEN rating <= 2 THEN 1 ELSE 0 END) AS negative,
                   SUM(CASE WHEN rating = 1 THEN 1 ELSE 0 END) AS one_star
            FROM reviews{where}
            """,
            parameters,
        ).fetchone()
        total = int(row["total"] or 0)
        negatives = int(row["negative"] or 0)
        result: Dict[str, Any] = {
            "review_sample": {
                "count": total,
                "average_rating": round(float(row["avg_rating"]), 2) if row["avg_rating"] is not None else None,
                "negative_count_ratings_1_2": negatives,
                "negative_share_percent": round(100 * negatives / total, 2) if total else None,
                "one_star_count": int(row["one_star"] or 0),
            },
            "filters": filters or {"platform": "all", "app_role": "all"},
        }
        theme_rows = conn.execute(f"SELECT assigned_themes FROM reviews{where}", parameters).fetchall()
        theme_counts: Dict[str, int] = {}
        for theme_row in theme_rows:
            try:
                themes = json.loads(theme_row["assigned_themes"] or "[]")
            except (TypeError, ValueError):
                themes = []
            if isinstance(themes, list):
                for theme in themes:
                    if theme:
                        theme_counts[str(theme)] = theme_counts.get(str(theme), 0) + 1
        result["themes_by_review_count"] = [
            {"theme": key, "review_count": value}
            for key, value in sorted(theme_counts.items(), key=lambda item: (-item[1], item[0]))[:10]
        ]
        comparison_where, comparison_parameters = build_where(include_platform=False)
        platform_rows = conn.execute(
            f"""
            SELECT platform, COUNT(*) AS total, AVG(rating) AS avg_rating,
                   SUM(CASE WHEN rating <= 2 THEN 1 ELSE 0 END) AS negative
            FROM reviews{comparison_where}
            GROUP BY platform ORDER BY platform
            """,
            comparison_parameters,
        ).fetchall()
        result["platform_comparison"] = [
            {
                "platform": item["platform"],
