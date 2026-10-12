from __future__ import annotations

import json
import sqlite3
from typing import Any, Dict, List, Optional


class QueryTools:
    def __init__(self, db_path: str):
        self.db_path = db_path

    def _conn(self):
        return sqlite3.connect(self.db_path)

    def get_rating_summary(self, platform: str = "all", start_date: Optional[str] = None, end_date: Optional[str] = None) -> Dict[str, Any]:
        conn = self._conn()
        query = "SELECT COUNT(*), AVG(rating), SUM(CASE WHEN rating <= 2 THEN 1 ELSE 0 END) FROM reviews WHERE 1=1"
        params = []
        if platform != "all":
            query += " AND platform = ?"
            params.append(platform)
        if start_date:
            query += " AND review_date >= ?"
            params.append(start_date)
        if end_date:
            query += " AND review_date <= ?"
            params.append(end_date)
        row = conn.execute(query, params).fetchone()
        total = int(row[0] or 0)
        avg = float(row[1] or 0.0)
        neg = int(row[2] or 0)
        return {"total_reviews": total, "average_rating": round(avg, 2), "negative_reviews": neg, "negative_share_pct": round((neg / total) * 100, 2) if total else 0.0}

    def get_rating_trend(self, platform: str = "all", days: int = 30) -> List[Dict[str, Any]]:
        conn = self._conn()
        query = "SELECT substr(review_date, 1, 10) as d, AVG(rating), COUNT(*) FROM reviews WHERE review_date IS NOT NULL"
        params = []
        if platform != "all":
            query += " AND platform = ?"
            params.append(platform)
        query += " GROUP BY substr(review_date, 1, 10) ORDER BY d DESC LIMIT ?"
        params.append(days)
        rows = conn.execute(query, params).fetchall()
        return [{"date": row[0], "avg_rating": round(float(row[1] or 0.0), 2), "review_count": int(row[2])} for row in rows]

    def get_review_counts(self, platform: str = "all") -> Dict[str, Any]:
        conn = self._conn()
        query = "SELECT platform, COUNT(*) FROM reviews"
        params = []
        if platform != "all":
            query += " WHERE platform = ?"
            params.append(platform)
        query += " GROUP BY platform"
        by_platform = {platform_name: int(count) for platform_name, count in conn.execute(query, params).fetchall()}
        return {"by_platform": by_platform, "total": sum(by_platform.values())}

    def get_negative_review_share(self, platform: str = "all") -> float:
        summary = self.get_rating_summary(platform)
        return summary["negative_share_pct"]

    def get_top_themes(self, platform: str = "all", limit: int = 5) -> List[Dict[str, Any]]:
        conn = self._conn()
        query = "SELECT assigned_themes FROM reviews WHERE assigned_themes IS NOT NULL"
        params = []
        if platform != "all":
            query += " AND platform = ?"
            params.append(platform)
        rows = conn.execute(query, params).fetchall()
        counts = {}
        for row in rows:
            try:
                themes = json.loads(row[0])
            except Exception:
                themes = []
            if isinstance(themes, list):
                for theme in themes:
                    counts[theme] = counts.get(theme, 0) + 1
        return [{"theme": theme, "count": count} for theme, count in sorted(counts.items(), key=lambda x: (-x[1], x[0]))[:limit]]

    def compare_platforms(self) -> Dict[str, Any]:
