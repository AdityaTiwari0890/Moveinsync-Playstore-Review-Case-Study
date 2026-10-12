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
