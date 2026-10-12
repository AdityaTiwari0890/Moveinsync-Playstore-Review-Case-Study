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
