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
