from __future__ import annotations

import json
import sqlite3
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Tuple

from ratings_watch.config import (
    ALERT_COOLDOWN_HOURS,
    ALERT_MIN_SAMPLE,
    ALERT_NEG_RATE_JUMP_THRESHOLD,
    ALERT_ONE_STAR_SPIKE_THRESHOLD,
    ALERT_RATING_DROP_THRESHOLD,
    ALERT_THEME_SPIKE_THRESHOLD,
)
from ratings_watch.database import upsert_alert


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _cooldown_until() -> str:
    return (_now() + timedelta(hours=ALERT_COOLDOWN_HOURS)).isoformat()


def _period_reviews(conn: sqlite3.Connection, platform: str, days_ago_start: int, days_ago_end: int):
    today = _now().date()
    start_date = (today - timedelta(days=days_ago_start)).isoformat()
    end_date = (today - timedelta(days=days_ago_end)).isoformat()
    if days_ago_end == 0:
        end_date = (today + timedelta(days=1)).isoformat()
    return conn.execute(
        """
        SELECT review_id, rating, assigned_themes
        FROM reviews
        WHERE platform=? AND date(review_date) >= ? AND date(review_date) < ?
        """,
        (platform, start_date, end_date),
    ).fetchall()


def evaluate_rating_drop(
    db_path: str,
    threshold: float = ALERT_RATING_DROP_THRESHOLD,
) -> List[Dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    alerts = []
    try:
        platforms = [row[0] for row in conn.execute("SELECT DISTINCT platform FROM rating_snapshots").fetchall()]
        for platform in platforms:
            snapshots = conn.execute(
                """
                SELECT id, official_rating, snapshot_at FROM rating_snapshots
                WHERE platform=? AND official_rating IS NOT NULL
                ORDER BY snapshot_at DESC, id DESC LIMIT 2
                """,
