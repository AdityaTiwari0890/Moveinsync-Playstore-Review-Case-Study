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
                (platform,),
            ).fetchall()
            if len(snapshots) < 2:
                continue
            current, previous = snapshots[0], snapshots[1]
            delta = float(previous[1]) - float(current[1])
            if delta > threshold:
                alerts.append(
                    {
                        "alert_type": "rating_drop",
                        "platform": platform,
                        "baseline_value": float(previous[1]),
                        "current_value": float(current[1]),
                        "threshold": threshold,
                        "severity": "high" if delta > threshold * 1.5 else "medium",
                        "supporting_ids": [f"snapshot:{previous[0]}", f"snapshot:{current[0]}"],
                        "suggested_action": "Verify the rating change against the original store listings, then inspect recent reviews and release or operations evidence.",
                        "alert_key": f"rating_drop:{platform}",
                        "cooldown_until": _cooldown_until(),
                    }
                )
        return alerts
    finally:
        conn.close()


def evaluate_negative_share_spike(
    db_path: str,
