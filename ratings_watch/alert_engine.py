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
    threshold: float = ALERT_NEG_RATE_JUMP_THRESHOLD,
    min_sample: int = ALERT_MIN_SAMPLE,
) -> List[Dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    alerts = []
    try:
        platforms = [row[0] for row in conn.execute("SELECT DISTINCT platform FROM reviews").fetchall()]
        for platform in platforms:
            current = _period_reviews(conn, platform, 7, 0)
            baseline = _period_reviews(conn, platform, 14, 7)
            if len(current) < min_sample or len(baseline) < min_sample:
                continue
            current_share = 100 * sum(row[1] <= 2 for row in current) / len(current)
            baseline_share = 100 * sum(row[1] <= 2 for row in baseline) / len(baseline)
            delta = current_share - baseline_share
            if delta > threshold:
                ids = [row[0] for row in current if row[1] <= 2][:20]
                alerts.append(
                    {
                        "alert_type": "negative_share_spike",
                        "platform": platform,
                        "baseline_value": round(baseline_share, 2),
                        "current_value": round(current_share, 2),
                        "threshold": threshold,
                        "severity": "high" if delta > threshold * 1.5 else "medium",
                        "supporting_ids": ids,
                        "suggested_action": "Compare negative reviews in the current seven-day window with the preceding seven-day baseline.",
                        "alert_key": f"negative_share_spike:{platform}",
                        "cooldown_until": _cooldown_until(),
                    }
                )
        return alerts
    finally:
        conn.close()


def evaluate_one_star_spike(
    db_path: str,
    threshold: float = ALERT_ONE_STAR_SPIKE_THRESHOLD,
    min_sample: int = ALERT_MIN_SAMPLE,
) -> List[Dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    alerts = []
    try:
        platforms = [row[0] for row in conn.execute("SELECT DISTINCT platform FROM reviews").fetchall()]
        for platform in platforms:
            current = _period_reviews(conn, platform, 7, 0)
            baseline = _period_reviews(conn, platform, 14, 7)
            if len(current) < min_sample or len(baseline) < min_sample:
                continue
            current_ones = [row[0] for row in current if row[1] == 1]
            baseline_ones = sum(row[1] == 1 for row in baseline)
            current_share = 100 * len(current_ones) / len(current)
            baseline_share = 100 * baseline_ones / len(baseline)
            delta = current_share - baseline_share
            if delta > threshold:
                alerts.append(
                    {
                        "alert_type": "one_star_spike",
                        "platform": platform,
                        "baseline_value": round(baseline_share, 2),
                        "current_value": round(current_share, 2),
                        "threshold": threshold,
                        "severity": "high" if delta > threshold * 1.5 else "medium",
                        "supporting_ids": current_ones[:20],
                        "suggested_action": "Inspect the current seven-day one-star review sample against the preceding seven days.",
                        "alert_key": f"one_star_spike:{platform}",
                        "cooldown_until": _cooldown_until(),
                    }
                )
        return alerts
    finally:
        conn.close()


def _theme_counts(rows: List[Tuple[Any, ...]]) -> Tuple[int, Dict[str, List[str]]]:
    negative = [row for row in rows if row[1] is not None and row[1] <= 2]
    counts: Dict[str, List[str]] = {}
    for review_id, _rating, serialized in negative:
        try:
            themes = json.loads(serialized or "[]")
        except (TypeError, ValueError):
            themes = []
        if isinstance(themes, list):
            for theme in themes:
                if theme:
                    counts.setdefault(str(theme), []).append(str(review_id))
    return len(negative), counts


def evaluate_theme_spike(
    db_path: str,
    threshold: float = ALERT_THEME_SPIKE_THRESHOLD,
    min_sample: int = ALERT_MIN_SAMPLE,
) -> List[Dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    alerts = []
    try:
        platforms = [row[0] for row in conn.execute("SELECT DISTINCT platform FROM reviews").fetchall()]
        for platform in platforms:
            current_rows = _period_reviews(conn, platform, 7, 0)
            baseline_rows = _period_reviews(conn, platform, 14, 7)
            current_total, current_themes = _theme_counts(current_rows)
            baseline_total, baseline_themes = _theme_counts(baseline_rows)
            if current_total < min_sample or baseline_total < min_sample:
                continue
            for theme, current_ids in current_themes.items():
                current_share = 100 * len(current_ids) / current_total
                baseline_share = 100 * len(baseline_themes.get(theme, [])) / baseline_total
                delta = current_share - baseline_share
                if delta > threshold:
                    alerts.append(
                        {
                            "alert_type": "theme_spike",
                            "platform": platform,
