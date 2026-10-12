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
