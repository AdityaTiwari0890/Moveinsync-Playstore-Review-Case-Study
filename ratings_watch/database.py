"""
ratings_watch/database.py
SQLite database initialisation, migrations, and CRUD helpers.
Uses parameterised queries throughout. Safe to re-run on existing databases.
"""

from __future__ import annotations
import json
import sqlite3
import threading
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Generator, List, Optional

from ratings_watch.config import DB_PATH, resolve_project_path

# Thread-local connection storage for safe multi-threaded Streamlit use
_local = threading.local()

DDL_STATEMENTS = [
    # ── Collection runs ──────────────────────────────────────────────────────
    """
    CREATE TABLE IF NOT EXISTS collection_runs (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        started_at    TEXT NOT NULL,
        completed_at  TEXT,
        platform      TEXT NOT NULL,
        status        TEXT NOT NULL DEFAULT 'running',
        fetched       INTEGER DEFAULT 0,
        inserted      INTEGER DEFAULT 0,
        updated       INTEGER DEFAULT 0,
        skipped       INTEGER DEFAULT 0,
        failed        INTEGER DEFAULT 0,
        error_message TEXT,
        notes         TEXT
    )
    """,
    # ── Reviews ──────────────────────────────────────────────────────────────
    """
    CREATE TABLE IF NOT EXISTS reviews (
        id                     INTEGER PRIMARY KEY AUTOINCREMENT,
        review_id              TEXT NOT NULL,
        platform               TEXT NOT NULL,
        store_id               TEXT,
        app_name               TEXT,
        app_role               TEXT,
        reviewer_name          TEXT,
        rating                 INTEGER,
        review_title           TEXT,
        review_text            TEXT,
        review_date            TEXT,
        app_version            TEXT,
        developer_reply        TEXT,
        developer_reply_date   TEXT,
        helpful_votes          INTEGER DEFAULT 0,
        assigned_themes        TEXT DEFAULT '[]',
        classification_method  TEXT,
        heuristic_confidence   TEXT,
        is_unclassified        INTEGER NOT NULL DEFAULT 0,
        source_url             TEXT,
        source_collection_timestamp TEXT,
        first_collected_at     TEXT NOT NULL,
        last_updated_at        TEXT NOT NULL,
        collection_run_id      INTEGER,
        is_imported            INTEGER DEFAULT 0,
        UNIQUE(platform, review_id)
    )
    """,
    # ── Rating snapshots ─────────────────────────────────────────────────────
    """
    CREATE TABLE IF NOT EXISTS rating_snapshots (
        id                  INTEGER PRIMARY KEY AUTOINCREMENT,
        platform            TEXT NOT NULL,
        store_id            TEXT NOT NULL,
        app_name            TEXT,
        official_rating     REAL,
        rating_source       TEXT NOT NULL,
        total_ratings_count INTEGER,
        snapshot_at         TEXT NOT NULL,
        collection_run_id   INTEGER,
        metadata_json       TEXT
    )
    """,
    # ── Alerts ───────────────────────────────────────────────────────────────
    """
    CREATE TABLE IF NOT EXISTS alerts (
        id               INTEGER PRIMARY KEY AUTOINCREMENT,
        alert_key        TEXT UNIQUE NOT NULL,
        alert_type       TEXT NOT NULL,
        platform         TEXT,
        theme            TEXT,
        severity         TEXT NOT NULL DEFAULT 'medium',
        baseline_value   REAL,
        current_value    REAL,
        threshold        REAL,
        detected_at      TEXT NOT NULL,
        supporting_ids   TEXT DEFAULT '[]',
        suggested_action TEXT,
        acknowledged     INTEGER DEFAULT 0,
        acknowledged_at  TEXT,
        resolved         INTEGER DEFAULT 0,
        resolved_at      TEXT,
        cooldown_until   TEXT
    )
    """,
    # ── Action items (Step 3 tracker) ────────────────────────────────────────
    """
    CREATE TABLE IF NOT EXISTS action_items (
        id                INTEGER PRIMARY KEY AUTOINCREMENT,
        source            TEXT DEFAULT 'step3_xlsx',
        priority          TEXT,
        theme             TEXT,
        category          TEXT,
        owner_team        TEXT,
        effort_estimate   TEXT,
        success_metric    TEXT,
        status            TEXT DEFAULT 'Open',
        status_updated_at TEXT,
        notes             TEXT,
        original_row_json TEXT
    )
    """,
    # ── App config ───────────────────────────────────────────────────────────
    """
    CREATE TABLE IF NOT EXISTS app_config (
        key        TEXT PRIMARY KEY,
        value      TEXT,
        updated_at TEXT NOT NULL
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS app_registry (
        platform      TEXT NOT NULL,
        store_id      TEXT NOT NULL,
        app_name      TEXT,
        app_role      TEXT,
        first_seen_at TEXT NOT NULL,
        last_seen_at  TEXT NOT NULL,
        PRIMARY KEY (platform, store_id)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS schema_metadata (
        key   TEXT PRIMARY KEY,
        value TEXT NOT NULL
    )
    """,
    # ── Indexes ──────────────────────────────────────────────────────────────
    "CREATE INDEX IF NOT EXISTS idx_reviews_platform_date   ON reviews(platform, review_date)",
    "CREATE INDEX IF NOT EXISTS idx_reviews_rating          ON reviews(rating)",
    "CREATE INDEX IF NOT EXISTS idx_reviews_app_version     ON reviews(app_version)",
    "CREATE INDEX IF NOT EXISTS idx_reviews_run             ON reviews(collection_run_id)",
    "CREATE INDEX IF NOT EXISTS idx_snapshots_platform_at   ON rating_snapshots(platform, snapshot_at)",
    "CREATE INDEX IF NOT EXISTS idx_alerts_type_platform    ON alerts(alert_type, platform)",
]


def get_connection(db_path: str = DB_PATH) -> sqlite3.Connection:
    """Return a thread-local SQLite connection (creates if needed)."""
    db_path = str(resolve_project_path(db_path))
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    if not hasattr(_local, "connections"):
        _local.connections = {}
    if db_path not in _local.connections:
        conn = sqlite3.connect(db_path, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        _local.connections[db_path] = conn
    return _local.connections[db_path]


def close_connection(db_path: str = DB_PATH) -> None:
    if hasattr(_local, "connections") and db_path in _local.connections:
        _local.connections[db_path].close()
        del _local.connections[db_path]


def close_all_connections() -> None:
    if hasattr(_local, "connections"):
        for db_path in list(_local.connections.keys()):
            try:
                _local.connections[db_path].close()
            except Exception:
                pass
        _local.connections.clear()


@contextmanager
def transaction(db_path: str = DB_PATH) -> Generator[sqlite3.Connection, None, None]:
    """Context manager that commits on success, rolls back on exception."""
    conn = get_connection(db_path)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise


def init_db(db_path: str = DB_PATH) -> None:
    """Create all tables and indexes if they don't already exist. Safe to re-run."""
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    conn = get_connection(db_path)
    for stmt in DDL_STATEMENTS:
        conn.execute(stmt)
    review_columns = {row["name"] for row in conn.execute("PRAGMA table_info(reviews)")}
    if "source_collection_timestamp" not in review_columns:
        conn.execute("ALTER TABLE reviews ADD COLUMN source_collection_timestamp TEXT")
    if "is_unclassified" not in review_columns:
        conn.execute("ALTER TABLE reviews ADD COLUMN is_unclassified INTEGER NOT NULL DEFAULT 0")
    conn.execute(
        "INSERT INTO schema_metadata(key, value) VALUES('schema_version', '3') "
