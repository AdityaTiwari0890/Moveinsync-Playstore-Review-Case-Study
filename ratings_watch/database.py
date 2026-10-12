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
