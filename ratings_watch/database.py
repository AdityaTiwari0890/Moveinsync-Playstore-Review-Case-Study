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
