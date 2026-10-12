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
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value"
    )
    now = datetime.now(timezone.utc).isoformat()
    conn.execute(
        """
        INSERT OR IGNORE INTO app_registry (platform, store_id, app_name, app_role, first_seen_at, last_seen_at)
        SELECT platform, store_id, MAX(app_name), MAX(app_role), ?, ?
        FROM reviews WHERE store_id IS NOT NULL AND store_id != ''
        GROUP BY platform, store_id
        """,
        (now, now),
    )
    conn.commit()


# ── Review CRUD ──────────────────────────────────────────────────────────────

def upsert_review(review: Dict[str, Any], run_id: Optional[int] = None,
                  is_imported: bool = False, db_path: str = DB_PATH) -> str:
    """
    Insert or update a review record.
    Returns 'inserted', 'updated', or 'skipped'.
    Deduplication key: (platform, review_id).
    """
    now = datetime.now(timezone.utc).isoformat()
    themes = review.get("assigned_themes", [])
    themes_json = json.dumps(themes) if isinstance(themes, list) else (themes or "[]")

    with transaction(db_path) as conn:
        existing = conn.execute(
            """SELECT rating, review_text, review_date, app_version, developer_reply,
                      developer_reply_date, assigned_themes, classification_method,
                      heuristic_confidence, app_name, app_role, store_id, source_url,
                      source_collection_timestamp, is_unclassified
               FROM reviews WHERE platform=? AND review_id=?""",
            (review.get("platform"), review.get("review_id")),
        ).fetchone()
        if review.get("platform") and review.get("store_id"):
            conn.execute(
                """
                INSERT INTO app_registry (platform, store_id, app_name, app_role, first_seen_at, last_seen_at)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(platform, store_id) DO UPDATE SET
                    app_name=COALESCE(excluded.app_name, app_registry.app_name),
                    app_role=COALESCE(excluded.app_role, app_registry.app_role),
                    last_seen_at=excluded.last_seen_at
                """,
                (review.get("platform"), review.get("store_id"), review.get("app_name"),
                 review.get("app_role"), now, now),
            )

        if existing:
            # Refresh only changed source fields; retain previously known optional metadata.
            changed = (
                existing["rating"] != review.get("rating")
                or existing["review_text"] != review.get("review_text")
                or existing["review_date"] != review.get("review_date")
                or existing["app_version"] != review.get("app_version")
                or existing["developer_reply"] != review.get("developer_reply")
                or existing["developer_reply_date"] != review.get("developer_reply_date")
                or existing["assigned_themes"] != themes_json
                or existing["classification_method"] != review.get("classification_method")
                or existing["heuristic_confidence"] != review.get("heuristic_confidence")
                or existing["app_name"] != review.get("app_name")
                or existing["app_role"] != review.get("app_role")
                or existing["store_id"] != review.get("store_id")
                or existing["source_url"] != review.get("source_url")
                or existing["source_collection_timestamp"] != review.get("source_collection_timestamp")
                or bool(existing["is_unclassified"]) != bool(review.get("is_unclassified", False))
            )
            if changed:
                conn.execute(
                    """UPDATE reviews SET rating=COALESCE(?, rating),
                       store_id=COALESCE(?, store_id), app_name=COALESCE(?, app_name),
                       app_role=COALESCE(?, app_role), review_text=COALESCE(?, review_text),
                       review_date=COALESCE(?, review_date), app_version=COALESCE(?, app_version),
                       developer_reply=COALESCE(?, developer_reply),
                       developer_reply_date=COALESCE(?, developer_reply_date),
                       assigned_themes=?, classification_method=COALESCE(?, classification_method),
                       heuristic_confidence=COALESCE(?, heuristic_confidence),
                       source_url=COALESCE(?, source_url),
                       source_collection_timestamp=COALESCE(?, source_collection_timestamp),
                       is_unclassified=?,
                       last_updated_at=?,
                       collection_run_id=COALESCE(?, collection_run_id)
                       WHERE platform=? AND review_id=?""",
                    (
                        review.get("rating"), review.get("store_id"), review.get("app_name"),
                        review.get("app_role"), review.get("review_text"), review.get("review_date"),
                        review.get("app_version"),
                        review.get("developer_reply"), review.get("developer_reply_date"),
                        themes_json, review.get("classification_method"),
                        review.get("heuristic_confidence"), review.get("source_url"),
                        review.get("source_collection_timestamp"), int(bool(review.get("is_unclassified", False))), now, run_id,
                        review.get("platform"), review.get("review_id"),
                    ),
                )
                return "updated"
            return "skipped"

        conn.execute(
            """INSERT INTO reviews (
                review_id, platform, store_id, app_name, app_role, reviewer_name,
                rating, review_title, review_text, review_date, app_version,
                developer_reply, developer_reply_date, helpful_votes,
                assigned_themes, classification_method, heuristic_confidence,
                is_unclassified,
                source_url, source_collection_timestamp, first_collected_at, last_updated_at,
                collection_run_id, is_imported
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                review.get("review_id"), review.get("platform"), review.get("store_id"),
                review.get("app_name"), review.get("app_role"), review.get("reviewer_name"),
                review.get("rating"), review.get("review_title"), review.get("review_text"),
                review.get("review_date"), review.get("app_version"),
                review.get("developer_reply"), review.get("developer_reply_date"),
                review.get("helpful_votes", 0),
                themes_json, review.get("classification_method"),
                review.get("heuristic_confidence"), int(bool(review.get("is_unclassified", False))),
                review.get("source_url"),
                review.get("source_collection_timestamp"),
                now, now, run_id, int(is_imported),
            ),
        )
        return "inserted"


def bulk_upsert_reviews(reviews: List[Dict[str, Any]], run_id: Optional[int] = None,
                        is_imported: bool = False, db_path: str = DB_PATH) -> Dict[str, int]:
    """Upsert a list of reviews. Returns counts of inserted/updated/skipped."""
    counts = {"inserted": 0, "updated": 0, "skipped": 0, "failed": 0}
    for r in reviews:
        try:
            result = upsert_review(r, run_id=run_id, is_imported=is_imported, db_path=db_path)
            counts[result] += 1
        except Exception:
            counts["failed"] += 1
    return counts


# ── Collection runs ──────────────────────────────────────────────────────────

def start_collection_run(platform: str, db_path: str = DB_PATH) -> int:
    """Create a new collection run record and return its ID."""
    now = datetime.now(timezone.utc).isoformat()
    with transaction(db_path) as conn:
        cur = conn.execute(
            "INSERT INTO collection_runs (started_at, platform, status) VALUES (?,?,?)",
            (now, platform, "running"),
        )
        return cur.lastrowid


def finish_collection_run(run_id: int, status: str, counts: Dict[str, int],
                          error: Optional[str] = None, notes: Optional[str] = None,
                          db_path: str = DB_PATH) -> None:
    now = datetime.now(timezone.utc).isoformat()
    with transaction(db_path) as conn:
        conn.execute(
            """UPDATE collection_runs SET completed_at=?, status=?,
               fetched=?, inserted=?, updated=?, skipped=?, failed=?,
               error_message=?, notes=? WHERE id=?""",
            (now, status,
             counts.get("fetched", 0), counts.get("inserted", 0),
             counts.get("updated", 0), counts.get("skipped", 0),
             counts.get("failed", 0), error, notes, run_id),
        )


# ── Rating snapshots ─────────────────────────────────────────────────────────

def save_rating_snapshot(platform: str, store_id: str, app_name: str,
                          official_rating: Optional[float], rating_source: str,
                          total_ratings_count: Optional[int] = None,
                          run_id: Optional[int] = None,
                          metadata: Optional[Dict] = None,
                          db_path: str = DB_PATH) -> int:
    now = datetime.now(timezone.utc).isoformat()
    meta_json = json.dumps(metadata) if metadata else None
    with transaction(db_path) as conn:
        cur = conn.execute(
            """INSERT INTO rating_snapshots
               (platform, store_id, app_name, official_rating, rating_source,
                total_ratings_count, snapshot_at, collection_run_id, metadata_json)
               VALUES (?,?,?,?,?,?,?,?,?)""",
            (platform, store_id, app_name, official_rating, rating_source,
             total_ratings_count, now, run_id, meta_json),
        )
        return cur.lastrowid


# ── Alert CRUD ───────────────────────────────────────────────────────────────

def upsert_alert(alert: Dict[str, Any], db_path: str = DB_PATH) -> str:
    """Insert or skip (if in cooldown) an alert. Returns 'inserted' or 'skipped'."""
    now = datetime.now(timezone.utc).isoformat()
    with transaction(db_path) as conn:
        existing = conn.execute(
            "SELECT id, cooldown_until, resolved FROM alerts WHERE alert_key=?",
            (alert["alert_key"],),
        ).fetchone()

        if existing:
            # Skip if in cooldown and not yet resolved
            cooldown = existing["cooldown_until"]
            if not existing["resolved"] and cooldown and cooldown > now:
                return "skipped"
            # Re-open if cooldown expired
            conn.execute(
                """UPDATE alerts SET detected_at=?, baseline_value=?, current_value=?,
                   threshold=?, supporting_ids=?, cooldown_until=?,
                   acknowledged=0, acknowledged_at=NULL, resolved=0, resolved_at=NULL
                   WHERE alert_key=?""",
                (now, alert.get("baseline_value"), alert.get("current_value"),
                 alert.get("threshold"),
                 json.dumps(alert.get("supporting_ids", [])),
                 alert.get("cooldown_until"), alert["alert_key"]),
            )
            return "inserted"

        conn.execute(
            """INSERT INTO alerts (alert_key, alert_type, platform, theme, severity,
               baseline_value, current_value, threshold, detected_at,
               supporting_ids, suggested_action, cooldown_until)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            (alert["alert_key"], alert["alert_type"], alert.get("platform"),
             alert.get("theme"), alert.get("severity", "medium"),
             alert.get("baseline_value"), alert.get("current_value"),
             alert.get("threshold"), now,
             json.dumps(alert.get("supporting_ids", [])),
             alert.get("suggested_action"), alert.get("cooldown_until")),
        )
        return "inserted"


def acknowledge_alert(alert_id: int, db_path: str = DB_PATH) -> None:
    now = datetime.now(timezone.utc).isoformat()
    with transaction(db_path) as conn:
        conn.execute(
            "UPDATE alerts SET acknowledged=1, acknowledged_at=? WHERE id=?",
            (now, alert_id),
        )


def resolve_alert(alert_id: int, db_path: str = DB_PATH) -> None:
    now = datetime.now(timezone.utc).isoformat()
    with transaction(db_path) as conn:
        conn.execute(
            "UPDATE alerts SET resolved=1, resolved_at=? WHERE id=?",
            (now, alert_id),
        )


# ── Action items ─────────────────────────────────────────────────────────────

def upsert_action_item(item: Dict[str, Any], db_path: str = DB_PATH) -> str:
    """Insert or skip if same theme+priority already exists."""
    with transaction(db_path) as conn:
        existing = conn.execute(
            "SELECT id FROM action_items WHERE theme=? AND priority=? AND source=?",
            (item.get("theme"), item.get("priority"), item.get("source", "step3_xlsx")),
        ).fetchone()
        if existing:
            return "skipped"
        now = datetime.now(timezone.utc).isoformat()
        conn.execute(
            """INSERT INTO action_items
               (source, priority, theme, category, owner_team, effort_estimate,
                success_metric, status, status_updated_at, notes, original_row_json)
