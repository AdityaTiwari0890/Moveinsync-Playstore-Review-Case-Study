from __future__ import annotations

import json
import math
import sqlite3
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, List, Optional

from ratings_watch.config import resolve_project_path


def _connect(path: str) -> sqlite3.Connection:
    resolved_path = resolve_project_path(path)
    resolved_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(resolved_path), timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS documents (
            document_id TEXT PRIMARY KEY,
            content TEXT NOT NULL,
            content_hash TEXT NOT NULL,
            metadata_json TEXT NOT NULL,
            embedding_json TEXT NOT NULL,
            indexed_at TEXT NOT NULL
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_source ON documents(json_extract(metadata_json, '$.source_type'))")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS index_state (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
        """
    )
    conn.commit()
    return conn


def get_index_status(path: str, pending_count: Optional[int] = None) -> Dict[str, Any]:
    resolved_path = str(resolve_project_path(path))
    try:
        conn = _connect(resolved_path)
    except sqlite3.Error as exc:
        return {
            "document_count": 0,
            "chunk_count": 0,
            "persisted_record_count": 0,
            "review_count": 0,
            "review_chunk_count": 0,
            "analysis_count": 0,
            "analysis_chunk_count": 0,
            "recommendation_count": 0,
            "recommendation_chunk_count": 0,
            "report_count": 0,
            "report_chunk_count": 0,
            "persisted_vector_count": 0,
            "invalid_vector_count": 0,
            "counts_by_source": {},
            "chunks_by_source": {},
            "last_indexed_at": None,
            "last_error": None,
            "embedding_model": None,
            "pending_count": pending_count,
            "path": resolved_path,
            "store_error": f"Vector store could not be opened: {exc}",
        }
    try:
        vector_rows = conn.execute(
            "SELECT metadata_json, embedding_json FROM documents"
        ).fetchall()
        state = dict(conn.execute("SELECT key, value FROM index_state").fetchall())
        chunk_counts: Dict[str, int] = {}
        document_ids: Dict[str, set[str]] = {}
        review_ids = set()
        persisted_vectors = 0
        for row in vector_rows:
            try:
                vector = json.loads(row["embedding_json"])
                metadata = json.loads(row["metadata_json"])
                if (
