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
                    isinstance(vector, list)
                    and vector
                    and all(isinstance(value, (int, float)) and math.isfinite(value) for value in vector)
                    and isinstance(metadata, dict)
                ):
                    persisted_vectors += 1
                    source_type = str(metadata.get("source_type") or "unknown")
                    chunk_counts[source_type] = chunk_counts.get(source_type, 0) + 1
                    document_ids.setdefault(source_type, set()).add(
                        str(metadata.get("parent_document_id") or row["metadata_json"])
                    )
                    if source_type == "review" and metadata.get("review_id") is not None:
                        review_ids.add((metadata.get("platform"), metadata["review_id"]))
            except (TypeError, ValueError):
                continue
        persisted_records = len(vector_rows)
        counts = {source: len(ids) for source, ids in document_ids.items()}
        document_count = sum(counts.values())
        return {
            "document_count": document_count,
            "chunk_count": persisted_vectors,
            "persisted_record_count": persisted_records,
            "review_count": len(review_ids),
            "review_chunk_count": chunk_counts.get("review", 0),
            "analysis_count": counts.get("analysis", 0),
            "analysis_chunk_count": chunk_counts.get("analysis", 0),
            "recommendation_count": counts.get("recommendation", 0),
            "recommendation_chunk_count": chunk_counts.get("recommendation", 0),
            "report_count": counts.get("report", 0),
            "report_chunk_count": chunk_counts.get("report", 0),
            "persisted_vector_count": persisted_vectors,
            "invalid_vector_count": persisted_records - persisted_vectors,
            "counts_by_source": counts,
            "chunks_by_source": chunk_counts,
            "last_indexed_at": state.get("last_indexed_at"),
            "last_error": state.get("last_error"),
            "embedding_model": state.get("embedding_model"),
            "pending_count": pending_count,
            "path": str(resolve_project_path(path)),
            "store_error": None,
        }
    finally:
        conn.close()


def get_document_hashes(path: str) -> Dict[str, str]:
    conn = _connect(path)
    try:
        return {row["document_id"]: row["content_hash"] for row in conn.execute("SELECT document_id, content_hash FROM documents")}
    finally:
        conn.close()


def write_documents(
    path: str,
    documents: Iterable[Dict[str, Any]],
    embedding_model: str,
    rebuild: bool = False,
    active_documents: Optional[Iterable[Dict[str, Any]]] = None,
) -> int:
    now = datetime.now(timezone.utc).isoformat()
    documents = list(documents)
    active_documents = list(active_documents) if active_documents is not None else documents
    conn = _connect(path)
    try:
        with conn:
            if rebuild:
                conn.execute("DELETE FROM documents")
            conn.executemany(
                """
                INSERT INTO documents (document_id, content, content_hash, metadata_json, embedding_json, indexed_at)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(document_id) DO UPDATE SET
                    content=excluded.content,
                    content_hash=excluded.content_hash,
                    metadata_json=excluded.metadata_json,
                    embedding_json=excluded.embedding_json,
                    indexed_at=excluded.indexed_at
                """,
                [
                    (
                        doc["document_id"],
                        doc["content"],
                        doc["content_hash"],
                        json.dumps(doc["metadata"], ensure_ascii=False),
