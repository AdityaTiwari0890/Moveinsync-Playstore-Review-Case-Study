from __future__ import annotations

import hashlib
import json
import sqlite3
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import pandas as pd

from ratings_watch.config import (
    EXISTING_WORKBOOK,
    STEP3_ACTION_PLAN,
    STEP3_MD,
    VECTOR_STORE_PATH,
    resolve_project_path,
)
from ratings_watch.database import DB_PATH
from ratings_watch.llm_assistant import LLMAssistant
from ratings_watch.normalization import normalize_review
from ratings_watch.theme_classifier import RatingsWatchClassifier
from ratings_watch.vector_store import (
    get_document_hashes,
    get_index_status,
    clear_index_error,
    record_index_error,
    write_documents,
)

_CHUNK_SIZE = 1400
_CHUNK_OVERLAP = 160


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list, tuple)):
        return json.dumps(value, ensure_ascii=False, default=str)
    if pd.isna(value):
        return ""
    return str(value).strip()


def _digest(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _document(document_id: str, content: str, metadata: Dict[str, Any]) -> Dict[str, Any]:
    normalized = content.strip()
    return {
        "document_id": document_id,
        "content": normalized,
        "content_hash": _digest(normalized + json.dumps(metadata, sort_keys=True, default=str)),
        "metadata": metadata,
    }


def _add_chunked_document(
    documents: Dict[str, Dict[str, Any]],
    document_id: str,
    content: str,
    metadata: Dict[str, Any],
) -> None:
    chunks = _split_text(content)
    for index, chunk in enumerate(chunks, start=1):
        chunk_id = document_id if len(chunks) == 1 else f"{document_id}:chunk:{index}"
        chunk_metadata = {
            **metadata,
            "parent_document_id": document_id,
            "chunk_index": index,
            "chunk_count": len(chunks),
        }
        documents[chunk_id] = _document(chunk_id, chunk, chunk_metadata)


def _split_text(text: str) -> List[str]:
    paragraphs = [part.strip() for part in text.split("\n") if part.strip()]
    chunks: List[str] = []
    current = ""
    for paragraph in paragraphs:
        while len(paragraph) > _CHUNK_SIZE:
            if current:
                chunks.append(current)
                current = ""
            chunks.append(paragraph[:_CHUNK_SIZE])
            paragraph = paragraph[_CHUNK_SIZE - _CHUNK_OVERLAP :]
        if len(current) + len(paragraph) + 1 > _CHUNK_SIZE and current:
            chunks.append(current)
            current = current[-_CHUNK_OVERLAP :] + "\n" + paragraph
        else:
            current = f"{current}\n{paragraph}".strip()
    if current:
        chunks.append(current)
    return chunks


def discover_documents(
    workbook_path: str = EXISTING_WORKBOOK,
    step3_md_path: str = STEP3_MD,
    action_plan_path: str = STEP3_ACTION_PLAN,
    db_path: str = DB_PATH,
) -> List[Dict[str, Any]]:
    documents: Dict[str, Dict[str, Any]] = {}
    workbook = resolve_project_path(workbook_path)
    step3_md = resolve_project_path(step3_md_path)
    action_plan = resolve_project_path(action_plan_path)
    database = resolve_project_path(db_path)
    classifier = RatingsWatchClassifier()
    if workbook.exists():
        with pd.ExcelFile(workbook) as excel:
            for sheet in excel.sheet_names:
