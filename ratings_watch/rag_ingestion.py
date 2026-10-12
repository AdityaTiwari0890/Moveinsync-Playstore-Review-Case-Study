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
