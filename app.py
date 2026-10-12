from __future__ import annotations

import json
import sqlite3
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

try:
    import plotly.express as px
    import streamlit as st
except ImportError:  # pragma: no cover - exercised when dashboard dependencies are absent
    st = None
    px = None

from ratings_watch.config import (
    EXISTING_WORKBOOK,
    IOS_APP_ID,
    OLLAMA_EMBEDDING_MODEL,
    OLLAMA_MODEL,
    STEP3_ACTION_PLAN,
    STEP3_MD,
    VECTOR_STORE_PATH,
    load_settings,
)
from ratings_watch.database import (
    acknowledge_alert,
    get_connection,
    init_db,
    resolve_alert,
    update_action_status,
)
from ratings_watch.monitor import collect_reviews, import_existing_workbook
from ratings_watch.rag_ingestion import count_pending_documents, index_knowledge_base
from ratings_watch.retriever import ask_analyst
from ratings_watch.llm_assistant import LLMAssistant
from ratings_watch.rag_ingestion import inspect_knowledge_sources
from ratings_watch.vector_store import get_index_status

PAGES = [
    "Overview",
    "Rating Trends",
    "Review Explorer",
    "Complaint Intelligence",
    "Alerts",
    "Action Plan",
    "AI Analysis Assistant",
    "Knowledge Base",
    "Data Collection Status",
]


def _rows(db_path: str, query: str, parameters: Tuple[Any, ...] = ()) -> List[Dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        return [dict(row) for row in conn.execute(query, parameters)]
    finally:
        conn.close()


def _filter_options(db_path: str) -> Dict[str, List[str]]:
    theme_values = set()
