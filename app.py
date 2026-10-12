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
    for row in _rows(db_path, "SELECT assigned_themes FROM reviews WHERE assigned_themes IS NOT NULL"):
        theme_values.update(_theme_list(row["assigned_themes"]))
    return {
        "platforms": [row["platform"] for row in _rows(db_path, "SELECT DISTINCT platform FROM reviews ORDER BY platform")],
        "apps": [row["app_name"] for row in _rows(db_path, "SELECT DISTINCT app_name FROM reviews WHERE app_name IS NOT NULL ORDER BY app_name")],
        "roles": [row["app_role"] for row in _rows(db_path, "SELECT DISTINCT app_role FROM reviews WHERE app_role IS NOT NULL ORDER BY app_role")],
        "versions": [row["app_version"] for row in _rows(db_path, "SELECT DISTINCT app_version FROM reviews WHERE app_version IS NOT NULL ORDER BY app_version")],
        "store_ids": [row["store_id"] for row in _rows(db_path, "SELECT DISTINCT store_id FROM reviews WHERE store_id IS NOT NULL ORDER BY store_id")],
        "themes": sorted(theme_values),
    }


def _selected_filters(db_path: str) -> Dict[str, Any]:
    options = _filter_options(db_path)
    st.sidebar.header("Analysis filters")
    platform = st.sidebar.selectbox("Platform", ["All platforms"] + options["platforms"])
    app_name = st.sidebar.selectbox("App", ["All apps"] + options["apps"])
    role = st.sidebar.selectbox("App role", ["All roles"] + options["roles"])
    store_id = st.sidebar.selectbox("Store identifier", ["All identifiers"] + options["store_ids"])
    theme = st.sidebar.selectbox("Complaint theme", ["All themes"] + options["themes"])
    date_range = st.sidebar.date_input("Review date range", value=(), key="rw_date_range")
    version = st.sidebar.selectbox("App version", ["All versions"] + options["versions"])
    filters = {
        "platform": None if platform == "All platforms" else platform,
        "app_name": None if app_name == "All apps" else app_name,
        "app_role": None if role == "All roles" else role,
        "store_id": None if store_id == "All identifiers" else store_id,
        "theme": None if theme == "All themes" else theme,
        "app_version": None if version == "All versions" else version,
        "start_date": date_range[0].isoformat() if isinstance(date_range, (tuple, list)) and len(date_range) > 0 else None,
        "end_date": date_range[-1].isoformat() if isinstance(date_range, (tuple, list)) and len(date_range) > 1 else None,
    }
    st.sidebar.caption("All sample metrics below use collected review records and the selected filters.")
    return filters


def _review_query(db_path: str, filters: Dict[str, Any], limit: Optional[int] = None) -> List[Dict[str, Any]]:
    clauses = []
    parameters: List[Any] = []
    for key in ("platform", "app_name", "app_role", "app_version", "store_id"):
        if filters.get(key):
            clauses.append(f"{key} = ?")
            parameters.append(filters[key])
    if filters.get("start_date"):
        clauses.append("review_date >= ?")
        parameters.append(filters["start_date"])
    if filters.get("end_date"):
        clauses.append("review_date <= ?")
        parameters.append(filters["end_date"])
    query = "SELECT * FROM reviews" + (" WHERE " + " AND ".join(clauses) if clauses else "") + " ORDER BY review_date DESC"
    if limit is not None:
        query += " LIMIT ?"
        parameters.append(limit)
    return _rows(db_path, query, tuple(parameters))


def _theme_list(raw: Any) -> List[str]:
    try:
        values = json.loads(raw or "[]") if isinstance(raw, str) else raw
    except (ValueError, TypeError):
        values = []
    return [str(value) for value in values] if isinstance(values, list) else []


def _filtered_review_data(db_path: str, filters: Dict[str, Any]) -> pd.DataFrame:
    rows = _review_query(db_path, filters)
