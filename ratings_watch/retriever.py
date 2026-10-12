from __future__ import annotations

import json
import re
import sqlite3
from datetime import date, timedelta
from typing import Any, Dict, List, Optional

from ratings_watch.config import DB_PATH, VECTOR_STORE_PATH
from ratings_watch.llm_assistant import LLMAssistant
from ratings_watch.vector_store import get_index_status, search_documents


def _question_filters(question: str) -> Dict[str, Any]:
    lowered = question.lower()
    filters: Dict[str, Any] = {}
    mentions_android = bool(re.search(r"\bandroid\b", lowered))
    mentions_ios = bool(re.search(r"\bios\b|iphone|ipad", lowered))
    if mentions_android != mentions_ios:
        filters["platform"] = "Android" if mentions_android else "iOS"
    if "driver" in lowered:
        filters["app_role"] = "driver"
    elif "employee" in lowered:
        filters["app_role"] = "employee"
    iso_dates = re.findall(r"\b\d{4}-\d{2}-\d{2}\b", question)
    if len(iso_dates) == 1:
        filters["date_after"] = iso_dates[0]
    elif len(iso_dates) >= 2:
