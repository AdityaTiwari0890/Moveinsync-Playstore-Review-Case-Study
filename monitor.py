from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd

from ratings_watch.alert_engine import collect_alerts
from ratings_watch.action_tracking import import_step3_actions
from ratings_watch.collection_service import CollectionService
from ratings_watch.config import (
    EXISTING_WORKBOOK,
    IOS_APP_ID,
    STEP3_ACTION_PLAN,
    STEP3_MD,
    VECTOR_STORE_PATH,
    load_settings,
    resolve_project_path,
)
from ratings_watch.database import bulk_upsert_reviews, get_connection, init_db
from ratings_watch.normalization import normalize_reviews
from ratings_watch.rag_ingestion import index_knowledge_base, inspect_knowledge_sources
from ratings_watch.llm_assistant import LLMAssistant
from ratings_watch.theme_classifier import RatingsWatchClassifier
