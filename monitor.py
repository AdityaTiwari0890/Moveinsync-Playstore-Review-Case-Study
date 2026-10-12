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
from ratings_watch.vector_store import get_index_status


def import_existing_workbook(workbook_path: str, db_path: str, dry_run: bool = False) -> Dict[str, Any]:
    workbook_path = str(resolve_project_path(workbook_path))
    db_path = str(resolve_project_path(db_path))
    init_db(db_path)
    if not Path(workbook_path).exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    with pd.ExcelFile(workbook_path) as xls:
        if "Combined_Reviews" not in xls.sheet_names:
            raise ValueError("Workbook does not contain Combined_Reviews; import cannot proceed.")
        df = pd.read_excel(xls, sheet_name="Combined_Reviews")
    rows = [row.to_dict() for _, row in df.iterrows()]
    normalized = normalize_reviews(rows)
    classifier = RatingsWatchClassifier()
    for review in normalized:
        if not review.get("assigned_themes") and review.get("review_text"):
            classified = classifier.classify(review["review_text"], rating=review.get("rating"))
            review["assigned_themes"] = classified["themes"]
            review["classification_method"] = classified["classification_method"]
            review["heuristic_confidence"] = classified["heuristic_confidence"]
            review["is_unclassified"] = classified["is_unclassified"]
    rejected = len(rows) - len(normalized)
    if dry_run:
        return {
            "source_rows": len(rows),
            "valid_rows": len(normalized),
            "rejected_missing_or_invalid_platform_id_or_rating": rejected,
            "action": "dry_run",
        }
    result = bulk_upsert_reviews(normalized, run_id=None, is_imported=True, db_path=db_path)
    return {"source_rows": len(rows), "valid_rows": len(normalized), "rejected": rejected, **result}


def collect_reviews(
    platform: str,
    app_id: str,
    app_name: str,
    count: int,
    db_path: str,
    app_role: str = "employee",
) -> Dict[str, Any]:
    init_db(db_path)
    service = CollectionService(db_path)
    return service.collect_platform(platform, app_id, app_name=app_name, count=count, app_role=app_role)


def check_alerts(db_path: str) -> List[Dict[str, Any]]:
    return collect_alerts(db_path)


def main() -> None:
    settings = load_settings()
    parser = argparse.ArgumentParser(description="MOVEINSYNC Ratings Watch monitoring CLI")
    subparsers = parser.add_subparsers(dest="command", required=True)

    import_parser = subparsers.add_parser("import-existing", help="Import Step 1/2 workbook into SQLite")
