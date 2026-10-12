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
    import_parser.add_argument("--input", default=EXISTING_WORKBOOK)
    import_parser.add_argument("--db", default=settings.db_path)
    import_parser.add_argument("--dry-run", action="store_true")

    collect_parser = subparsers.add_parser("collect", help="Collect ratings from a configured source")
    collect_parser.add_argument("--platform", choices=["android", "ios"], default="android")
    collect_parser.add_argument("--app-id", default=None)
    collect_parser.add_argument("--app-name", default=None)
    collect_parser.add_argument("--app-role", choices=["employee", "driver", "competitor"], default="employee")
    collect_parser.add_argument("--count", type=int, default=200)
    collect_parser.add_argument("--db", default=settings.db_path)

    snapshot_parser = subparsers.add_parser("snapshot", help="Save a rating snapshot")
    snapshot_parser.add_argument("--platform", default="android")
    snapshot_parser.add_argument("--app-id", default=settings.default_android_package)
    snapshot_parser.add_argument("--app-name", default=settings.current_app_name)
    snapshot_parser.add_argument("--rating", type=float, default=None)
    snapshot_parser.add_argument("--db", default=settings.db_path)

    alert_parser = subparsers.add_parser("check-alerts", help="Evaluate and store alerts")
    alert_parser.add_argument("--db", default=settings.db_path)

    index_parser = subparsers.add_parser("index-knowledge-base", help="Incrementally index analysis, reviews and actions")
    index_parser.add_argument("--db", default=settings.db_path)
    index_parser.add_argument("--vector-store", default=VECTOR_STORE_PATH)
    index_parser.add_argument("--rebuild", action="store_true")

    subparsers.add_parser("ai-check", help="Test hosted/local chat inference and embeddings independently")

    status_parser = subparsers.add_parser("status", help="Show database, source and knowledge-base status")
    status_parser.add_argument("--db", default=settings.db_path)
    status_parser.add_argument("--vector-store", default=VECTOR_STORE_PATH)

    monitor_parser = subparsers.add_parser("run-monitor", help="Collect, evaluate alerts, then update the knowledge base")
    monitor_parser.add_argument("--platform", choices=["android", "ios"], default="android")
    monitor_parser.add_argument("--app-id", default=None)
    monitor_parser.add_argument("--app-name", default=None)
    monitor_parser.add_argument("--app-role", choices=["employee", "driver", "competitor"], default="employee")
    monitor_parser.add_argument("--count", type=int, default=200)
    monitor_parser.add_argument("--db", default=settings.db_path)
    monitor_parser.add_argument("--vector-store", default=VECTOR_STORE_PATH)

    step3_parser = subparsers.add_parser("import-step3", help="Import Step 3 recommendations into the tracking database")
    step3_parser.add_argument("--db", default=settings.db_path)

    args = parser.parse_args()

    if hasattr(args, "db"):
        init_db(args.db)

    if args.command == "import-existing":
        result = import_existing_workbook(args.input, args.db, dry_run=args.dry_run)
        print(json.dumps(result, indent=2, ensure_ascii=False))
    elif args.command == "collect":
        app_id = args.app_id or (IOS_APP_ID if args.platform == "ios" else settings.default_android_package)
        app_name = args.app_name or settings.current_app_name
        result = collect_reviews(args.platform, app_id, app_name, args.count, args.db, app_role=args.app_role)
        print(json.dumps(result, indent=2, ensure_ascii=False))
    elif args.command == "snapshot":
        from ratings_watch.database import save_rating_snapshot
        save_rating_snapshot(
            platform=args.platform,
            store_id=args.app_id,
            app_name=args.app_name,
            official_rating=args.rating,
            rating_source="manual_snapshot",
            total_ratings_count=None,
            db_path=args.db,
        )
        print({"status": "snapshot_saved", "platform": args.platform, "official_rating": args.rating})
    elif args.command == "check-alerts":
        result = check_alerts(args.db)
        print(json.dumps(result, indent=2, ensure_ascii=False))
    elif args.command == "index-knowledge-base":
        try:
            result = index_knowledge_base(
                rebuild=args.rebuild,
                workbook_path=EXISTING_WORKBOOK,
                step3_md_path=STEP3_MD,
                action_plan_path=STEP3_ACTION_PLAN,
                db_path=args.db,
                vector_path=args.vector_store,
            )
            print(json.dumps(result, indent=2, ensure_ascii=False))
        except Exception as exc:
            source_status = inspect_knowledge_sources(
                EXISTING_WORKBOOK, STEP3_MD, STEP3_ACTION_PLAN, args.db
            )
            index_status = get_index_status(args.vector_store)
            try:
                index_error = json.loads(index_status.get("last_error") or "{}")
            except (TypeError, ValueError):
                index_error = {"detail": index_status.get("last_error")}
            print(json.dumps({
                "status": "failed",
                "error": str(exc),
                "sources": source_status,
                "chunks_generated": source_status["total_chunks_generated"],
                "failed_documents": index_error.get("failed_document_count", 0),
                "not_persisted_documents": max(
                    0, source_status["total_chunks_generated"] - index_status["chunk_count"]
                ),
                "index_error_details": index_error,
                "knowledge_base": index_status,
            }, indent=2, ensure_ascii=False))
            raise SystemExit(1)
    elif args.command == "ai-check":
        result = LLMAssistant().test_connection()
        print(json.dumps(result, indent=2, ensure_ascii=False))
        if not result["ready"]:
            raise SystemExit(1)
    elif args.command == "status":
        conn = get_connection(args.db)
        try:
            source_status = inspect_knowledge_sources(
                EXISTING_WORKBOOK, STEP3_MD, STEP3_ACTION_PLAN, args.db
            )
            vector_status = get_index_status(args.vector_store)
            ready = bool(
                source_status["workbook_exists"]
                and source_status["total_chunks_generated"] > 0
                and vector_status["document_count"] > 0
                and vector_status["persisted_vector_count"] == vector_status["chunk_count"]
                and vector_status["invalid_vector_count"] == 0
                and not vector_status.get("store_error")
                and not vector_status["last_error"]
            )
            ai_status = LLMAssistant().test_connection()
            status = {
                "reviews_by_platform": {
                    row["platform"]: row["count"]
                    for row in conn.execute("SELECT platform, COUNT(*) AS count FROM reviews GROUP BY platform")
                },
                "registered_app_targets": [
                    dict(row)
                    for row in conn.execute(
                        "SELECT platform, store_id, app_name, app_role, first_seen_at, last_seen_at FROM app_registry ORDER BY platform, app_name, app_role, store_id"
                    )
                ],
                "latest_collection_runs": [
                    dict(row)
                    for row in conn.execute(
                        "SELECT platform, status, started_at, completed_at, fetched, inserted, updated, skipped, failed, error_message FROM collection_runs ORDER BY started_at DESC LIMIT 10"
