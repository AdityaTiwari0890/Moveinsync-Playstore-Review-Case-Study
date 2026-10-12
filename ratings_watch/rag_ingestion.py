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
                frame = pd.read_excel(excel, sheet_name=sheet)
                if sheet in {"Combined_Reviews", "Android_Reviews", "iOS_Reviews"}:
                    for index, row in frame.iterrows():
                        raw = row.to_dict()
                        review = normalize_review(raw)
                        if not review:
                            continue
                        source_text = raw.get("Review Text")
                        if source_text is None or pd.isna(source_text):
                            source_text = raw.get("review_text", raw.get("content"))
                        if source_text is None or pd.isna(source_text):
                            source_text = review.get("review_text") or ""
                        else:
                            source_text = str(source_text)
                        review["review_text"] = source_text
                        if not review.get("assigned_themes") and review.get("review_text"):
                            classification = classifier.classify(review["review_text"], review.get("rating"))
                            review["assigned_themes"] = classification["themes"]
                            review["heuristic_confidence"] = classification["heuristic_confidence"]
                            review["classification_method"] = classification["classification_method"]
                            review["is_unclassified"] = classification["is_unclassified"]
                        review_id = f"review:{review['platform']}:{review['review_id']}"
                        themes = review.get("assigned_themes") or []
                        text = review.get("review_text") or ""
                        content = (
                            f"Review ID: {review['review_id']}\nPlatform: {review['platform']}\n"
                            f"App: {review.get('app_name') or 'Unavailable'}\n"
                            f"Role: {review.get('app_role') or 'Unavailable'}\nRating: {review['rating']}\n"
                            f"Date: {review.get('review_date') or 'Unavailable'}\n"
                            f"Version: {review.get('app_version') or 'Unavailable'}\n"
                            f"Themes: {', '.join(map(str, themes)) or 'Unclassified / no theme assigned'}\n"
                            f"Review: {text}"
                        )
                        _add_chunked_document(
                            documents,
                            review_id,
                            content,
                            {
                                "source_type": "review",
                                "source_name": workbook.name,
                                "source_path": str(workbook),
                                "sheet": sheet,
                                "row": int(index) + 2,
                                "review_id": review["review_id"],
                                "platform": review["platform"],
                                "app_name": review.get("app_name"),
                                "app_role": review.get("app_role"),
                                "rating": review["rating"],
                                "review_date": review.get("review_date"),
                                "app_version": review.get("app_version"),
                                "themes": themes,
                                "confidence": review.get("heuristic_confidence"),
                                "classification_method": review.get("classification_method"),
                                "is_unclassified": review.get("is_unclassified", False),
                                "original_review_text": text,
                            },
                        )
                    continue
                for index, row in frame.iterrows():
                    fields = [(str(key), _text(value)) for key, value in row.to_dict().items()]
                    fields = [(key, value) for key, value in fields if value]
                    if not fields:
                        continue
                    body = "\n".join(f"{key}: {value}" for key, value in fields)
                    doc_id = f"workbook:{workbook}:{sheet}:{int(index) + 2}"
                    _add_chunked_document(
                        documents,
                        doc_id,
                        f"Workbook: {workbook.name}\nSheet: {sheet}\nSource row: {int(index) + 2}\n{body}",
                        {
                            "source_type": "analysis",
                            "source_name": workbook.name,
                            "source_path": str(workbook),
                            "sheet": sheet,
                            "row": int(index) + 2,
                        },
                    )

    for path, doc_type in ((step3_md, "recommendation"),):
        if not path.exists():
            continue
        content = path.read_text(encoding="utf-8", errors="replace")
        section = "Document"
        section_chunks: Dict[str, int] = {}
        buffer: List[str] = []
        for line in content.splitlines():
            if line.startswith("#"):
                if buffer:
                    _add_report_chunks(documents, path, doc_type, section, "\n".join(buffer), section_chunks)
                section = line.lstrip("#").strip() or "Document"
                buffer = [line]
            else:
                buffer.append(line)
        if buffer:
            _add_report_chunks(documents, path, doc_type, section, "\n".join(buffer), section_chunks)

    report_path = workbook.parent / "MOVEIN_SYNC_ANALYSIS_REPORT.md"
    if report_path.exists():
        content = report_path.read_text(encoding="utf-8", errors="replace")
        sections: Dict[str, int] = {}
        _add_report_chunks(documents, report_path, "report", "Full report", content, sections)

    action_path = action_plan
    if action_path.exists():
        with pd.ExcelFile(action_path) as excel:
            for sheet in excel.sheet_names:
                frame = pd.read_excel(excel, sheet_name=sheet)
                for index, row in frame.iterrows():
                    fields = [(str(key), _text(value)) for key, value in row.to_dict().items()]
                    fields = [(key, value) for key, value in fields if value]
                    if not fields:
                        continue
                    body = "\n".join(f"{key}: {value}" for key, value in fields)
                    doc_id = f"action-plan:{action_path.name}:{sheet}:{int(index) + 2}"
                    _add_chunked_document(
                        documents,
                        doc_id,
                        f"Action plan: {action_path.name}\nSheet: {sheet}\n{body}",
                        {
                            "source_type": "recommendation",
                            "source_name": action_path.name,
                            "source_path": str(action_path),
                            "sheet": sheet,
                            "row": int(index) + 2,
                        },
                    )

    db = database
    if db.exists():
        conn = sqlite3.connect(str(db))
        conn.row_factory = sqlite3.Row
        try:
            workbook_review_metadata = {
                str(item["metadata"].get("parent_document_id") or item["document_id"]): item["metadata"]
                for item in documents.values()
                if item["metadata"].get("source_type") == "review"
            }
            for row in conn.execute(
                "SELECT review_id, platform, app_name, app_role, rating, review_text, review_date, app_version, assigned_themes, heuristic_confidence, is_unclassified, source_url, source_collection_timestamp FROM reviews"
            ):
                doc_id = f"review:{row['platform']}:{row['review_id']}"
                themes = json.loads(row["assigned_themes"] or "[]")
                content = (
                    f"Review ID: {row['review_id']}\nPlatform: {row['platform']}\n"
                    f"App: {row['app_name'] or 'Unavailable'}\nRole: {row['app_role'] or 'Unavailable'}\n"
                    f"Rating: {row['rating']}\nDate: {row['review_date'] or 'Unavailable'}\n"
                    f"Version: {row['app_version'] or 'Unavailable'}\nThemes: {', '.join(themes) or 'Unclassified / no theme assigned'}\n"
                    f"Review: {row['review_text'] or ''}"
                )
                metadata = {
                    "source_type": "review",
                    "source_name": "ratings_watch.db",
                    "source_path": str(db),
                    "review_id": row["review_id"],
                    "platform": row["platform"],
                    "app_name": row["app_name"],
                    "app_role": row["app_role"],
                    "rating": row["rating"],
                    "review_date": row["review_date"],
                    "source_collection_timestamp": row["source_collection_timestamp"],
                    "app_version": row["app_version"],
                    "themes": themes,
                    "confidence": row["heuristic_confidence"],
                    "is_unclassified": bool(row["is_unclassified"]),
                    "source_url": row["source_url"],
                    "original_review_text": row["review_text"] or "",
                }
                if doc_id in workbook_review_metadata:
                    metadata.update(workbook_review_metadata[doc_id])
                    metadata.update({"confidence": row["heuristic_confidence"], "source_url": row["source_url"]})
                _add_chunked_document(documents, doc_id, content, metadata)

            for row in conn.execute(
                "SELECT id, alert_type, platform, theme, severity, baseline_value, current_value, threshold, detected_at, suggested_action FROM alerts"
            ):
                content = "Monitoring alert\n" + "\n".join(f"{key}: {row[key]}" for key in row.keys() if row[key] is not None)
                doc_id = f"alert:{row['id']}"
                _add_chunked_document(
                    documents,
                    doc_id,
                    content,
                    {"source_type": "alert", "source_name": "ratings_watch.db", "alert_id": row["id"], "platform": row["platform"], "theme": row["theme"]},
                )
            for row in conn.execute("SELECT id, priority, theme, category, owner_team, effort_estimate, success_metric, status, notes FROM action_items"):
                content = "Step 3 action item\n" + "\n".join(f"{key}: {row[key]}" for key in row.keys() if row[key] is not None)
                doc_id = f"action:{row['id']}"
                _add_chunked_document(
                    documents,
                    doc_id,
                    content,
                    {"source_type": "recommendation", "source_name": "ratings_watch.db", "action_id": row["id"], "theme": row["theme"]},
                )
        finally:
            conn.close()
    return list(documents.values())


def _add_report_chunks(
    documents: Dict[str, Dict[str, Any]],
    path: Path,
    source_type: str,
    section: str,
    text: str,
    counters: Dict[str, int],
) -> None:
    chunks = _split_text(text)
    counter_key = f"{path.name}:{section}"
    base = counters.get(counter_key, 0)
    for offset, chunk in enumerate(chunks):
        number = base + offset + 1
        doc_id = f"{source_type}:{path.name}:{section}:{number}"
        documents[doc_id] = _document(
            doc_id,
            chunk,
            {
                "source_type": source_type,
                "source_name": path.name,
                "source_path": str(path),
                "section": section,
                "chunk": number,
                "parent_document_id": f"{source_type}:{path.name}:{section}",
                "chunk_index": offset + 1,
                "chunk_count": len(chunks),
            },
        )
    counters[counter_key] = base + len(chunks)


def inspect_knowledge_sources(
    workbook_path: str = EXISTING_WORKBOOK,
    step3_md_path: str = STEP3_MD,
    action_plan_path: str = STEP3_ACTION_PLAN,
    db_path: str = DB_PATH,
) -> Dict[str, Any]:
    workbook = resolve_project_path(workbook_path)
    recommendation_path = resolve_project_path(step3_md_path)
    action_path = resolve_project_path(action_plan_path)
    report_path = workbook.parent / "MOVEIN_SYNC_ANALYSIS_REPORT.md"
    errors: List[str] = []
    missing_sources: List[str] = []
    sheet_names: List[str] = []
    review_rows = 0
    valid_review_rows = 0
    review_columns: List[str] = []

    if not workbook.exists():
        errors.append(f"Primary analysis workbook was not found: {workbook}")
    else:
        try:
            with pd.ExcelFile(workbook) as excel:
                sheet_names = list(excel.sheet_names)
                if "Combined_Reviews" not in sheet_names:
                    errors.append("Primary analysis workbook has no Combined_Reviews sheet.")
                else:
                    review_frame = pd.read_excel(excel, sheet_name="Combined_Reviews")
                    review_rows = len(review_frame)
                    review_columns = [str(column) for column in review_frame.columns]
                    valid_review_rows = sum(
                        normalize_review(row.to_dict()) is not None
                        for _, row in review_frame.iterrows()
                    )
        except Exception as exc:
            errors.append(f"Could not read primary analysis workbook: {exc}")

    optional_sources = [
        (recommendation_path, "Step 3 recommendation Markdown"),
        (action_path, "Step 3 action-plan workbook"),
        (report_path, "analysis report Markdown"),
    ]
    for path, label in optional_sources:
        if not path.exists():
            missing_sources.append(f"{label}: {path}")

    try:
        documents = discover_documents(workbook, recommendation_path, action_path, db_path)
    except Exception as exc:
        documents = []
        errors.append(f"Knowledge-source discovery failed: {exc}")
    chunk_counts: Dict[str, int] = {}
    document_ids: Dict[str, set[str]] = {}
    for document in documents:
        source_type = str(document["metadata"].get("source_type") or "unknown")
        chunk_counts[source_type] = chunk_counts.get(source_type, 0) + 1
        document_ids.setdefault(source_type, set()).add(
            str(document["metadata"].get("parent_document_id") or document["document_id"])
        )
    document_counts = {source: len(ids) for source, ids in document_ids.items()}
    total_documents = sum(document_counts.values())

    return {
        "workbook_path": str(workbook),
        "workbook_exists": workbook.exists(),
        "workbook_sheets": sheet_names,
        "review_columns": review_columns,
        "review_source_rows": review_rows,
        "valid_review_source_rows": valid_review_rows,
        "invalid_review_source_rows": review_rows - valid_review_rows,
        "review_documents": document_counts.get("review", 0),
        "analysis_documents": document_counts.get("analysis", 0),
        "recommendation_documents": document_counts.get("recommendation", 0),
        "report_documents": document_counts.get("report", 0),
        "total_documents_generated": total_documents,
        "total_chunks_generated": len(documents),
        "counts_by_source": document_counts,
        "chunks_by_source": chunk_counts,
        "missing_sources": missing_sources,
        "errors": errors,
    }


def index_knowledge_base(
    rebuild: bool = False,
    workbook_path: str = EXISTING_WORKBOOK,
    step3_md_path: str = STEP3_MD,
    action_plan_path: str = STEP3_ACTION_PLAN,
    db_path: str = DB_PATH,
    vector_path: str = VECTOR_STORE_PATH,
    assistant: Optional[LLMAssistant] = None,
) -> Dict[str, Any]:
    client = assistant or LLMAssistant()
    workbook_path = str(resolve_project_path(workbook_path))
    step3_md_path = str(resolve_project_path(step3_md_path))
    action_plan_path = str(resolve_project_path(action_plan_path))
    db_path = str(resolve_project_path(db_path))
    vector_path = str(resolve_project_path(vector_path))
    source_status = inspect_knowledge_sources(
        workbook_path, step3_md_path, action_plan_path, db_path
    )
    documents = discover_documents(workbook_path, step3_md_path, action_plan_path, db_path)
    if not documents:
        message = "; ".join(source_status["errors"]) or (
            "No review, report, recommendation, or monitoring documents were found to index."
        )
        record_index_error(vector_path, message)
        raise ValueError(message)
