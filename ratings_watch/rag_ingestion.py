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
