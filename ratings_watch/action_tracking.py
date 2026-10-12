from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

from ratings_watch.database import upsert_action_item


def _clean(value: Any) -> Optional[str]:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    return text or None


def parse_step3_markdown(md_path: str) -> List[Dict[str, Any]]:
    path = Path(md_path)
    if not path.exists():
        return []
    content = path.read_text(encoding="utf-8", errors="replace")
    items = []
    for line in content.splitlines():
        if line.startswith("### Action"):
            title = line.replace("### ", "", 1).strip()
            items.append(
                {
                    "theme": title,
                    "category": "Step 3 recommendation",
                    "priority": "Unspecified",
                    "notes": f"Imported from {path.name}",
                    "source": "step3_md",
                    "original_row": {"title": title, "source_path": str(path)},
                }
            )
    return items


def _priority_matrix_rows(path: Path) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    with pd.ExcelFile(path) as excel:
        for sheet_name in excel.sheet_names:
            frame = pd.read_excel(excel, sheet_name=sheet_name, header=None)
            header_index = None
            headers: List[str] = []
            for index, row in frame.iterrows():
                values = [_clean(value) for value in row.tolist()]
                normalized = {value.casefold() for value in values if value}
                if {"priority", "theme"}.issubset(normalized):
                    header_index = index
                    headers = [value or f"column_{column}" for column, value in enumerate(values)]
                    break
            if header_index is None:
                continue
            for _, row in frame.iloc[header_index + 1 :].iterrows():
                record = {
                    headers[column]: row.iloc[column]
                    for column in range(min(len(headers), len(row)))
                }
                record = {key: value for key, value in record.items() if _clean(value) is not None}
                theme = _clean(record.get("Theme"))
                priority = _clean(record.get("Priority"))
                if not theme or not priority:
                    continue
                rows.append(
                    {
                        "priority": priority,
                        "theme": theme,
                        "category": _clean(record.get("Category")) or "Recommendation",
                        "owner_team": _clean(record.get("Owner Team")),
                        "effort_estimate": _clean(record.get("Effort Estimate") or record.get("Effort Estimate ")),
                        "success_metric": _clean(record.get("Suggested Success Metric")),
                        "status": _clean(record.get("Action Status")) or "Open",
                        "notes": f"Imported from {path.name} / {sheet_name}",
                        "source": "step3_xlsx",
                        "original_row": {
                            key: (None if _clean(value) is None else str(value))
                            for key, value in record.items()
                        },
                    }
                )
    return rows


def load_step3_recommendations(md_path: str, xlsx_path: Optional[str] = None) -> List[Dict[str, Any]]:
    items = parse_step3_markdown(md_path) if md_path else []
    if xlsx_path and Path(xlsx_path).exists():
        items.extend(_priority_matrix_rows(Path(xlsx_path)))
    return items


def import_step3_actions(md_path: str, xlsx_path: Optional[str] = None, db_path: str = "ratings_watch.db") -> int:
    inserted = 0
    for item in load_step3_recommendations(md_path, xlsx_path):
        if item.get("theme"):
            if upsert_action_item(item, db_path=db_path) == "inserted":
                inserted += 1
    return inserted
