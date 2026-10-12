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
