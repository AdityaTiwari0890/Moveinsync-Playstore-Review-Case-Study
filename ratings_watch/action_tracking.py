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


