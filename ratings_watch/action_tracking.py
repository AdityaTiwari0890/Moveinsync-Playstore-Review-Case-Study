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
