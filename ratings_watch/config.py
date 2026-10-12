"""
ratings_watch/config.py
Configuration loader for MOVEINSYNC Ratings Watch.
Reads from environment variables / .env file.
All non-LLM features work without an API key.
"""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

# Load .env from the project root (parent of this file's directory)
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_ENV_FILE = _PROJECT_ROOT / ".env"
load_dotenv(dotenv_path=_ENV_FILE, override=False)


def _first_env(*names: str, default: str = "") -> str:
    for name in names:
        value = os.getenv(name)
        if value is not None and value.strip() != "":
            return value.strip()
    return default


def _as_int(value: str, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _as_float(value: str, default: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def resolve_project_path(value: str | Path) -> Path:
    """Resolve relative configuration paths from the project root, not the caller's CWD."""
    path = Path(value).expanduser()
    return (path if path.is_absolute() else _PROJECT_ROOT / path).resolve()


def is_placeholder_api_key(value: str) -> bool:
    normalized = value.strip().lower()
    return not normalized or normalized in {
        "your_ollama_api_key",
        "your_api_key",
        "replace_me",
        "changeme",
