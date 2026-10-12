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
        "placeholder",
    }


# ── Project root ─────────────────────────────────────────────────────────────
PROJECT_ROOT: Path = _PROJECT_ROOT

# ── Database ────────────────────────────────────────────────────────────────
DB_PATH: str = str(resolve_project_path(
    _first_env("DATABASE_PATH", "RATINGS_WATCH_DB", "DB_PATH", default="ratings_watch.db")
))

# ── Android app ─────────────────────────────────────────────────────────────
ANDROID_PACKAGE_ID: str = _first_env("ANDROID_PACKAGE_ID", "APP_STORE_ID", "ANDROID_APP_ID", "ANDROID_APP_PACKAGE_ID", default="com.moveinsync.ets")
ANDROID_APP_NAME: str = _first_env("ANDROID_APP_NAME", default="MoveInSync")
ANDROID_APP_ROLE: str = _first_env("ANDROID_APP_ROLE", default="employee")

# ── iOS app ──────────────────────────────────────────────────────────────────
IOS_APP_ID: str = _first_env("IOS_APP_ID", "IOS_STORE_ID", default="962949521")
IOS_APP_NAME: str = _first_env("IOS_APP_NAME", default="MoveInSync")
IOS_APP_ROLE: str = _first_env("IOS_APP_ROLE", default="employee")

# ── Collection ───────────────────────────────────────────────────────────────
COLLECTION_COUNTRY: str = _first_env("COLLECTION_COUNTRY", default="in")
COLLECTION_LANG: str = _first_env("COLLECTION_LANG", default="en")
COLLECTION_COUNT: int = _as_int(_first_env("COLLECTION_COUNT", default="200"), 200)
MAX_RETRIES: int = _as_int(_first_env("MAX_RETRIES", default="3"), 3)
DEFAULT_PLATFORM: str = _first_env("DEFAULT_PLATFORM", "DEFAULT_APP_PLATFORM", default="all")

# ── Existing workbook (Step 1/2 output) ──────────────────────────────────────
EXISTING_WORKBOOK: str = str(resolve_project_path(_first_env(
    "EXISTING_WORKBOOK",
    default="App_Review_Intelligence_MoveInSync_20261012_173118_Analyzed.xlsx",
)))
STEP3_MD: str = str(resolve_project_path(
    _first_env("STEP3_MD", default="MOVEIN_SYNC_STEP3_RECOMMENDATIONS.md")
))
STEP3_ACTION_PLAN: str = str(resolve_project_path(
    _first_env("STEP3_ACTION_PLAN", default="MOVEIN_SYNC_ACTION_PLAN.xlsx")
))

# ── Alert thresholds ─────────────────────────────────────────────────────────
ALERT_RATING_DROP_THRESHOLD: float = _as_float(_first_env("ALERT_RATING_DROP_THRESHOLD", "RATING_DROP_THRESHOLD", default="0.3"), 0.3)
ALERT_NEG_RATE_JUMP_THRESHOLD: float = _as_float(_first_env("ALERT_NEG_RATE_JUMP_THRESHOLD", "NEGATIVE_RATE_JUMP_THRESHOLD", default="10.0"), 10.0)
ALERT_ONE_STAR_SPIKE_THRESHOLD: float = _as_float(_first_env("ALERT_ONE_STAR_SPIKE_THRESHOLD", "ONE_STAR_SPIKE_THRESHOLD", default="20.0"), 20.0)
ALERT_THEME_SPIKE_THRESHOLD: float = _as_float(_first_env("ALERT_THEME_SPIKE_THRESHOLD", "THEME_SPIKE_THRESHOLD", default="25.0"), 25.0)
ALERT_MIN_SAMPLE: int = _as_int(_first_env("MIN_ALERT_SAMPLE_SIZE", "ALERT_MIN_SAMPLE", "MIN_SAMPLE_SIZE", default="15"), 15)
ALERT_COOLDOWN_HOURS: int = _as_int(_first_env("ALERT_COOLDOWN_HOURS", "ALERT_COOLDOWN_HOURS", default="48"), 48)

# ── Ollama / LLM ────────────────────────────────────────────────────────────
OLLAMA_BASE_URL: str = _first_env("OLLAMA_BASE_URL", "OLLAMA_HOST", "OLLAMA_API_BASE_URL", "OLLAMA_API_URL", default="https://ollama.com")
OLLAMA_MODEL: str = _first_env("OLLAMA_MODEL", default="gemma4:31b")
OLLAMA_EMBEDDING_MODEL: str = _first_env("OLLAMA_EMBEDDING_MODEL", default="embeddinggemma")
OLLAMA_API_KEY: str = _first_env("OLLAMA_API_KEY", "API_KEY", default="")
LLM_ENABLED: bool = bool(OLLAMA_BASE_URL)
