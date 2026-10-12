from __future__ import annotations

from typing import Any, Dict, List, Optional

import requests

from ratings_watch.config import (
    OLLAMA_API_KEY,
    OLLAMA_BASE_URL,
    OLLAMA_EMBEDDING_MODEL,
    OLLAMA_MODEL,
    is_placeholder_api_key,
)


class LLMAssistant:
    """Ollama Cloud/local client; credentials stay server-side."""

    def __init__(
        self,
        api_key: Optional[str] = None,
        base_url: str = OLLAMA_BASE_URL,
