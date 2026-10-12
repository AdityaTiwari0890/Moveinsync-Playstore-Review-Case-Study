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
        model: str = OLLAMA_MODEL,
        embedding_model: str = OLLAMA_EMBEDDING_MODEL,
    ):
        self.api_key = OLLAMA_API_KEY if api_key is None else api_key
        normalized_url = base_url.rstrip("/")
        while normalized_url.endswith("/api/api"):
            normalized_url = normalized_url[:-4]
        if not normalized_url.endswith("/api"):
            normalized_url += "/api"
        self.base_url = normalized_url
        self.model = model
        self.embedding_model = embedding_model

    def _is_local(self) -> bool:
        return "localhost" in self.base_url.lower() or "127.0.0.1" in self.base_url.lower()

    def configuration_error(self, operation: str = "chat") -> Optional[str]:
        if not self.base_url:
            return "OLLAMA_BASE_URL is not configured."
        if operation == "chat" and not self.model:
            return "OLLAMA_MODEL is not configured."
        if operation == "embedding" and not self.embedding_model:
            return "OLLAMA_EMBEDDING_MODEL is not configured."
