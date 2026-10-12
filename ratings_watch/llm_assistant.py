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
        if not self._is_local() and is_placeholder_api_key(self.api_key):
            return "Hosted Ollama is configured, but API_KEY is missing or is still a placeholder. Add a valid key to the project .env file."
        return None

    def available(self) -> bool:
        return self.configuration_error("chat") is None

    def _headers(self) -> Dict[str, str]:
        headers = {"Content-Type": "application/json"}
        if not self._is_local() and not is_placeholder_api_key(self.api_key):
            headers["Authorization"] = f"Bearer {self.api_key}"
        return headers

    def _response_error(self, response: requests.Response, operation: str) -> str:
        detail = ""
        try:
            body = response.json()
            if isinstance(body, dict):
                detail = str(body.get("error") or body.get("message") or "")
        except (ValueError, TypeError):
            pass
        if self.api_key:
