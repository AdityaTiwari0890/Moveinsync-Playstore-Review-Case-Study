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
            detail = detail.replace(self.api_key, "[redacted]")
        detail = detail.replace("\r", " ").replace("\n", " ").strip()[:400]
        if response.status_code in (401, 403):
            return f"Ollama {operation} authentication failed (HTTP {response.status_code}). Verify the hosted API key."
        if response.status_code == 404:
            suffix = f" Server response: {detail}" if detail else ""
            return f"Ollama {operation} endpoint or configured model was not found (HTTP 404). Verify the model name and service access.{suffix}"
        suffix = f" Server response: {detail}" if detail else ""
        return f"Ollama {operation} request failed (HTTP {response.status_code}).{suffix}"

    def ask(
        self,
        prompt: str,
        tools: Optional[List[Dict[str, Any]]] = None,
        timeout: int = 45,
        system_prompt: Optional[str] = None,
    ) -> Dict[str, Any]:
        configuration_error = self.configuration_error("chat")
        if configuration_error:
            return {"ok": False, "error": configuration_error}
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})
        payload: Dict[str, Any] = {"model": self.model, "messages": messages, "stream": False}
        if tools:
            payload["tools"] = tools
        try:
            response = requests.post(
                f"{self.base_url}/chat",
                headers=self._headers(),
                json=payload,
                timeout=timeout,
            )
            if response.status_code != 200:
                return {"ok": False, "error": self._response_error(response, "chat")}
            payload = response.json()
            if not isinstance(payload, dict):
                return {"ok": False, "error": "Ollama chat returned an invalid response object."}
            message = payload.get("message")
            content = message.get("content") if isinstance(message, dict) else None
            if not isinstance(content, str):
                return {"ok": False, "error": "Ollama chat returned an invalid response: message.content is missing."}
            return {"ok": True, "answer": content.strip()}
        except requests.Timeout:
            return {"ok": False, "error": "Ollama chat request timed out. Check connectivity or try again."}
        except requests.RequestException:
            return {"ok": False, "error": "Could not connect to Ollama chat. Check OLLAMA_BASE_URL and network connectivity."}
        except (ValueError, TypeError):
            return {"ok": False, "error": "Ollama chat returned invalid JSON or an unsupported response format."}

    def embed(self, texts: List[str], timeout: int = 60) -> List[List[float]]:
        configuration_error = self.configuration_error("embedding")
        if configuration_error:
            raise RuntimeError(configuration_error)
        if not texts:
            return []
        try:
            response = requests.post(
                f"{self.base_url}/embed",
                headers=self._headers(),
                json={"model": self.embedding_model, "input": texts},
                timeout=timeout,
            )
            if response.status_code != 200:
                raise RuntimeError(self._response_error(response, "embedding"))
            payload = response.json()
            if not isinstance(payload, dict):
