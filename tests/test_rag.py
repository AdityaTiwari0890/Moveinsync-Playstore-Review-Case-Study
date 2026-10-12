import json
import os
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from ratings_watch.action_tracking import import_step3_actions
from ratings_watch.config import PROJECT_ROOT, resolve_project_path
from ratings_watch.database import close_all_connections, init_db
from ratings_watch.llm_assistant import LLMAssistant
from ratings_watch.monitor import import_existing_workbook
from ratings_watch.rag_ingestion import discover_documents, index_knowledge_base
from ratings_watch.retriever import ask_analyst
from ratings_watch.vector_store import get_document_hashes, get_index_status, search_documents


class FakeOllama:
    embedding_model = "test-embeddings"

    def __init__(self, fail=False):
        self.fail = fail
        self.embed_calls = 0
        self.system_prompt = ""

    def available(self):
        return True

    def configuration_error(self, operation="chat"):
        return None

    def embed(self, texts):
        self.embed_calls += 1
        if self.fail:
            raise RuntimeError("simulated embedding error")
        return [[1.0, 2.0] for _ in texts]

    def ask(self, prompt, system_prompt=None, **kwargs):
        self.system_prompt = system_prompt or ""
        return {"ok": True, "answer": "Login reports appear in the collected evidence [R1] [R999]."}


class RagTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
