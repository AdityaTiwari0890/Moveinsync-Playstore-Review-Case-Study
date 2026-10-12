import os
import sqlite3
import tempfile
import unittest
from unittest.mock import ANY, patch

import pandas as pd

from ratings_watch.database import close_all_connections, init_db, save_rating_snapshot, upsert_alert
from ratings_watch.metrics import get_rating_summary, get_negative_review_share
from ratings_watch.monitor import import_existing_workbook
from ratings_watch.alert_engine import collect_alerts
from ratings_watch.llm_assistant import LLMAssistant
from ratings_watch.normalization import normalize_review
from ratings_watch.collection_service import CollectionService
from ratings_watch.review_sources import AppleAppSource
from ratings_watch.theme_classifier import RatingsWatchClassifier


class Step4Tests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.tempdir.name, "ratings_watch.db")
        init_db(self.db_path)
