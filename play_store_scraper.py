"""
Google Play Store Scraper Module
Extracts customer reviews and searches apps using google-play-scraper.
"""

from typing import List, Dict, Any, Optional, Callable
import datetime
import hashlib
import time
import re
from urllib.parse import quote
from google_play_scraper import reviews, Sort, search
from google_play_scraper.features.search import Formats, Regex, get, json


def search_play_store(query: str, country: str = "in", lang: str = "en", limit: int = 5) -> List[Dict[str, Any]]:
    """
    Search Google Play Store for candidate applications matching a query.
    Handles extraction edge cases to ensure valid package IDs.
    """
    candidates: List[Dict[str, Any]] = []
    seen_ids = set()
