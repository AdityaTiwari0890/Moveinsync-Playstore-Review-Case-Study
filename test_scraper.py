"""
Acceptance & Integration Test Suite for App Review Intelligence Scraper
Tests Play Store scraper, App Store scraper, Metadata collector,
Business Analysis calculations, and Excel 6-sheet export integrity.
"""

import unittest
import os
import datetime
import openpyxl

from play_store_scraper import search_play_store, scrape_play_store_reviews
from app_store_scraper import search_app_store, scrape_app_store_reviews
from metadata import get_play_store_metadata, get_app_store_metadata
from analysis import (
    calculate_sample_summary,
    calculate_monthly_trends,
    calculate_segment_comparisons,
    extract_issue_themes,
)
from excel_exporter import export_intelligence_workbook


