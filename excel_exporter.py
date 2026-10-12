"""
Excel Exporter Engine
Generates a multi-sheet Microsoft Excel workbook (.xlsx) using pandas and openpyxl.
Includes 6 standardized sheets: Android_Reviews, iOS_Reviews, Combined_Reviews,
App_Summary, Issue_Themes, and Collection_Log with tables, filters, frozen panes,
auto-fitted column widths, and wrapped text.
"""

from typing import List, Dict, Any, Optional
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd


# Color Palette
NAVY_HEADER = "1B365D"
SECTION_HEADER = "2C4C64"
WHITE_TEXT = "FFFFFF"
BORDER_GRAY = "D9D9D9"


def export_intelligence_workbook(
    android_reviews: List[Dict[str, Any]],
    ios_reviews: List[Dict[str, Any]],
    combined_reviews: List[Dict[str, Any]],
    metadata_list: List[Dict[str, Any]],
    summary_metrics: List[Dict[str, Any]],
    monthly_trends: List[Dict[str, Any]],
    segment_comparisons: Dict[str, Any],
    issue_themes: List[Dict[str, Any]],
    collection_logs: List[Dict[str, Any]],
    output_filepath: str,
) -> str:
    """
    Build the complete 6-sheet formatted Excel workbook.
    """
    wb = openpyxl.Workbook()
    # Remove default empty sheet
    default_sheet = wb.active
    if default_sheet is not None:
        wb.remove(default_sheet)

    # Header style helper
    header_font = Font(name="Calibri", size=11, bold=True, color=WHITE_TEXT)
    header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    section_fill = PatternFill(start_color=SECTION_HEADER, end_color=SECTION_HEADER, fill_type="solid")
    regular_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style="thin", color=BORDER_GRAY),
        right=Side(style="thin", color=BORDER_GRAY),
        top=Side(style="thin", color=BORDER_GRAY),
        bottom=Side(style="thin", color=BORDER_GRAY),
    )

    # -------------------------------------------------------------
    # 1. Sheet 1: Android_Reviews
    # -------------------------------------------------------------
    ws_android = wb.create_sheet(title="Android_Reviews")
    android_cols = [
        "App Name", "Store ID", "Platform", "App Role", "Review ID",
        "Reviewer Name", "Rating", "Review Text", "Review Date",
        "App Version", "Developer Reply", "Developer Reply Date",
        "Helpful Votes", "Collection Timestamp", "Source URL"
    ]
    ws_android.append(android_cols)
    for r in android_reviews:
        ws_android.append([
            r.get("app_name"), r.get("store_id"), r.get("platform"), r.get("app_role"),
            r.get("review_id"), r.get("reviewer_name"), r.get("rating"),
            r.get("review_text"), r.get("review_date"), r.get("app_version"),
            r.get("developer_reply"), r.get("developer_reply_date"),
            r.get("helpful_votes"), r.get("collection_timestamp"), r.get("source_url")
        ])

    # -------------------------------------------------------------
    # 2. Sheet 2: iOS_Reviews
    # -------------------------------------------------------------
    ws_ios = wb.create_sheet(title="iOS_Reviews")
    ios_cols = [
        "App Name", "Store ID", "Platform", "App Role", "Review ID",
        "Reviewer Name", "Rating", "Review Title", "Review Text", "Review Date",
        "App Version", "Developer Reply", "Developer Reply Date",
        "Helpful Votes", "Collection Timestamp", "Source URL"
    ]
    ws_ios.append(ios_cols)
    for r in ios_reviews:
        ws_ios.append([
            r.get("app_name"), r.get("store_id"), r.get("platform"), r.get("app_role"),
            r.get("review_id"), r.get("reviewer_name"), r.get("rating"),
            r.get("review_title"), r.get("review_text"), r.get("review_date"),
            r.get("app_version"), r.get("developer_reply"), r.get("developer_reply_date"),
            r.get("helpful_votes"), r.get("collection_timestamp"), r.get("source_url")
        ])

    # -------------------------------------------------------------
    # 3. Sheet 3: Combined_Reviews
    # -------------------------------------------------------------
    ws_combined = wb.create_sheet(title="Combined_Reviews")
    combined_cols = [
        "App Name", "Store ID", "Platform", "App Role", "Review ID",
        "Reviewer Name", "Rating", "Review Title", "Review Text", "Review Date",
        "App Version", "Developer Reply", "Developer Reply Date",
        "Helpful Votes", "Collection Timestamp", "Source URL"
    ]
    ws_combined.append(combined_cols)
    for r in combined_reviews:
        ws_combined.append([
            r.get("app_name"), r.get("store_id"), r.get("platform"), r.get("app_role"),
            r.get("review_id"), r.get("reviewer_name"), r.get("rating"),
            r.get("review_title"), r.get("review_text"), r.get("review_date"),
            r.get("app_version"), r.get("developer_reply"), r.get("developer_reply_date"),
            r.get("helpful_votes"), r.get("collection_timestamp"), r.get("source_url")
        ])

    # -------------------------------------------------------------
    # 4. Sheet 4: App_Summary
    # -------------------------------------------------------------
    ws_summary = wb.create_sheet(title="App_Summary")

    # Section 1: Public Store Listing Metadata
    ws_summary.append(["SECTION 1: PUBLIC STORE LISTING METADATA"])
    ws_summary.cell(row=1, column=1).font = Font(name="Calibri", size=12, bold=True, color=WHITE_TEXT)
    ws_summary.cell(row=1, column=1).fill = section_fill
    meta_cols = [
        "Store Title", "Developer", "Store ID", "Platform", "Category",
        "Installs", "Store Rating", "Total Store Ratings", "Current Version",
        "Last Updated Date", "Store URL", "Limitation Note"
    ]
    ws_summary.append(meta_cols)
    for m in metadata_list:
        ws_summary.append([
            m.get("title"), m.get("developer"), m.get("store_id"), m.get("platform"),
            m.get("category"), m.get("installs"), m.get("store_rating"),
            m.get("store_ratings_count"), m.get("current_version"),
            m.get("last_updated"), m.get("store_url"), m.get("limitation_note")
        ])

    ws_summary.append([])  # Spacer

    # Section 2: Scraped Sample Review Analytics
    r_idx = ws_summary.max_row + 1
    ws_summary.cell(row=r_idx, column=1, value="SECTION 2: SCRAPED SAMPLE REVIEW METRICS")
    ws_summary.cell(row=r_idx, column=1).font = Font(name="Calibri", size=12, bold=True, color=WHITE_TEXT)
    ws_summary.cell(row=r_idx, column=1).fill = section_fill

    sample_cols = [
        "App Name", "Platform", "App Role", "Requested Count", "Collected Count",
        "Target Reached", "Sample Avg Rating", "1★ Count", "1★ %", "2★ Count", "2★ %",
        "3★ Count", "3★ %", "4★ Count", "4★ %", "5★ Count", "5★ %",
        "Negative % (1★+2★)", "Positive % (4★+5★)", "Earliest Review Date",
        "Most Recent Review Date", "Collection Date"
    ]
    ws_summary.append(sample_cols)
    for sm in summary_metrics:
        ws_summary.append([
            sm.get("app_name"), sm.get("platform"), sm.get("app_role"),
            sm.get("requested_count"), sm.get("total_reviews"), sm.get("target_reached"),
            sm.get("avg_rating"),
            sm.get("star_1_count"), f"{sm.get('star_1_pct')}%",
            sm.get("star_2_count"), f"{sm.get('star_2_pct')}%",
            sm.get("star_3_count"), f"{sm.get('star_3_pct')}%",
            sm.get("star_4_count"), f"{sm.get('star_4_pct')}%",
            sm.get("star_5_count"), f"{sm.get('star_5_pct')}%",
            f"{sm.get('negative_pct')}%", f"{sm.get('positive_pct')}%",
            sm.get("earliest_date"), sm.get("most_recent_date"), sm.get("collection_date")
        ])

    ws_summary.append([])  # Spacer

    # Section 3: Segment Comparison Matrices
    r_idx = ws_summary.max_row + 1
    ws_summary.cell(row=r_idx, column=1, value="SECTION 3: SEGMENT COMPARISON MATRICES")
    ws_summary.cell(row=r_idx, column=1).font = Font(name="Calibri", size=12, bold=True, color=WHITE_TEXT)
    ws_summary.cell(row=r_idx, column=1).fill = section_fill

    ws_summary.append(["Segment", "Sub-Segment", "Collected Reviews", "Sample Avg Rating", "Negative % (1★+2★)", "Positive % (4★+5★)"])
    evd = segment_comparisons.get("employee_vs_driver", {})
    p_comp = segment_comparisons.get("platform_comparison", {})
    mvc = segment_comparisons.get("moveinsync_vs_competitors", {})

    ws_summary.append(["Role Comparison", "Employee Apps", evd.get("employee", {}).get("count", 0), evd.get("employee", {}).get("avg_rating"), f"{evd.get('employee', {}).get('negative_pct', 0.0)}%", f"{evd.get('employee', {}).get('positive_pct', 0.0)}%"])
    ws_summary.append(["Role Comparison", "Driver Apps", evd.get("driver", {}).get("count", 0), evd.get("driver", {}).get("avg_rating"), f"{evd.get('driver', {}).get('negative_pct', 0.0)}%", f"{evd.get('driver', {}).get('positive_pct', 0.0)}%"])
    ws_summary.append(["Platform Comparison", "Android", p_comp.get("android", {}).get("count", 0), p_comp.get("android", {}).get("avg_rating"), f"{p_comp.get('android', {}).get('negative_pct', 0.0)}%", f"{p_comp.get('android', {}).get('positive_pct', 0.0)}%"])
    ws_summary.append(["Platform Comparison", "iOS", p_comp.get("ios", {}).get("count", 0), p_comp.get("ios", {}).get("avg_rating"), f"{p_comp.get('ios', {}).get('negative_pct', 0.0)}%", f"{p_comp.get('ios', {}).get('positive_pct', 0.0)}%"])
    ws_summary.append(["Market Benchmark", "MoveInSync Ecosystem", mvc.get("moveinsync", {}).get("count", 0), mvc.get("moveinsync", {}).get("avg_rating"), f"{mvc.get('moveinsync', {}).get('negative_pct', 0.0)}%", f"{mvc.get('moveinsync', {}).get('positive_pct', 0.0)}%"])
    ws_summary.append(["Market Benchmark", "Competitor Apps", mvc.get("competitors", {}).get("count", 0), mvc.get("competitors", {}).get("avg_rating"), f"{mvc.get('competitors', {}).get('negative_pct', 0.0)}%", f"{mvc.get('competitors', {}).get('positive_pct', 0.0)}%"])

    ws_summary.append([])  # Spacer

    # Section 4: Monthly Review Trajectory
    r_idx = ws_summary.max_row + 1
    ws_summary.cell(row=r_idx, column=1, value="SECTION 4: MONTHLY REVIEW VOLUME & RATING TRAJECTORY")
    ws_summary.cell(row=r_idx, column=1).font = Font(name="Calibri", size=12, bold=True, color=WHITE_TEXT)
    ws_summary.cell(row=r_idx, column=1).fill = section_fill

    monthly_cols = ["Year-Month", "App Name", "Platform", "Review Count", "Sample Avg Rating", "Negative Count (1★+2★)", "Negative %"]
    ws_summary.append(monthly_cols)
    for mt in monthly_trends:
        ws_summary.append([
            mt.get("year_month"), mt.get("app_name"), mt.get("platform"),
            mt.get("review_count"), mt.get("avg_rating"),
            mt.get("negative_reviews_count"), f"{mt.get('negative_percentage')}%"
        ])

    # -------------------------------------------------------------
    # 5. Sheet 5: Issue_Themes
    # -------------------------------------------------------------
    ws_themes = wb.create_sheet(title="Issue_Themes")
    theme_cols = [
        "Theme Name", "Type / Sentiment", "Review Count", "% of Total Reviews",
        "Average Rating", "Representative Review IDs", "Example Review Excerpts",
        "Classification Methodology"
    ]
    ws_themes.append(theme_cols)
    for it in issue_themes:
        ws_themes.append([
            it.get("theme"), it.get("type"), it.get("review_count"),
            f"{it.get('percentage_of_total')}%", it.get("average_rating"),
            it.get("representative_review_ids"), it.get("example_excerpts"),
            it.get("classification_method")
        ])

    # -------------------------------------------------------------
    # 6. Sheet 6: Collection_Log
    # -------------------------------------------------------------
    ws_logs = wb.create_sheet(title="Collection_Log")
    log_cols = [
        "Timestamp", "App Name", "Platform", "Store ID", "Status",
        "Requested Count", "Retrieved Count", "Data Source Used",
        "Message / Audit Details", "Known API Limitations"
    ]
    ws_logs.append(log_cols)
    for lg in collection_logs:
        ws_logs.append([
            lg.get("timestamp"), lg.get("app_name"), lg.get("platform"),
            lg.get("store_id"), lg.get("status"), lg.get("requested_count"),
            lg.get("retrieved_count"), lg.get("data_source"),
            lg.get("message"), lg.get("limitation_note")
        ])

    # -------------------------------------------------------------
    # Post-Processing: Table Formatting, Freezing, Column Auto-Fit
    # -------------------------------------------------------------
    sheets_to_format_as_table = [
        (ws_android, "Table_Android_Reviews", len(android_reviews)),
        (ws_ios, "Table_iOS_Reviews", len(ios_reviews)),
        (ws_combined, "Table_Combined_Reviews", len(combined_reviews)),
        (ws_themes, "Table_Issue_Themes", len(issue_themes)),
        (ws_logs, "Table_Collection_Log", len(collection_logs)),
    ]

    for ws, table_name, data_count in sheets_to_format_as_table:
        max_col = ws.max_column
        max_row = max(ws.max_row, 2)  # Excel tables require at least 1 data row

        # Style header row
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Freeze header row
        ws.freeze_panes = "A2"

        # Auto-filter and Excel table creation
        tab_range = f"A1:{get_column_letter(max_col)}{max_row}"
        tab = Table(displayName=table_name, ref=tab_range)
        tab_style = TableStyleInfo(
            name="TableStyleLight9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        tab.tableStyleInfo = tab_style
        ws.add_table(tab)

    # Style App_Summary sheet headers
    for row in ws_summary.iter_rows():
        for cell in row:
            cell.font = regular_font
            cell.border = thin_border
            # Section titles styling
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("SECTION "):
                cell.font = Font(name="Calibri", size=11, bold=True, color=WHITE_TEXT)
                cell.fill = section_fill
            # Sub-table headers
            elif cell.row in (2, 8, 17, 26):
                cell.font = header_font
                cell.fill = header_fill

    ws_summary.freeze_panes = "A3"

    # Auto-adjust column widths with max clamping across all sheets
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                cell.font = regular_font
                cell.border = thin_border
                cell.alignment = Alignment(vertical="top", wrap_text=True)

                if cell.value:
                    val_str = str(cell.value)
                    # Check first line length for wrapped strings
                    first_line = val_str.split("\n")[0]
                    max_len = max(max_len, min(len(first_line), 80))

            # Apply clamped width (min 12, max 65)
            ws.column_dimensions[col_letter].width = max(12, min(max_len + 3, 65))

    wb.save(output_filepath)
    return output_filepath
