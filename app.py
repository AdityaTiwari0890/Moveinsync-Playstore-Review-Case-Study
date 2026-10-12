from __future__ import annotations

import json
import sqlite3
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

try:
    import plotly.express as px
    import streamlit as st
except ImportError:  # pragma: no cover - exercised when dashboard dependencies are absent
    st = None
    px = None

from ratings_watch.config import (
    EXISTING_WORKBOOK,
    IOS_APP_ID,
    OLLAMA_EMBEDDING_MODEL,
    OLLAMA_MODEL,
    STEP3_ACTION_PLAN,
    STEP3_MD,
    VECTOR_STORE_PATH,
    load_settings,
)
from ratings_watch.database import (
    acknowledge_alert,
    get_connection,
    init_db,
    resolve_alert,
    update_action_status,
)
from ratings_watch.monitor import collect_reviews, import_existing_workbook
from ratings_watch.rag_ingestion import count_pending_documents, index_knowledge_base
from ratings_watch.retriever import ask_analyst
from ratings_watch.llm_assistant import LLMAssistant
from ratings_watch.rag_ingestion import inspect_knowledge_sources
from ratings_watch.vector_store import get_index_status

PAGES = [
    "Overview",
    "Rating Trends",
    "Review Explorer",
    "Complaint Intelligence",
    "Alerts",
    "Action Plan",
    "AI Analysis Assistant",
    "Knowledge Base",
    "Data Collection Status",
]


def _rows(db_path: str, query: str, parameters: Tuple[Any, ...] = ()) -> List[Dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        return [dict(row) for row in conn.execute(query, parameters)]
    finally:
        conn.close()


def _filter_options(db_path: str) -> Dict[str, List[str]]:
    theme_values = set()
    for row in _rows(db_path, "SELECT assigned_themes FROM reviews WHERE assigned_themes IS NOT NULL"):
        theme_values.update(_theme_list(row["assigned_themes"]))
    return {
        "platforms": [row["platform"] for row in _rows(db_path, "SELECT DISTINCT platform FROM reviews ORDER BY platform")],
        "apps": [row["app_name"] for row in _rows(db_path, "SELECT DISTINCT app_name FROM reviews WHERE app_name IS NOT NULL ORDER BY app_name")],
        "roles": [row["app_role"] for row in _rows(db_path, "SELECT DISTINCT app_role FROM reviews WHERE app_role IS NOT NULL ORDER BY app_role")],
        "versions": [row["app_version"] for row in _rows(db_path, "SELECT DISTINCT app_version FROM reviews WHERE app_version IS NOT NULL ORDER BY app_version")],
        "store_ids": [row["store_id"] for row in _rows(db_path, "SELECT DISTINCT store_id FROM reviews WHERE store_id IS NOT NULL ORDER BY store_id")],
        "themes": sorted(theme_values),
    }


def _selected_filters(db_path: str) -> Dict[str, Any]:
    options = _filter_options(db_path)
    st.sidebar.header("Analysis filters")
    platform = st.sidebar.selectbox("Platform", ["All platforms"] + options["platforms"])
    app_name = st.sidebar.selectbox("App", ["All apps"] + options["apps"])
    role = st.sidebar.selectbox("App role", ["All roles"] + options["roles"])
    store_id = st.sidebar.selectbox("Store identifier", ["All identifiers"] + options["store_ids"])
    theme = st.sidebar.selectbox("Complaint theme", ["All themes"] + options["themes"])
    date_range = st.sidebar.date_input("Review date range", value=(), key="rw_date_range")
    version = st.sidebar.selectbox("App version", ["All versions"] + options["versions"])
    filters = {
        "platform": None if platform == "All platforms" else platform,
        "app_name": None if app_name == "All apps" else app_name,
        "app_role": None if role == "All roles" else role,
        "store_id": None if store_id == "All identifiers" else store_id,
        "theme": None if theme == "All themes" else theme,
        "app_version": None if version == "All versions" else version,
        "start_date": date_range[0].isoformat() if isinstance(date_range, (tuple, list)) and len(date_range) > 0 else None,
        "end_date": date_range[-1].isoformat() if isinstance(date_range, (tuple, list)) and len(date_range) > 1 else None,
    }
    st.sidebar.caption("All sample metrics below use collected review records and the selected filters.")
    return filters


def _review_query(db_path: str, filters: Dict[str, Any], limit: Optional[int] = None) -> List[Dict[str, Any]]:
    clauses = []
    parameters: List[Any] = []
    for key in ("platform", "app_name", "app_role", "app_version", "store_id"):
        if filters.get(key):
            clauses.append(f"{key} = ?")
            parameters.append(filters[key])
    if filters.get("start_date"):
        clauses.append("review_date >= ?")
        parameters.append(filters["start_date"])
    if filters.get("end_date"):
        clauses.append("review_date <= ?")
        parameters.append(filters["end_date"])
    query = "SELECT * FROM reviews" + (" WHERE " + " AND ".join(clauses) if clauses else "") + " ORDER BY review_date DESC"
    if limit is not None:
        query += " LIMIT ?"
        parameters.append(limit)
    return _rows(db_path, query, tuple(parameters))


def _theme_list(raw: Any) -> List[str]:
    try:
        values = json.loads(raw or "[]") if isinstance(raw, str) else raw
    except (ValueError, TypeError):
        values = []
    return [str(value) for value in values] if isinstance(values, list) else []


def _filtered_review_data(db_path: str, filters: Dict[str, Any]) -> pd.DataFrame:
    rows = _review_query(db_path, filters)
    data = pd.DataFrame(rows)
    if filters.get("theme") and not data.empty:
        data = data[data["assigned_themes"].apply(lambda value: filters["theme"] in _theme_list(value))]
    return data


def _theme_rows(data: pd.DataFrame, negative_only: bool = False) -> pd.DataFrame:
    if data.empty:
        return pd.DataFrame(columns=["theme", "review_count", "average_rating"])
    if negative_only and "rating" in data:
        data = data[data["rating"] <= 2]
    flattened = []
    for _, row in data.iterrows():
        for theme in _theme_list(row.get("assigned_themes")):
            flattened.append({"theme": theme, "rating": row.get("rating"), "platform": row.get("platform")})
    if not flattened:
        return pd.DataFrame(columns=["theme", "review_count", "average_rating"])
    return (
        pd.DataFrame(flattened)
        .groupby("theme", as_index=False)
        .agg(review_count=("rating", "size"), average_rating=("rating", "mean"))
        .sort_values("review_count", ascending=False)
    )


def _render_chat(db_path: str, vector_path: str, key_suffix: str, filters: Optional[Dict[str, Any]] = None) -> None:
    chat_key = f"rw_chat_{key_suffix}"
    st.session_state.setdefault(chat_key, [])
    for item in st.session_state[chat_key]:
        with st.chat_message(item["role"]):
            st.markdown(item["content"])
            if item.get("evidence"):
                with st.expander(f"Evidence ({len(item['evidence'])} sources)"):
                    for evidence in item["evidence"]:
                        metadata = evidence["metadata"]
                        identifier = metadata.get("review_id") or metadata.get("source_name")
                        st.markdown(
                            f"**[{evidence['citation']}] {identifier}** · "
                            f"{metadata.get('platform') or metadata.get('sheet') or metadata.get('section') or metadata.get('source_type')}"
                        )
                        st.caption(evidence["content"][:1200])
            if item.get("metrics"):
                with st.expander("Verified structured metrics and filters"):
                    st.json({"filters": item.get("filters"), "metrics": item["metrics"]})
    prompt = st.chat_input("Ask about ratings, complaints, evidence, or recommendations", key=f"chat_input_{key_suffix}")
    if prompt:
        previous = st.session_state[chat_key]
        st.session_state[chat_key].append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner("Retrieving evidence and checking structured metrics..."):
                result = ask_analyst(
                    prompt,
                    db_path=db_path,
                    vector_path=vector_path,
                    conversation=previous,
                    context_filters=filters,
                )
            if result.get("ok"):
                st.markdown(result["answer"])
                if result.get("evidence"):
                    with st.expander(f"Evidence ({len(result['evidence'])} sources)"):
                        for evidence in result["evidence"]:
                            metadata = evidence["metadata"]
                            identifier = metadata.get("review_id") or metadata.get("source_name")
                            st.markdown(
                                f"**[{evidence['citation']}] {identifier}** · "
                                f"{metadata.get('platform') or metadata.get('sheet') or metadata.get('section') or metadata.get('source_type')}"
                            )
                            st.caption(evidence["content"][:1200])
                with st.expander("Verified structured metrics and filters"):
                    st.json({"filters": result.get("filters"), "metrics": result.get("metrics")})
                st.session_state[chat_key].append(
                    {
                        "role": "assistant",
                        "content": result["answer"],
                        "evidence": result.get("evidence", []),
                        "metrics": result.get("metrics"),
                        "filters": result.get("filters"),
                    }
                )
            else:
                st.error(result.get("error", "The assistant could not answer this question."))


def _overview(db_path: str, vector_path: str, filters: Dict[str, Any]) -> None:
    st.title("MOVEINSYNC Ratings Watch")
    st.caption("Evidence-led review analytics. Collected review averages and published store ratings are reported separately.")
    data = _filtered_review_data(db_path, filters)
    total = len(data)
    average = float(data["rating"].mean()) if total and "rating" in data else None
    negative_count = int((data["rating"] <= 2).sum()) if total else 0
    negative_share = negative_count * 100.0 / total if total else None
    alerts = _rows(db_path, "SELECT COUNT(*) AS count FROM alerts WHERE resolved=0")
    latest_run = _rows(db_path, "SELECT platform, completed_at, status FROM collection_runs WHERE status IN ('success','partial') ORDER BY completed_at DESC LIMIT 1")
    snapshots = _rows(
        db_path,
        """SELECT s.platform, s.official_rating, s.total_ratings_count, s.rating_source, s.snapshot_at
           FROM rating_snapshots s JOIN (
             SELECT platform, MAX(snapshot_at) AS latest FROM rating_snapshots
             WHERE official_rating IS NOT NULL GROUP BY platform
           ) latest ON s.platform=latest.platform AND s.snapshot_at=latest.latest ORDER BY s.platform""",
    )
    cols = st.columns(5)
    cols[0].metric("Collected reviews", f"{total:,}", help="Review records in SQLite after the current filters.")
    cols[1].metric("Sample average", f"{average:.2f} / 5" if average is not None else "N/A", help="Arithmetic mean of collected review ratings, not a store listing rating.")
    cols[2].metric("Negative sample share", f"{negative_share:.1f}%" if negative_share is not None else "N/A", help="Percent of selected collected review records rated 1–2 stars.")
    cols[3].metric("Active alerts", int(alerts[0]["count"]))
    cols[4].metric("Latest collection", latest_run[0]["completed_at"] if latest_run else "N/A")
    if snapshots:
        st.subheader("Latest recorded store-rating snapshots")
        snapshot_cols = st.columns(min(4, len(snapshots)))
        for index, snapshot in enumerate(snapshots):
            snapshot_cols[index % len(snapshot_cols)].metric(
                f"{snapshot['platform']} snapshot",
                f"{snapshot['official_rating']:.2f}",
                help=f"Recorded source: {snapshot['rating_source']}; timestamp: {snapshot['snapshot_at']}",
            )
    else:
        st.info("No verified published store-rating snapshots are available. The collected-review average above is not a substitute.")
    if data.empty:
        st.warning("No review records match the current filters. Import the analyzed workbook or collect supported platform data.")
    else:
        left, right = st.columns(2)
        with left:
            st.subheader("Review volume by publication month")
            monthly = data.copy()
            monthly["month"] = pd.to_datetime(monthly["review_date"], errors="coerce").dt.to_period("M").astype(str)
            volume = monthly.dropna(subset=["month"]).groupby("month").size().rename("reviews").reset_index()
            if not volume.empty:
                st.plotly_chart(px.bar(volume, x="month", y="reviews", title="Reviews by review date"), use_container_width=True)
            else:
                st.info("Review publication dates are unavailable for this selection.")
        with right:
            st.subheader("Collected review rating distribution")
            distribution = data.groupby("rating").size().rename("review_count").reset_index()
            st.plotly_chart(px.bar(distribution, x="rating", y="review_count", title="Collected review sample"), use_container_width=True)
        st.subheader("Most frequent classified themes")
        themes = _theme_rows(data, negative_only=True).head(10)
        if themes.empty:
            st.info("No classified negative-review themes are available for this selection.")
        else:
            st.plotly_chart(px.bar(themes, x="review_count", y="theme", orientation="h"), use_container_width=True)
        comparisons = (
            data.groupby(["app_name", "app_role", "store_id", "platform"], dropna=False)
            .agg(review_count=("rating", "size"), sample_average=("rating", "mean"), negative_count=("rating", lambda values: int((values <= 2).sum())))
            .reset_index()
        )
        comparisons["negative_share_percent"] = comparisons["negative_count"] / comparisons["review_count"] * 100
        st.subheader("App and platform comparison")
        st.caption("Comparisons describe the available collected samples; they are not causal comparisons and may have different sample sizes.")
        st.dataframe(comparisons, use_container_width=True, hide_index=True)
    st.subheader("Ask MOVEINSYNC AI")
    st.caption("Answers are based on indexed project evidence and deterministic SQLite metrics. AI needs a working Ollama model and a current index.")
    _render_chat(db_path, vector_path, "overview", filters)


def _rating_trends(db_path: str, filters: Dict[str, Any]) -> None:
    st.title("Rating Trends")
    st.caption("Review-date charts aggregate imported and collected records by their publication date; they are not official store-rating histories.")
    data = _filtered_review_data(db_path, filters)
    if data.empty:
        st.info("No review data matches these filters.")
        return
    data = data.copy()
    data["review_date"] = pd.to_datetime(data["review_date"], errors="coerce")
    data = data.dropna(subset=["review_date"])
    if data.empty:
        st.info("No valid review publication dates are available.")
        return
    data["month"] = data["review_date"].dt.to_period("M").astype(str)
    trend = data.groupby("month").agg(average_rating=("rating", "mean"), review_count=("rating", "size")).reset_index()
    left, right = st.columns(2)
    with left:
        st.plotly_chart(px.line(trend, x="month", y="average_rating", markers=True, title="Average of collected reviews by publication month"), use_container_width=True)
    with right:
        st.plotly_chart(px.bar(trend, x="month", y="review_count", title="Review volume by publication month"), use_container_width=True)
    data["negative"] = data["rating"] <= 2
    neg = data.groupby("month").agg(negative_share=("negative", "mean")).reset_index()
    neg["negative_share"] *= 100
    st.plotly_chart(px.line(neg, x="month", y="negative_share", markers=True, title="Negative-review share (1–2 stars)"), use_container_width=True)
    st.subheader("Platform and app sample comparison")
    comparison = (
        _filtered_review_data(db_path, {**filters, "platform": None})
        .groupby(["app_name", "app_role", "store_id", "platform"], dropna=False)
        .agg(review_count=("rating", "size"), sample_average=("rating", "mean"), negative_share=("rating", lambda values: (values <= 2).mean() * 100))
        .reset_index()
    )
    st.dataframe(comparison, use_container_width=True, hide_index=True)
    if "app_version" in data:
        versions = data[data["app_version"].notna()].groupby("app_version").agg(
            review_count=("rating", "size"),
            sample_average=("rating", "mean"),
            negative_share_percent=("rating", lambda values: (values <= 2).mean() * 100),
        ).reset_index()
        if not versions.empty:
            st.subheader("App-version sample comparison")
            st.caption("Version labels come from review records and are not release-time causal analysis.")
            st.dataframe(versions, use_container_width=True, hide_index=True)
    snapshot_rows = _rows(
        db_path,
        "SELECT platform, official_rating, rating_source, snapshot_at FROM rating_snapshots WHERE official_rating IS NOT NULL ORDER BY snapshot_at",
    )
    if len(snapshot_rows) >= 2:
        snapshot_data = pd.DataFrame(snapshot_rows)
        st.plotly_chart(px.line(snapshot_data, x="snapshot_at", y="official_rating", color="platform", markers=True, title="Recorded published-rating snapshots"), use_container_width=True)
    else:
        st.info("Historical store-rating trend will appear after at least two genuine rating snapshots have been recorded.")


def _review_explorer(db_path: str, filters: Dict[str, Any]) -> None:
    st.title("Review Explorer")
    data = _filtered_review_data(db_path, filters)
    if data.empty:
        st.info("No reviews match the current filters.")
        return
    search = st.text_input("Search original review text")
    if search:
        data = data[data["review_text"].fillna("").str.contains(search, case=False, regex=False)]
    rating_options = sorted(int(value) for value in data["rating"].dropna().unique())
    selected_ratings = st.multiselect("Rating", rating_options, default=rating_options)
    if selected_ratings:
        data = data[data["rating"].isin(selected_ratings)]
    if not data.empty:
        display_fields = [
            field for field in (
                "review_id", "platform", "app_name", "app_role", "rating", "review_date",
                "app_version", "assigned_themes", "heuristic_confidence", "review_text",
                "is_unclassified",
                "developer_reply", "source_url", "first_collected_at", "last_updated_at", "is_imported",
                "source_collection_timestamp",
            ) if field in data.columns
        ]
        st.caption(f"{len(data):,} matching reviews (display capped at 2,000). Review text is unchanged from the source.")
        export = data[display_fields].copy()
        export["assigned_themes"] = export["assigned_themes"].apply(lambda item: ", ".join(_theme_list(item)))
        st.download_button("Export filtered reviews (CSV)", export.to_csv(index=False).encode("utf-8-sig"), "filtered_reviews.csv", "text/csv")
        st.dataframe(export.head(2000), use_container_width=True, hide_index=True)
        if "review_id" in data.columns:
            selected_id = st.selectbox("Review detail", data["review_id"].astype(str).tolist())
            selected = data[data["review_id"].astype(str) == selected_id].iloc[0]
            with st.expander("Full review and source metadata", expanded=True):
                st.write(str(selected.get("review_text") or ""))
                st.json(
                    {
                        key: (None if pd.isna(selected.get(key)) else selected.get(key))
                        for key in (                                                "review_id", "platform", "app_name", "app_role", "rating", "review_date", "app_version", "assigned_themes", "heuristic_confidence", "is_unclassified", "source_url", "developer_reply", "developer_reply_date", "source_collection_timestamp", "first_collected_at", "last_updated_at", "is_imported")
                    }
                )
    else:
        st.info("No reviews remain after the search and rating filters.")


def _complaints(db_path: str, filters: Dict[str, Any]) -> None:
    st.title("Complaint Intelligence")
    data = _filtered_review_data(db_path, filters)
    themes = _theme_rows(data, negative_only=True)
    if themes.empty:
        st.info("No theme-classified negative reviews match the current filters.")
        return
    negative_total = int((data["rating"] <= 2).sum()) if not data.empty else 0
