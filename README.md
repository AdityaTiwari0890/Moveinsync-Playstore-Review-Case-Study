# App Review Intelligence Scraper

An enterprise-ready Python solution to scrape, analyze, and benchmark mobile application customer reviews from the **Google Play Store** and **Apple App Store**. Designed specifically for **MOVEIN SYNC** employee-side apps, driver-side apps, and industry competitors, with full reusability for any mobile application.

> Review datasets, SQLite databases, and generated reports are intentionally excluded from this repository because they can contain reviewer identifiers and verbatim customer comments. Provide the source workbooks locally when running the analysis.

---

## Key Features

### Step 1: Intelligent Multi-App Scraper
1. **Multi-App & Role-Based Collection**:
   - Collect reviews across multiple applications in a single run.
   - Categorize each app by role: `employee`, `driver`, or `competitor`.
   - Supports independent review targets per app (e.g., 500 reviews, up to 10,000).
2. **Interactive Search & Confirmation Protocol**:
   - Searches live store listings and displays candidates with developer name, store ID, and rating.
   - Never makes unconfirmed guesses; requires explicit confirmation.
3. **Configurable Date Range Filtering**:
   - Filter reviews by the most recent **6 months**, **12 months**, or a custom `YYYY-MM-DD` date range.
4. **Initial 6-Sheet Formatted Workbook**:
   - Exports raw datasets (`Android_Reviews`, `iOS_Reviews`, `Combined_Reviews`, `App_Summary`, `Issue_Themes`, `Collection_Log`).

### Step 2: Enterprise App Review Analysis Engine
1. **Data Normalization & Deduplication**:
   - Ingests existing scraper workbooks, normalizes columns, parses dates, and deduplicates review IDs without modifying raw sheets.
2. **Statistical Rating Reliability**:
   - Sample sizes ($N$) displayed beside all metrics.
   - **Wilson Score 95% Confidence Intervals** calculated for negative review proportions.
   - Configurable sample thresholds flagging low-sample versions ($N < 15$).
   - Segment comparisons: Employee vs. Driver, Android vs. iOS, MoveInSync vs. Competitors.
3. **11-Category Multi-Label Complaint Taxonomy**:
   - Classifies reviews into: *Cab delays and no-shows*, *Booking and OTP failures*, *GPS and location tracking*, *App crashes and performance*, *Login and account access*, *UI and usability*, *Roster and scheduling*, *Driver-side experience*, *Billing and payment*, *Customer support*, and *Other or unclear*.
   - Multi-labeling enabled; heuristic confidence scored (`High`, `Medium`, `Low`); explicit `Unclassified` state for ambiguous text.
4. **Pareto 80/20 Negative Defect Analysis**:
   - Evaluates 1★ and 2★ negative reviews to identify the vital few themes contributing to ~70–80% of customer dissatisfaction.
   - Generates native openpyxl Pareto BarCharts and Cumulative Percentage lines.
5. **Version Trajectory & Regression Tracking**:
   - Tracks rating and negative rate trajectories across consecutive releases ($N \ge 15$).
   - Flags drops ($\Delta \text{Rating} \le -0.30★$ or $\Delta \text{Neg Rate} \ge +15.0\%$) with review ID evidence and excerpts.
   - Transparently documents findings as observational associations, not causal proofs.
6. **Action-Priority Matrix (`Action_Priorities`)**:
   - Ranks root causes into P1–P3 priorities with user impact severity, evidence confidence, estimated effort, recommended fixes, and success metrics.
7. **Executive Markdown Report**:
   - Compiles comprehensive executive intelligence report (`MOVEIN_SYNC_ANALYSIS_REPORT.md`).

---

## Project Structure

```
App Review Intelligence Scraper & Analysis Engine
├── main.py                     # CLI interface & interactive terminal wizard
├── play_store_scraper.py       # Google Play Store scraper (continuation tokens, batching, retry)
├── app_store_scraper.py        # Apple App Store scraper (RSS multi-page + SerpApi optional)
├── metadata.py                 # Public store listing metadata collector
├── classifier.py               # 11-category multi-label classifier with evidence tracing
├── analysis_engine.py          # Step 2 analytical engine (stats, Pareto, versions, priorities, charts)
├── excel_exporter.py           # Openpyxl 6-sheet raw Excel exporter
├── test_scraper.py             # Unit & integration tests for scraper
├── test_analysis_engine.py     # Unit & integration tests for analysis engine (8 tests)
├── MOVEIN_SYNC_ANALYSIS_REPORT.md # Generated executive analysis report
├── requirements.txt            # Pinned dependencies
├── .env.example                # Environment variables template
