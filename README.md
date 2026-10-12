# App Review Intelligence Scraper

An enterprise-ready Python solution to scrape, analyze, and benchmark mobile application customer reviews from the **Google Play Store** and **Apple App Store**. Designed specifically for **MOVEIN SYNC** employee-side apps, driver-side apps, and industry competitors, with full reusability for any mobile application.

> Review datasets, SQLite databases, and generated reports are intentionally excluded from this repository because they can contain reviewer identifiers and verbatim customer comments. Provide the source workbooks locally when running the analysis.

Report - https://drive.google.com/file/d/10fzLKUOrdGiXHgQp12I6hxeaHLjaURKM/view?usp=sharing
Reviews with analysis - https://docs.google.com/spreadsheets/d/1kIvhi35SveEbQI1DqjfE51ZOhAOuPTZG/edit?usp=sharing&ouid=118233679977606631667&rtpof=true&sd=true
Case Study resources - https://drive.google.com/drive/folders/1zPzSy4mqzQubb9BnXEfaz8cA5fHsMNE2?usp=sharing
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
     <img width="1530" height="790" alt="image" src="https://github.com/user-attachments/assets/ac9a03dd-58c0-4a2b-9e3a-de91207f61f3" />
     <img width="1535" height="787" alt="image" src="https://github.com/user-attachments/assets/c83cb7e3-5ce3-4715-bb69-018651fdaaa4" />
     <img width="1535" height="801" alt="image" src="https://github.com/user-attachments/assets/82a27233-13b2-4cb6-a00e-066902bef08d" />
     <img width="1535" height="801" alt="image" src="https://github.com/user-attachments/assets/fb770abe-9683-465d-a1e0-feec9e62fce5" />
     <img width="1535" height="793" alt="image" src="https://github.com/user-attachments/assets/c6239cfe-e207-4518-9c8d-b2b344a9c6c8" />
     <img width="1535" height="793" alt="image" src="https://github.com/user-attachments/assets/fb085fea-026a-4e3a-b2e7-ded2d8c45fc6" />
     <img width="1535" height="791" alt="image" src="https://github.com/user-attachments/assets/ba419ca0-8d55-491b-b745-d655d01ef338" />
     <img width="1535" height="791" alt="image" src="https://github.com/user-attachments/assets/3d81bc1d-6675-4302-bee7-954d1339e344" />
     <img width="1535" height="792" alt="image" src="https://github.com/user-attachments/assets/adb3a1c5-9c64-48b5-8950-1032061cd0b8" />
     <img width="1535" height="792" alt="image" src="https://github.com/user-attachments/assets/f943ec23-4317-4f78-81b9-498d1f0dff61" />
     <img width="1535" height="790" alt="image" src="https://github.com/user-attachments/assets/36ae85e3-5e0a-44ad-be87-b218930c507d" />
     <img width="1535" height="790" alt="image" src="https://github.com/user-attachments/assets/579f5290-7ff9-4f63-9cc2-311133e2f8a2" />
     <img width="1532" height="791" alt="image" src="https://github.com/user-attachments/assets/b40a6574-00dd-4b34-811e-037ae5cfe718" />
     <img width="1535" height="790" alt="image" src="https://github.com/user-attachments/assets/c6c8a794-db23-457a-bca8-f88b18f8b18e" />
     <img width="1535" height="790" alt="image" src="https://github.com/user-attachments/assets/1145f2f3-5ff5-4a0e-8397-4cfa7e93ca3f" />
     

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
<img width="1528" height="762" alt="image" src="https://github.com/user-attachments/assets/af8d8ceb-60e5-4897-b17e-76c9f069b0e6" />
<img width="1535" height="771" alt="image" src="https://github.com/user-attachments/assets/8c5a5896-2a98-4262-bbca-331c6baf69f3" />
<img width="1535" height="772" alt="image" src="https://github.com/user-attachments/assets/5e8d3c9a-bed8-43b1-9d27-bb875aecb63f" />
<img width="1532" height="748" alt="image" src="https://github.com/user-attachments/assets/5cbbaea6-d63a-4321-a012-7822b8199b35" />
<img width="1531" height="772" alt="image" src="https://github.com/user-attachments/assets/61da627a-f931-4a49-a8cf-016825f6d1b1" />
<img width="1532" height="766" alt="image" src="https://github.com/user-attachments/assets/0a6c1291-6d10-4385-ae0f-b8c9a459d51e" />
<img width="1532" height="766" alt="image" src="https://github.com/user-attachments/assets/1127a4c6-fc07-4b7f-afe2-0ecf345fb297" />

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
└── README.md                   # User guide & PowerShell execution commands
```

---

## Getting Started (Windows PowerShell)

Follow these exact commands in **Windows PowerShell**:

### 1. Create and Activate Virtual Environment

```powershell
# Navigate to the project root directory
cd "c:\Users\aktiw\Downloads\Performance Analyzer"

# Create a virtual environment named 'venv'
python -m venv venv

# Activate the virtual environment
.\venv\Scripts\Activate.ps1
```

> *Tip: If PowerShell prevents script execution, run:*
> `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`

### 2. Install Dependencies

```powershell
pip install -r requirements.txt
```

### 3. Optional Environment Configuration

If you have a SerpApi key for extended iOS review scraping beyond Apple's 500-review RSS limit:

```powershell
Copy-Item .env.example .env
# Open .env and insert your SERPAPI_API_KEY=your_key_here
```

---

## Execution Guide

### Option A: Interactive Terminal Wizard (Recommended)

Run the script without arguments to enter the interactive setup wizard:

```powershell
python main.py
```

The wizard will interactively prompt you for:
1. Application Name (e.g., `MoveInSync`, `Routematic`, `Uber`).
2. Role Category (`employee`, `driver`, `competitor`).
3. Platform (`Both Android & iOS`, `Android only`, `iOS only`).
4. Store URLs / Package IDs (or auto-search matching store listings).
5. Number of reviews (default: 500, min: 100, max: 10,000).
6. Option to queue additional apps (e.g., driver app or competitors) in the same run.
7. Confirmation to proceed and output workbook path.

---

### Option B: Command-Line (CLI) Mode

Run directly with command-line arguments for automated or scripted execution:

#### 1. MoveInSync Employee App (Android Only, 500 Reviews)
```powershell
python main.py --app "MoveInSync" --play-url "com.moveinsync.ets" --platform android --count 500 --output "MoveInSync_Employee_Reviews.xlsx"
```

#### 2. MoveInSync Driver App (Android Only, Last 6 Months)
```powershell
python main.py --app "MoveInSync Driver" --play-url "com.moveinsync.driver.dco" --role driver --platform android --months 6 --output "MoveInSync_Driver_6M.xlsx"
```

#### 3. Both Android & iOS with Candidate Auto-Search
```powershell
python main.py --app "MoveInSync" --platform both --count 200 --output "MoveInSync_CrossPlatform.xlsx"
```

#### 4. Date Range Filter (e.g., Last 12 Months)
```powershell
python main.py --app "MoveInSync" --play-url "com.moveinsync.ets" --months 12 --count 1000
```

#### 5. Custom Date Range
```powershell
python main.py --app "MoveInSync" --play-url "com.moveinsync.ets" --start-date "2026-01-01" --end-date "2026-10-12"
```

---

### Option C: Step 2 — Running the Analysis Engine

To analyze any existing scraper output workbook (such as `App_Review_Intelligence_MoveInSync_20261012_173118.xlsx`):

#### 1. Standard Analysis Run
```powershell
python analysis_engine.py --input "App_Review_Intelligence_MoveInSync_20261012_173118.xlsx"
```
*Outputs:*
- Enhanced 11-sheet workbook: `App_Review_Intelligence_MoveInSync_20261012_173118_Analyzed.xlsx`
- Executive report: `MOVEIN_SYNC_ANALYSIS_REPORT.md`

#### 2. Custom Output Path & Sample Threshold
```powershell
python analysis_engine.py --input "App_Review_Intelligence_MoveInSync_20261012_173118.xlsx" --output "MoveInSync_Intelligence_Report.xlsx" --min-sample 20
```

---

## Standardized Review Schema

All collected reviews conform to the following schema across both platforms:

| Field | Description | Example |
|---|---|---|
| `App Name` | Application display name | `MoveInSync` |
| `Store ID` | Package name (Android) or Numeric ID (iOS) | `com.moveinsync.ets` / `962949521` |
| `Platform` | Operating system platform | `Android` / `iOS` |
| `App Role` | Segment classification | `employee` / `driver` / `competitor` |
| `Review ID` | Platform review ID or deterministic hash | `<store-review-id>` |
| `Reviewer Name` | Name supplied by the store, when available | `<reviewer-name>` |
| `Rating` | Star rating (1 to 5) | `5` |
| `Review Title` | Review title (iOS only; blank for Android) | `Smooth commute app` |
| `Review Text` | Full customer review commentary | `<review text>` |
| `Review Date` | Review date and time (`YYYY-MM-DD HH:MM:SS`) | `<review timestamp>` |
| `App Version` | App version active when review was written | `<app version>` |
| `Developer Reply` | Text of developer's reply (if present) | `<developer reply>` |
| `Developer Reply Date` | Date of developer reply (if present) | `<reply timestamp>` |
| `Helpful Votes` | Number of thumbs-up / helpful votes | `<vote count>` |
| `Collection Timestamp`| Exact timestamp when scraping occurred | `<collection timestamp>` |
| `Source URL` | Direct store listing URL | `https://play.google.com/store/apps/details?id=...` |

---

## Enhanced Excel Workbook Layout (11 Sheets)

### Original Scraper Sheets (Preserved Intact)
1. **`Android_Reviews`**: Raw data table for all scraped Android reviews.
2. **`iOS_Reviews`**: Raw data table for all scraped iOS reviews.
3. **`Combined_Reviews`**: Unified master dataset with all platforms, sorted newest first.
4. **`App_Summary`**: Public store metadata, sample metrics, segment matrices, and monthly counts.
5. **`Issue_Themes`**: Scraper-level keyword themes.
6. **`Collection_Log`**: Scraper audit trail and API limitations.

### Step 2 Analysis Engine Sheets (Appended with Charts)
7. **`Rating_Trends`**: Rating distributions, Wilson score 95% confidence intervals, monthly trajectory, embedded openpyxl **Star Rating BarChart**, and **Monthly Trajectory LineChart**.
8. **`Theme_Analysis`**: 11-category complaint breakdown (Total mentions, % of all reviews, average rating, employee vs. driver breakdown, and clean excerpts).
9. **`Version_Analysis`**: Software version metrics ($N$, average rating, negative %, Wilson CI, regression flags, and evidence quotes).
10. **`Pareto_Analysis`**: 1★ & 2★ negative review defect ranking, negative share %, cumulative share %, Pareto vital few classification, and embedded **Pareto BarChart**.
11. **`Action_Priorities`**: Executive prioritization matrix with P1–P3 rankings, user impact severity, evidence confidence, estimated effort, recommended fixes, and success metrics.

---

## Running Automated Tests

### 1. Run Step 2 Analysis Engine Tests (8 Tests)
```powershell
python test_analysis_engine.py
```

### 2. Run Step 1 Scraper Tests (6 Tests)
```powershell
python test_scraper.py
```

Expected output:
```text
Ran 8 tests in ~1s -> OK
Ran 6 tests in ~5s -> OK
```

---

## Known API Limitations

1. **Apple App Store Public RSS Limits**:
   - Apple's official customer-review RSS feed is limited to **10 pages × 50 reviews = 500 reviews per country**.
   - If more than 500 iOS reviews are requested, the scraper collects all 500 available reviews and documents this limitation in the `Collection_Log`. Optional SerpApi support can be enabled via `.env` to pull additional iOS reviews.
2. **Historical Version Release Notes**:
   - Public store pages only display the *current* version release date. Neither Google nor Apple expose historical release date changelogs on public listings without developer console authentication. This is formally noted in `metadata.py` and the `App_Summary` sheet.

---

## Step 4 — MOVEINSYNC Ratings Watch

Step 4 is an additive local analytical and monitoring workspace. It reads, but does not overwrite, the Step 1–3 workbooks and reports. SQLite stores review records, registered app targets, collection runs, published-rating snapshots, alerts, and tracked Step 3 actions. A separate SQLite vector store persists Ollama embeddings for hybrid retrieval.

### Quick start

```powershell
cd "C:\Users\aktiw\Downloads\Performance Analyzer"
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
if (-not (Test-Path .env)) { Copy-Item .env.example .env }
notepad .env
python monitor.py import-existing --input "App_Review_Intelligence_MoveInSync_20261012_173118_Analyzed.xlsx"
python monitor.py import-step3
python monitor.py status
python monitor.py ai-check
streamlit run app.py
```

The analyzed workbook currently contains 1,115 collected reviews across Android and iOS, including MoveInSync employee and driver app targets. The targets and counts are discovered from the workbook; this is a historical sample, not a claim that a live collection just ran. In the existing workbook, the iOS employee target is `962949521`, the iOS driver target is `389801252`, Android employee targets include `com.moveinsync.ets` and `field.moveinsync.ets`, and the Android driver target is `com.moveinsync.driver.dco`.

### Dashboard pages and analytics

- **Overview:** filtered sample KPIs, latest recorded published-rating snapshots (if available), app/platform comparison, review volume, rating distribution, negative-review themes, and a contextual AI panel.
- **Rating Trends:** review-date monthly aggregates, review volume, negative share, actual recorded store-rating snapshots, and app-version sample comparisons. Review-date aggregates are not official-rating history.
- **Review Explorer:** searchable and filterable source reviews with metadata and CSV export.
- **Complaint Intelligence:** negative-theme frequency, overlapping multi-label share, platform matrix, and supporting review IDs/excerpts.
- **Alerts:** alert history with acknowledge and resolve controls.
- **Action Plan:** import and track the existing Step 3 priority matrix without editing the source workbook.
- **AI Analysis Assistant:** hybrid semantic retrieval plus deterministic SQLite metrics, with visible evidence snippets and source labels.
- **Knowledge Base:** incremental index/rebuild controls, indexed-document counts, pending updates, model, storage path, last indexing time, and errors.
- **Data Collection Status:** app registry, source coverage, collection attempts, workbook import, and manual store collection.

### Configuration and hosted Ollama

The checked-in `.env.example` contains placeholders only. Do not overwrite an existing `.env`; copy the template only if no `.env` exists, then edit it locally. `API_KEY` is the canonical hosted Ollama secret. The Python server sends it only as `Authorization: Bearer <API_KEY>`; it is never displayed in the dashboard or included in logs. A missing or placeholder key is rejected before any hosted request is sent.

```powershell
if (-not (Test-Path .env)) { Copy-Item .env.example .env }
notepad .env
```

Supported settings include:

```text
API_KEY=your_ollama_api_key
OLLAMA_BASE_URL=https://ollama.com/api
OLLAMA_MODEL=gemma4:31b
OLLAMA_EMBEDDING_MODEL=embeddinggemma
DATABASE_PATH=ratings_watch.db
VECTOR_STORE_PATH=ratings_watch_vectors.db
APP_STORE_ID=com.moveinsync.ets
IOS_APP_ID=962949521
RATING_DROP_THRESHOLD=0.5
THEME_SPIKE_THRESHOLD=25.0
MIN_ALERT_SAMPLE_SIZE=10
```

For local Ollama, use `OLLAMA_BASE_URL=http://localhost:11434` and install the configured chat and embedding models locally; a local API key is not required. The hosted endpoint follows Ollama's documented direct API (`https://ollama.com/api/chat` and `/api/embed`). Verify that both configured models are available to your Ollama account before indexing; the application does not silently substitute models. Documentation: [Ollama API authentication](https://docs.ollama.com/api/authentication), [Ollama chat API](https://docs.ollama.com/api/chat), and [Ollama embeddings API](https://docs.ollama.com/api/embed).

If Ollama is unavailable, collection, imports, metrics, alerts, review exploration, and action tracking still work. Semantic indexing and AI answers require the configured embedding and chat endpoints.

### RAG indexing and assistant

The project implements retrieval-augmented generation, not fine-tuning. It embeds workbook review records, each populated analysis workbook sheet, Step 2/3 Markdown reports, the Step 3 action-plan workbook, and current SQLite reviews, alerts, and actions. Stable document IDs and content hashes support incremental updates. The local SQLite vector store retains embedding arrays and metadata; retrieval ranks candidates by cosine similarity and filters by available platform, app, role, theme, and date metadata. Exact counts and rating metrics are computed by fixed, parameterized SQLite queries; the model cannot execute SQL or mutate data. The assistant receives only a bounded set of retrieved snippets and a compact metrics bundle. Retrieved review text is treated as untrusted data.

Relative values for `DATABASE_PATH`, `VECTOR_STORE_PATH`, `EXISTING_WORKBOOK`, `STEP3_MD`, and `STEP3_ACTION_PLAN` are resolved from the project directory, independent of the shell's current directory. `ratings_watch_vectors.db` is a SQLite database containing persisted document content, source metadata, and real embedding arrays. Readiness counts include only persisted documents with valid vectors; discovered records alone are not reported as indexed.

```powershell
# Import the historical review sample and Step 3 actions first.
python monitor.py import-existing
python monitor.py import-step3

# Test chat and embeddings independently; the command never displays the API key.
python monitor.py ai-check

# Build or incrementally update the semantic index; --rebuild replaces it only after embeddings succeed.
python monitor.py index-knowledge-base
python monitor.py index-knowledge-base --rebuild
python monitor.py status
```

The dashboard's **Knowledge Base** page offers **Test Ollama Connection**, **Index Knowledge Base**, **Rebuild Knowledge Base**, and **Refresh Knowledge-Base Status** controls. An absent/placeholder hosted key blocks requests and indexing with a clear error. Chat and embedding checks execute their actual API operations when configured. A rebuild embeds the complete replacement before the SQLite transaction replaces the existing index.

The index can be created only when the configured embedding service is reachable. The first real Ollama indexing run may send review/report text to the configured Ollama endpoint. Use a local Ollama endpoint if the corpus must remain entirely on this machine.

### Collection and monitoring commands

The Step 4 adapters reuse the existing paginated Step 1 Google Play collector and Apple's public customer-review RSS collector. The existing public iOS feed has a limit of up to 500 reviews per country. iOS live collection can return fewer records than requested. Each source's success is represented by the recorded collection run; merely configuring an ID is not proof of success. Public listing metadata is queried independently and a published-rating snapshot is saved only when the metadata lookup returns a rating. A sample mean is never saved as a store rating.

```powershell
# Collect from one known registered target (repeat with another verified app ID/role as needed).
python monitor.py collect --platform android --app-id "com.moveinsync.ets" --app-name "MoveInSync" --app-role employee --count 200
python monitor.py collect --platform ios --app-id "962949521" --app-name "MoveInSync" --app-role employee --count 200

# Evaluate and store alerts, inspect status, or run collect -> alert evaluation -> incremental indexing.
python monitor.py check-alerts
python monitor.py status
python monitor.py run-monitor --platform android --app-id "com.moveinsync.ets" --app-name "MoveInSync" --app-role employee --count 200
```

Alerts compare timestamped snapshots and equivalent seven-day review windows with minimum sample safeguards. Identical active alerts are deduplicated and held through their configured cooldown. Theme and negative-share comparisons use negative review records; multi-label themes can overlap. A resolved alert is not proof that an issue was fixed.

### Windows Task Scheduler (daily monitoring)

The following PowerShell creates a user-level task that runs independently of Streamlit. Review the target ID and time before executing; task creation is intentionally left to the operator.

```powershell
$project = "C:\Users\aktiw\Downloads\Performance Analyzer"
$python = Join-Path $project "venv\Scripts\python.exe"
$arguments = "`"$project\monitor.py`" run-monitor --platform android --app-id `"com.moveinsync.ets`" --app-name `"MoveInSync`" --app-role employee --count 200"
$action = New-ScheduledTaskAction -Execute $python -Argument $arguments -WorkingDirectory $project
$trigger = New-ScheduledTaskTrigger -Daily -At 7:00AM
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName "MOVEINSYNC Ratings Watch" -Action $action -Trigger $trigger -Principal $principal

# Verify and optionally run once:
Get-ScheduledTask -TaskName "MOVEINSYNC Ratings Watch"
Start-ScheduledTask -TaskName "MOVEINSYNC Ratings Watch"
Get-ScheduledTaskInfo -TaskName "MOVEINSYNC Ratings Watch"
```

`run-monitor` reports collection, alert, and indexing outcomes separately. A missing Ollama endpoint can make the knowledge-base update partial without stopping review collection or alert evaluation.

### Tests and known limitations

```powershell
python -m unittest tests.test_step4 tests.test_rag
python test_analysis_engine.py
```

- Historical iOS reviews are present in the workbook and both existing Step 1 public collectors are reused. Live store access was not re-run as part of these tests.
- Official rating snapshots appear only after the store metadata lookup provides a rating; the historic analyzed workbook is not treated as snapshot history.
- Step 3 is imported from the actual `Priority_Matrix` table, which has title/header rows; source-row fields are retained for evidence and tracking.
- SQLite cosine ranking is intended for this local-scale corpus. It loads candidate vectors into process memory for ranking and is not a multi-user or large-corpus vector database.
- Semantic retrieval quality and hosted-model behavior still require the operator's real Ollama model/endpoint. Automated RAG tests use deterministic mocked embeddings/chat and do not claim hosted service connectivity.
- Current source platform data covers MoveInSync employee and driver variants in the workbook; the registry discovers those real IDs. Additional products are supported by adding/importing their verified identifiers and records.


