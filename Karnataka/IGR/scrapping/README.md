# Karnataka IGR (Kaveri Online Services) EC Scrapers

Automated, resilient web scrapers designed to extract property transaction records and Encumbrance Certificates (EC) from the **Karnataka Kaveri Online Services Portal** ([kaverionline.karnataka.gov.in](https://kaverionline.karnataka.gov.in/)).

This repository provides two specialized scraping architectures tailored for different operational requirements:

1. **[IGR_Karnataka_SingleBrowser.py](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_SingleBrowser.py)**: Single-instance scraper featuring **automated account rotation** to bypass the 25-search daily quota per citizen account, fine-grained CLI overrides, and granular audit logging.
2. **[IGR_Karnataka_MultiBrowser.py](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_MultiBrowser.py)**: Multi-process concurrent scraper orchestrating **multiple heterogeneous browser engines** (Chrome, Edge, Firefox), dynamic village claiming with inter-process locks, an isolated Excel-writer process, and a live ANSI terminal dashboard.

---

## Table of Contents

- [Architectural Comparison](#architectural-comparison)
- [Key Features](#key-features)
  - [Automated 2-Factor Authentication (OTP via IMAP)](#automated-2-factor-authentication-otp-via-imap)
  - [Dual-Engine Hybrid Captcha Solver](#dual-engine-hybrid-captcha-solver)
  - [Property Type 6x Batch Optimization](#property-type-6x-batch-optimization)
  - [Smart Village Completion Heuristic](#smart-village-completion-heuristic)
  - [Concurrency-Safe Atomic Excel Storage](#concurrency-safe-atomic-excel-storage)
  - [Resilience and Auto-Recovery](#resilience-and-auto-recovery)
- [File & Directory Structure](#file--directory-structure)
- [Prerequisites & System Setup](#prerequisites--system-setup)
- [Configuration](#configuration)
  - [1. User Accounts Configuration (`accounts.json`)](#1-user-accounts-configuration-accountsjson)
  - [2. Scraper Settings & Target Hierarchy](#2-scraper-settings--target-hierarchy)
- [Usage Guide](#usage-guide)
  - [Running SingleBrowser Scraper](#running-singlebrowser-scraper)
  - [Running MultiBrowser Scraper](#running-multibrowser-scraper)
- [Data Outputs & Telemetry](#data-outputs--telemetry)
- [Troubleshooting & FAQs](#troubleshooting--faqs)

---

## Architectural Comparison

| Dimension | [SingleBrowser Scraper](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_SingleBrowser.py) | [MultiBrowser Scraper](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_MultiBrowser.py) |
| :--- | :--- | :--- |
| **Execution Model** | Single process, single browser window | Multi-process (`multiprocessing`), multi-worker |
| **Browser Engines** | Google Chrome | Heterogeneous: Chrome, Edge, Firefox simultaneously |
| **Account Management** | **Automated Account Rotation**: cycles through accounts every 25 searches (`ACCOUNT_SEARCH_LIMIT = 25`) | **Dedicated Account per Worker**: each worker operates under its own citizen account |
| **Work Allocation** | Sequential survey stepping with rich CLI argument controls | Dynamic village claiming via shared memory lock (`village_lock`) |
| **Excel Persistence** | Direct atomic replace via `.tmp.xlsx` | Dedicated `ExcelSaverProcess` consuming from an IPC queue (`mp.Queue()`) |
| **Deduplication** | In-memory + Excel document reference ID checking | Real-time global cache (`manager.dict()`) synchronized across all workers |
| **Progress Tracking** | [single_progress.json](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/single_progress.json) + [single_search_history.json](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/single_search_history.json) | Per-worker state files (`igr_last_scraped_N.txt`) + global summary JSON/XLSX |
| **User Interface** | Timestamped colored console log with detailed search breakdown | Live ASCII/ANSI dashboard displaying real-time telemetry across workers |
| **Best Used For** | Targeted single-village scraping, controlled batch jobs, unattended overnight runs on single accounts | High-throughput bulk extraction across entire hoblis/talukas on multi-core PCs |

---

## Key Features

### Automated 2-Factor Authentication (OTP via IMAP)
Kaveri Online Services mandates an email OTP upon citizen login. Both scrapers implement [fetch_otp](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_SingleBrowser.py#L796):
- Connects securely via SSL to Gmail's IMAP server (`imap.gmail.com:993`) using Google App Passwords.
- Evaluates email arrival timestamps (`after_time` delta) to prevent stale OTP re-use.
- Parses verification codes from official sender templates via regular expressions (`\b\d{6}\b`).
- In multi-process mode, an IMAP process lock prevents concurrent mailbox access race conditions.

### Dual-Engine Hybrid Captcha Solver
Kaveri visual captchas use noisy backgrounds, distortion lines, and inconsistent character kerning. The scrapers use an intelligent multi-stage pipeline:
1. **Preprocessing ([keep_black_text_only](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_SingleBrowser.py#L560))**: Extracts character glyphs based on HSV/luminance thresholds, strips colored noise artifacts, and applies Otsu binarization and morphological dilation.
2. **Shape Heuristics ([_recover_missing_narrow_chars](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_SingleBrowser.py#L465))**: Analyzes connected component bounding boxes, baseline clusters, and J-hook curves to detect easily misidentified characters (e.g., `j`, `i`, `l`, `1`).
3. **Primary OCR (`ddddocr`)**: Ultra-fast deep-learning OCR model.
4. **Fallback OCR (`pytesseract`)**: Automatically triggers if `ddddocr` yields non-alphanumeric lengths or confidence dips.
5. **Auto-Refresh**: If portal validation fails, the scraper triggers [refresh_search_captcha](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_SingleBrowser.py#L728) without restarting the entire search workflow.

### Property Type 6x Batch Optimization
The portal form allows citizens to query multiple property number types using an **"Add More"** button, up to a portal limit of 6 entries per search:
- Both scripts dynamically inspect all property type dropdown options (e.g., *Survey Number*, *Site Number*, *Flat Number*, *House Number*, *Municipal No.*, etc.).
- Groups options into batches of 6 (`PROPERTY_BATCH_SIZE = 6`).
- Pre-populates all 6 rows with the active survey number before submission, **increasing search throughput by 600%** per search request.
- Discovered property types are permanently indexed in [Property_Types_Reference.xlsx](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/Property_Types_Reference.xlsx).

### Smart Village Completion Heuristic
Government revenue records do not provide an API to query the maximum survey number in a village.
- The scrapers systematically iterate through survey numbers starting from 1 (or configured resume point).
- Maintains a consecutive empty counter (`consecutive_empty`).
- When **50 consecutive survey numbers** (`MAX_CONSECUTIVE_EMPTY = 50`) return 0 transaction records across all property batches, the scraper determines that all valid land parcels in the village have been exhausted.
- Flags the village as `COMPLETED` in the summary trackers and seamlessly advances to the next village.

### Concurrency-Safe Atomic Excel Storage
Data corruption and file permission lockups are common when writing to Excel on Windows:
- **Atomic File Writing**: Writes first to `[filename].tmp.xlsx` and calls `os.replace` to prevent partial corruptions during unexpected terminations.
- **Office Lock Handler**: If a user has the target Excel spreadsheet open in Microsoft Excel, the scraper catches `PermissionError`, prints a non-blocking warning, pauses, and automatically resumes saving once the file is closed.
- **Multiprocess Queue**: In [IGR_Karnataka_MultiBrowser.py](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_MultiBrowser.py), workers do not write directly to disk; instead, they send records over `mp.Queue()` to [excel_saver_process_main](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_MultiBrowser.py#L3633).

### Resilience and Auto-Recovery
- **Angular Overlay Watchdog**: Detects and waits for portal spinner overlays (`div.loader`, `ngx-spinner`) with graceful timeouts.
- **Modal Clearing**: Auto-dismisses portal announcements, maintenance popups, and search status dialogs.
- **Session Expiry Recovery**: Re-authenticates and navigates back to the current search page with state restoration.

---

## File & Directory Structure

```text
d:/Karnataka_IGR_Scrapping/Anekal Kasba Hobli 2/SingleBrowser/
│
├── IGR_Karnataka_SingleBrowser.py     # Single-process scraper with account rotation & CLI
├── IGR_Karnataka_MultiBrowser.py      # Multi-process concurrent scraper with live dashboard
│
├── accounts.json                      # Citizen login & Gmail IMAP credentials pool
├── requirements.txt                   # Python package dependencies
│
├── Dodda_Hagade.xlsx                  # Main scraped records dataset (SingleBrowser default output)
├── Anekall Kasba Hoblli2.xlsx         # Main scraped records dataset (MultiBrowser default output)
│
├── Single_Village_Scraping_Summary.xlsx # Village completion summary (Excel)
├── single_village_scraping_summary.json # Village completion summary (JSON)
├── Single_Search_History.xlsx         # Granular per-survey audit trail (Excel)
├── single_search_history.json         # Granular per-survey audit trail (JSON)
├── single_progress.json               # Checkpoint state (account, village, survey, batch)
│
├── Property_Types_Reference.xlsx      # Auto-indexed property types reference (Excel)
├── property_types_reference.json      # Auto-indexed property types reference (JSON)
│
├── captcha_samples/                   # Debug directory for failed/successful captcha images
└── tmp_chrome_cache/                  # Browser user profile and cache isolation
```

---

## Prerequisites & System Setup

### 1. Python Environment
Python **3.10**, **3.11**, or **3.12** is recommended.
> [!NOTE]
> `ddddocr` currently supports Python `< 3.13`. On Python 3.13+, the scraper automatically falls back to `pytesseract`.

Install required Python libraries:
```bash
pip install -r requirements.txt
```

### 2. Tesseract OCR Engine (Windows)
1. Download and install Tesseract OCR from [UB-Mannheim/tesseract](https://github.com/UB-Mannheim/tesseract/wiki).
2. Default installation path: `C:\Program Files\Tesseract-OCR\tesseract.exe`.
3. If installed in an alternative location, update `CANDIDATE_TESSERACT_PATHS` in the scripts.

### 3. Web Browsers & WebDrivers
- **Google Chrome**: Ensure Chrome is updated to the latest stable version.
- **Microsoft Edge**: Standard Windows installation (used by MultiBrowser Worker-1).
- **Mozilla Firefox**: Optional (used by MultiBrowser Worker-3).
- Drivers are managed automatically by Selenium 4.x / `webdriver-manager`.

---

## Configuration

### 1. User Accounts Configuration (`accounts.json`)
The scrapers load citizen accounts and their matching Gmail IMAP credentials from [accounts.json](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/accounts.json):

```json
[
  {
    "label": "Account-1",
    "login_email": "citizen_user1@example.com",
    "login_password": "YourKaveriPassword",
    "imap_email": "citizen_user1@example.com",
    "imap_app_password": "abcd efgh ijkl mnop"
  },
  {
    "label": "Account-2",
    "login_email": "citizen_user2@example.com",
    "login_password": "YourKaveriPassword",
    "imap_email": "citizen_user2@example.com",
    "imap_app_password": "qrst uvwx yzab cdef"
  }
]
```

> [!IMPORTANT]
> `imap_app_password` **must be a 16-character Google App Password**, not your ordinary Google account login password. To generate one:
> 1. Visit your [Google Account Security Settings](https://myaccount.google.com/security).
> 2. Enable **2-Step Verification**.
> 3. Navigate to **App Passwords**, generate a password for "Mail" / "Windows Computer", and paste the 16-letter code into `imap_app_password`.

### 2. Scraper Settings & Target Hierarchy
Both scrapers expose high-level constants near the top of the file:

```python
# Location Hierarchy
TARGET_DISTRICT_NAME = "Basavanagudi"   # Registration District
TARGET_DISTRICT_INDEX = 3               # Fallback dropdown index
TARGET_TALUKA_NAME = "Anekal"           # Taluka
TARGET_TALUKA_INDEX = 1                 # Fallback dropdown index
TARGET_HOBLI_OPTION_INDEX = 3           # Option index in Hobli dropdown (3 = option[3])

# SingleBrowser Target Village Settings
TARGET_VILLAGE_NAME = "Dodda Hagade"    # Target village name or None for all
TARGET_VILLAGE_OPTION_INDEX = 9         # 1-based dropdown option index
TARGET_VILLAGE_ONLY = True              # True = exit after village; False = continue to next
ACCOUNT_SEARCH_LIMIT = 25               # Daily search quota per account
MAX_CONSECUTIVE_EMPTY = 50              # Empty survey streak to declare village completed
```

---

## Usage Guide

### Running SingleBrowser Scraper

#### Method A: Basic Run (Uses configuration defaults)
```powershell
python IGR_Karnataka_SingleBrowser.py
```
Automatically loads checkpoint from [single_progress.json](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/single_progress.json) and resumes execution.

#### Method B: Command-Line Overrides (CLI)
The SingleBrowser scraper supports command-line arguments to override configured defaults without modifying source code:

| Argument | Shorthand | Description | Example |
| :--- | :--- | :--- | :--- |
| `--village` | `-v` | Jump directly to a village by name | `python IGR_Karnataka_SingleBrowser.py -v "Dodda Hagade"` |
| `--village-index` | `-vi` | Jump directly to village by 1-based dropdown index | `python IGR_Karnataka_SingleBrowser.py -vi 5` |
| `--village-only` | | Scrape only the targeted village and stop | `python IGR_Karnataka_SingleBrowser.py -v "Dodda Hagade" --village-only` |
| `--survey` | `-s` | Force starting survey number | `python IGR_Karnataka_SingleBrowser.py -s 100` |
| `--account` | `-a` | Force starting account index (1-based) | `python IGR_Karnataka_SingleBrowser.py -a 2` |
| `--hobli-index` | `-hi` | Target specific Hobli dropdown index | `python IGR_Karnataka_SingleBrowser.py -hi 3` |
| `--list-villages` | | Connects, lists all available villages in the Hobli, and exits | `python IGR_Karnataka_SingleBrowser.py --list-villages` |

**Common Scenarios:**
- **Inspect all villages in the Hobli:**
  ```powershell
  python IGR_Karnataka_SingleBrowser.py --list-villages
  ```
- **Scrape a specific village from Survey 1 to completion:**
  ```powershell
  python IGR_Karnataka_SingleBrowser.py -v "Dodda Hagade" --village-only -s 1
  ```
- **Force account rotation to start with Account-3:**
  ```powershell
  python IGR_Karnataka_SingleBrowser.py -a 3
  ```

---

### Running MultiBrowser Scraper

Launch concurrent scraping workers across multiple browser engines:

```powershell
python IGR_Karnataka_MultiBrowser.py
```

#### What Happens at Launch:
1. **Deduplication Preload**: Inspects existing records in the output Excel sheet ([Anekall Kasba Hoblli2.xlsx](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/Anekall%20Kasba%20Hoblli2.xlsx)) and preloads all Document Reference IDs into a shared hash table.
2. **Dedicated Excel Saver**: Spawns [excel_saver_process_main](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_MultiBrowser.py#L3633) in the background.
3. **Staggered Worker Launches**: Launches workers with a 15-second stagger delay (`STAGGER_DELAY_SECONDS = 15`) to prevent concurrent OTP requests from colliding in the same mailbox.
4. **Dynamic Village Allocation**: Workers query the Hobli dropdown and claim unassigned villages using an inter-process lock (`village_lock`).
5. **Live Dashboard**: The terminal transitions to a colorized progress monitor:

```text
==================================================================================================
📊 LIVE KAVERI SCRAPER PROGRESS MONITOR [16:45:10] | Saved Records: 1420 | Completed: 3/28 Villages
--------------------------------------------------------------------------------------------------
Worker             Engine   Status         Current Village          Survey #   Streak     Records
--------------------------------------------------------------------------------------------------
Worker-1-Edge      EDGE     SCRAPING       Dodda Hagade             42         0          312 docs
Worker-2-Chrome    CHROME   SCRAPING       Aravantike Pura          18         4          195 docs
Worker-3-firefox   FIREFOX  SCRAPING       Maranayakanahalli        77         12         84 docs
--------------------------------------------------------------------------------------------------
🏆 Completed Villages (3): Samanduru(450 docs), Marsur(280 docs), Haragadde(689 docs)
==================================================================================================
```

#### Graceful Termination:
Press `Ctrl + C` in the terminal. The parent process safely terminates browser child processes, drains the data queue to flush remaining records into Excel, and safely shuts down.

---

## Data Outputs & Telemetry

### 1. Primary Scraped Datasets (`*.xlsx`)
Contains all extracted transaction fields including:
- **Registration Details**: Document Reference Number (`ದಸ್ತಾವೇಜಿನ ಉಲ್ಲೇಖ`), Registration Date, Execution Date, Presentation Date.
- **Property Details**: Survey Number, Property Type, Extent / Area, Boundaries (East, West, North, South).
- **Financial Details**: Consideration Amount, Market Value, Stamp Duty Paid, Registration Fee.
- **Parties**: Executant (Seller / Transferor) names, Claimant (Buyer / Transferee) names.
- **Location Meta**: District, Taluka, Hobli, Village.

### 2. Village Scraping Summary (`*Village_Scraping_Summary.xlsx` / `.json`)
Maintains high-level progress per village:
- `Village Name`, `Hobli`, `Taluka`, `District`
- `Last Survey Searched`: Highest survey number checked.
- `Last Survey Scraped`: Highest survey number that yielded transactions.
- `Total Records`: Total transactions extracted for the village.
- `Status`: `IN_PROGRESS` or `COMPLETED`.

### 3. Granular Search History (`Single_Search_History.xlsx` / `.json`)
*(SingleBrowser only)* Records a line-item entry for every batch search submitted to the portal:
- Timestamp, Account Label, Village, Survey Number, Batch Index, Records Found, Search Status (`SUCCESS`, `EMPTY`, `SESSION_TIMEOUT`, `CAPTCHA_ERROR`).

---

## Troubleshooting & FAQs

### Q1: IMAP OTP authentication fails or times out
- **Symptom**: Console logs `❌ OTP not found via IMAP within timeout`.
- **Cause**: Incorrect Gmail App Password, 2-Step Verification disabled, or emails arriving with unexpected delay.
- **Fix**:
  1. Confirm `imap_app_password` in [accounts.json](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/accounts.json) is an active 16-character Google App Password.
  2. Verify that Kaveri registration emails are arriving in the inbox of `imap_email`.
  3. Increase OTP wait retries in [fetch_otp](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_SingleBrowser.py#L796) (`retries=15`, `delay=6`).

### Q2: Captcha recognition fails repeatedly
- **Symptom**: Scraper continuously clicks "Refresh Captcha" or receives invalid captcha alerts.
- **Cause**: Missing Tesseract binaries or Python version incompatibility with `ddddocr`.
- **Fix**:
  1. If running Python 3.13+, install Tesseract OCR and verify `tesseract --version` in terminal.
  2. Inspect failed captcha images saved in `captcha_samples/` to evaluate noise levels.

### Q3: `PermissionError` when saving to Excel
- **Symptom**: `⚠️ [FILE LOCKED] 'Dodda_Hagade.xlsx' is open in Microsoft Excel!`
- **Behavior**: The scraper catches this error and pauses without losing data.
- **Fix**: Close the Excel spreadsheet. The scraper will detect the file release and automatically complete the write operation.

### Q4: Browser driver crashes or disconnects unexpectedly
- **Symptom**: `WebDriverException: chrome not reachable` or connection reset.
- **Fix**: Both scripts feature automatic crash handlers:
  - In [SingleBrowser](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_SingleBrowser.py), [is_driver_alive](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_SingleBrowser.py#L241) checks browser health and rebuilds the driver if closed.
  - In [MultiBrowser](file:///d:/Karnataka_IGR_Scrapping/Anekal%20Kasba%20Hobli%202/SingleBrowser/IGR_Karnataka_MultiBrowser.py), worker processes recover up to 30 crash iterations (`max_worker_crashes = 30`) before aborting.

---

## License & Usage Disclaimer
This software is developed strictly for automated research, compliance verification, and public land record aggregation under permitted fair-use guidelines of publicly accessible registry records. Ensure all scraping activities comply with local terms of service, server load policies, and data usage regulations.
