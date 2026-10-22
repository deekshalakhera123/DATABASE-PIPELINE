# 🏙️ Dubai_Srapper01.py — Documentation & User Guide

A dedicated, standalone Python automation script to extract official real estate open data from the [Dubai Land Department Portal](https://dubailand.gov.ae/en/open-data/real-estate-data/#/) and automatically upload the exported CSV files to Google Drive (via OAuth 2.0 or Service Account).

Everything is self-contained in **`Dubai_Srapper01.py`**—no auxiliary helper scripts are needed.

---

## 📌 Table of Contents

- [Overview & Capabilities](#-overview--capabilities)
- [Prerequisites](#-prerequisites)
- [Quick Start Installation](#-quick-start-installation)
- [How to Run `Dubai_Srapper01.py`](#-how-to-run-dubai_srapper01py)
  - [1. Interactive Mode](#1-interactive-mode-default)
  - [2. Command-Line Arguments Mode](#2-command-line-arguments-mode-automated)
  - [CLI Flags Reference](#cli-flags-reference)
- [Configuration Settings (Inside `Dubai_Srapper01.py`)](#-configuration-settings-inside-dubai_srapper01py)
  - [Date & Scraping Defaults](#1-date--scraping-defaults)
  - [Google Drive Cloud Storage Setup](#2-google-drive-cloud-storage-setup)
    - [Option A: Service Account (Recommended for Automated Tasks)](#option-a-service-account-recommended-for-automated-tasks)
    - [Option B: OAuth 2.0 (Browser Login)](#option-b-oauth-20-browser-login)
    - [Option C: Local Only (Offline / No Google Drive)](#option-c-local-only-offline--no-google-drive)
- [Code Architecture (Section-by-Section)](#-code-architecture-section-by-section)
- [How CAPTCHA Handling Works](#-how-captcha-handling-works)
- [Visual Highlighting & UX](#-visual-highlighting--ux)
- [Troubleshooting & FAQs](#-troubleshooting--faqs)

---

## 🚀 Overview & Capabilities

`Dubai_Srapper01.py` automates the end-to-end extraction and cloud synchronization process:

1. **Launches Chrome** with optimized download and anti-detection configurations.
2. **Navigates to the DLD Portal** and switches to the specified data module (`Transactions`, `Rents`, `Project`, `Valuations`, `Building`, or `Developer`).
3. **Fills Date Filters** (`From Date` and `To Date`) with input event dispatching.
4. **Handles Google reCAPTCHA**:
   - Clicks the checkbox inside the reCAPTCHA iframe.
   - Detects auto-resolves (instant green tick).
   - If an image challenge puzzle appears, enters an active polling loop with a live terminal countdown giving you time to solve it manually in Chrome.
5. **Submits Search** and dynamically waits for the database to process and return results.
6. **Downloads the CSV Dataset** to your local `downloads/` folder (verifying file completion and ignoring partial `.crdownload` files).
7. **Uploads to Google Drive** via Google Drive API v3 (Service Account or OAuth 2.0) and optionally copies to a local Google Drive Desktop sync folder.

---

## 📋 Prerequisites

- **Python**: Version `3.8` or newer.
- **Google Chrome**: Installed on your system.
- **Google Cloud Project** *(Optional, only if uploading to Google Drive)*:
  - Enabled **Google Drive API**.

---

## 📦 Quick Start Installation

1. **Activate your Python environment**:
   - **Windows (PowerShell)**:
     ```powershell
     .\.venv\Scripts\Activate.ps1
     ```
   - **Windows (Command Prompt)**:
     ```cmd
     .\.venv\Scripts\activate.bat
     ```
   - **Linux / macOS**:
     ```bash
     source .venv/bin/activate
     ```

2. **Install Required Libraries**:
   ```bash
   pip install selenium webdriver-manager google-api-python-client google-auth-oauthlib google-auth-httplib2
   ```
   *(Or if using `requirements.txt`: `pip install -r requirements.txt`)*

---

## 💻 How to Run `Dubai_Srapper01.py`

### 1. Interactive Mode (Default)

Run the script without arguments. An interactive terminal menu guides you through dates, tab selection, and headless mode with pre-filled defaults:

```bash
python Dubai_Srapper01.py
```

**Terminal Prompt Example**:
```text
===========================================================================
 🏙️   DUBAI LAND DEPARTMENT - REAL ESTATE DATA SCRAPER & UPLOADER
===========================================================================

📋 [Terminal Interactive Setup]
---------------------------------------------------------------------------
👉 Enter FROM Date (DD/MM/YYYY) [Default: 01/01/2025]: 
👉 Enter TO Date   (DD/MM/YYYY) [Default: 15/09/2026]: 
👉 Run browser in Background / Headless mode? (y/N) [Default: N]: n
---------------------------------------------------------------------------
🎯 Selected Configuration:
   • Module        : Real Estate Transactions
   • From Date     : 01/01/2025
   • To Date       : 15/09/2026
   • Headless      : False
   • GDrive Upload : Enabled (Folder ID: 1ABCxyz...)
---------------------------------------------------------------------------
⚡ Press [ENTER] to start scraping...
```

---

### 2. Command-Line Arguments Mode (Automated)

You can pass arguments to bypass interactive prompts—ideal for task schedulers, background jobs, or batch runs:

- **Custom Date Range**:
  ```bash
  python Dubai_Srapper01.py --from-date 01/01/2024 --to-date 31/12/2024
  ```

- **Scrape Different Tabs (e.g. Rents, Valuations, Projects)**:
  ```bash
  python Dubai_Srapper01.py --from-date 01/06/2024 --to-date 01/06/2025 --tab Rents
  ```

- **Run in Background / Headless Mode**:
  ```bash
  python Dubai_Srapper01.py --from-date 01/01/2025 --to-date 01/03/2025 --headless
  ```

- **Run Non-Interactively with Script Defaults**:
  ```bash
  python Dubai_Srapper01.py --non-interactive
  ```

- **Custom CAPTCHA Timeout (e.g. 180 seconds)**:
  ```bash
  python Dubai_Srapper01.py --from-date 01/01/2025 --to-date 15/09/2026 --captcha-timeout 180
  ```

---

### CLI Flags Reference

| Parameter | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `--from-date` | `string` | `"01/01/2025"` | Start date in `DD/MM/YYYY` format. |
| `--to-date` | `string` | `"15/09/2026"` | End date in `DD/MM/YYYY` format. |
| `--tab` | `string` | `"Transactions"` | Target tab: `Transactions`, `Rents`, `Project`, `Valuations`, `Building`, `Developer`. |
| `--headless` | `flag` | `False` | Run Chrome in background headless mode. |
| `--non-interactive` | `flag` | `False` | Skip terminal input prompts and immediately run with config/CLI parameters. |
| `--captcha-timeout` | `int` | `120` | Max seconds to wait for manual image CAPTCHA solving. |

---

## ⚙️ Configuration Settings (Inside `Dubai_Srapper01.py`)

All customizable settings are located in **`[SECTION 2]`** (lines 80–158) of `Dubai_Srapper01.py`.

### 1. Date & Scraping Defaults

```python
# ------------------------------------------------------------------------------
# 2.1 Date Range Settings (Format: DD/MM/YYYY)
# ------------------------------------------------------------------------------
FROM_DATE = "01/01/2025"
TO_DATE = "15/09/2026"

# ------------------------------------------------------------------------------
# 2.2 Portal & Web Settings
# ------------------------------------------------------------------------------
WEBSITE_URL = "https://dubailand.gov.ae/en/open-data/real-estate-data/#/"

# Available: "Transactions", "Rents", "Project", "Valuations", "Building", "Developer"
TARGET_TAB = "Transactions"

# Run Chrome visibly (False) or in the background (True)
HEADLESS = False

# Maximum wait time in seconds for manual CAPTCHA solving
CAPTCHA_TIMEOUT = 120

# Local download folder name
DOWNLOAD_DIR_NAME = "downloads"
```

---

### 2. Google Drive Cloud Storage Setup

In `[SECTION 2.3]` of `Dubai_Srapper01.py`:

```python
GOOGLE_DRIVE_FOLDER_ID = ""  # <-- Paste your Google Drive Folder URL or raw ID here
AUTH_TYPE = "oauth"          # Options: "oauth" or "service_account"
```

> The script includes `extract_folder_id()`, which accepts **either** a raw folder ID (`1ABCxyz...`) or a full browser link (`https://drive.google.com/drive/folders/1ABCxyz...`).

#### Option A: Service Account (Recommended for Automated Tasks)
1. In the [Google Cloud Console](https://console.cloud.google.com/), enable the **Google Drive API**.
2. Go to **APIs & Services** ➔ **Credentials** ➔ **Create Credentials** ➔ **Service Account**.
3. Under the created service account, go to the **Keys** tab ➔ **Add Key** ➔ **Create new key** ➔ **JSON**.
4. Save the downloaded JSON file as `service_account.json` in the same directory as `Dubai_Srapper01.py`.
5. **Important**: Open your target Google Drive folder in your browser, click **Share**, and grant the Service Account email address (e.g. `your-sa@project.iam.gserviceaccount.com`) **Editor** permissions.
6. In `Dubai_Srapper01.py`, set:
   ```python
   AUTH_TYPE = "service_account"
   SERVICE_ACCOUNT_FILE = "service_account.json"
   ```

#### Option B: OAuth 2.0 (Browser Login)
1. In the [Google Cloud Console](https://console.cloud.google.com/), enable the **Google Drive API**.
2. Configure your **OAuth Consent Screen** (User Type: External; add your Gmail under **Test Users**).
3. Go to **Credentials** ➔ **Create Credentials** ➔ **OAuth Client ID** ➔ Application Type: **Desktop App**.
4. Download the JSON, rename it to `credentials.json`, and place it in the same directory as `Dubai_Srapper01.py`.
5. In `Dubai_Srapper01.py`, set:
   ```python
   AUTH_TYPE = "oauth"
   OAUTH_CREDENTIALS_FILE = "credentials.json"
   ```
6. On the first run, a browser tab opens asking you to sign in with your Google account. A `token.json` file is saved locally to allow future automatic logins.

#### Option C: Local Only (Offline / No Google Drive)
- If `GOOGLE_DRIVE_FOLDER_ID` is left empty (`""`), cloud upload is skipped automatically.
- All exported CSV files are saved safely in the local `./downloads/` folder.

#### Optional: Local Google Drive Sync Path
If you use Google Drive for Desktop app, you can also specify a local folder path to automatically copy exported files into:
```python
LOCAL_GDRIVE_PATH = r"C:\Users\YourUser\Google Drive\MyFolder"
```

---

## 🏗️ Code Architecture (Section-by-Section)

`Dubai_Srapper01.py` is organized into clean, indexed sections:

| Section | Lines | Description |
| :--- | :--- | :--- |
| **`[SECTION 1]`** | `42–78` | Imports (Selenium, Google API client, OAuthlib) and Windows UTF-8 stdout encoding setup. |
| **`[SECTION 2]`** | `80–158` | User Configuration: dates, target tab, headless mode, timeouts, and Google Drive credentials. |
| **`[SECTION 3]`** | `160–332` | Google Drive Engine: `extract_folder_id()`, `get_drive_service()`, and `upload_file_to_gdrive()`. |
| **`[SECTION 4]`** | `334–450` | Visual Highlights & CAPTCHA detection: `highlight_and_scroll()`, `unhighlight()`, `click_visible()`, `is_recaptcha_solved()`. |
| **`[SECTION 5]`** | `452–592` | reCAPTCHA handler (`handle_recaptcha`), Chrome WebDriver setup (`get_driver`), and download tracker (`wait_for_new_download`). |
| **`[SECTION 6]`** | `594–668` | CLI Presentation: `print_banner()`, `validate_date()`, and interactive terminal input menu (`get_terminal_inputs()`). |
| **`[SECTION 7]`** | `670–868` | Core Automation Engine: `run_scraper()` orchestrating browser startup, tab navigation, date entry, search submission, CSV download, and Google Drive upload. |
| **`[SECTION 8]`** | `870–904` | CLI argument parser (`argparse`) and main execution entrypoint. |

---

## 🧩 How CAPTCHA Handling Works

The script contains a dual-layer reCAPTCHA resolver:

1. **Auto-Trigger**: Locates the reCAPTCHA iframe, scrolls to it, highlights it in blue, and clicks the checkbox.
2. **Auto-Pass Check**: Checks via DOM tokens (`grecaptcha.getResponse()`, hidden textareas, and `aria-checked="true"`) whether the challenge immediately passed.
3. **Manual Challenge Loop**:
   - If Google presents an image challenge puzzle (e.g., "Select all traffic lights"), the terminal displays:
     ```text
     ===========================================================================
      🧩 [MANUAL CAPTCHA ACTION REQUIRED]
      👉 An Image Selection CAPTCHA challenge appeared in the Chrome browser!
      👉 Please solve the image puzzle directly in your open Chrome window.
      ⏳ Script is actively waiting for resolution (Timeout: 120s)...
     ===========================================================================
     ```
   - The script actively checks every 0.8s for your solution.
   - As soon as the green checkmark appears, it immediately resumes without waiting for the timeout!

---

## 🎨 Visual Highlighting & UX

When running in visible mode (`HEADLESS = False`), `Dubai_Srapper01.py` applies glowing visual highlights to elements before interacting with them:
- **Blue Border / Glow**: Selecting tabs or locating CAPTCHA iframe.
- **Orange Border / Glow**: Entering From/To dates.
- **Green Border / Glow**: Clicking the Search button.
- **Teal / Cyan Border / Glow**: Clicking the "Download as CSV" export button.

This visual feedback makes debugging and live monitoring effortless.

---

## ❓ Troubleshooting & FAQs

### Q: Why does the browser stay open after clicking the checkbox?
- **Answer**: An image challenge puzzle was likely displayed. Switch to the open Chrome window and complete the puzzle. The terminal will show a live countdown and resume as soon as the puzzle is solved.

### Q: Can I run this completely in the background?
- **Answer**: Yes! Pass `--headless` in the CLI:
  ```bash
  python Dubai_Srapper01.py --from-date 01/01/2025 --to-date 01/06/2025 --headless
  ```
  *(Note: If Google triggers an image CAPTCHA in headless mode, manual solving is not possible, so visible mode is recommended for stability).*

### Q: How do I change the default dates permanently?
- **Answer**: Open `Dubai_Srapper01.py` and modify `FROM_DATE` and `TO_DATE` in `[SECTION 2.1]` (lines 87–88).

### Q: Google Drive error: "File not found" or "403 Forbidden"?
- **Answer**: If using a Service Account, make sure you shared your Google Drive folder with your service account email (`xxx@yyy.iam.gserviceaccount.com`) with **Editor** permissions.
