# Dubai Land Department Real Estate Data Scraper & Google Drive Uploader

A robust, production-ready Python automation tool for extracting real estate datasets from the **Dubai Land Department (DLD) Open Data Portal** ([https://dubailand.gov.ae/en/open-data/real-estate-data/#/](https://dubailand.gov.ae/en/open-data/real-estate-data/#/)) and automatically synchronizing downloaded CSV files with **Google Drive**.

---

## 🌟 Highlights & Capabilities

- **Standalone All-In-One Script**: Run everything from a single file (`dubai_scraper_all_in_one.py`) or use the modular configuration setup.
- **Dual Authentication for Google Drive**: Full support for both **OAuth 2.0** (interactive personal Google Account) and **Service Accounts** (automated background/server workflows).
- **Works Without Google Drive**: If you don't configure Google Drive, the scraper seamlessly defaults to local-only mode, saving all CSVs into the `downloads/` directory.
- **Smart reCAPTCHA Resolution**: Automatically triggers and verifies Google reCAPTCHA. If an image selection challenge appears, the script pauses with an interactive countdown timer, allowing you to solve the puzzle in the Chrome browser before automatically resuming.
- **Visual Action Highlighting**: Smoothly scrolls to each form element and applies a glowing highlight during automation so you can monitor what the script is doing in real time.
- **Dynamic Fast Synchronization**: Uses explicit Selenium `WebDriverWait` state detection to react the instant data loads, eliminating rigid delays.

---

## 📁 Project Structure

```
Dubai1/
├── dubai_scraper_all_in_one.py   # 🌟 Primary standalone script (recommended & easiest to share)
├── config.py                     # Central configuration file for modular execution
├── dubai_scraper.py              # Modular automation scraper
├── gdrive_uploader.py            # Google Drive upload integration module
├── requirements.txt              # Required Python dependencies
├── credentials.example.json      # Reference template for OAuth 2.0 Client credentials
├── service_account.example.json  # Reference template for Google Cloud Service Account
├── .gitignore                    # Prevents accidental sharing of credentials & tokens
├── README.md                     # Complete documentation & setup instructions
└── downloads/                    # Local directory where downloaded CSV files are saved
```

---

## ⚡ Quick Start (In 3 Steps)

### Step 1: Install Dependencies
Open your terminal/command prompt in this folder and run:
```bash
pip install -r requirements.txt
```
*(Dependencies: `selenium`, `webdriver-manager`, `google-api-python-client`, `google-auth-oauthlib`, `google-auth-httplib2`)*

### Step 2: Configure Your Settings
Open **`dubai_scraper_all_in_one.py`** (or **`config.py`**) and review **SECTION 2**:
- Set your preferred date range (`FROM_DATE` and `TO_DATE`).
- (Optional) Paste your **Google Drive Folder URL / ID** into `GOOGLE_DRIVE_FOLDER_ID`.
- (Optional) Set `AUTH_TYPE` to `"oauth"` or `"service_account"`.

*(If you leave `GOOGLE_DRIVE_FOLDER_ID = ""`, the script will save CSVs locally in `downloads/` and skip Google Drive upload without any error).*

### Step 3: Run the Scraper
```bash
python dubai_scraper_all_in_one.py
```

---

## 🔑 Google Drive Setup Guide

You can choose one of three modes depending on your requirements:

### Mode 1: Local Only (No Google Drive Needed)
If you only need CSV files saved to your computer:
1. In `dubai_scraper_all_in_one.py` (or `config.py`), leave:
   ```python
   GOOGLE_DRIVE_FOLDER_ID = ""
   ```
2. Run the script. Downloaded datasets will be preserved locally inside the `downloads/` folder.

---

### Mode 2: Service Account (Recommended for automated & server environments)
A Service Account acts as an automated bot account and does not require interactive browser logins.

1. **Enable Google Drive API**:
   - Go to the [Google Cloud Console](https://console.cloud.google.com/).
   - Create a project (or select an existing one).
   - Navigate to **APIs & Services** > **Library**, search for **Google Drive API**, and click **Enable**.
2. **Create Service Account**:
   - Go to **APIs & Services** > **Credentials** > **Create Credentials** > **Service Account**.
   - Enter a name (e.g. `dld-scraper-bot`) and complete the creation.
3. **Generate & Download JSON Key**:
   - Click on your newly created Service Account.
   - Go to the **Keys** tab > **Add Key** > **Create new key** > Select **JSON**.
   - Download the file, rename it to **`service_account.json`**, and place it in the project root directory.
   *(See `service_account.example.json` for reference)*.
4. **Share Your Google Drive Folder (Crucial Step)**:
   - Open your destination Google Drive folder in your browser.
   - Click **Share**.
   - Paste the Service Account email address (e.g., `dld-scraper-bot@your-project.iam.gserviceaccount.com`).
   - Assign the role **Editor** and click **Send**.
5. **Configure Scraper**:
   - In `dubai_scraper_all_in_one.py` (or `config.py`), set:
     ```python
     AUTH_TYPE = "service_account"
     GOOGLE_DRIVE_FOLDER_ID = "https://drive.google.com/drive/folders/YOUR_FOLDER_ID_HERE"
     ```

---

### Mode 3: OAuth 2.0 (For personal Google Accounts)
OAuth allows you to upload files directly into your personal Google Drive using your Google login.

1. **Enable Google Drive API**:
   - In [Google Cloud Console](https://console.cloud.google.com/), enable the **Google Drive API**.
2. **Configure OAuth Consent Screen**:
   - Go to **APIs & Services** > **OAuth consent screen**.
   - Choose **External** user type.
   - Fill in the required app name and support email.
   - If the app is in "Testing" status, add your personal Google email under **Test Users**.
3. **Create OAuth Client ID**:
   - Go to **APIs & Services** > **Credentials** > **Create Credentials** > **OAuth client ID**.
   - Select Application Type: **Desktop app**.
   - Download the JSON credentials file, rename it to **`credentials.json`**, and place it in the project root directory.
   *(See `credentials.example.json` for reference)*.
4. **Configure Scraper**:
   - In `dubai_scraper_all_in_one.py` (or `config.py`), set:
     ```python
     AUTH_TYPE = "oauth"
     GOOGLE_DRIVE_FOLDER_ID = "https://drive.google.com/drive/folders/YOUR_FOLDER_ID_HERE"
     ```
5. **First-Time Authorization**:
   - When you run the script for the first time, your browser will open asking you to sign in with your Google Account and grant permission.
   - Once authorized, a file named **`token.json`** is automatically created to store the session for future automatic runs.

---

## 🏃 Execution Modes

### 1. Interactive Terminal Mode (Default)
Prompts you interactively for start date, end date, data tab, and browser visibility:
```bash
python dubai_scraper_all_in_one.py
```

### 2. Command-Line Arguments Mode
Run automated tasks with custom parameters:
```bash
# Scrape transactions for a custom date range
python dubai_scraper_all_in_one.py --from-date 01/01/2025 --to-date 15/09/2026

# Scrape a specific data tab
python dubai_scraper_all_in_one.py --tab Rents --from-date 01/06/2025 --to-date 01/01/2026

# Run in background (headless) without terminal prompts
python dubai_scraper_all_in_one.py --non-interactive --headless
```

### 3. Available CLI Flags
| Flag | Description | Default |
| :--- | :--- | :--- |
| `--from-date` | Start date in `DD/MM/YYYY` format | `01/01/2025` |
| `--to-date` | End date in `DD/MM/YYYY` format | `15/09/2026` |
| `--tab` | Data module (`Transactions`, `Rents`, `Project`, `Valuations`, `Building`, `Developer`) | `Transactions` |
| `--headless` | Run Google Chrome in background without opening a visible window | `False` |
| `--non-interactive` | Skip terminal prompts and run using CLI args or configuration defaults | `False` |
| `--captcha-timeout` | Timeout in seconds to wait for manual CAPTCHA solving | `120` |

### 4. Running the Modular Version
If you prefer using separate modules:
```bash
# Edit settings in config.py first, then run:
python dubai_scraper.py
```

---

## 🧩 Supported DLD Data Modules

You can target any of the following tabs on the Dubai Land Department portal:
- **`Transactions`**: Property sales, mortgages, and gift transactions.
- **`Rents`**: Rental contract registrations (Ejari datasets).
- **`Project`**: Status and details of registered real estate development projects.
- **`Valuations`**: Property valuation records.
- **`Building`**: Building details, permits, and structural classifications.
- **`Developer`**: Licensed real estate developers and project portfolios.

---

## 🛡️ Security Guidelines

- **Never share credentials publicly**: Do not commit or send `credentials.json`, `token.json`, or `service_account.json` to public Git repositories or unauthorized individuals.
- **`.gitignore` Included**: A pre-configured `.gitignore` file is provided in this repository to prevent credentials, active session tokens, and downloaded data from being committed to version control.
