# ADREC Abu Dhabi Property Transaction Downloader

Production-ready standalone Python Selenium automation script for downloading Abu Dhabi real estate transaction data from the official ADREC website and automatically uploading it to Google Drive.

---

## 📋 Requirements

- **Python 3.8+** installed on your system.
- **Google Chrome** browser installed.
- A Google Cloud project with the Google Drive API enabled (only needed for Drive uploads).

---

## 🚀 Quick Setup (3 Minutes)

### Step 1: Install Required Python Packages
Open your terminal/command prompt and run:

```bash
pip install selenium google-api-python-client google-auth google-auth-oauthlib google-auth-httplib2
```
*(Or if you have `requirements.txt`: `pip install -r requirements.txt`)*

---

## 🔐 Configure Google Drive (Optional)

The script contains no Google credentials. To enable Google Drive uploads:

1. In Google Cloud Console, enable the **Google Drive API** for your project.
2. Configure the OAuth consent screen and add your Google account as a test user
   if the app is in testing mode.
3. Create an OAuth client ID for a **Desktop app** and download its JSON file.
4. Save that file as `credentials.json` in the same folder as
   `adrec_downloader.py`. Do not commit or share this file.
5. Run the script. The first run opens Google's sign-in flow and creates a local
   `token.json` file for subsequent runs. Do not commit or share that file either.

For uploads, specify your Drive destination with `--folder-id` or set the
`ADREC_GOOGLE_DRIVE_FOLDER_ID` environment variable. If no folder is specified,
the script downloads the CSV locally and skips the upload.

Example:

```bash
python adrec_downloader.py --folder-id YOUR_GOOGLE_DRIVE_FOLDER_ID
```

---

## 🏃 How to Run the Script

### Option 1: Interactive Mode (Terminal Prompts)
Run the script directly:

```bash
python adrec_downloader.py
```

You will see:
```text
=================================================================
   ADREC Abu Dhabi Property Transaction Downloader
=================================================================

Enter From Date (YYYY-MM-DD): 2026-01-01
Enter To Date (YYYY-MM-DD): 2026-10-08
```

---

### Option 2: Automated Mode (Command-Line Arguments)
For automated tasks or cron jobs without typing dates:

```bash
python adrec_downloader.py --from-date 2026-01-01 --to-date 2026-10-08
```

Optional flags:
- `--headless`: Runs Chrome in the background without opening a browser window.
- `--folder-id YOUR_FOLDER_ID`: Uploads to a custom Google Drive folder ID.
- `--cleanup`: Deletes the local CSV file after successful Google Drive upload.

---

## 🔄 Complete Automation Workflow

```text
Start
  ↓
User enters From Date & To Date
  ↓
Open ADREC Website (https://adrec.gov.ae/en/property_and_index/adrec-dashboard)
  ↓
Switch into Dashboard Iframe Context
  ↓
Scroll to 'Recent Sales' section
  ↓
Enter From Date & To Date into Datepickers
  ↓
Set all 7 Dropdowns to 'All' (Asset Type, Property Type, Transaction Type, District, Community, Project, Layout)
  ↓
Click Export & Wait for CSV Download
  ↓
Upload CSV file to Google Drive folder
  ↓
Verify Upload & Print Google Drive File ID
```

---

## ❓ Troubleshooting

- **Google Sign-in "Access blocked: App has not completed the Google verification process"**:
  Make sure your Gmail account is added under **Test users** in the Google Cloud Console OAuth consent screen, or sign in with the authorized account.
- **Chrome fails to launch**:
  Ensure you have the latest version of Google Chrome installed. Selenium 4 automatically manages ChromeDriver.
- **Wrong Date Range**:
  Dates must follow the `YYYY-MM-DD` format (e.g. `2026-01-01`), and `From Date` must be less than or equal to `To Date`.
