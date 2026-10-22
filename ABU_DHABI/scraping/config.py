# ==============================================================================
# DUBAI LAND DEPARTMENT SCRAPER & GOOGLE DRIVE UPLOAD CONFIGURATION
# ==============================================================================

# ------------------------------------------------------------------------------
# 1. DATE SETTINGS (Format: DD/MM/YYYY)
# ------------------------------------------------------------------------------
# Easily modify the start and end dates below:
FROM_DATE = "01/01/2025"
TO_DATE = "15/09/2026"

# ------------------------------------------------------------------------------
# 2. WEBSITE SETTINGS
# ------------------------------------------------------------------------------
WEBSITE_URL = "https://dubailand.gov.ae/en/open-data/real-estate-data/#/"

# Target Data Tab to scrape:
# Available options: "Transactions", "Rents", "Project", "Valuations", "Building", "Developer"
TARGET_TAB = "Transactions"

# Run browser in headless mode (False = opens visible Chrome window, True = background)
HEADLESS = False

# Maximum time (in seconds) to wait for user to solve image CAPTCHA challenge if presented
CAPTCHA_TIMEOUT = 120

# ------------------------------------------------------------------------------
# 3. GOOGLE DRIVE SETTINGS
# ------------------------------------------------------------------------------
# Paste your Shared Google Drive Folder ID or full URL here.
# Example Folder URL: https://drive.google.com/drive/folders/1ABC123xyz456...
# Example Folder ID: 1ABC123xyz456...
GOOGLE_DRIVE_FOLDER_ID = "https://drive.google.com/drive/folders/15p5cTByYqE-9i74seXQGXXhYRnmlc1Zc?usp=drive_link"

# Authentication Method: "service_account" or "oauth"
# - "service_account": Place your downloaded Google Cloud service_account.json in this directory.
# - "oauth": Place your credentials.json in this directory (will prompt browser login on first run).
AUTH_TYPE = "oauth"

# Credential file names
SERVICE_ACCOUNT_FILE = "service_account.json"
OAUTH_CREDENTIALS_FILE = "credentials.json"
OAUTH_TOKEN_FILE = "token.json"

# (Optional) Local Google Drive sync path (if you use Google Drive for Desktop app)
# Example: r"G:\Shared drives\MyTeamFolder" or r"C:\Users\Username\Google Drive\MyFolder"
LOCAL_GDRIVE_PATH = ""

# ------------------------------------------------------------------------------
# 4. DOWNLOAD DIRECTORY
# ------------------------------------------------------------------------------
DOWNLOAD_DIR_NAME = "downloads"
