import sys
import os

def configure_console_utf8():
    """Ensure UTF-8 console encoding across all platforms and subprocesses."""
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    if hasattr(sys.stderr, 'reconfigure'):
        try:
            sys.stderr.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass

configure_console_utf8()

def enable_windows_ansi():
    """Enables ANSI escape sequence support on Windows terminal / PowerShell."""
    if os.name == 'nt':
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            hStdOut = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE = -11
            mode = ctypes.c_ulong()
            kernel32.GetConsoleMode(hStdOut, ctypes.byref(mode))
            mode.value |= 0x0004  # ENABLE_VIRTUAL_TERMINAL_PROCESSING
            kernel32.SetConsoleMode(hStdOut, mode)
        except Exception:
            pass

enable_windows_ansi()

import io
import base64
import json
import re
import time
import math
from collections import Counter
import imaplib
import email
from email import utils as email_utils
from datetime import datetime, timedelta
from typing import Optional, List, Tuple, Dict, Any

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoAlertPresentException,
    TimeoutException,
    WebDriverException,
    NoSuchElementException,
    ElementClickInterceptedException
)
from selenium.webdriver.chrome.options import Options as ChromeOptions

from PIL import Image, ImageOps
import pytesseract
import cv2
import numpy as np

try:
    import ddddocr
except Exception:
    ddddocr = None

from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook


# ======================================================================================
# GLOBAL CONFIGURATION & DEDICATED OUTPUT FILES
# ======================================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 1. Accounts Configuration
ACCOUNTS_FILE = os.path.join(BASE_DIR, "accounts.json")

# 2. Main Scraped Transactions Output File
OUTPUT_FILE = os.path.join(BASE_DIR, "Dodda_Hagade.xlsx")

# 3. Dedicated Progress State File
PROGRESS_FILE = os.path.join(BASE_DIR, "single_progress.json")

# 4. Dedicated Village Scraping Summary Files
VILLAGE_SUMMARY_JSON = os.path.join(BASE_DIR, "single_village_scraping_summary.json")
VILLAGE_SUMMARY_XLSX = os.path.join(BASE_DIR, "Single_Village_Scraping_Summary.xlsx")

# 5. Dedicated Property Types Reference Files
PROPERTY_TYPES_JSON = os.path.join(BASE_DIR, "property_types_reference.json")
PROPERTY_TYPES_XLSX = os.path.join(BASE_DIR, "Property_Types_Reference.xlsx")

# 6. Dedicated Property Number & Search History Log Files
SEARCH_HISTORY_JSON = os.path.join(BASE_DIR, "single_search_history.json")
SEARCH_HISTORY_XLSX = os.path.join(BASE_DIR, "Single_Search_History.xlsx")

# Business constraints
ACCOUNT_SEARCH_LIMIT = 25     # Maximum searches allowed per account
PROPERTY_BATCH_SIZE = 6       # Maximum property types selectable in one search form
MAX_CONSECUTIVE_EMPTY = 50    # Consecutive empty survey numbers to mark village completed
SKIP_COMPLETED_VILLAGES = True

# Target Hierarchy Configuration (User-Configurable)
TARGET_DISTRICT_NAME = "Basavanagudi"   # Registration District (Basavanagudi / Bengaluru Urban)
TARGET_DISTRICT_INDEX = 3                # Dropdown option index fallback (3 in Kaveri)
TARGET_TALUKA_NAME = "Anekal"            # Taluka (Anekal)
TARGET_TALUKA_INDEX = 1                  # Dropdown option index fallback (1 in Kaveri)
TARGET_HOBLI_OPTION_INDEX = 3            # 3 -> option[3] (e.g. Attibele Hobli 1 or Anekal Kasaba Hobli 2). Set to None for all hoblis.

# Target Village Configuration (User-Configurable)
# Set TARGET_VILLAGE_NAME to jump directly to a village by name (e.g., "Aravantike Pura" or partial match).
# Or set TARGET_VILLAGE_OPTION_INDEX to jump by 1-based dropdown index (e.g., 1 for 1st village, 5 for 5th).
# Set to "" or None to iterate all villages normally.
TARGET_VILLAGE_NAME = "Dodda Hagade" # e.g. "Aravantike Pura" or None/"" for all villages
TARGET_VILLAGE_OPTION_INDEX = 9      # 1-based index (e.g. 1, 2, 5) or None
TARGET_VILLAGE_ONLY = True              # If True, scrapes ONLY the target village and exits. If False, jumps to target village and continues.
TARGET_SURVEY_START = None          # Optional: force start survey number (e.g. 1 or 50). None = auto-resume.
TARGET_ACCOUNT_INDEX = None           # Optional: force start account index (1-based, e.g. 1 for Account-1). None = auto-resume.
LIST_VILLAGES_ONLY = False               # If True, connects, prints all available villages in the hobli, and exits.

DEFAULT_ACCOUNTS = [
    {
        "label": "Account-1",
        "login_email": "",
        "login_password": "",
        "imap_email": "",
        "imap_app_password": ""
    },
    {
        "label": "Account-2",
        "login_email": "",
        "login_password": "",
        "imap_email": "",
        "imap_app_password": ""
    },
    {
        "label": "Account-3",
        "login_email": "",
        "login_password": "",
        "imap_email": "",
        "imap_app_password": ""
    }
]

IMAP_SERVER = "imap.gmail.com"
CANDIDATE_TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
    os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]
TESSERACT_PATH = next((p for p in CANDIDATE_TESSERACT_PATHS if os.path.exists(p)), r"C:\Program Files\Tesseract-OCR\tesseract.exe")
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

DISTRICT_XPATH = "//app-ec-search-citizen//form//div[1]/div[1]//select | (//form//select)[1]"
TALUKA_XPATH = "//app-ec-search-citizen//form//div[1]/div[2]//select | (//form//select)[2]"
HOBLI_XPATH = "//app-ec-search-citizen//form//div[1]/div[3]//select | (//form//select)[3]"
VILLAGE_XPATH = "//app-ec-search-citizen//form//div[1]/div[4]//select | (//form//select)[4]"

# ANSI Color & Style Palettes for Clean Terminal UX
RESET = "\033[0m"
BOLD = "\033[1m"
DIM = "\033[2m"
COLOR_CYAN = "\033[96m"
COLOR_BLUE = "\033[94m"
COLOR_GREEN = "\033[92m"
COLOR_MAGENTA = "\033[95m"
COLOR_YELLOW = "\033[93m"
COLOR_RED = "\033[91m"
COLOR_GRAY = "\033[90m"
COLOR_WHITE = "\033[97m"

def log(label: str, message: str):
    """Atomic, color-coded, timestamped console print."""
    now_str = datetime.now().strftime("%H:%M:%S")
    clean_label = f"[{label}]"
    formatted = f"{COLOR_GRAY}[{now_str}]{RESET} {COLOR_CYAN}{BOLD}{clean_label:<20}{RESET} {message}\n"
    try:
        sys.stdout.write(formatted)
        sys.stdout.flush()
    except Exception:
        print(f"[{now_str}] [{label}] {message}", flush=True)


def load_accounts() -> List[Dict[str, Any]]:
    """Loads accounts from accounts.json or falls back to DEFAULT_ACCOUNTS."""
    if os.path.exists(ACCOUNTS_FILE):
        try:
            with open(ACCOUNTS_FILE, "r", encoding="utf-8") as f:
                accounts = json.load(f)
            if isinstance(accounts, list) and len(accounts) > 0:
                log("System", f"📋 Loaded {len(accounts)} accounts from accounts.json")
                return accounts
        except Exception as e:
            log("System", f"⚠️ Error reading accounts.json: {e}")
    log("System", f"📋 Using {len(DEFAULT_ACCOUNTS)} default built-in accounts")
    return DEFAULT_ACCOUNTS


# ======================================================================================
# BROWSER DRIVER FACTORY (CHROME)
# ======================================================================================
def build_browser_driver(label="Browser"):
    """Creates an isolated Google Chrome WebDriver instance with anti-automation flags."""
    log(label, "🌐 Initializing Google Chrome WebDriver...")
    options = ChromeOptions()
    options.page_load_strategy = "eager"
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--start-maximized")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-animations")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--disable-backgrounding-occluded-windows")
    cache_dir = os.path.join(BASE_DIR, "tmp_chrome_cache")
    os.makedirs(cache_dir, exist_ok=True)
    options.add_argument(f"--disk-cache-dir={cache_dir}")
    options.add_argument("--disk-cache-size=104857600")  # Limit cache to 100MB
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(options=options)
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"}
        )
    except Exception:
        pass
    return driver


def is_driver_alive(driver):
    """Checks if the WebDriver instance is active and responsive."""
    try:
        if driver is None:
            return False
        _ = driver.current_window_handle
        return True
    except Exception:
        return False


# ======================================================================================
# POPUPS, LOADERS & MODAL HELPERS
# ======================================================================================
def wait_for_loader(driver, timeout=20):
    try:
        is_visible = driver.execute_script("""
            const el = document.querySelector('.loader-overlay');
            if (!el) return false;
            const style = window.getComputedStyle(el);
            return style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0';
        """)
        if not is_visible:
            return
        WebDriverWait(driver, timeout, poll_frequency=0.05).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, 'loader-overlay'))
        )
    except Exception:
        pass


def close_landing_popup(driver):
    """Closes initial landing page announcement popup if present."""
    try:
        popups = driver.find_elements(By.CSS_SELECTOR, "div.popup-container.open, .popup-container, .announcement-popup, div[role='dialog']")
        for popup in popups:
            if popup.is_displayed():
                try:
                    close_btn = popup.find_element(By.XPATH, ".//button | .//span[text()='X'] | .//span[contains(@class, 'close')] | .//a[contains(@class, 'close')]")
                    driver.execute_script("arguments[0].click();", close_btn)
                    time.sleep(0.2)
                except Exception:
                    pass
    except Exception:
        pass



def close_result_modal(driver):
    """Close result dialogs/modals/backdrops that block interaction."""
    try:
        has_modal = driver.execute_script("""
            const selectors = ['.modal-backdrop', '.cdk-overlay-backdrop', '.swal2-container', '[role=\"dialog\"]', '.mat-dialog-container', 'mat-dialog-container'];
            for (let s of selectors) {
                const els = document.querySelectorAll(s);
                for (let el of els) {
                    const st = window.getComputedStyle(el);
                    if (st.display !== 'none' && st.visibility !== 'hidden' && el.offsetHeight > 0) return true;
                }
            }
            return false;
        """)
        if not has_modal:
            return True

        for _ in range(4):
            modal_selectors = [
                ".modal-backdrop",
                ".cdk-overlay-backdrop",
                ".swal2-container",
                "[role='dialog']",
                ".mat-dialog-container",
                "mat-dialog-container"
            ]
            found_any = False
            for selector in modal_selectors:
                try:
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    for el in elements:
                        if el.is_displayed():
                            found_any = True
                            try:
                                close_btn = el.find_element(By.XPATH, ".//button[normalize-space()='OK' or normalize-space()='Close' or normalize-space()='Done' or normalize-space()='Yes' or contains(@class, 'close') or contains(text(), 'X')]")
                                close_btn.click()
                            except Exception:
                                try:
                                    el.click()
                                except Exception:
                                    pass
                            time.sleep(0.1)
                except Exception:
                    pass

            if not found_any:
                return True
            time.sleep(0.1)
    except Exception:
        pass
    return False


def wait_for_page_ready(driver, timeout=12):
    """Wait until overlays and loaders are clear."""
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            wait_for_loader(driver, timeout=5)
            is_blocked = driver.execute_script("""
                const overlays = document.querySelectorAll('.modal-backdrop, .cdk-overlay-backdrop, .swal2-container');
                for (let el of overlays) {
                    const st = window.getComputedStyle(el);
                    if (st.display !== 'none' && st.visibility !== 'hidden' && el.offsetHeight > 0) return true;
                }
                return false;
            """)
            if not is_blocked:
                return True
        except Exception:
            pass
        time.sleep(0.05)
    return False


def wait_for_dashboard(driver, timeout=30):
    """Return only when the authenticated dashboard is actually visible."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: "dashboard" in d.current_url.lower()
            or "kaveri-dashboard" in d.current_url.lower()
        )
        return True
    except Exception:
        try:
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "app-kaveri-dashboard"))
            )
            return True
        except Exception:
            return False


def safe_get(driver, url, max_retries=3, wait_after=5, label=""):
    """driver.get(url) with automatic retry and refresh."""
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            driver.get(url)
            WebDriverWait(driver, 20).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(wait_after)
            return True
        except Exception as e:
            last_exc = e
            log(label, f"⚠️ Failed to load '{url}' (attempt {attempt}/{max_retries}): {e}. Refreshing...")
            try:
                driver.refresh()
                time.sleep(wait_after)
            except Exception:
                pass
    log(label, f"❌ Failed to load '{url}' after {max_retries} attempts: {last_exc}")
    return False


# ======================================================================================
# CAPTCHA PROCESSING (OCR) - HIGH-SPEED HYBRID ENGINE
# ======================================================================================
CAPTCHA_SAMPLES_DIR = os.path.join(BASE_DIR, "captcha_samples")
CAPTCHA_SUCCESS_DIR = os.path.join(CAPTCHA_SAMPLES_DIR, "successful")
CAPTCHA_FAILED_DIR = os.path.join(CAPTCHA_SAMPLES_DIR, "failed")
os.makedirs(CAPTCHA_SUCCESS_DIR, exist_ok=True)
os.makedirs(CAPTCHA_FAILED_DIR, exist_ok=True)


def save_captcha_sample(img_bytes, predicted_text, is_success, context="search", label="Worker"):
    try:
        if not img_bytes or not isinstance(img_bytes, bytes):
            return
        target_dir = CAPTCHA_SUCCESS_DIR if is_success else CAPTCHA_FAILED_DIR
        clean_pred = re.sub(r'[^A-Za-z0-9_]', '', str(predicted_text or 'EMPTY'))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:19]
        clean_label = re.sub(r'[^A-Za-z0-9_]', '', str(label))
        status_prefix = "SUCCESS" if is_success else "FAIL"
        filename = f"{status_prefix}_{clean_pred}_{timestamp}_{context}_{clean_label}.png"
        filepath = os.path.join(target_dir, filename)
        with open(filepath, "wb") as f:
            f.write(img_bytes)
    except Exception:
        pass


_LOCAL_DDDDOCR = None

def get_ddddocr_instance():
    """Returns a singleton instance of ddddocr."""
    global _LOCAL_DDDDOCR
    if _LOCAL_DDDDOCR is None:
        try:
            if ddddocr is not None:
                _LOCAL_DDDDOCR = ddddocr.DdddOcr(show_ad=False)
            else:
                _LOCAL_DDDDOCR = False
        except Exception:
            _LOCAL_DDDDOCR = False
    return _LOCAL_DDDDOCR if _LOCAL_DDDDOCR is not False else None


def _find_j_hook_clusters(above_baseline, j_hook_thresh=28, merge_gap=3):
    j_hook_cols = [x for x, v in enumerate(above_baseline) if v > j_hook_thresh]
    clusters = []
    cur = None
    for x in j_hook_cols:
        if cur is None:
            cur = [x, x]
        elif x - cur[1] <= merge_gap:
            cur[1] = x
        else:
            clusters.append((cur[0], cur[1]))
            cur = [x, x]
    if cur:
        clusters.append((cur[0], cur[1]))
    return clusters


def _recover_missing_narrow_chars(cl_result, raw_result, cleaned_img, img_bytes):
    n_missing = 6 - len(cl_result)
    if n_missing <= 0:
        return cl_result

    im_orig = cv2.imdecode(np.frombuffer(img_bytes, np.uint8), cv2.IMREAD_COLOR)
    if im_orig is None:
        return cl_result
    h, w = im_orig.shape[:2]
    ocr = get_ddddocr_instance()
    if ocr is None:
        return cl_result

    if len(raw_result) == 6 and len(cl_result) == 5 and raw_result[1:] == cl_result:
        for sx in [1.2, 1.3, 1.5]:
            s = cv2.resize(im_orig, (int(w * sx), h), interpolation=cv2.INTER_LANCZOS4)
            _, buf = cv2.imencode(".png", s)
            r = re.sub(r"[^A-Z0-9]", "", ocr.classification(buf.tobytes()).strip().upper())
            if len(r) == 6 and r[1:] == cl_result:
                return r
        return raw_result

    six_results = []
    for sx in [1.2, 1.3, 1.5, 1.8, 2.0]:
        for src in [im_orig, cleaned_img]:
            s = cv2.resize(src, (int(w * sx), h), interpolation=cv2.INTER_LANCZOS4)
            _, buf = cv2.imencode(".png", s)
            r = re.sub(r"[^A-Z0-9]", "", ocr.classification(buf.tobytes()).strip().upper())
            if len(r) == 6:
                six_results.append(r)

    if six_results:
        voted = []
        for pos in range(6):
            votes = Counter(r[pos] for r in six_results if pos < len(r))
            if len(cl_result) > pos and cl_result[pos] == "0":
                if votes.get("O", 0) > votes.get("0", 0):
                    votes["0"] = votes["O"] + 1
            if pos == 5 and votes.get("H", 0) > 0:
                if (cl_result and cl_result[-1] == "I") or any(r[-1] == "I" for r in six_results):
                    votes["I"] = votes.get("I", 0) + 2
            voted.append(votes.most_common(1)[0][0] if votes else "?")
        result = "".join(voted)

        n_cl = len(cl_result)
        if n_cl >= 3:
            mismatches = sum(1 for i in range(n_cl) if i < len(result) and result[i] != cl_result[i])
            if mismatches > 1:
                six_results = []

        if six_results:
            if raw_result and result[-1] == "O" and raw_result[-1] == "0":
                result = result[:-1] + "0"
            return result

    gray = cv2.cvtColor(cleaned_img, cv2.COLOR_BGR2GRAY)
    col_counts = (gray < 128).sum(axis=0)
    baseline = 10
    above = (np.array(col_counts) - baseline).clip(0).astype(float)

    j_clusters = _find_j_hook_clusters(above)
    n_j_img = len(j_clusters)
    n_j_cl = cl_result.count("J")
    correction = list(cl_result)

    if n_j_img > n_j_cl > 0:
        j_idx = correction.index("J")
        correction.insert(j_idx + 1, "J")
    elif n_j_img == n_j_cl > 0:
        for jc_start, jc_end in j_clusters:
            check_start = max(0, jc_start - 18)
            pre_above = above[check_start:jc_start]
            max_pre = max(pre_above) if len(pre_above) > 0 else 0
            high_pre = sum(1 for v in pre_above if v >= 22)
            if 22 <= max_pre <= 28 and high_pre >= 2:
                j_idx = correction.index("J")
                correction.insert(j_idx, "I")
                break
    elif n_j_img > 0 and n_j_cl == 0:
        for jc_start, jc_end in j_clusters:
            jc_center = (jc_start + jc_end) / 2.0
            slot_6 = jc_center / w * 6.0
            insert_at = min(math.ceil(slot_6), len(correction))
            correction.insert(insert_at, "J")
            break

    while len(correction) < 6:
        correction.append("I")

    result = "".join(correction[:6])
    if raw_result and result[-1] == "O" and raw_result[-1] == "0":
        result = result[:-1] + "0"
    return result


def keep_black_text_only(image_source, label="default"):
    """High-speed, high-accuracy hybrid CAPTCHA solver."""
    if isinstance(image_source, bytes):
        img_bytes = image_source
    elif isinstance(image_source, str) and os.path.exists(image_source):
        try:
            with open(image_source, "rb") as f:
                img_bytes = f.read()
        except Exception:
            img_bytes = None
    else:
        img_bytes = None

    if not img_bytes:
        return ""

    ocr = get_ddddocr_instance()
    im = None
    cleaned_img = None
    cleaned_bytes = None

    try:
        im = cv2.imdecode(np.frombuffer(img_bytes, np.uint8), cv2.IMREAD_COLOR)
        if im is not None:
            dark_mask = (im[:, :, 0] < 125) & (im[:, :, 1] < 125) & (im[:, :, 2] < 125)
            if np.sum(dark_mask) < 650:
                return ""

            b = im[:, :, 0].astype(int)
            g = im[:, :, 1].astype(int)
            r = im[:, :, 2].astype(int)
            chroma = np.maximum(np.maximum(r, g), b) - np.minimum(np.minimum(r, g), b)
            is_text = (im.mean(axis=2) < 135) & (chroma < 20)

            cleaned_img = np.full_like(im, 255)
            cleaned_img[is_text] = 0

            _, buf = cv2.imencode(".png", cleaned_img)
            cleaned_bytes = buf.tobytes()
    except Exception:
        pass

    cand_cl = ""
    cand_raw = ""
    cand_tess = ""

    if ocr is not None and cleaned_bytes is not None:
        try:
            raw_text = ocr.classification(cleaned_bytes)
            cand_cl = re.sub(r"[^A-Z0-9]", "", str(raw_text).strip().upper())
        except Exception:
            pass

    if ocr is not None:
        try:
            raw_text = ocr.classification(img_bytes)
            cand_raw = re.sub(r"[^A-Z0-9]", "", str(raw_text).strip().upper())
        except Exception:
            pass

    config = "--psm 7 --oem 1 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    try:
        target_mat = cleaned_img if cleaned_img is not None else im
        if target_mat is not None:
            gray = cv2.cvtColor(target_mat, cv2.COLOR_BGR2GRAY)
            scaled = cv2.resize(gray, None, fx=2.5, fy=2.5, interpolation=cv2.INTER_LANCZOS4)
            raw_text = pytesseract.image_to_string(scaled, config=config)
            cand_tess = re.sub(r"[^A-Z0-9]", "", raw_text.strip().upper())
    except Exception:
        pass

    best_ddd_len = max(len(cand_cl), len(cand_raw))
    chosen = ""
    if len(cand_cl) == 6:
        chosen = cand_cl
    elif len(cand_raw) == 6:
        chosen = cand_raw
    elif best_ddd_len < 4:
        if len(cand_tess) == 6:
            chosen = cand_tess
        else:
            try:
                if target_mat is not None:
                    gray = cv2.cvtColor(target_mat, cv2.COLOR_BGR2GRAY)
                    blur = cv2.GaussianBlur(gray, (3, 3), 0)
                    _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                    padded = cv2.copyMakeBorder(thresh, 15, 15, 15, 20, cv2.BORDER_CONSTANT, value=255)
                    scaled_pad = cv2.resize(padded, None, fx=2.5, fy=2.5, interpolation=cv2.INTER_LANCZOS4)
                    raw_text_pad = pytesseract.image_to_string(scaled_pad, config=config)
                    cand_pad = re.sub(r"[^A-Z0-9]", "", raw_text_pad.strip().upper())
                    if len(cand_pad) == 6:
                        chosen = cand_pad
            except Exception:
                pass

    if len(chosen) == 6:
        voted = list(chosen)
        for i in range(6):
            if len(cand_cl) == 6 and cand_cl[i] == "1" and voted[i] == "I":
                voted[i] = "1"
            if voted[i] == "L":
                if len(cand_raw) == 6 and cand_raw[i] == "I":
                    voted[i] = "I"
                elif len(cand_tess) == 6 and cand_tess[i] == "I":
                    voted[i] = "I"
            elif voted[i] == "5" and len(cand_raw) == 6 and cand_raw[i] == "S":
                voted[i] = "S"
            elif voted[i] == "S" and len(cand_raw) == 6 and cand_raw[i] == "5":
                pass

        chosen = "".join(voted)
        if len(cand_raw) == 6 and len(cand_cl) == 5 and cand_raw[1:] == cand_cl and chosen[0] == cand_raw[0] and cleaned_img is not None:
            fixed = _recover_missing_narrow_chars(cand_cl, cand_raw, cleaned_img, img_bytes)
            if len(fixed) == 6:
                return fixed
        return chosen

    if cleaned_img is not None:
        base = cand_cl if len(cand_cl) >= len(cand_raw) else cand_raw
        if not base:
            base = cand_tess
        recovered = _recover_missing_narrow_chars(base, cand_raw, cleaned_img, img_bytes)
        if recovered:
            return recovered[:6]

    fallback = chosen or cand_cl or cand_raw or cand_tess
    if fallback and len(fallback) == 6:
        return fallback
    elif fallback and len(fallback) > 6:
        return fallback[:6]
    return fallback


def get_captcha_bytes(driver, captcha_element):
    """Extracts lossless CAPTCHA image bytes directly from browser memory."""
    try:
        src = captcha_element.get_attribute("src") or ""
        if src.startswith("data:image"):
            data_b64 = src.split(",", 1)[1]
            raw_bytes = base64.b64decode(data_b64)
            if raw_bytes and len(raw_bytes) > 200:
                return raw_bytes
    except Exception:
        pass

    try:
        canvas_b64 = driver.execute_script("""
            var img = arguments[0];
            var canvas = document.createElement('canvas');
            canvas.width = img.naturalWidth || img.width || 200;
            canvas.height = img.naturalHeight || img.height || 60;
            var ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0);
            return canvas.toDataURL('image/png').substring(22);
        """, captcha_element)
        if canvas_b64 and len(canvas_b64) > 200:
            raw_bytes = base64.b64decode(canvas_b64)
            if raw_bytes and len(raw_bytes) > 200:
                return raw_bytes
    except Exception:
        pass

    try:
        return captcha_element.screenshot_as_png
    except Exception:
        return None


def refresh_search_captcha(driver, captcha_element, timeout=5):
    """Refreshes the CAPTCHA image and waits for the new image to render."""
    try:
        old_png = None
        try:
            old_png = captcha_element.screenshot_as_png
        except Exception:
            pass

        clicked = False
        refresh_selectors = [
            "./following-sibling::*[self::button or self::a or self::i or self::span or self::mat-icon]",
            "./parent::*//*[self::button or self::a or self::i or self::span or self::mat-icon][not(self::img)]",
            "./ancestor::div[1]//*[self::button or self::a or self::i or self::span or self::mat-icon][not(self::img)]",
            "./ancestor::div[2]//*[contains(@class, 'refresh') or contains(@class, 'reload') or contains(text(), 'Refresh') or .//mat-icon[contains(text(), 'refresh')]]",
            "//form//button[.//mat-icon[text()='refresh'] or contains(@class, 'refresh')]",
            "//app-ec-search-citizen//button[.//mat-icon[text()='refresh'] or contains(@class, 'refresh')]",
            "//img[@alt='Loaded Image']/following-sibling::*",
            "//img[@alt='Loaded Image']/parent::*//button",
            "//app-login//button[.//mat-icon[text()='refresh'] or contains(@class, 'refresh')]"
        ]
        for xpath in refresh_selectors:
            try:
                buttons = captcha_element.find_elements(By.XPATH, xpath) if not xpath.startswith("//") else driver.find_elements(By.XPATH, xpath)
                for btn in buttons:
                    if btn.is_displayed():
                        try:
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                            driver.execute_script("arguments[0].click();", btn)
                            clicked = True
                            break
                        except Exception:
                            try:
                                btn.click()
                                clicked = True
                                break
                            except Exception:
                                pass
                if clicked:
                    break
            except Exception:
                pass

        if not clicked:
            try:
                driver.execute_script("arguments[0].click();", captcha_element)
            except Exception:
                pass

        if old_png:
            end_time = time.time() + timeout
            while time.time() < end_time:
                time.sleep(0.05)
                try:
                    if captcha_element.screenshot_as_png != old_png:
                        return True
                except Exception:
                    pass
        else:
            time.sleep(0.3)
    except Exception as e:
        log("System", f"⚠️ Error refreshing CAPTCHA: {e}")
    return False


# ======================================================================================
# IMAP OTP FETCHING
# ======================================================================================
def fetch_otp(retries=12, delay=5, after_time=None, label="", imap_email=None, imap_password=None):
    if after_time:
        after_time = after_time - timedelta(minutes=3)
    log(label, f"📧 Connecting to IMAP mailbox: {imap_email}")

    def try_fetch_once():
        try:
            mail = imaplib.IMAP4_SSL(IMAP_SERVER)
            mail.login(imap_email, imap_password)
            mail.select("inbox")
            status, data = mail.search(None, '(FROM "dsrnotify@karnataka.gov.in")')
            mail_ids = data[0].split()
            if not mail_ids:
                return None
            for i in reversed(mail_ids[-5:]):
                status, msg_data = mail.fetch(i, "(RFC822)")
                for part in msg_data:
                    if isinstance(part, tuple):
                        msg = email.message_from_bytes(part[1])
                        date_tuple = email_utils.parsedate_tz(msg['Date'])
                        if date_tuple is None:
                            continue
                        msg_time = datetime.fromtimestamp(email_utils.mktime_tz(date_tuple))
                        if after_time and msg_time < after_time:
                            continue
                        payload = msg.get_payload(decode=True)
                        if not isinstance(payload, bytes):
                            continue
                        html = payload.decode(errors='ignore')
                        otp = re.findall(r'\b\d{4,6}\b', html)
                        if otp:
                            return otp[0]
        except Exception as e:
            log(label, f"❌ OTP fetch error: {e}")
        return None

    for attempt in range(retries):
        otp = try_fetch_once()
        if otp:
            log(label, f"🔑 OTP '{otp}' retrieved successfully!")
            return otp
        log(label, f"⏳ Waiting for OTP email (Attempt {attempt+1}/{retries})...")
        time.sleep(delay)
    log(label, f"❌ OTP not received after {retries} retries.")
    return None


# ======================================================================================
# LOGIN & LOGOUT FLOWS
# ======================================================================================
def perform_login(driver, account: Dict[str, Any], label="Login"):
    """Logs into the Kaveri portal using the specified account credentials."""
    login_email = account["login_email"]
    login_password = account["login_password"]
    imap_email = account.get("imap_email", login_email)
    imap_password = account.get("imap_app_password")

    log(label, f"🔐 Logging in as: {login_email}...")
    safe_get(driver, "https://kaveri.karnataka.gov.in/landing-page", label=label)
    login_start_time = datetime.now()

    try:
        max_attempts = 10
        attempt = 0
        captcha_success = False
        consecutive_login_fails = 0

        while attempt < max_attempts and not captcha_success:
            attempt += 1
            log(label, f"🔐 CAPTCHA Attempt {attempt} of {max_attempts}")
            try:
                is_modal_open = False
                try:
                    user_el = driver.find_element(By.XPATH, '//*[@id="userName"]')
                    if user_el.is_displayed():
                        is_modal_open = True
                except Exception:
                    is_modal_open = False

                if not is_modal_open:
                    close_landing_popup(driver)
                    WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="headerMain"]/ul/span/li[2]/button'))
                    ).click()
                    log(label, "Clicked header Login button.")
                    time.sleep(0.5)

                username_input = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="userName"]'))
                )
                if username_input.get_attribute("value") != login_email:
                    username_input.send_keys(Keys.CONTROL + "a")
                    username_input.send_keys(Keys.BACKSPACE)
                    username_input.clear()
                    username_input.send_keys(login_email)
                    driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));"
                        "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                        username_input
                    )

                password_input = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))
                )
                password_input.send_keys(Keys.CONTROL + "a")
                password_input.send_keys(Keys.BACKSPACE)
                password_input.clear()
                password_input.send_keys(login_password)
                driver.execute_script(
                    "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));"
                    "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                    password_input
                )

                wait_for_loader(driver, timeout=5)
                captcha_xpath = "//img[@alt='Loaded Image']"
                captcha_element = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, captcha_xpath))
                )

                if attempt > 1:
                    refresh_search_captcha(driver, captcha_element)
                    wait_for_loader(driver, timeout=5)

                captcha_bytes = get_captcha_bytes(driver, captcha_element)
                captcha_text = keep_black_text_only(captcha_bytes, label=label)
                log(label, f"Extracted CAPTCHA: '{captcha_text}'")

                if not captcha_text or len(captcha_text) != 6:
                    log(label, f"⚠️ CAPTCHA OCR not 6 chars ('{captcha_text}'). Retrying...")
                    save_captcha_sample(captcha_bytes, captcha_text, is_success=False, context="login_ocr_fail", label=label)
                    refresh_search_captcha(driver, captcha_element)
                    time.sleep(0.5)
                    continue

                captcha_input = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="loginCaptchaUserInput1"]'))
                )
                captcha_input.clear()
                captcha_input.send_keys(captcha_text)
                driver.execute_script(
                    "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));"
                    "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                    captcha_input
                )

                time.sleep(0.5)
                driver.find_element(By.XPATH, "//button[text()='Log In']").click()

                login_verdict = None
                poll_start = time.time()
                poll_end = poll_start + 4.5
                while time.time() < poll_end:
                    try:
                        alert = driver.switch_to.alert
                        alert_text = alert.text
                        alert.accept()
                        log(label, f"Alert detected: {alert_text}")
                        if any(w in alert_text.lower() for w in ["multiple", "session"]):
                            time.sleep(1.0)
                            login_verdict = "session_cleared"
                            break
                        elif any(w in alert_text.lower() for w in ["captcha", "invalid"]):
                            login_verdict = "captcha_invalid"
                            break
                        else:
                            time.sleep(1.0)
                            break
                    except Exception:
                        pass

                    try:
                        otp_el = driver.find_element(By.XPATH, '//*[@id="otp"]')
                        if otp_el.is_displayed():
                            login_verdict = "success"
                            break
                    except Exception:
                        pass

                    if (time.time() - poll_start) > 1.0:
                        try:
                            err_elements = driver.find_elements(
                                By.XPATH,
                                "//*[contains(text(), 'Captcha is not valid') or contains(text(), 'Invalid captcha') or contains(text(), 'Captcha does not match')]"
                            )
                            for err in err_elements:
                                if err.is_displayed():
                                    login_verdict = "captcha_invalid"
                                    break
                            if login_verdict:
                                break
                        except Exception:
                            pass

                    time.sleep(0.15)

                if login_verdict == "success":
                    captcha_success = True
                    save_captcha_sample(captcha_bytes, captcha_text, is_success=True, context="login", label=label)
                    log(label, "✅ CAPTCHA verified — OTP modal opened.")
                    break
                elif login_verdict == "session_cleared":
                    log(label, "🔄 Active session cleared by Kaveri. Re-submitting...")
                    time.sleep(1.0)
                    continue
                elif login_verdict == "captcha_invalid":
                    save_captcha_sample(captcha_bytes, captcha_text, is_success=False, context="login_invalid", label=label)
                    log(label, "⚠️ CAPTCHA invalid. Refreshing...")
                    consecutive_login_fails += 1
                    try:
                        stale_captcha = driver.find_element(By.XPATH, captcha_xpath)
                        refresh_search_captcha(driver, stale_captcha)
                    except Exception:
                        pass
                    time.sleep(0.5)
                    continue
                else:
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="otp"]'))
                        )
                        captcha_success = True
                        save_captcha_sample(captcha_bytes, captcha_text, is_success=True, context="login", label=label)
                        log(label, "✅ CAPTCHA verified — OTP modal opened.")
                        break
                    except Exception:
                        save_captcha_sample(captcha_bytes, captcha_text, is_success=False, context="login_timeout", label=label)
                        consecutive_login_fails += 1
                        time.sleep(0.5)
                        continue

            except (TimeoutException, WebDriverException) as e:
                log(label, f"⚠️ CAPTCHA attempt failed: {e}. Retrying...")
                consecutive_login_fails += 1
                if consecutive_login_fails % 4 == 0:
                    try:
                        driver.refresh()
                        wait_for_loader(driver)
                    except Exception:
                        pass
                time.sleep(1)
                continue

        if not captcha_success:
            raise Exception("Failed to solve CAPTCHA after maximum attempts")

        time.sleep(10)
        otp = fetch_otp(after_time=login_start_time, label=label, imap_email=imap_email, imap_password=imap_password)
        if not otp:
            log(label, "❌ OTP fetch failed.")
            return False

        otp_input = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="otp"]'))
        )
        otp_input.clear()
        otp_input.send_keys(otp)
        time.sleep(2)
        log(label, "Entered OTP.")

        driver.find_element(By.XPATH, "//input[@id='otp']/ancestor::form//button[not(contains(text(), 'Resend'))]").click()
        log(label, "⏳ Waiting for dashboard to load after OTP...")
        if not wait_for_dashboard(driver):
            raise TimeoutException("Dashboard did not load after OTP")
        time.sleep(3)
        log(label, f"✅ Logged in successfully as {login_email}.")
        return True

    except Exception as e:
        log(label, f"❌ Error during login: {e}")
        return False


def perform_logout(driver, label="Logout"):
    """Logs out of the current active session in Kaveri portal."""
    log(label, "🚪 Logging out from current account...")
    try:
        if not is_driver_alive(driver):
            return True
        close_result_modal(driver)
        wait_for_loader(driver)

        profile_triggers = [
            "//li[contains(@class, 'user')]//a",
            "//div[contains(@class, 'user-profile')]",
            "//span[contains(text(), 'Hello') or contains(text(), '@')]/parent::*",
            "//*[contains(text(), 'Hello')]",
            "//img[contains(@src, 'profile') or contains(@class, 'profile') or contains(@src, 'user')]/parent::*",
            "//*[@id='headerMain']//ul//li[last()]",
        ]
        dropdown_opened = False
        for xpath in profile_triggers:
            try:
                elems = driver.find_elements(By.XPATH, xpath)
                for el in elems:
                    if el.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                        try:
                            el.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", el)
                        dropdown_opened = True
                        time.sleep(0.5)
                        break
                if dropdown_opened:
                    break
            except Exception:
                pass

        logout_candidates = [
            "//button[normalize-space()='Logout' or contains(text(), 'Logout')]",
            "//a[normalize-space()='Logout' or contains(text(), 'Logout')]",
            "//span[normalize-space()='Logout' or contains(text(), 'Logout')]",
            "//li[normalize-space()='Logout' or contains(text(), 'Logout')]",
            "//*[contains(text(), 'Logout')]",
        ]
        logged_out = False
        for xpath in logout_candidates:
            try:
                elems = driver.find_elements(By.XPATH, xpath)
                for el in elems:
                    if el.is_displayed():
                        try:
                            el.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", el)
                        logged_out = True
                        break
                if logged_out:
                    break
            except Exception:
                pass

        try:
            alert = driver.switch_to.alert
            alert.accept()
        except Exception:
            pass

        time.sleep(1)
        try:
            driver.delete_all_cookies()
        except Exception:
            pass

        safe_get(driver, "https://kaveri.karnataka.gov.in/landing-page", label=label)
        log(label, "✅ Logged out cleanly.")
        return True
    except Exception as e:
        log(label, f"⚠️ Error during logout: {e}. Resetting session.")
        return False


# ======================================================================================
# FORM & SEARCH NAVIGATION
# ======================================================================================
def ensure_property_boundary_selected(driver, label=""):
    """Ensures Non-Agricultural and Property/Survey boundary radio button is selected on the EC form."""
    try:
        wait_for_loader(driver)
        # 1. Ensure 'Non - Agricultural' is selected
        non_agri_selectors = [
            "//label[contains(., 'Non - Agricultural') or contains(., 'Non-Agricultural')]",
            "//input[@value='NON_AGRICULTURAL' or contains(@id, 'nonAgri') or contains(@id, 'Non')]/following-sibling::label",
            "//form/div[3]//label[contains(., 'Non')]",
        ]
        for xpath in non_agri_selectors:
            try:
                elems = driver.find_elements(By.XPATH, xpath)
                for el in elems:
                    if el.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                        driver.execute_script("arguments[0].click();", el)
                        wait_for_loader(driver)
                        time.sleep(0.3)
                        break
            except Exception:
                pass

        # 2. Ensure 'Property / Survey Number' radio is selected under Property Details
        prop_survey_selectors = [
            "/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[4]/div[1]/div[2]/label",
            "//form/div[4]/div[1]/div[2]/label",
            "//form/div[4]//label[contains(text(), 'Property / Survey') or contains(text(), 'Property/Survey') or contains(text(), 'Property Number')]",
            "//label[normalize-space()='Property / Survey Number' or normalize-space()='Property/Survey Number' or normalize-space()='Property Number']",
            "//input[@id='propertySurveyNo' or @value='property' or @value='PROPERTY_NO' or @value='2']/following-sibling::label",
        ]
        for xpath in prop_survey_selectors:
            try:
                elems = driver.find_elements(By.XPATH, xpath)
                for el in elems:
                    if el.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                        driver.execute_script("arguments[0].click();", el)
                        wait_for_loader(driver)
                        time.sleep(0.4)
                        return True
            except Exception:
                pass
    except Exception as e:
        if label:
            log(label, f"⚠️ Error ensuring property boundary radio selected: {e}")
    return False


def get_property_number_type_options(driver) -> List[Tuple[str, str]]:
    wait_for_loader(driver)
    time.sleep(0.5)
    ensure_property_boundary_selected(driver)
    try:
        selects = get_property_type_select_elements(driver)
        first_select = selects[0] if selects else None

        if not first_select:
            first_select = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//form//div[4]//select | (//form//select)[position() > 4]"))
            )

        select_obj = Select(first_select)
        valid_options = []
        for opt in select_obj.options:
            txt = opt.text.strip()
            val = opt.get_attribute("value")
            if txt and not any(p in txt.lower() for p in ["select", "choose", "--"]) and val not in ["", "0", None]:
                valid_options.append((val, txt))

        if not valid_options:
            for i, opt in enumerate(select_obj.options):
                if i > 0 and opt.text.strip():
                    valid_options.append((opt.get_attribute("value"), opt.text.strip()))

        return valid_options
    except Exception as e:
        log("System", f"⚠️ Error detecting Property Number Types: {e}")
        return []


def find_add_more_button(driver):
    candidates = [
        "//form//div[4]//*[contains(text(), 'Add More') or contains(text(), 'add more')]",
        "//button[contains(., 'Add More')]",
        "//a[contains(., 'Add More')]",
        "//span[contains(., 'Add More')]",
        "//*[normalize-space(text())='+ Add More' or normalize-space(text())='Add More' or contains(text(), 'Add More')]",
    ]
    for xpath in candidates:
        try:
            elems = driver.find_elements(By.XPATH, xpath)
            for el in elems:
                if el.is_displayed():
                    return el
        except Exception:
            pass
    return None


def get_property_type_select_elements(driver) -> List[Any]:
    """
    Returns ALL <select> elements corresponding to Property Number Type dropdown rows.
    Handles all 1 to 6 dynamic rows in the form.
    """
    # Strategy 1: JS query directly for all non-hierarchy form selects
    try:
        selects = driver.execute_script("""
            const allSelects = Array.from(document.querySelectorAll("form select, app-ec-search-citizen select"));
            return allSelects.filter(s => {
                if (!s || s.offsetParent === null) return false;
                const id = (s.id || '').toLowerCase();
                const name = (s.name || s.getAttribute('formcontrolname') || '').toLowerCase();
                const isLoc = id.includes('district') || id.includes('taluk') || id.includes('hobli') || id.includes('village')
                           || name.includes('district') || name.includes('taluk') || name.includes('hobli') || name.includes('village');
                return !isLoc;
            });
        """)
        if selects and len(selects) > 0:
            return selects
    except Exception:
        pass

    # Strategy 2: XPath querying all selects in div[4] (the property details container)
    try:
        elems = driver.find_elements(By.XPATH, "//app-ec-search-citizen//form//div[4]//select | //form/div[4]//select")
        vis = [e for e in elems if e.is_displayed()]
        if vis:
            return vis
    except Exception:
        pass

    # Strategy 3: Form selects with index > 4 (skipping District[1], Taluka[2], Hobli[3], Village[4])
    try:
        all_form_selects = driver.find_elements(By.XPATH, "//form//select")
        vis_selects = [s for s in all_form_selects if s.is_displayed()]
        if len(vis_selects) > 4:
            return vis_selects[4:]
    except Exception:
        pass

    return []


def select_dropdown_option(driver, select_elem, text=None, value=None):
    """
    Robustly selects an option in a Kaveri / Angular <select> by text or value,
    using Selenium native selection + DOM option clicks + Angular event dispatching.
    """
    if not select_elem:
        return False

    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_elem)
    except Exception:
        pass
    time.sleep(0.05)

    # 1. Wait briefly for options to populate if select is empty
    end_wait = time.time() + 3
    while time.time() < end_wait:
        try:
            opts = select_elem.find_elements(By.TAG_NAME, "option")
            if len(opts) > 1:
                break
        except Exception:
            pass
        time.sleep(0.1)

    selected = False
    clean_text = str(text).strip() if text else ""
    clean_val = str(value).strip() if value else ""

    # Strategy 1: Selenium Select by visible text
    if clean_text:
        try:
            sel = Select(select_elem)
            sel.select_by_visible_text(clean_text)
            selected = True
        except Exception:
            pass

    # Strategy 2: Selenium Select by value
    if not selected and clean_val:
        try:
            sel = Select(select_elem)
            sel.select_by_value(clean_val)
            selected = True
        except Exception:
            pass

    # Strategy 3: Click matching <option> element directly (Selenium native click & JS click)
    if not selected:
        try:
            for opt in select_elem.find_elements(By.TAG_NAME, "option"):
                opt_txt = opt.text.strip()
                opt_v = opt.get_attribute("value") or ""
                if (clean_text and (opt_txt.lower() == clean_text.lower() or clean_text.lower() in opt_txt.lower())) or \
                   (clean_val and opt_v == clean_val):
                    try:
                        opt.click()
                        selected = True
                        break
                    except Exception:
                        try:
                            driver.execute_script("arguments[0].selected = true; arguments[0].click();", opt)
                            selected = True
                            break
                        except Exception:
                            pass
        except Exception:
            pass

    # Strategy 4: Angular JavaScript selection matching text or value + dispatching events
    try:
        driver.execute_script("""
            var sel = arguments[0];
            var targetText = (arguments[1] || '').trim().toLowerCase();
            var targetVal = (arguments[2] || '').trim();
            
            var matchedIdx = -1;
            for (var i = 0; i < sel.options.length; i++) {
                var opt = sel.options[i];
                var optTxt = (opt.text || '').trim().toLowerCase();
                var optVal = (opt.value || '').trim();
                
                if (targetText && optTxt === targetText) {
                    matchedIdx = i;
                    break;
                }
                if (targetVal && optVal === targetVal) {
                    matchedIdx = i;
                    break;
                }
            }
            if (matchedIdx === -1 && targetText) {
                for (var i = 0; i < sel.options.length; i++) {
                    var optTxt = (sel.options[i].text || '').trim().toLowerCase();
                    if (optTxt.includes(targetText) || targetText.includes(optTxt)) {
                        matchedIdx = i;
                        break;
                    }
                }
            }
            
            if (matchedIdx !== -1) {
                sel.selectedIndex = matchedIdx;
                sel.options[matchedIdx].selected = true;
                sel.value = sel.options[matchedIdx].value;
                sel.dispatchEvent(new Event('input', { bubbles: true }));
                sel.dispatchEvent(new Event('change', { bubbles: true }));
                sel.dispatchEvent(new Event('blur', { bubbles: true }));
            }
        """, select_elem, clean_text, clean_val)
    except Exception:
        pass

    # Final event trigger
    try:
        driver.execute_script("""
            var sel = arguments[0];
            sel.dispatchEvent(new Event('input', { bubbles: true }));
            sel.dispatchEvent(new Event('change', { bubbles: true }));
            sel.dispatchEvent(new Event('blur', { bubbles: true }));
        """, select_elem)
    except Exception:
        pass

    return True


def reset_property_details_rows(driver):
    """Removes extra property type rows until only 1 row remains without toggling agricultural radio."""
    try:
        wait_for_loader(driver)
        for _ in range(12):
            current_selects = get_property_type_select_elements(driver)
            if len(current_selects) <= 1:
                return True

            remove_btns = driver.find_elements(By.XPATH, "//form//div[4]//button[normalize-space()='Remove' or contains(text(), 'Remove')] | //button[normalize-space()='Remove' or contains(text(), 'Remove')] | //a[normalize-space()='Remove' or contains(text(), 'Remove')]")
            vis_remove = [b for b in remove_btns if b.is_displayed()]
            if not vis_remove:
                break
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", vis_remove[-1])
                driver.execute_script("arguments[0].click();", vis_remove[-1])
                time.sleep(0.15)
                wait_for_loader(driver)
            except Exception:
                pass

        current_selects = get_property_type_select_elements(driver)
        return len(current_selects) <= 1
    except Exception as e:
        log("System", f"⚠️ Error resetting property rows: {e}")
        return False


def is_batch_already_setup(driver, batch_options):
    """Checks if the existing dropdown rows in the DOM already match the batch options."""
    if not batch_options:
        return False
    try:
        selects = get_property_type_select_elements(driver)
        if len(selects) != len(batch_options):
            return False
        for idx, (val, txt) in enumerate(batch_options):
            curr_txt = Select(selects[idx]).first_selected_option.text.strip()
            if not curr_txt or any(p in curr_txt.lower() for p in ["select", "choose", "--"]):
                return False
            if txt.strip().lower() not in curr_txt.lower() and curr_txt.lower() not in txt.strip().lower():
                return False
        return True
    except Exception:
        return False


def setup_property_type_batch(driver, batch_options):
    if not batch_options:
        return False

    wait_for_loader(driver)
    time.sleep(0.1)

    target_count = len(batch_options)
    current_selects = get_property_type_select_elements(driver)

    # 1. If more selects than target_count, remove extra rows from the bottom
    while len(current_selects) > target_count:
        remove_btns = driver.find_elements(By.XPATH, "//form//div[4]//button[normalize-space()='Remove' or contains(text(), 'Remove')] | //button[normalize-space()='Remove' or contains(text(), 'Remove')] | //a[normalize-space()='Remove' or contains(text(), 'Remove')]")
        vis_remove = [b for b in remove_btns if b.is_displayed()]
        if not vis_remove:
            break
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", vis_remove[-1])
            driver.execute_script("arguments[0].click();", vis_remove[-1])
            time.sleep(0.15)
            wait_for_loader(driver)
        except Exception:
            break
        current_selects = get_property_type_select_elements(driver)

    # 2. If fewer selects than target_count, click + Add More
    max_add_tries = 10
    add_tries = 0
    while len(current_selects) < target_count and add_tries < max_add_tries:
        add_tries += 1
        add_more_btn = find_add_more_button(driver)
        if not add_more_btn:
            break

        prev_len = len(current_selects)
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", add_more_btn)
            driver.execute_script("arguments[0].click();", add_more_btn)
        except Exception:
            try:
                add_more_btn.click()
            except Exception:
                pass

        end_time = time.time() + 2.5
        while time.time() < end_time:
            current_selects = get_property_type_select_elements(driver)
            if len(current_selects) > prev_len:
                break
            time.sleep(0.1)

    time.sleep(0.15)
    current_selects = get_property_type_select_elements(driver)

    # 3. Select each property type into its corresponding dropdown row
    for idx, (val, txt) in enumerate(batch_options):
        if idx < len(current_selects):
            select_dropdown_option(driver, current_selects[idx], text=txt, value=val)
            time.sleep(0.08)

    # 4. Verification & retry pass
    time.sleep(0.2)
    current_selects = get_property_type_select_elements(driver)
    for idx, (val, txt) in enumerate(batch_options):
        if idx < len(current_selects):
            try:
                sel_text = Select(current_selects[idx]).first_selected_option.text.strip()
                if not sel_text or any(p in sel_text.lower() for p in ["select", "choose", "--"]):
                    select_dropdown_option(driver, current_selects[idx], text=txt, value=val)
            except Exception:
                select_dropdown_option(driver, current_selects[idx], text=txt, value=val)

    return True


def select_date(driver, day, month, year, date_type="first"):
    try:
        date_picker_xpaths = {
            "first": '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[8]/div[1]/span/mat-datepicker-toggle/button',
            "second": '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[8]/div[2]/span/mat-datepicker-toggle/button'
        }
        wait_for_loader(driver)
        calendar_button = None
        if date_type in date_picker_xpaths:
            try:
                calendar_button = driver.find_element(By.XPATH, date_picker_xpaths[date_type])
            except Exception:
                pass

        if not calendar_button:
            toggles = driver.find_elements(By.XPATH, "//mat-datepicker-toggle//button")
            vis = [t for t in toggles if t.is_displayed()]
            calendar_button = vis[0] if date_type == "first" and vis else (vis[1] if len(vis) > 1 else None)

        if not calendar_button:
            return False

        calendar_button.click()

        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, 'mat-datepicker-content'))
        )
        calendar_dialog = driver.find_element(By.CLASS_NAME, 'mat-datepicker-content')

        year_button = WebDriverWait(calendar_dialog, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, 'mat-calendar-period-button'))
        )
        year_button.click()

        years = WebDriverWait(calendar_dialog, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, 'mat-calendar-body-cell'))
        )
        for y in years:
            if y.text == str(year):
                y.click()
                break

        months = WebDriverWait(calendar_dialog, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, 'mat-calendar-body-cell'))
        )
        for m in months:
            if m.text.upper()[:3] == month.upper()[:3]:
                m.click()
                break

        days = WebDriverWait(calendar_dialog, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.mat-calendar-body-cell:not(.mat-calendar-body-disabled)'))
        )
        for d in days:
            if d.text == str(day):
                d.click()
                break

        time.sleep(0.5)
        try:
            WebDriverWait(driver, 5).until(
                EC.invisibility_of_element_located((By.CLASS_NAME, 'mat-datepicker-content'))
            )
        except Exception:
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            except Exception:
                pass
        wait_for_loader(driver)
        time.sleep(0.5)
        return True
    except Exception as e:
        log("System", f"Error selecting {date_type} date: {str(e)}")
        return False


def get_selected_dropdown_text(driver, xpath) -> str:
    xpath = re.sub(r'/option\[\d+\]', '', xpath)
    fallback_index = 1
    if "div[2]" in xpath or "div[1]/div[2]" in xpath:
        fallback_index = 2
    elif "div[3]" in xpath or "div[1]/div[3]" in xpath:
        fallback_index = 3
    elif "div[4]" in xpath or "div[1]/div[4]" in xpath:
        fallback_index = 4

    try:
        select_elements = driver.find_elements(By.XPATH, xpath)
        if not select_elements:
            select_elements = driver.find_elements(By.XPATH, f"(//form//select)[{fallback_index}]")
        if select_elements:
            select_element = select_elements[0]
            selected_option = Select(select_element).first_selected_option
            text = selected_option.text.strip()
            if text and not any(marker in text.lower() for marker in ("select", "choose", "--")):
                return text
    except Exception:
        pass
    try:
        elem = driver.find_element(By.XPATH, xpath)
        text = driver.execute_script("return arguments[0].options[arguments[0].selectedIndex].text;", elem).strip()
        if text and not any(marker in text.lower() for marker in ("select", "choose", "--")):
            return text
    except Exception:
        pass
    return ""


def select_dropdown_value(driver, xpath, value):
    xpath = re.sub(r'/option\[\d+\]', '', xpath)
    wait_for_page_ready(driver)
    wait_for_loader(driver)
    
    fallback_index = 1
    if "div[2]" in xpath or "div[1]/div[2]" in xpath:
        fallback_index = 2
    elif "div[3]" in xpath or "div[1]/div[3]" in xpath:
        fallback_index = 3
    elif "div[4]" in xpath or "div[1]/div[4]" in xpath:
        fallback_index = 4

    select_element = None
    try:
        select_elements = driver.find_elements(By.XPATH, xpath)
        if not select_elements:
            select_elements = driver.find_elements(By.XPATH, f"(//form//select)[{fallback_index}]")
        if select_elements:
            select_element = select_elements[0]
        else:
            select_element = WebDriverWait(driver, 10).until(
                lambda d: d.find_elements(By.XPATH, xpath) or d.find_elements(By.XPATH, f"(//form//select)[{fallback_index}]")
            )[0]
    except Exception:
        select_element = driver.find_element(By.XPATH, f"(//form//select)[{fallback_index}]")

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_element)
    time.sleep(0.2)

    end_time = time.time() + 10
    while time.time() < end_time:
        try:
            val_strs = [str(o.get_attribute("value")).strip() for o in Select(select_element).options]
            if str(value).strip() in val_strs:
                break
        except Exception:
            pass
        time.sleep(0.3)

    try:
        Select(select_element).select_by_value(str(value))
    except Exception:
        pass

    driver.execute_script("""
        var sel = arguments[0];
        var val = arguments[1];
        if (val !== undefined && val !== null) {
            sel.value = val;
        }
        sel.dispatchEvent(new Event('input', { bubbles: true }));
        sel.dispatchEvent(new Event('change', { bubbles: true }));
        sel.dispatchEvent(new Event('blur', { bubbles: true }));
    """, select_element, str(value))
    wait_for_loader(driver)
    time.sleep(0.8)


def get_dropdown_options(driver, xpath, timeout=15, dedup_by_name=True, label=None) -> List[Tuple[str, str]]:
    xpath = re.sub(r'/option\[\d+\]', '', xpath)
    wait_for_loader(driver)
    
    fallback_index = 1
    if "div[2]" in xpath or "div[1]/div[2]" in xpath:
        fallback_index = 2
    elif "div[3]" in xpath or "div[1]/div[3]" in xpath:
        fallback_index = 3
    elif "div[4]" in xpath or "div[1]/div[4]" in xpath:
        fallback_index = 4

    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            select_elements = driver.find_elements(By.XPATH, xpath)
            if not select_elements:
                select_elements = driver.find_elements(By.XPATH, f"(//form//select)[{fallback_index}]")
            if not select_elements:
                time.sleep(0.4)
                continue
            select_element = select_elements[0]
            options = []
            seen_names = set()
            for option in Select(select_element).options:
                text = option.text.strip()
                value = option.get_attribute("value")
                if text and value not in (None, "", "0") and not any(
                    marker in text.lower() for marker in ("select", "choose", "--")
                ):
                    norm = text.lower()
                    if dedup_by_name:
                        if norm in seen_names:
                            continue
                        seen_names.add(norm)
                    options.append((value, text))
            if options:
                return options
        except Exception:
            pass
        time.sleep(0.4)
    return []


def navigate_to_search_page(driver, label="Browser"):
    log(label, "🧭 Navigating to EC Search Citizen page...")
    close_landing_popup(driver)
    close_result_modal(driver)
    wait_for_loader(driver)
    wait_for_page_ready(driver)

    is_on_search = False
    try:
        if driver.find_elements(By.TAG_NAME, "app-ec-search-citizen") or "ec-search" in driver.current_url.lower():
            is_on_search = True
    except Exception:
        pass

    if not is_on_search:
        # 1. Click dashboard link / button
        wait_for_loader(driver)
        dashboard_clicked = False
        dashboard_selectors = [
            "/html/body/div[1]/app-root/div/div/app-kaveri-dashboard/div/div[1]/div/div[1]/div[3]/div/button/a",
            "//app-kaveri-dashboard//div[1]//div[3]//button/a",
            "//app-kaveri-dashboard//button/a",
            "//a[contains(text(), 'Services') or contains(text(), 'Start') or contains(@href, 'application')]",
            "//button[contains(., 'Start') or contains(., 'Application') or contains(., 'Service') or contains(., 'ಸೇವೆ')]",
            "//app-kaveri-dashboard//button",
        ]
        for xpath in dashboard_selectors:
            try:
                elems = driver.find_elements(By.XPATH, xpath)
                for el in elems:
                    if el.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                        try:
                            el.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", el)
                        dashboard_clicked = True
                        log(label, "Clicked Start a New Application button on dashboard.")
                        break
                if dashboard_clicked:
                    break
            except Exception:
                pass

        wait_for_loader(driver)
        time.sleep(1)

        # 2. Click document / application type image (Encumbrance Certificate)
        app_type_clicked = False
        app_type_selectors = [
            "/html/body/div[1]/app-root/div/div/app-kaveri-dashboard/div/div[2]/div/app-application-type/div/div/div[2]/div[2]/div[2]/img",
            "//app-application-type//div[2]//div[2]//img",
            "//app-application-type//div[contains(., 'Encumbrance') or contains(., 'EC') or contains(., 'ಋಣಭಾರ')]//img",
            "(//app-application-type//img)[2]",
            "//app-application-type//img",
        ]
        for xpath in app_type_selectors:
            try:
                elems = driver.find_elements(By.XPATH, xpath)
                for el in elems:
                    if el.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                        try:
                            el.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", el)
                        app_type_clicked = True
                        log(label, "Clicked Encumbrance Certificate (EC) card.")
                        break
                if app_type_clicked:
                    break
            except Exception:
                pass

        wait_for_loader(driver)
        time.sleep(1)

        # 3. Dismiss prerequisite modals (Modal 1, Modal 2, Modal 3)
        for dialog_step in range(1, 5):
            wait_for_loader(driver)
            time.sleep(1)
            dialog_btns = driver.find_elements(
                By.XPATH,
                "//app-perquisite//button | //mat-dialog-container//button | //div[@role='dialog']//button | //button[normalize-space()='Continue' or normalize-space()='OK' or normalize-space()='Agree' or normalize-space()='Proceed' or contains(text(), 'ಮುಂದುವರೆಯಿರಿ') or contains(text(), 'ಸಮ್ಮತಿಸುತ್ತೇನೆ')]"
            )
            if not dialog_btns:
                break
            btn_clicked = False
            for btn in reversed(dialog_btns):
                if btn.is_displayed():
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                        driver.execute_script("arguments[0].click();", btn)
                        btn_clicked = True
                        log(label, f"Dismissed prerequisite dialog step {dialog_step}.")
                        break
                    except Exception:
                        pass
            if not btn_clicked:
                break
            wait_for_loader(driver)
            time.sleep(1)

    # 4. Wait for EC Search page elements to appear
    wait_for_loader(driver)
    WebDriverWait(driver, 30).until(
        lambda d: d.find_elements(By.XPATH, DISTRICT_XPATH) or d.find_elements(By.XPATH, "(//form//select)[1]")
    )

    # 5. Select Target District & Taluka with full AJAX synchronization
    select_district_and_taluka(
        driver,
        target_district_name=TARGET_DISTRICT_NAME,
        target_district_index=TARGET_DISTRICT_INDEX,
        target_taluka_name=TARGET_TALUKA_NAME,
        target_taluka_index=TARGET_TALUKA_INDEX,
        label=label
    )

    # 6. Ensure Property / Survey Number radio
    ensure_property_boundary_selected(driver, label=label)

    # 7. Set initial dates
    try:
        now = datetime.now()
        curr_day = now.day
        curr_month = now.strftime("%b").upper()
        curr_year = now.year
        select_date(driver, 1, "JAN", 2020, "first")
        select_date(driver, curr_day, curr_month, curr_year, "second")
    except Exception:
        pass

    log(label, "✅ Reached EC Search page; District and Taluka ready.")


def select_district_and_taluka(
    driver,
    target_district_name=TARGET_DISTRICT_NAME,
    target_district_index=TARGET_DISTRICT_INDEX,
    target_taluka_name=TARGET_TALUKA_NAME,
    target_taluka_index=TARGET_TALUKA_INDEX,
    label="Browser"
):
    """
    Robustly selects the target District (Basavanagudi / Bengaluru Urban) and Taluka (Anekal)
    matching by name first, with configurable option index fallbacks (e.g. index 3 for District, index 1 for Taluka).
    """
    wait_for_page_ready(driver)
    wait_for_loader(driver)

    # 1. District Selection
    district_elem = WebDriverWait(driver, 25).until(
        lambda d: (
            d.find_element(By.XPATH, DISTRICT_XPATH)
            if d.find_elements(By.XPATH, DISTRICT_XPATH)
            else d.find_element(By.XPATH, "(//form//select)[1]")
        )
    )

    curr_dist = get_selected_dropdown_text(driver, DISTRICT_XPATH) or ""
    is_district_valid = curr_dist and any(
        k in curr_dist.lower() for k in [
            target_district_name.lower(), "basavanagudi", "basavanagaudi", "ಬಸವನಗುಡಿ", "bengaluru urban", "bangalore urban", "ಬೆಂಗಳೂರು ನಗರ", "bengaluru", "bangalore"
        ]
    )

    if is_district_valid:
        log(label, f"📍 District already set to: '{curr_dist}'")
    else:
        log(label, f"🔍 Waiting for District dropdown options to populate from Kaveri server...")
        end_time = time.time() + 25
        while time.time() < end_time:
            wait_for_loader(driver, timeout=2)
            opts = Select(district_elem).options
            if len(opts) > 1:
                break
            time.sleep(0.5)

        select_obj = Select(district_elem)
        target_dist_val = None
        target_dist_text = ""

        # Search by keywords
        for opt in select_obj.options:
            txt = opt.text.strip()
            val = opt.get_attribute("value")
            if not val or val in ["", "0", None]:
                continue
            if any(k in txt.lower() for k in [target_district_name.lower(), "basavanagudi", "basavanagaudi", "ಬಸವನಗುಡಿ", "bengaluru urban", "bangalore urban", "ಬೆಂಗಳೂರು ನಗರ"]):
                target_dist_val = val
                target_dist_text = txt
                break

        # Fallback to index
        if not target_dist_val:
            if target_district_index is not None and len(select_obj.options) > target_district_index:
                target_dist_val = select_obj.options[target_district_index].get_attribute("value")
                target_dist_text = select_obj.options[target_district_index].text.strip()
                log(label, f"📍 District matched by index [{target_district_index}]: '{target_dist_text}'")
            elif len(select_obj.options) > 1:
                target_dist_val = select_obj.options[1].get_attribute("value")
                target_dist_text = select_obj.options[1].text.strip()

        if target_dist_val:
            log(label, f"📍 Selecting District: '{target_dist_text}' (value: {target_dist_val})")
            select_dropdown_value(driver, DISTRICT_XPATH, target_dist_val)
        else:
            avail = [o.text.strip() for o in Select(district_elem).options]
            log(label, f"⚠️ District '{target_district_name}' not found in {avail}!")

    wait_for_loader(driver)
    time.sleep(1)

    # 2. Taluka Selection
    taluka_elem = WebDriverWait(driver, 25).until(
        lambda d: (
            d.find_element(By.XPATH, TALUKA_XPATH)
            if d.find_elements(By.XPATH, TALUKA_XPATH)
            else d.find_element(By.XPATH, "(//form//select)[2]")
        )
    )

    curr_tal = get_selected_dropdown_text(driver, TALUKA_XPATH) or ""
    hobli_opts = get_dropdown_options(driver, HOBLI_XPATH, timeout=2)

    if curr_tal and any(k in curr_tal.lower() for k in [target_taluka_name.lower(), "anekal", "ಆನೇಕಲ್"]) and hobli_opts:
        log(label, f"📍 Taluka already set to: '{curr_tal}' ({len(hobli_opts)} Hoblis loaded)")
    else:
        log(label, f"🔍 Waiting for Taluka dropdown options for District...")
        end_time = time.time() + 25
        while time.time() < end_time:
            wait_for_loader(driver, timeout=2)
            opts = Select(taluka_elem).options
            if len(opts) > 1:
                break
            time.sleep(0.5)

        select_obj = Select(taluka_elem)
        target_tal_val = None
        target_tal_text = ""

        # Search by keywords
        for opt in select_obj.options:
            txt = opt.text.strip()
            val = opt.get_attribute("value")
            if not val or val in ["", "0", None]:
                continue
            if any(k in txt.lower() for k in [target_taluka_name.lower(), "anekal", "ಆನೇಕಲ್"]):
                target_tal_val = val
                target_tal_text = txt
                break

        # Fallback to index
        if not target_tal_val:
            if target_taluka_index is not None and len(select_obj.options) > target_taluka_index:
                target_tal_val = select_obj.options[target_taluka_index].get_attribute("value")
                target_tal_text = select_obj.options[target_taluka_index].text.strip()
                log(label, f"📍 Taluka matched by index [{target_taluka_index}]: '{target_tal_text}'")
            elif len(select_obj.options) > 1:
                target_tal_val = select_obj.options[1].get_attribute("value")
                target_tal_text = select_obj.options[1].text.strip()

        if target_tal_val:
            log(label, f"📍 Selecting Taluka: '{target_tal_text}' (value: {target_tal_val})")
            select_dropdown_value(driver, TALUKA_XPATH, target_tal_val)
        else:
            avail = [o.text.strip() for o in Select(taluka_elem).options]
            log(label, f"⚠️ Taluka '{target_taluka_name}' not found in {avail}!")

    wait_for_loader(driver)
    time.sleep(1)

    # 3. Wait for Hobli options to populate
    end_hobli_wait = time.time() + 15
    while time.time() < end_hobli_wait:
        if get_dropdown_options(driver, HOBLI_XPATH, timeout=2):
            break
        time.sleep(0.5)


# ======================================================================================
# DATA PERSISTENCE & DEDUPLICATION
# ======================================================================================
def extract_valid_records_from_table(
    table_html, survey_value, batch_idx, batch_types_str="",
    district_name="", taluka_name="", hobli_name="", village_name=""
) -> Optional[pd.DataFrame]:
    if not table_html:
        return None

    df = None
    try:
        soup = BeautifulSoup(str(table_html), 'html.parser')
        target_table = soup.find('table') or soup
        dfs = pd.read_html(io.StringIO(str(target_table)))
        if dfs and not dfs[0].empty:
            df = dfs[0]
    except Exception:
        try:
            soup = BeautifulSoup(str(table_html), 'html.parser')
            table = soup.find('table') if '<table' in str(table_html) else soup
            if table:
                rows = []
                for tr in table.find_all('tr'):
                    cells = [c.get_text(strip=True) for c in tr.find_all(['th', 'td'])]
                    if cells:
                        rows.append(cells)
                if len(rows) >= 2:
                    headers = rows[0]
                    max_cols = max(len(r) for r in rows)
                    if len(headers) < max_cols:
                        headers = headers + [f"Col_{i}" for i in range(len(headers), max_cols)]
                    data = [r + [''] * (len(headers) - len(r)) for r in rows[1:]]
                    df = pd.DataFrame(data, columns=headers[:len(data[0]) if data else len(headers)])
        except Exception:
            return None

    if df is None or df.empty:
        return None

    try:
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [' '.join(str(c) for c in col).strip() for col in df.columns.values]
        df.columns = [str(c).strip() for c in df.columns]

        if len(df) <= 1:
            all_text = " ".join(df.astype(str).values.flatten()).lower()
            if any(marker in all_text for marker in ["no data", "no record", "record not found", "not available", "ಮಾಹಿತಿ ಲಭ್ಯವಿಲ್ಲ"]):
                return None

        doc_col = next(
            (c for c in df.columns if any(k in str(c).lower() for k in [
                'ದಸ್ತಾವೇಜಿನ ಉಲ್ಲೇಖ', 'ದಸ್ತಾವೇಜು', 'document reference', 'doc ref', 'registration no', 'doc no'
            ])),
            None
        )
        if not doc_col:
            return None

        df[doc_col] = df[doc_col].ffill()
        invalid_doc_words = ['select', 'no data', 'record not', 'null', 'nan', '-- select --']
        valid_mask = df[doc_col].dropna().astype(str).str.strip().apply(
            lambda d: len(d) >= 3 and not any(w == d.lower() or f"{w} " in d.lower() for w in invalid_doc_words)
        )
        if valid_mask.any():
            df = df[valid_mask].copy()
        else:
            return None

        if df.empty:
            return None

        df['District'] = district_name
        df['Taluka'] = taluka_name
        df['Hobli/Town'] = hobli_name
        df['Village'] = village_name
        df['Survey Value'] = survey_value
        df['Batch'] = batch_idx
        if batch_types_str:
            df['Property Types Searched'] = batch_types_str
        return df
    except Exception:
        return None


def save_df_to_excel(df: pd.DataFrame, excel_file: str):
    """Saves DataFrame to excel_file atomically with deduplication against existing records."""
    if df is None or df.empty:
        return 0

    doc_col = next(
        (c for c in df.columns if any(k in str(c).lower() for k in [
            'ದಸ್ತಾವೇಜಿನ ಉಲ್ಲೇಖ', 'ದಸ್ತಾವೇಜು', 'document reference', 'doc ref', 'registration no', 'doc no'
        ])),
        None
    )
    if not doc_col:
        return 0

    temp_excel = f"{excel_file}.tmp.xlsx"
    junk_cols = [c for c in df.columns if any(k in str(c).lower() for k in ['enter one', 'unnamed', 'add more'])]
    if junk_cols:
        df = df.drop(columns=junk_cols)

    saved_count = 0
    if os.path.exists(excel_file):
        try:
            existing_df = pd.read_excel(excel_file)
            existing_doc_col = next(
                (c for c in existing_df.columns if any(k in str(c).lower() for k in [
                    'ದಸ್ತಾವೇಜಿನ ಉಲ್ಲೇಖ', 'ದಸ್ತಾವೇಜು', 'document reference', 'doc ref', 'registration no', 'doc no'
                ])),
                None
            )
            if existing_doc_col and doc_col:
                existing_refs = set(existing_df[existing_doc_col].dropna().astype(str).str.strip())
                new_rows = df[~df[doc_col].astype(str).str.strip().isin(existing_refs)]
            else:
                new_rows = df
        except Exception:
            new_rows = df

        if not new_rows.empty:
            combined_df = pd.concat([existing_df, new_rows], ignore_index=True)
            junk = [c for c in combined_df.columns if any(k in str(c).lower() for k in ['enter one', 'unnamed', 'add more'])]
            if junk:
                combined_df = combined_df.drop(columns=junk)
            combined_df.to_excel(temp_excel, index=False, engine='openpyxl')
            while True:
                try:
                    os.replace(temp_excel, excel_file)
                    saved_count = len(new_rows)
                    break
                except (PermissionError, OSError):
                    log("Excel", f"⚠️ Excel file locked. Please close '{os.path.basename(excel_file)}' to save. Waiting...")
                    time.sleep(2)
    else:
        df.to_excel(temp_excel, index=False, engine='openpyxl')
        while True:
            try:
                os.replace(temp_excel, excel_file)
                saved_count = len(df)
                break
            except (PermissionError, OSError):
                time.sleep(2)

    return saved_count


# ======================================================================================
# PROPERTY TYPES REFERENCE STORAGE
# ======================================================================================
def save_property_types_reference(batches: List[List[Tuple[str, str]]]):
    """Saves all detected property number types and their assigned batches to JSON & Excel."""
    if not batches:
        return

    flat_records = []
    item_idx = 1
    for b_idx, batch in enumerate(batches, start=1):
        for val, txt in batch:
            flat_records.append({
                "Batch Number": f"Batch {b_idx}",
                "Batch Index": b_idx,
                "Property Number Type Index": item_idx,
                "Property Type Option Value": str(val),
                "Property Type Name": str(txt),
            })
            item_idx += 1

    tmp_json = f"{PROPERTY_TYPES_JSON}.tmp"
    try:
        with open(tmp_json, "w", encoding="utf-8") as f:
            json.dump(flat_records, f, indent=2, ensure_ascii=False)
        os.replace(tmp_json, PROPERTY_TYPES_JSON)
        log("System", f"📑 Saved {len(flat_records)} property types to {os.path.basename(PROPERTY_TYPES_JSON)}")
    except Exception as e:
        log("System", f"⚠️ Error saving {os.path.basename(PROPERTY_TYPES_JSON)}: {e}")

    try:
        df = pd.DataFrame(flat_records)
        tmp_xlsx = f"{PROPERTY_TYPES_XLSX}.tmp.xlsx"
        df.to_excel(tmp_xlsx, index=False, engine="openpyxl")
        os.replace(tmp_xlsx, PROPERTY_TYPES_XLSX)
        log("System", f"📑 Saved property types reference sheet to {os.path.basename(PROPERTY_TYPES_XLSX)}")
    except Exception as e:
        log("System", f"⚠️ Error saving {os.path.basename(PROPERTY_TYPES_XLSX)}: {e}")


# ======================================================================================
# PROPERTY NUMBER & SEARCH HISTORY LOGGING (JSON & EXCEL)
# ======================================================================================
def log_search_history(
    account_label: str,
    hobli_name: str,
    village_name: str,
    survey_number: int,
    batch_idx: int,
    total_batches: int,
    batch_types: List[Tuple[str, str]],
    records_found: int,
    status: str
):
    """Logs each search query execution with property types and numbers to JSON & Excel."""
    entry = {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Account": account_label,
        "Hobli": hobli_name,
        "Village": village_name,
        "Property Number (Survey #)": survey_number,
        "Batch": f"Batch {batch_idx}/{total_batches}",
        "Batch Size": len(batch_types),
        "Property Types Searched": ", ".join(txt for _, txt in batch_types),
        "Records Found": records_found,
        "Status": status
    }

    # JSON History
    history_list = []
    if os.path.exists(SEARCH_HISTORY_JSON):
        try:
            with open(SEARCH_HISTORY_JSON, "r", encoding="utf-8") as f:
                history_list = json.load(f)
            if not isinstance(history_list, list):
                history_list = []
        except Exception:
            history_list = []

    history_list.append(entry)
    tmp_json = f"{SEARCH_HISTORY_JSON}.tmp"
    try:
        with open(tmp_json, "w", encoding="utf-8") as f:
            json.dump(history_list, f, indent=2, ensure_ascii=False)
        os.replace(tmp_json, SEARCH_HISTORY_JSON)
    except Exception:
        pass

    # Excel History
    tmp_xlsx = f"{SEARCH_HISTORY_XLSX}.tmp.xlsx"
    try:
        if os.path.exists(SEARCH_HISTORY_XLSX):
            try:
                edf = pd.read_excel(SEARCH_HISTORY_XLSX)
                df = pd.concat([edf, pd.DataFrame([entry])], ignore_index=True)
            except Exception:
                df = pd.DataFrame([entry])
        else:
            df = pd.DataFrame([entry])

        df.to_excel(tmp_xlsx, index=False, engine="openpyxl")
        os.replace(tmp_xlsx, SEARCH_HISTORY_XLSX)
    except Exception:
        pass


# ======================================================================================
# PROGRESS STATE & VILLAGE SUMMARY
# ======================================================================================
def load_rotation_progress() -> Dict[str, Any]:
    """Loads fine-grained resumption state for account rotation."""
    default_state = {
        "account_index": 0,
        "searches_used_in_account": 0,
        "hobli_index": 0,
        "village_index": 0,
        "village_name": "",
        "property_value": 1,
        "batch_index": 0,
        "consecutive_empty": 0,
        "current_survey_had_records": False
    }
    if not os.path.exists(PROGRESS_FILE):
        return default_state
    try:
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return {
                "account_index": max(0, int(data.get("account_index", 0))),
                "searches_used_in_account": max(0, int(data.get("searches_used_in_account", 0))),
                "hobli_index": max(0, int(data.get("hobli_index", 0))),
                "village_index": max(0, int(data.get("village_index", 0))),
                "village_name": str(data.get("village_name", "")),
                "property_value": max(1, int(data.get("property_value", 1))),
                "batch_index": max(0, int(data.get("batch_index", 0))),
                "consecutive_empty": max(0, int(data.get("consecutive_empty", 0))),
                "current_survey_had_records": bool(data.get("current_survey_had_records", False)),
            }
    except Exception:
        pass
    return default_state


def save_rotation_progress(state: Dict[str, Any]):
    """Atomically saves the fine-grained state."""
    tmp = f"{PROGRESS_FILE}.tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
        os.replace(tmp, PROGRESS_FILE)
    except Exception:
        pass


def load_village_summary() -> Dict[str, Any]:
    if os.path.exists(VILLAGE_SUMMARY_JSON):
        try:
            with open(VILLAGE_SUMMARY_JSON, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_village_summary(summary_data):
    tmp_json = f"{VILLAGE_SUMMARY_JSON}.tmp"
    try:
        with open(tmp_json, 'w', encoding='utf-8') as f:
            json.dump(summary_data, f, indent=2, ensure_ascii=False)
        os.replace(tmp_json, VILLAGE_SUMMARY_JSON)
    except Exception:
        pass

    if summary_data:
        try:
            rows = list(summary_data.values())
            sdf = pd.DataFrame(rows)
            cols_map = [
                ('district', 'District'),
                ('taluka', 'Taluka'),
                ('hobli', 'Hobli'),
                ('village', 'Village'),
                ('scraped_by', 'Scraped By'),
                ('last_scraped_value', 'Last Scraped Value'),
                ('last_searched_value', 'Last Searched Value'),
                ('total_records', 'Total Records'),
                ('status', 'Status'),
                ('last_updated', 'Last Updated')
            ]
            for col, _ in cols_map:
                if col not in sdf.columns:
                    sdf[col] = ''
            sdf = sdf[[col for col, _ in cols_map]]
            sdf.columns = [header for _, header in cols_map]
            tmp_xlsx = f"{VILLAGE_SUMMARY_XLSX}.tmp.xlsx"
            sdf.to_excel(tmp_xlsx, index=False, engine='openpyxl')
            try:
                os.replace(tmp_xlsx, VILLAGE_SUMMARY_XLSX)
            except Exception:
                pass
        except Exception:
            pass


def update_village_summary(village_name, hobli_name, taluka_name, district_name, searched_val=None, scraped_val=None, new_recs_count=0, is_completed=False, label=""):
    if not village_name:
        return
    v_key = str(village_name).strip()
    summary = load_village_summary()
    record = summary.get(v_key, {
        'district': str(district_name or 'Basavanagudi'),
        'taluka': str(taluka_name or 'Anekal'),
        'hobli': str(hobli_name or 'Anekal Kasaba Hobli'),
        'village': v_key,
        'scraped_by': str(label or ''),
        'last_scraped_value': 0,
        'last_searched_value': 0,
        'total_records': 0,
        'status': 'IN_PROGRESS',
        'last_updated': ''
    })
    if searched_val is not None:
        record['last_searched_value'] = max(int(record.get('last_searched_value', 0)), int(searched_val))
    if scraped_val is not None:
        record['last_scraped_value'] = max(int(record.get('last_scraped_value', 0)), int(scraped_val))
    if new_recs_count > 0:
        record['total_records'] = int(record.get('total_records', 0)) + int(new_recs_count)
    if is_completed:
        record['status'] = 'COMPLETED'
    record['last_updated'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    summary[v_key] = record
    save_village_summary(summary)


# ======================================================================================
# CORE BATCH SEARCH EXECUTION
# ======================================================================================
def execute_single_batch_search(driver, survey_value: int, batch_idx: int, label="Scraper") -> Tuple[bool, Optional[str]]:
    """
    Executes search for a given survey value and property batch.
    Targets ONLY the property details inputs (leaves boundary & measurement untouched).
    Uses the exact hybrid CAPTCHA solver logic from igr_karnataka.py.
    """
    wait_for_loader(driver)
    close_result_modal(driver)
    wait_for_page_ready(driver)

    # 1. Fill survey value ONLY into Property Details inputs (PRESENT PROPERTY NUMBER DETAILS)
    driver.execute_script("""
        const val = arguments[0];
        const inps = Array.from(document.querySelectorAll("form div:nth-child(4) input, app-ec-search-citizen form div:nth-child(4) input, form input"))
            .filter(inp => {
                if (!inp || inp.offsetParent === null) return false;
                const ph = (inp.placeholder || '').toLowerCase();
                const name = (inp.getAttribute('formcontrolname') || '').toLowerCase();
                const id = (inp.id || '').toLowerCase();
                const isDate = ph.includes('date') || ph.includes('dd/mm') || name.includes('date') || id.includes('date');
                const isCaptcha = ph.includes('captcha') || name.includes('captcha') || id.includes('captcha');
                return !isDate && !isCaptcha && inp.type !== 'radio' && inp.type !== 'checkbox' && inp.type !== 'hidden';
            });
        for (let i = 0; i < inps.length; i++) {
            inps[i].value = val;
            inps[i].dispatchEvent(new Event('input', { bubbles: true }));
            inps[i].dispatchEvent(new Event('change', { bubbles: true }));
            inps[i].dispatchEvent(new Event('blur', { bubbles: true }));
        }
    """, str(survey_value))

    inputs = driver.find_elements(By.XPATH, "//form//div[4]//input[not(contains(@placeholder, 'DD/MM')) and not(@type='radio') and not(@type='checkbox')]")
    if inputs:
        for inp in inputs:
            try:
                if inp.is_displayed() and inp.get_attribute("value") != str(survey_value):
                    driver.execute_script(
                        "arguments[0].value = arguments[1];"
                        "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));"
                        "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                        inp,
                        str(survey_value),
                    )
            except Exception:
                pass

    # 2. Ensure date fields are populated
    dates_ok = driver.execute_script("""
        const inps = Array.from(document.querySelectorAll("form input")).filter(inp => {
            const ph = (inp.placeholder || '').toLowerCase();
            const name = (inp.getAttribute('formcontrolname') || '').toLowerCase();
            return ph.includes('dd/mm/yyyy') || ph.includes('date') || name.includes('date');
        });
        return inps.length >= 2 && inps[0].value.trim() !== '' && inps[1].value.trim() !== '';
    """)
    if not dates_ok:
        now = datetime.now()
        select_date(driver, 1, "JAN", 2020, "first")
        select_date(driver, now.day, now.strftime("%b").upper(), now.year, "second")

    # 3. Solve Search CAPTCHA
    max_attempts = 15
    attempt = 0
    captcha2_success = False
    no_records_confirmed = False

    while not captcha2_success and attempt < max_attempts:
        attempt += 1
        try:
            wait_for_loader(driver)

            # Find Captcha image (Kaveri uses alt='Loaded Image' for both login and search)
            captcha2_element = None
            img_candidates = [
                "//app-ec-search-citizen//img[@alt='Loaded Image']",
                "//form//img[@alt='Loaded Image']",
                "//form//img[contains(@src, 'data:image') and not(contains(@src, 'logo')) and not(contains(@src, 'emblem')) and not(contains(@src, 'icon'))]",
                "//form//div[contains(@class, 'captcha') or contains(@class, 'Captcha')]//img",
                "//app-ec-search-citizen//img[contains(@src, 'data:image')]",
                "//img[@alt='Loaded Image']",
                "/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[9]/div[1]/img"
            ]
            for xpath in img_candidates:
                try:
                    elems = driver.find_elements(By.XPATH, xpath)
                    for el in elems:
                        if el.is_displayed():
                            src = el.get_attribute("src") or ""
                            if src.startswith("data:image") or el.size.get('width', 0) > 40:
                                captcha2_element = el
                                break
                    if captcha2_element:
                        break
                except Exception:
                    pass

            if not captcha2_element:
                img_list = driver.find_elements(By.XPATH, "//form//img")
                for img in img_list:
                    if img.is_displayed() and img.size.get('width', 0) > 40:
                        captcha2_element = img
                        break

            if not captcha2_element:
                raise NoSuchElementException("Search captcha image not found")

            # Find Captcha input
            captcha2_input = None
            input_candidates = [
                "//form//input[@id='loginCaptchaUserInput1' or @id='captcha' or @id='captchaInput' or @id='searchCaptcha']",
                "//form//input[contains(@placeholder, 'Captcha') or contains(@placeholder, 'CAPTCHA') or contains(@placeholder, 'captcha')]",
                "//form//input[contains(@formcontrolname, 'captcha') or contains(@formcontrolname, 'Captcha')]",
                "//app-ec-search-citizen//img[@alt='Loaded Image']/ancestor::div[contains(@class, 'row') or contains(@class, 'form')]//input",
                "//app-ec-search-citizen//img[@alt='Loaded Image']/following::input[1]",
                "(//form//input[not(@type='radio') and not(@type='checkbox') and not(@type='hidden')])[last()]",
                "/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[9]/div[2]/input"
            ]
            for xpath in input_candidates:
                try:
                    elems = driver.find_elements(By.XPATH, xpath)
                    for el in elems:
                        if el.is_displayed() and el.is_enabled():
                            captcha2_input = el
                            break
                    if captcha2_input:
                        break
                except Exception:
                    pass

            if not captcha2_input:
                raise NoSuchElementException("Search captcha input not found")

            captcha2_input.click()

            if attempt > 1:
                refresh_search_captcha(driver, captcha2_element)
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", captcha2_element)
                except Exception:
                    pass

            try:
                captcha2_bytes = get_captcha_bytes(driver, captcha2_element)
                captcha2_text = keep_black_text_only(captcha2_bytes, label=label)
            except Exception:
                captcha2_image_path = os.path.abspath(f"captcha2_{label}.png")
                captcha2_element.screenshot(captcha2_image_path)
                with open(captcha2_image_path, "rb") as f:
                    captcha2_bytes = f.read()
                captcha2_text = keep_black_text_only(captcha2_image_path, label=label)

            if not captcha2_text or len(captcha2_text) != 6:
                log(label, f"⚠️ CAPTCHA OCR failed for Survey #{survey_value} (attempt {attempt}/{max_attempts}); refreshing CAPTCHA and retrying.")
                save_captcha_sample(captcha2_bytes, captcha2_text, is_success=False, context="search_ocr_fail", label=label)
                refresh_search_captcha(driver, captcha2_element)
                time.sleep(0.05)
                continue

            captcha2_input.clear()
            captcha2_input.send_keys(captcha2_text)
            driver.execute_script(
                "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));"
                "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                captcha2_input
            )

            entered_captcha = captcha2_input.get_attribute("value") or ""
            if entered_captcha.strip().upper() != captcha2_text:
                driver.execute_script(
                    "arguments[0].value = arguments[1];"
                    "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));"
                    "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                    captcha2_input,
                    captcha2_text,
                )

            # Click Search button
            search_btn = None
            btn_candidates = [
                "//button[normalize-space()='Search' or normalize-space(text())='Search' or contains(., 'Search')]",
                "//form//button[contains(@class, 'search') or contains(@class, 'primary')][not(contains(., 'Reset')) and not(contains(., 'Clear')) and not(contains(., 'Remove')) and not(contains(., 'Add More'))]",
                "//button[@type='submit']",
                "(//form//button)[last()]"
            ]
            for xpath in btn_candidates:
                try:
                    elems = driver.find_elements(By.XPATH, xpath)
                    for el in elems:
                        if el.is_displayed() and el.is_enabled():
                            search_btn = el
                            break
                    if search_btn:
                        break
                except Exception:
                    pass

            if not search_btn:
                try:
                    search_btn = driver.execute_script("""
                        const btns = Array.from(document.querySelectorAll("form button, button"));
                        return btns.find(b => b.innerText.trim().toLowerCase() === 'search' && b.offsetParent !== null) || null;
                    """)
                except Exception:
                    pass

            if not search_btn:
                raise NoSuchElementException("Search button not found on page")

            try:
                results_container = driver.find_element(By.ID, "PdfData")
                driver.execute_script("arguments[0].innerHTML = '';", results_container)
            except NoSuchElementException:
                pass

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", search_btn)
            time.sleep(0.05)
            try:
                search_btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", search_btn)
            time.sleep(0.05)
            wait_for_loader(driver, timeout=25)

            # Check browser alert
            is_alert_no_records = False
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                alert.accept()
                if any(w in alert_text.lower() for w in ["captcha", "invalid", "incorrect"]):
                    log(label, f"⚠️ CAPTCHA rejected by browser alert for Survey #{survey_value}: {alert_text}")
                    save_captcha_sample(captcha2_bytes, captcha2_text, is_success=False, context="search_alert", label=label)
                    refresh_search_captcha(driver, captcha2_element)
                    time.sleep(0.2)
                    continue
                elif any(w in alert_text.lower() for w in ["no record", "no data"]):
                    is_alert_no_records = True
            except (NoAlertPresentException, TimeoutException):
                pass

            if is_alert_no_records:
                captcha2_success = True
                save_captcha_sample(captcha2_bytes, captcha2_text, is_success=True, context="search_alert_no_rec", label=label)
                no_records_confirmed = True
                break

            # Check modal error or 'No Records' dialog
            is_captcha_error = False
            is_modal_no_records = False
            try:
                modals = driver.find_elements(By.XPATH, "//mat-dialog-container | //div[contains(@class, 'swal2-modal')] | //div[contains(@class, 'modal-content')] | //div[@role='dialog' or @role='alertdialog']")
                for modal in modals:
                    if modal.is_displayed():
                        modal_text = modal.text.strip().lower()
                        if any(w in modal_text for w in ["invalid captcha", "captcha is not valid", "incorrect captcha", "enter valid captcha", "captcha does not match"]):
                            is_captcha_error = True
                            save_captcha_sample(captcha2_bytes, captcha2_text, is_success=False, context="search_modal_err", label=label)
                            try:
                                close_btn = modal.find_element(By.XPATH, ".//button | .//span[contains(@class, 'close')]")
                                close_btn.click()
                            except Exception:
                                pass
                            break
                        elif any(w in modal_text for w in ["no record", "no data", "record not found"]):
                            is_modal_no_records = True
                            try:
                                close_btn = modal.find_element(By.XPATH, ".//button | .//span[contains(@class, 'close')]")
                                close_btn.click()
                            except Exception:
                                pass
                            break
            except Exception:
                pass

            if is_captcha_error:
                log(label, f"⚠️ CAPTCHA rejected for Survey #{survey_value} (attempt {attempt}/{max_attempts}); refreshing and retrying.")
                refresh_search_captcha(driver, captcha2_element)
                time.sleep(0.2)
                continue

            if is_modal_no_records:
                captcha2_success = True
                save_captcha_sample(captcha2_bytes, captcha2_text, is_success=True, context="search_modal_no_rec", label=label)
                no_records_confirmed = True
                break

            try:
                inline_errors = driver.find_elements(
                    By.XPATH,
                    "//div[contains(@class, 'error') or contains(@class, 'danger')]"
                    " | //span[contains(@class, 'error') or contains(@class, 'danger')]"
                    " | //mat-error"
                )
                for error_element in inline_errors:
                    error_text = error_element.text.strip().lower()
                    if error_element.is_displayed() and any(
                        word in error_text for word in ["captcha", "invalid", "incorrect"]
                    ):
                        log(label, f"⚠️ CAPTCHA inline error for Survey #{survey_value}: {error_element.text.strip()}")
                        save_captcha_sample(captcha2_bytes, captcha2_text, is_success=False, context="search_inline_err", label=label)
                        refresh_search_captcha(driver, captcha2_element)
                        is_captcha_error = True
                        break
            except Exception:
                pass

            if is_captcha_error:
                time.sleep(0.2)
                continue

            captcha2_success = True
            save_captcha_sample(captcha2_bytes, captcha2_text, is_success=True, context="search", label=label)
            break

        except (TimeoutException, WebDriverException, NoSuchElementException) as e:
            log(label, f"⚠️ Search CAPTCHA attempt {attempt}/{max_attempts} failed for Survey #{survey_value}: {e}; retrying.")
            try:
                if captcha2_element and captcha2_element.is_displayed():
                    refresh_search_captcha(driver, captcha2_element)
                else:
                    refreshes = driver.find_elements(By.XPATH, "//form//button[.//mat-icon[text()='refresh'] or contains(@class, 'refresh')]")
                    if refreshes:
                        driver.execute_script("arguments[0].click();", refreshes[0])
            except Exception:
                pass
            time.sleep(0.3)
            continue

    if not captcha2_success:
        raise TimeoutException(f"Search CAPTCHA not solved after {max_attempts} attempts for Survey #{survey_value}")

    if no_records_confirmed:
        close_result_modal(driver)
        return True, None

    # Wait for table HTML in #PdfData
    wait_for_loader(driver, timeout=25)
    table_html = None
    end_table_wait = time.time() + 25
    while time.time() < end_table_wait:
        try:
            table_html = driver.execute_script("""
                const tables = document.querySelectorAll("#PdfData table, div[id*='PdfData'] table");
                for (let tbl of tables) {
                    if (tbl.offsetHeight > 0 && tbl.querySelectorAll("tr").length > 1) {
                        return tbl.outerHTML;
                    }
                }
                return null;
            """)
            if table_html:
                break
        except Exception:
            pass
        time.sleep(0.1)

    close_result_modal(driver)
    wait_for_page_ready(driver)
    return True, table_html


# ======================================================================================
# MAIN ACCOUNT ROTATION & SCRAPING WORKFLOW
# ======================================================================================
def run_rotation_scraper(
    target_village_name: Optional[str] = None,
    target_village_option_index: Optional[int] = None,
    target_village_only: Optional[bool] = None,
    target_survey_start: Optional[int] = None,
    target_account_index: Optional[int] = None,
    list_villages_only: Optional[bool] = None,
):
    accounts = load_accounts()
    if not accounts:
        log("System", "❌ No accounts configured. Exiting.")
        return

    # 1. Parse CLI arguments if provided on command line
    cli_args = {}
    try:
        import argparse
        parser = argparse.ArgumentParser(description="Karnataka IGR Kaveri EC Scraper (Single Browser)")
        parser.add_argument("--village", "-v", type=str, default=None, help="Target village name to jump to (e.g. 'Aravantike Pura')")
        parser.add_argument("--village-index", "-vi", type=int, default=None, help="Target village option index (1-based, e.g. 1)")
        parser.add_argument("--village-only", action="store_true", default=None, help="Scrape only the target village and exit")
        parser.add_argument("--survey", "-s", type=int, default=None, help="Starting survey number (e.g. 1, 50)")
        parser.add_argument("--account", "-a", type=int, default=None, help="Starting account number (1-based, e.g. 1 for Account-1)")
        parser.add_argument("--list-villages", action="store_true", default=None, help="List all villages in targeted hobli and exit")
        parser.add_argument("--hobli-index", "-hi", type=int, default=None, help="Target hobli option index fallback (e.g. 3)")
        args, _ = parser.parse_known_args()
        if args.village is not None:
            cli_args["village"] = args.village
        if args.village_index is not None:
            cli_args["village_index"] = args.village_index
        if args.village_only is True:
            cli_args["village_only"] = True
        if args.survey is not None:
            cli_args["survey"] = args.survey
        if args.account is not None:
            cli_args["account"] = args.account
        if args.list_villages is True:
            cli_args["list_villages"] = True
        if args.hobli_index is not None:
            cli_args["hobli_index"] = args.hobli_index
    except Exception:
        pass

    # Resolve active configurations: function param > CLI arg > Global config
    active_village_name = (
        target_village_name
        if target_village_name is not None
        else cli_args.get("village", TARGET_VILLAGE_NAME)
    )
    active_village_option_index = (
        target_village_option_index
        if target_village_option_index is not None
        else cli_args.get("village_index", TARGET_VILLAGE_OPTION_INDEX)
    )
    active_village_only = (
        target_village_only
        if target_village_only is not None
        else cli_args.get("village_only", TARGET_VILLAGE_ONLY)
    )
    active_survey_start = (
        target_survey_start
        if target_survey_start is not None
        else cli_args.get("survey", TARGET_SURVEY_START)
    )
    active_account_index = (
        target_account_index
        if target_account_index is not None
        else cli_args.get("account", TARGET_ACCOUNT_INDEX)
    )
    active_list_villages = (
        list_villages_only
        if list_villages_only is not None
        else cli_args.get("list_villages", LIST_VILLAGES_ONLY)
    )
    active_hobli_opt = cli_args.get("hobli_index", TARGET_HOBLI_OPTION_INDEX)

    progress = load_rotation_progress()
    if active_account_index is not None:
        account_idx = (int(active_account_index) - 1) % len(accounts)
        searches_used = 0
        log("AccountManager", f"📍 Starting account forced to #{active_account_index} ({accounts[account_idx]['label']}) via configuration/CLI")
    else:
        account_idx = progress["account_index"] % len(accounts)
        searches_used = progress["searches_used_in_account"]
        if searches_used >= ACCOUNT_SEARCH_LIMIT:
            account_idx = (account_idx + 1) % len(accounts)
            searches_used = 0

    hobli_idx = progress["hobli_index"]
    village_idx = progress["village_index"]
    current_survey = progress["property_value"]
    current_batch_idx = progress["batch_index"]
    consecutive_empty = progress["consecutive_empty"]
    survey_had_records = progress["current_survey_had_records"]

    if active_village_name or active_village_option_index is not None:
        target_repr = active_village_name if active_village_name else f"Index {active_village_option_index}"
        mode_str = "ONLY this village" if active_village_only else "jump directly to this village and continue"
        log("System", f"🎯 Target Village configured: '{target_repr}' ({mode_str}).")
    else:
        log("System", f"🚀 Resuming Scraper: Account [{accounts[account_idx]['login_email']}] ({searches_used}/{ACCOUNT_SEARCH_LIMIT} searches used) | Survey #{current_survey} (Batch {current_batch_idx+1})")

    driver = build_browser_driver(label="Main")

    current_account = accounts[account_idx]
    if not perform_login(driver, current_account, label=current_account["label"]):
        log("System", f"❌ Initial login failed for account {current_account['login_email']}. Trying next account...")
        account_idx = (account_idx + 1) % len(accounts)
        searches_used = 0
        current_account = accounts[account_idx]
        if not is_driver_alive(driver):
            try:
                driver.quit()
            except Exception:
                pass
            driver = build_browser_driver(label="Main")
        if not perform_login(driver, current_account, label=current_account["label"]):
            log("System", "❌ Subsequent login failed. Terminating.")
            driver.quit()
            return

    navigate_to_search_page(driver, label=current_account["label"])

    hobli_options = get_dropdown_options(driver, HOBLI_XPATH, timeout=15)
    if not hobli_options:
        log("System", "⚠️ Hoblis not loaded yet. Re-verifying District and Taluka selection...")
        select_district_and_taluka(
            driver,
            target_district_name=TARGET_DISTRICT_NAME,
            target_district_index=TARGET_DISTRICT_INDEX,
            target_taluka_name=TARGET_TALUKA_NAME,
            target_taluka_index=TARGET_TALUKA_INDEX,
            label=current_account["label"]
        )
        hobli_options = get_dropdown_options(driver, HOBLI_XPATH, timeout=15)

    if not hobli_options:
        log("System", "⛔ No Hoblis found. Exiting.")
        driver.quit()
        return

    district_name = get_selected_dropdown_text(driver, DISTRICT_XPATH) or TARGET_DISTRICT_NAME
    taluka_name = get_selected_dropdown_text(driver, TALUKA_XPATH) or TARGET_TALUKA_NAME
    log("System", f"📍 Location Hierarchy: District='{district_name}', Taluka='{taluka_name}'")

    target_opt = active_hobli_opt
    target_hobli_idx = (target_opt - 2) if (target_opt and target_opt >= 2) else 0
    assigned_hoblis = [target_hobli_idx] if (0 <= target_hobli_idx < len(hobli_options)) else list(range(len(hobli_options)))

    if target_opt and 0 <= target_hobli_idx < len(hobli_options):
        log("System", f"🎯 Targeted Hobli option[{target_opt}]: '{hobli_options[target_hobli_idx][1]}' (value: {hobli_options[target_hobli_idx][0]}).")
    else:
        log("System", f"🧩 Iterating over all {len(assigned_hoblis)} Hoblis.")

    for h_idx in assigned_hoblis:
        hobli_val, hobli_name = hobli_options[h_idx]
        select_dropdown_value(driver, HOBLI_XPATH, hobli_val)

        village_options = get_dropdown_options(driver, VILLAGE_XPATH, timeout=15, dedup_by_name=True)
        if not village_options:
            continue

        if active_list_villages:
            log("System", f"\n📜 VILLAGES LIST FOR HOBLI '{hobli_name}' ({len(village_options)} total):")
            for idx, (val, name) in enumerate(village_options):
                log("System", f"   [{idx + 1}] '{name}' (Dropdown Value: {val})")
            log("System", "ℹ️ LIST_VILLAGES_ONLY is enabled. Exiting after printing villages.")
            driver.quit()
            return

        target_v_idx = None
        has_target_village = bool(active_village_name or active_village_option_index is not None)

        if active_village_name:
            clean_target = str(active_village_name).strip().lower()
            # 1. Exact match
            for idx, (_, v_name) in enumerate(village_options):
                if v_name.strip().lower() == clean_target:
                    target_v_idx = idx
                    break
            # 2. Case-insensitive substring match
            if target_v_idx is None:
                for idx, (_, v_name) in enumerate(village_options):
                    if clean_target in v_name.strip().lower():
                        target_v_idx = idx
                        break

        if target_v_idx is None and active_village_option_index is not None:
            try:
                t_opt = int(active_village_option_index)
                if 1 <= t_opt <= len(village_options):
                    target_v_idx = t_opt - 1
                elif 0 <= t_opt < len(village_options):
                    target_v_idx = t_opt
            except (ValueError, TypeError):
                pass

        if target_v_idx is not None:
            target_val, target_name = village_options[target_v_idx]
            if active_village_only:
                assigned_villages = [target_v_idx]
                log("System", f"🎯 Targeted Single Village: '{target_name}' (Index: {target_v_idx + 1}/{len(village_options)}, Dropdown Value: {target_val}).")
            else:
                assigned_villages = list(range(target_v_idx, len(village_options)))
                log("System", f"🎯 Jumping directly to Target Village: '{target_name}' (Index: {target_v_idx + 1}/{len(village_options)}, Dropdown Value: {target_val}) and continuing with {len(assigned_villages)} villages.")
        else:
            if has_target_village:
                sought = active_village_name or f"Index {active_village_option_index}"
                log("System", f"⚠️ Target village '{sought}' not found in Hobli '{hobli_name}'.")
                sample_villages = [name for _, name in village_options[:10]]
                log("System", f"📋 First {len(sample_villages)} available villages in '{hobli_name}': {', '.join(sample_villages)}...")
                if active_village_only:
                    log("System", "⏩ active_village_only is True and target village was not found. Skipping hobli.")
                    continue
            start_v_idx = village_idx if h_idx == hobli_idx else 0
            assigned_villages = list(range(start_v_idx, len(village_options)))
            log("System", f"🧩 Iterating over {len(assigned_villages)} villages (starting at index {start_v_idx + 1}/{len(village_options)}).")

        for v_idx in assigned_villages:
            village_val, village_name = village_options[v_idx]
            v_key = village_name.strip()

            summary = load_village_summary()
            v_rec = summary.get(v_key, {})
            if SKIP_COMPLETED_VILLAGES and v_rec.get("status") == "COMPLETED":
                if has_target_village and v_idx == target_v_idx:
                    log("System", f"⏩ Target Village '{village_name}' is already marked COMPLETED in summary. Set SKIP_COMPLETED_VILLAGES = False to re-scrape. Skipping.")
                else:
                    log("System", f"⏩ Village '{village_name}' is already completed. Skipping.")
                continue

            log("System", f"\n🏘️ {'='*60}\n🏘️ STARTING VILLAGE: {village_name} ({v_idx+1}/{len(village_options)})\n🏘️ {'='*60}")

            select_dropdown_value(driver, VILLAGE_XPATH, village_val)
            ensure_property_boundary_selected(driver)

            # Get and batch all property types (6 per batch)
            all_property_types = get_property_number_type_options(driver)
            batches = [all_property_types[i:i + PROPERTY_BATCH_SIZE] for i in range(0, len(all_property_types), PROPERTY_BATCH_SIZE)] if all_property_types else []
            if not batches:
                for retry_pt in range(3):
                    log("System", f"⚠️ Retrying property types detection for {village_name} (attempt {retry_pt+1}/3)...")
                    ensure_property_boundary_selected(driver)
                    time.sleep(1.5)
                    all_property_types = get_property_number_type_options(driver)
                    if all_property_types:
                        batches = [all_property_types[i:i + PROPERTY_BATCH_SIZE] for i in range(0, len(all_property_types), PROPERTY_BATCH_SIZE)]
                        break

            if not batches:
                log("System", f"❌ Could not load Property Number Types for {village_name}. Skipping.")
                continue

            log("System", f"📋 Found {len(all_property_types)} Property Number Types -> Split into {len(batches)} batches of up to {PROPERTY_BATCH_SIZE}.")

            save_property_types_reference(batches)

            # Determine starting survey and batch
            is_same_as_saved_progress = (
                h_idx == hobli_idx and
                (progress.get("village_name", "").strip().lower() == v_key.lower() or (v_idx == village_idx and not has_target_village))
            )

            if active_survey_start is not None and (v_idx == target_v_idx or not has_target_village):
                survey_num = int(active_survey_start)
                batch_start = 0
                consecutive_empty = 0
                survey_had_records = False
                log("System", f"📍 Starting survey overridden to #{survey_num} by TARGET_SURVEY_START")
            elif is_same_as_saved_progress:
                survey_num = current_survey
                batch_start = current_batch_idx if survey_num == current_survey else 0
                log("System", f"📍 Resuming saved progress for '{village_name}' at Survey #{survey_num} (Batch {batch_start+1})")
            else:
                last_searched = v_rec.get("last_searched_value", 0)
                if last_searched and int(last_searched) > 0:
                    survey_num = int(last_searched) + 1
                    batch_start = 0
                    consecutive_empty = 0
                    survey_had_records = False
                    log("System", f"📍 Resuming village '{village_name}' from Survey #{survey_num} based on previous summary")
                else:
                    survey_num = 1
                    batch_start = 0
                    consecutive_empty = 0
                    survey_had_records = False

            while True:
                log("System", f"\n🔍 [{village_name}] Survey #{survey_num} | Empty streak: {consecutive_empty}/{MAX_CONSECUTIVE_EMPTY}")

                for b_idx in range(batch_start, len(batches)):
                    if searches_used >= ACCOUNT_SEARCH_LIMIT:
                        log("AccountManager", f"🛑 Account [{current_account['login_email']}] reached search limit ({searches_used}/{ACCOUNT_SEARCH_LIMIT})!")
                        perform_logout(driver, label=current_account["label"])

                        account_idx += 1
                        logged_in = False
                        while account_idx < len(accounts) and not logged_in:
                            current_account = accounts[account_idx]
                            searches_used = 0
                            log("AccountManager", f"🔄 Attempting login with Account #{account_idx+1}: {current_account['login_email']}")
                            if not is_driver_alive(driver):
                                try:
                                    driver.quit()
                                except Exception:
                                    pass
                                driver = build_browser_driver(label="Main")

                            if perform_login(driver, current_account, label=current_account["label"]):
                                logged_in = True
                            else:
                                log("AccountManager", f"⚠️ Direct login failed for {current_account['login_email']}; attempting clean browser restart...")
                                try:
                                    driver.quit()
                                except Exception:
                                    pass
                                driver = build_browser_driver(label="Main")
                                if perform_login(driver, current_account, label=current_account["label"]):
                                    logged_in = True
                                else:
                                    log("AccountManager", f"❌ Failed to login with {current_account['login_email']}. Trying next account...")
                                    account_idx += 1

                        if not logged_in or account_idx >= len(accounts):
                            log("AccountManager", "🏁 All configured accounts have been utilized or failed login. Terminating.")
                            try:
                                driver.quit()
                            except Exception:
                                pass
                            return

                        navigate_to_search_page(driver, label=current_account["label"])
                        select_dropdown_value(driver, HOBLI_XPATH, hobli_val)
                        select_dropdown_value(driver, VILLAGE_XPATH, village_val)
                        ensure_property_boundary_selected(driver)

                    batch = batches[b_idx]
                    batch_types_str = ", ".join(txt for _, txt in batch)
                    log("Scraper", f"⚙️ Setting up Batch {b_idx+1}/{len(batches)} ({len(batch)} types: {batch_types_str}) for Survey #{survey_num} [Account Searches: {searches_used+1}/{ACCOUNT_SEARCH_LIMIT}]...")

                    setup_property_type_batch(driver, batch)

                    search_ok, table_html = execute_single_batch_search(driver, survey_num, b_idx + 1, label=current_account["label"])
                    searches_used += 1

                    new_saved = 0
                    if table_html:
                        valid_df = extract_valid_records_from_table(
                            table_html, survey_num, b_idx + 1, batch_types_str=batch_types_str,
                            district_name=district_name, taluka_name=taluka_name,
                            hobli_name=hobli_name, village_name=village_name
                        )
                        if valid_df is not None and not valid_df.empty:
                            new_saved = save_df_to_excel(valid_df, OUTPUT_FILE)
                            if new_saved > 0:
                                survey_had_records = True
                                log("Scraper", f"⭐ [{village_name}] Survey #{survey_num} (Batch {b_idx+1}) -> Saved {new_saved} NEW records to Excel!")
                                update_village_summary(village_name, hobli_name, taluka_name, district_name, searched_val=survey_num, scraped_val=survey_num, new_recs_count=new_saved, label=current_account["label"])

                    search_status = f"SAVED_{new_saved}_RECORDS" if new_saved > 0 else ("NO_NEW_RECORDS" if table_html else "NO_RECORDS")

                    log_search_history(
                        account_label=current_account["label"],
                        hobli_name=hobli_name,
                        village_name=village_name,
                        survey_number=survey_num,
                        batch_idx=b_idx + 1,
                        total_batches=len(batches),
                        batch_types=batch,
                        records_found=new_saved,
                        status=search_status
                    )

                    save_rotation_progress({
                        "account_index": account_idx,
                        "searches_used_in_account": searches_used,
                        "hobli_index": h_idx,
                        "village_index": v_idx,
                        "village_name": village_name,
                        "property_value": survey_num,
                        "batch_index": b_idx + 1 if (b_idx + 1 < len(batches)) else 0,
                        "consecutive_empty": consecutive_empty,
                        "current_survey_had_records": survey_had_records
                    })

                # Completed all batches for survey_num
                if survey_had_records:
                    consecutive_empty = 0
                    log("Scraper", f"✅ [{village_name}] Survey #{survey_num} complete -> Records found. Consecutive empty streak reset to 0.")
                else:
                    consecutive_empty += 1
                    log("Scraper", f"⏩ [{village_name}] Survey #{survey_num} complete -> 0 records across all {len(batches)} batches. Empty streak: {consecutive_empty}/{MAX_CONSECUTIVE_EMPTY}")

                update_village_summary(village_name, hobli_name, taluka_name, district_name, searched_val=survey_num, scraped_val=survey_num if survey_had_records else None, label=current_account["label"])

                # Advance to next survey number
                survey_num += 1
                batch_start = 0
                survey_had_records = False

                save_rotation_progress({
                    "account_index": account_idx,
                    "searches_used_in_account": searches_used,
                    "hobli_index": h_idx,
                    "village_index": v_idx,
                    "village_name": village_name,
                    "property_value": survey_num,
                    "batch_index": 0,
                    "consecutive_empty": consecutive_empty,
                    "current_survey_had_records": False
                })

                if consecutive_empty >= MAX_CONSECUTIVE_EMPTY:
                    log("System", f"🎉 [{village_name}] VILLAGE COMPLETED! Reached {MAX_CONSECUTIVE_EMPTY} consecutive empty surveys.")
                    update_village_summary(village_name, hobli_name, taluka_name, district_name, searched_val=survey_num, is_completed=True, label=current_account["label"])
                    consecutive_empty = 0
                    current_survey = 1
                    current_batch_idx = 0
                    break

    log("System", "🎉 All assigned villages completed successfully!")
    driver.quit()


if __name__ == "__main__":
    run_rotation_scraper()
