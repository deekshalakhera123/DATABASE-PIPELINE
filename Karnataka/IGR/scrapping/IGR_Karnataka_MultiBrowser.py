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

import os
import io
import base64
import json
import re
import time
import queue
import itertools
import math
from collections import Counter
import imaplib
import email
from email import utils as email_utils
from datetime import datetime, timedelta
import multiprocessing as mp
import subprocess
import tempfile
import glob
from typing import Optional

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
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.firefox.service import Service as FirefoxService

from deep_translator import GoogleTranslator
from PIL import Image, ImageOps
import pytesseract
import cv2
import numpy as np
ddddocr = None
if sys.version_info < (3, 13):
    try:
        import ddddocr
    except Exception:
        ddddocr = None
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook


# ======================================================================================
# GLOBAL DEFAULT CONFIGURATION
# ======================================================================================
DEFAULT_EMAIL = "surajkumar@sigmavalue.co.in"
DEFAULT_EMAIL_PASSWORD = "gkze lcpy jduk vfvo"
IMAP_SERVER = "imap.gmail.com"
CANDIDATE_TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
    os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]
TESSERACT_PATH = next((p for p in CANDIDATE_TESSERACT_PATHS if os.path.exists(p)), r"C:\Program Files\Tesseract-OCR\tesseract.exe")
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "Anekall Kasba Hoblli2.xlsx")
VILLAGE_SUMMARY_JSON = os.path.join(BASE_DIR, "village_scraping_summary.json")
VILLAGE_SUMMARY_XLSX = os.path.join(BASE_DIR, "Village_Scraping_Summary.xlsx")
SAVE_DEBUG_SCREENSHOTS = False
MAX_CONSECUTIVE_EMPTY = 50
SKIP_COMPLETED_VILLAGES = True

# --------------------------------------------------------------------------------------
# TARGET LOCATION HIERARCHY CONFIGURATION (USER SELECTION)
# --------------------------------------------------------------------------------------
TARGET_DISTRICT_NAME = "Basavanagudi"   # Registration District (Basavanagudi / Bengaluru Urban)
TARGET_DISTRICT_INDEX = 3                # Dropdown option index fallback (3 in Kaveri)
TARGET_TALUKA_NAME = "Anekal"            # Taluka (Anekal)
TARGET_TALUKA_INDEX = 1                  # Dropdown option index fallback (1 in Kaveri)

# TARGET HOBLI BY OPTION[N] INDEX
# In the Kaveri Hobli dropdown (div[1]/div[3]):
#   2 -> option[2] (1st Hobli, e.g. Anekal Kasaba Hobli)
#   3 -> option[3] (2nd Hobli, e.g. Attibele Hobli 1)
#   4 -> option[4] (3rd Hobli, e.g. Attibele Hobli 2)
# Set to None to scrape all hoblis sequentially from resume index.
TARGET_HOBLI_OPTION_INDEX = 3

DISTRICT_XPATH = "//app-ec-search-citizen//form//div[1]/div[1]//select | (//form//select)[1]"
TALUKA_XPATH = "//app-ec-search-citizen//form//div[1]/div[2]//select | (//form//select)[2]"
HOBLI_XPATH = "//app-ec-search-citizen//form//div[1]/div[3]//select | (//form//select)[3]"
VILLAGE_XPATH = "//app-ec-search-citizen//form//div[1]/div[4]//select | (//form//select)[4]"


# ======================================================================================
# WORKER & MULTI-BROWSER CONFIGURATION
# ======================================================================================
# Configure each worker with its own browser engine, login account, and mailbox.
# Worker 1 uses Google Chrome; Worker 2 uses Microsoft Edge.
ACCOUNTS = [
    {
        "label": "Worker-1-Edge",
        "browser_type": "edge",                       # Microsoft Edge engine
        "login_email": "",
        "login_password": "",
        "excel_file": OUTPUT_FILE,                     # shared output file across workers
        "progress_file": "igr_last_scraped_1.txt",
        "imap_email": "",     # mailbox to fetch OTP from
        "imap_app_password": ""    # Gmail app password (not account password)
    },
    {
        "label": "Worker-2-Chrome",
        "browser_type": "chrome",                     # Google Chrome engine
        "login_email": "",
        "login_password": "",
        "excel_file": OUTPUT_FILE,                     # shared output file across workers
        "progress_file": "igr_last_scraped_2.txt",
        "imap_email": "",
        "imap_app_password": "",
    },
    {
        "label": "Worker-3-firefox",
        "browser_type": "firefox",                     # Firefox engine
        "login_email": "",
        "login_password": "",
        "excel_file": OUTPUT_FILE,                     # shared output file across workers
        "progress_file": "igr_last_scraped_3.txt",
        "imap_email": "",
        "imap_app_password": "",
    },
]

# Stagger delay between worker launches (seconds) to prevent OTP email collisions
STAGGER_DELAY_SECONDS = 25


# ANSI Color & Style Palettes for Clean Multi-Process Terminal UX
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

WORKER_COLORS = {
    "Worker-1-Edge": COLOR_CYAN,
    "Worker-2-Chrome": COLOR_GREEN,
    "Worker-3-firefox": COLOR_MAGENTA,
    "Worker-3-Firefox": COLOR_MAGENTA,
    "ExcelSaver": COLOR_YELLOW,
    "System": COLOR_WHITE,
    "Main": COLOR_WHITE,
}


def log(label, message):
    """Atomic, color-coded, timestamped, worker-labeled console print for clean multi-process logs."""
    now_str = datetime.now().strftime("%H:%M:%S")
    clean_label = f"[{label}]"
    color = WORKER_COLORS.get(label, COLOR_CYAN)
    formatted = f"{COLOR_GRAY}[{now_str}]{RESET} {color}{BOLD}{clean_label:<20}{RESET} {message}\n"
    try:
        sys.stdout.write(formatted)
        sys.stdout.flush()
    except Exception:
        print(f"[{now_str}] [{label}] {message}", flush=True)


def update_worker_status(shared_worker_status, label, engine=None, village=None, survey=None, empty_streak=None, records=None, state=None, detail=None):
    """Thread-safe and process-safe update of an active worker's real-time telemetry."""
    if shared_worker_status is None:
        return
    try:
        current = shared_worker_status.get(label, {})
        new_status = {
            "engine": str(engine if engine is not None else current.get("engine", "CHROME")),
            "village": str(village if village is not None else current.get("village", "-")),
            "survey": str(survey if survey is not None else current.get("survey", "-")),
            "empty_streak": str(empty_streak if empty_streak is not None else current.get("empty_streak", "-")),
            "records": int(records if records is not None else current.get("records", 0)),
            "state": str(state if state is not None else current.get("state", "RUNNING")),
            "detail": str(detail if detail is not None else current.get("detail", "")),
            "last_updated": time.time(),
        }
        shared_worker_status[label] = new_status
    except Exception:
        pass


def save_debug_screenshot(driver, file_path):
    if SAVE_DEBUG_SCREENSHOTS:
        driver.save_screenshot(file_path)


# Per-mailbox IMAP locks: workers sharing the same inbox are serialized;
# workers with their own mailbox run concurrently.
import threading as _threading
_imap_locks = {}
_imap_locks_guard = _threading.Lock()


def _get_imap_lock(imap_email):
    """Returns a lock specific to one mailbox so workers sharing the same
    inbox are serialized while workers with separate mailboxes run freely."""
    with _imap_locks_guard:
        if imap_email not in _imap_locks:
            _imap_locks[imap_email] = _threading.Lock()
        return _imap_locks[imap_email]


# ======================================================================================
# MULTI-BROWSER FACTORY WITH NATIVE ISOLATION & ANTI-AUTOMATION FLAGS
# ======================================================================================
def build_browser_driver(browser_type="chrome", label="Worker-1"):
    """
    Creates an isolated WebDriver instance.
    Selenium automatically allocates a fresh, collision-free temporary profile
    for each browser session, eliminating ProcessSingleton / WinError 32 lock crashes
    and ensuring 100% cookie and session isolation between parallel workers.
    """
    browser_type = browser_type.lower().strip()
    log(label, f"🌐 Initializing {browser_type.upper()} driver...")

    if browser_type == "chrome":
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
        options.add_argument("--disable-blink-features=AutomationControlled")
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

    elif browser_type == "edge":
        options = EdgeOptions()
        options.page_load_strategy = "eager"
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--start-maximized")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-animations")
        options.add_argument("--disable-renderer-backgrounding")
        options.add_argument("--disable-backgrounding-occluded-windows")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_experimental_option("detach", True)

        driver = webdriver.Edge(options=options)

        try:
            driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"}
            )
        except Exception:
            pass
        return driver

    elif browser_type == "firefox":
        options = FirefoxOptions()
        options.page_load_strategy = "eager"
        options.set_preference("dom.webdriver.enabled", False)
        options.set_preference("useAutomationExtension", False)
        options.set_preference("dom.webnotifications.enabled", False)

        firefox_binary = os.getenv("FIREFOX_BINARY")
        if firefox_binary:
            if not os.path.isfile(firefox_binary):
                raise FileNotFoundError(f"FIREFOX_BINARY does not exist: {firefox_binary}")
            options.binary_location = firefox_binary

        geckodriver_path = os.getenv("GECKODRIVER_PATH")
        if geckodriver_path:
            if not os.path.isfile(geckodriver_path):
                raise FileNotFoundError(f"GECKODRIVER_PATH does not exist: {geckodriver_path}")
            service = FirefoxService(executable_path=geckodriver_path)
        else:
            service = FirefoxService()

        driver = webdriver.Firefox(service=service, options=options)
        driver.maximize_window()
        return driver

    else:
        raise ValueError(f"Unsupported browser type: {browser_type}. Use 'chrome', 'edge', or 'firefox'.")



# ======================================================================================
# POPUPS & MODAL HELPERS
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
        popup = driver.find_element(By.CSS_SELECTOR, "div.popup-container.open")
        if popup:
            close_btn = popup.find_element(By.XPATH, ".//button | .//span[text()='X'] | .//span[contains(@class, 'close')] | .//a[contains(@class, 'close')]")
            close_btn.click()
            time.sleep(0.2)
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
            log(label, f"⚠️ Site failed to load '{url}' (attempt {attempt}/{max_retries}): {e}. Refreshing...")
            try:
                driver.refresh()
                time.sleep(wait_after)
            except Exception:
                pass
    log(label, f"❌ Failed to load '{url}' after {max_retries} attempts: {last_exc}")
    return False


def with_page_reload_retry(action_fn, driver, *args, max_retries=3, wait_after=5, label="", **kwargs):
    """Retries an action function upon timeouts by refreshing the page."""
    last_exc: Optional[BaseException] = None
    for attempt in range(1, max_retries + 1):
        try:
            return action_fn(driver, *args, **kwargs)
        except (TimeoutException, WebDriverException) as e:
            last_exc = e
            log(label, f"⚠️ '{getattr(action_fn, '__name__', 'action')}' failed (attempt {attempt}/{max_retries}): {e}")
            try:
                driver.refresh()
                wait_for_loader(driver)
                time.sleep(wait_after)
            except Exception:
                pass
        except Exception:
            raise
    if last_exc is not None:
        raise last_exc
    raise RuntimeError("Action failed without an exception")


# ======================================================================================
# CAPTCHA PROCESSING (OCR) - HIGH-SPEED HYBRID ENGINE
# ======================================================================================
CAPTCHA_SAMPLES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "captcha_samples")
CAPTCHA_SUCCESS_DIR = os.path.join(CAPTCHA_SAMPLES_DIR, "successful")
CAPTCHA_FAILED_DIR = os.path.join(CAPTCHA_SAMPLES_DIR, "failed")
os.makedirs(CAPTCHA_SUCCESS_DIR, exist_ok=True)
os.makedirs(CAPTCHA_FAILED_DIR, exist_ok=True)


def save_captcha_sample(img_bytes, predicted_text, is_success, context="search", label="Worker"):
    """
    Saves CAPTCHA image sample to 'captcha_samples/successful' or 'captcha_samples/failed'
    for ground-truth dataset curation and neural network / OCR fine-tuning.
    """
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
    """Returns a process-local instance of ddddocr for safe multiprocessing."""
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
    """
    Finds clusters of J-hook columns in the column-wise pixel count profile.
    J characters have a distinctive bottom hook producing taller columns
    (above_baseline > j_hook_thresh). Returns list of (start_col, end_col) tuples.
    """
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
    """
    Pass 5: Recovery engine for when ddddocr returns fewer than 6 characters.
    I and J are the narrowest characters in this captcha font and are frequently
    merged together or dropped entirely. Three strategies are used:

    Strategy A - First-char I/T confusion:
      When raw gives 6 chars and cl gives 5 chars where raw[1:]==cl, the raw
      result has a spurious first character (typically T instead of I).
      Confirm and correct it using x-stretched versions.

    Strategy B - Multi-stretch voting:
      Stretching horizontally forces narrow adjacent characters apart.
      We collect all 6-char results from multiple stretch factors and vote per position.

    Strategy C - J-hook pixel detection:
      When no stretch gives 6 chars, analyze the cleaned image column pixel profile.
      J characters leave a distinctive tall hook (>28px above baseline).
      Count J hooks in image vs J chars in result and insert missing Js (and trailing I).
    """
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

    # Strategy A: raw[1:] == cl -> first char of raw may be spurious (I->T confusion)
    if (len(raw_result) == 6 and len(cl_result) == 5
            and raw_result[1:] == cl_result):
        for sx in [1.2, 1.3, 1.5]:
            s = cv2.resize(im_orig, (int(w * sx), h), interpolation=cv2.INTER_LANCZOS4)
            _, buf = cv2.imencode(".png", s)
            r = re.sub(r"[^A-Z0-9]", "", ocr.classification(buf.tobytes()).strip().upper())
            if len(r) == 6 and r[1:] == cl_result:
                return r  # Corrected first char confirmed by stretch
        return raw_result  # Fall back to raw (first char suspect but best guess)

    # Strategy B: Multi-stretch voting
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
            # O/0 override: if cleaned result has '0' at this position -> prefer '0'
            if len(cl_result) > pos and cl_result[pos] == "0":
                if votes.get("O", 0) > votes.get("0", 0):
                    votes["0"] = votes["O"] + 1
            # I/H override at last position: if cl or any result says I -> prefer I over H
            if pos == 5 and votes.get("H", 0) > 0:
                if (cl_result and cl_result[-1] == "I") or any(
                    r[-1] == "I" for r in six_results
                ):
                    votes["I"] = votes.get("I", 0) + 2
            voted.append(votes.most_common(1)[0][0] if votes else "?")
        result = "".join(voted)

        # Cross-validate: if the stretch-voted result disagrees with cl_result on
        # more than 1 of the known positions, the stretch is scrambling the image.
        # In that case, fall through to Strategy C (pixel-based J-hook detection).
        n_cl = len(cl_result)
        if n_cl >= 3:
            mismatches = sum(
                1 for i in range(n_cl)
                if i < len(result) and result[i] != cl_result[i]
            )
            if mismatches > 1:
                # Stretch gave a scrambled result; skip it and use Strategy C
                six_results = []  # Clear so we fall through to Strategy C

        if six_results:  # Still valid after cross-validation
            # O/0 last-char fix
            if raw_result and result[-1] == "O" and raw_result[-1] == "0":
                result = result[:-1] + "0"
            return result

    # Strategy C: J-hook pixel detection fallback
    gray = cv2.cvtColor(cleaned_img, cv2.COLOR_BGR2GRAY)
    col_counts = (gray < 128).sum(axis=0)
    baseline = 10
    above = (np.array(col_counts) - baseline).clip(0).astype(float)

    j_clusters = _find_j_hook_clusters(above)
    n_j_img = len(j_clusters)
    n_j_cl = cl_result.count("J")
    correction = list(cl_result)

    # Case A: More J hooks in image than J chars in result -> duplicate J
    if n_j_img > n_j_cl > 0:
        j_idx = correction.index("J")
        correction.insert(j_idx + 1, "J")
    # Case B: J count matches but I bar precedes J -> IJ pair (I was merged into J)
    elif n_j_img == n_j_cl > 0:
        for jc_start, jc_end in j_clusters:
            check_start = max(0, jc_start - 18)
            pre_above = above[check_start:jc_start]
            max_pre = max(pre_above) if len(pre_above) > 0 else 0
            high_pre = sum(1 for v in pre_above if v >= 22)
            if max_pre >= 22 and max_pre <= 28 and high_pre >= 2:
                j_idx = correction.index("J")
                correction.insert(j_idx, "I")  # Insert I before J
                break
    # Case C: J exists in image but not in result -> insert J at correct position
    elif n_j_img > 0 and n_j_cl == 0:
        for jc_start, jc_end in j_clusters:
            jc_center = (jc_start + jc_end) / 2.0
            slot_6 = jc_center / w * 6.0  # 0..6 float position in 6-char grid
            insert_at = min(math.ceil(slot_6), len(correction))
            correction.insert(insert_at, "J")
            break

    # Fill remaining missing positions with 'I' (most commonly dropped narrow char)
    while len(correction) < 6:
        correction.append("I")

    result = "".join(correction[:6])
    if raw_result and result[-1] == "O" and raw_result[-1] == "0":
        result = result[:-1] + "0"
    return result


def keep_black_text_only(image_source, label="default"):
    """
    High-speed, high-accuracy hybrid CAPTCHA solver:
    - Strips colored strike-through lines using chromatic difference (eliminates dark green,
      purple, and cyan lines without leaving line artifacts).
    - Preserves true black/neutral text glyphs.
    - Correctly distinguishes letter 'O' from digit '0'.
    - Distinguishes '1' from 'I' (removes top strike-through line fusing across '1').
    - Multi-pass consensus: ddddocr (cleaned) -> ddddocr (raw) -> Tesseract (scaled cleaned)
      -> Tesseract (Otsu) -> Stretch voting + J-hook pixel recovery (Pass 5).
    Returns 6 alphanumeric uppercase characters.
    """
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

    # Preprocessing: Decode image and strip colored strike-through lines
    try:
        im = cv2.imdecode(np.frombuffer(img_bytes, np.uint8), cv2.IMREAD_COLOR)
        if im is not None:
            # Reject loading overlay / spinner images (<650 dark pixels)
            dark_mask = (im[:, :, 0] < 125) & (im[:, :, 1] < 125) & (im[:, :, 2] < 125)
            if np.sum(dark_mask) < 650:
                return ""

            # Colored noise lines have high chroma: max(RGB) - min(RGB) >= 20.
            # Black/gray text pixels have low chroma (chroma < 20) and low brightness (< 135).
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

    # Pass 1: ddddocr on cleaned line-stripped image
    if ocr is not None and cleaned_bytes is not None:
        try:
            raw_text = ocr.classification(cleaned_bytes)
            cand_cl = re.sub(r"[^A-Z0-9]", "", str(raw_text).strip().upper())
        except Exception:
            pass

    # Pass 2: ddddocr on raw image
    if ocr is not None:
        try:
            raw_text = ocr.classification(img_bytes)
            cand_raw = re.sub(r"[^A-Z0-9]", "", str(raw_text).strip().upper())
        except Exception:
            pass

    # Pass 3: Tesseract on scaled cleaned image
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

    # Candidate selection: prioritize cleaned ddddocr (free of strike-through lines)
    # IMPORTANT: Tesseract (Pass 3/4) is used ONLY as a last resort when ddddocr returns <4 chars.
    # When ddddocr returns 4-5 chars, it has identified most chars correctly and is more
    # char-accurate than Tesseract. In those cases, Pass 5 (stretch voting + J-hook detection)
    # is used to recover the missing narrow chars (I/J) instead of relying on Tesseract.
    best_ddd_len = max(len(cand_cl), len(cand_raw))

    chosen = ""
    if len(cand_cl) == 6:
        chosen = cand_cl
    elif len(cand_raw) == 6:
        chosen = cand_raw
    elif best_ddd_len < 4:
        # ddddocr is very unreliable here; trust Tesseract as fallback
        if len(cand_tess) == 6:
            chosen = cand_tess
        else:
            # Pass 4: Tesseract Otsu with border padding fallback
            try:
                if target_mat is not None:
                    gray = cv2.cvtColor(target_mat, cv2.COLOR_BGR2GRAY)
                    blur = cv2.GaussianBlur(gray, (3, 3), 0)
                    _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                    padded = cv2.copyMakeBorder(thresh, 15, 15, 15, 20, cv2.BORDER_CONSTANT, value=255)
                    scaled_pad = cv2.resize(
                        padded, None, fx=2.5, fy=2.5, interpolation=cv2.INTER_LANCZOS4
                    )
                    raw_text_pad = pytesseract.image_to_string(scaled_pad, config=config)
                    cand_pad = re.sub(r"[^A-Z0-9]", "", raw_text_pad.strip().upper())
                    if len(cand_pad) == 6:
                        chosen = cand_pad
            except Exception:
                pass
    # else: ddddocr gave 4-5 chars -> skip Tesseract, go to Pass 5 recovery

    # Character consensus and refinement (for 6-char candidates)
    if len(chosen) == 6:
        voted = list(chosen)
        for i in range(6):
            # 1 vs I: strike-through lines across top of '1' make raw image look like 'I'
            if len(cand_cl) == 6 and cand_cl[i] == "1" and voted[i] == "I":
                voted[i] = "1"

            # Letter L vs I refinement
            if voted[i] == "L":
                if len(cand_raw) == 6 and cand_raw[i] == "I":
                    voted[i] = "I"
                elif len(cand_tess) == 6 and cand_tess[i] == "I":
                    voted[i] = "I"

            # Digit 5 vs S refinement
            elif voted[i] == "5" and len(cand_raw) == 6 and cand_raw[i] == "S":
                voted[i] = "S"
            elif voted[i] == "S" and len(cand_raw) == 6 and cand_raw[i] == "5":
                pass

        chosen = "".join(voted)

        # First-char I/T confusion even in 6-char results:
        # If raw[1:] == cand_cl (5 chars) -> raw has spurious first char (T instead of I)
        if (len(cand_raw) == 6 and len(cand_cl) == 5
                and cand_raw[1:] == cand_cl
                and chosen[0] == cand_raw[0]
                and cleaned_img is not None):
            fixed = _recover_missing_narrow_chars(cand_cl, cand_raw, cleaned_img, img_bytes)
            if len(fixed) == 6:
                return fixed

        return chosen

    # Pass 5: Narrow-char recovery (I/J merger and drop detection)
    if cleaned_img is not None:
        base = cand_cl if len(cand_cl) >= len(cand_raw) else cand_raw
        if not base:
            base = cand_tess
        recovered = _recover_missing_narrow_chars(base, cand_raw, cleaned_img, img_bytes)
        if recovered:
            return recovered

    return chosen or cand_cl or cand_raw or cand_tess


def get_captcha_bytes(driver, captcha_element):
    """
    Extracts lossless CAPTCHA image bytes directly from browser memory:
    1. Checks for base64 data URI in the <img> 'src' attribute.
    2. Uses HTML5 Canvas draw & export for DOM-level pixel extraction.
    3. Falls back to screenshot_as_png if DOM extraction is unavailable.
    """
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


def refresh_search_captcha(driver, captcha_element, timeout=6):
    """Refreshes the CAPTCHA image and waits for the new image to render."""
    try:
        old_png = captcha_element.screenshot_as_png
        clicked = False

        # Try finding refresh button in the same container
        try:
            parent = captcha_element.find_element(By.XPATH, "..")
            buttons = parent.find_elements(By.XPATH, ".//*[self::button or self::a or self::i or self::span][not(self::img)]")
            for btn in buttons:
                if btn.is_displayed():
                    try:
                        driver.execute_script("arguments[0].click();", btn)
                        clicked = True
                        break
                    except Exception:
                        pass
        except Exception:
            pass

        if not clicked:
            try:
                driver.execute_script("arguments[0].click();", captcha_element)
            except Exception:
                pass

        # Ultra-fast polling: check every 0.03s instead of 0.1s
        end_time = time.time() + timeout
        while time.time() < end_time:
            time.sleep(0.03)
            try:
                if captcha_element.screenshot_as_png != old_png:
                    return True
            except Exception:
                pass
    except Exception as e:
        log("System", f"⚠️ Error refreshing CAPTCHA: {e}")
    return False


def page_down_before_captcha_screenshot(driver):
    """Move the viewport down before capturing a CAPTCHA during a retry."""
    try:
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.END)
        time.sleep(0.5)
    except Exception:
        pass


# ======================================================================================
# IMAP OTP FETCHING
# ======================================================================================
def fetch_otp(retries=12, delay=5, after_time=None, label="", imap_email=None, imap_password=None):
    imap_email = imap_email or DEFAULT_EMAIL
    imap_password = imap_password or DEFAULT_EMAIL_PASSWORD
    if after_time:
        after_time = after_time - timedelta(minutes=3)
    log(label, f"📧 Connecting to IMAP mailbox: {imap_email}")
    def try_fetch_once():
        # Serialize IMAP access per-mailbox: two workers sharing the SAME inbox could
        # otherwise race and pick up each other's OTP. Workers with their own separate
        # mailbox (different imap_email) don't block each other.
        with _get_imap_lock(imap_email):
            try:
                mail = imaplib.IMAP4_SSL(IMAP_SERVER)
                mail.login(imap_email, imap_password)
                mail.select("inbox")
                status, data = mail.search(None, '(FROM "dsrnotify@karnataka.gov.in")')
                mail_ids = data[0].split()
                if not mail_ids:
                    return None
                for i in reversed(mail_ids[-5:]):  # Check last 5 mails
                    status, msg_data = mail.fetch(i, "(RFC822)")
                    for part in msg_data:
                        if isinstance(part, tuple):
                            msg = email.message_from_bytes(part[1])
                            date_tuple = email_utils.parsedate_tz(msg['Date'])
                            if date_tuple is None:
                                continue
                            msg_time = datetime.fromtimestamp(email_utils.mktime_tz(date_tuple))

                            if after_time and msg_time < after_time:
                                continue  # Skip old OTPs

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
# LOGIN & RE-LOGIN FLOWS
# ======================================================================================
def launch_and_login(login_email, login_password, label="Worker-1", browser_type="chrome", imap_email=None, imap_password=None):
    driver = build_browser_driver(browser_type=browser_type, label=label)
    
    if not safe_get(driver, "https://kaveri.karnataka.gov.in/landing-page", label=label):
        driver.quit()
        return None
        
    log(label, "Launched browser and opened Kaveri landing page.")
    login_start_time = datetime.now()

    try:
        max_attempts = 10
        attempt = 0
        captcha_success = False
        consecutive_login_fails = 0  # Full reload every 3 consecutive login rejects

        while attempt < max_attempts and not captcha_success:
            attempt += 1
            log(label, f"🔐 CAPTCHA Attempt {attempt} of {max_attempts}")

            try:
                # 1. Open login modal only if not already open (avoids modal toggle flicker/blink)
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

                # Enter Username if not already filled
                username_input = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="userName"]'))
                )
                if not username_input.get_attribute("value"):
                    username_input.clear()
                    username_input.send_keys(login_email)

                # Enter Password if not already filled
                password_input = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))
                )
                if not password_input.get_attribute("value"):
                    password_input.clear()
                    password_input.send_keys(login_password)

                # Capture CAPTCHA Image with overlay guard
                wait_for_loader(driver, timeout=5)
                captcha_xpath = "//img[@alt='Loaded Image']"
                captcha_element = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, captcha_xpath))
                )

                if attempt > 1:
                    log(label, "🔄 Refreshing CAPTCHA image before retry...")
                    refresh_search_captcha(driver, captcha_element)
                    wait_for_loader(driver, timeout=5)

                # Lossless in-memory DOM / canvas extraction
                captcha_bytes = get_captcha_bytes(driver, captcha_element)
                captcha_text = keep_black_text_only(captcha_bytes, label=label)
                log(label, f"Extracted CAPTCHA: '{captcha_text}'")

                if not captcha_text or len(captcha_text) != 6:
                    log(label, f"⚠️ CAPTCHA OCR did not produce 6 characters ('{captcha_text}'). Retrying...")
                    save_captcha_sample(captcha_bytes, captcha_text, is_success=False, context="login_ocr_fail", label=label)
                    refresh_search_captcha(driver, captcha_element)
                    time.sleep(0.5)
                    continue

                # Input CAPTCHA
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

                # Click Log In Button
                time.sleep(0.5)
                driver.find_element(By.XPATH, "//button[text()='Log In']").click()

                # Fast polling (up to 4.5s) for Alert, OTP success, or "Captcha is not valid" failure
                login_verdict = None
                poll_start = time.time()
                poll_end = poll_start + 4.5
                while time.time() < poll_end:
                    # Check browser alert
                    try:
                        alert = driver.switch_to.alert
                        alert_text = alert.text
                        alert.accept()
                        log(label, f"Alert detected: {alert_text}")
                        if any(w in alert_text.lower() for w in ["multiple", "session"]):
                            log(label, "⚠️ Active session collision detected. Cleared session in portal.")
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

                    # Check OTP modal (Success)
                    try:
                        otp_el = driver.find_element(By.XPATH, '//*[@id="otp"]')
                        if otp_el.is_displayed():
                            login_verdict = "success"
                            break
                    except Exception:
                        pass

                    # Wait at least 1.0s before checking inline errors to prevent matching stale banners from prior attempts
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
                    log(label, "✅ CAPTCHA verified — OTP modal is now open.")
                    break
                elif login_verdict == "session_cleared":
                    log(label, "🔄 Active session cleared by Kaveri. Re-submitting credentials...")
                    time.sleep(1.0)
                    continue
                elif login_verdict == "captcha_invalid":
                    save_captcha_sample(captcha_bytes, captcha_text, is_success=False, context="login_invalid", label=label)
                    log(label, "⚠️ CAPTCHA invalid. Refreshing CAPTCHA in-place without page reload...")
                    consecutive_login_fails += 1
                    try:
                        stale_captcha = driver.find_element(By.XPATH, captcha_xpath)
                        refresh_search_captcha(driver, stale_captcha)
                    except Exception:
                        pass
                    time.sleep(0.5)
                    continue
                else:
                    # Final check for OTP field in case of network lag
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="otp"]'))
                        )
                        captcha_success = True
                        save_captcha_sample(captcha_bytes, captcha_text, is_success=True, context="login", label=label)
                        log(label, "✅ CAPTCHA verified — OTP modal is now open.")
                        break
                    except Exception:
                        save_captcha_sample(captcha_bytes, captcha_text, is_success=False, context="login_timeout", label=label)
                        log(label, "⚠️ CAPTCHA verification timed out. Retrying...")
                        consecutive_login_fails += 1
                        time.sleep(0.5)
                        continue

            except (TimeoutException, WebDriverException) as e:
                log(label, f"⚠️ CAPTCHA attempt failed (site load issue): {e}. Retrying...")
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
            driver.quit()
            return None

        otp_input = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="otp"]'))
        )
        otp_input.clear()
        otp_input.send_keys(otp)
        time.sleep(2)
        log(label, "Entered OTP.")

        # Click OTP submit button
        time.sleep(2)

        driver.find_element(By.XPATH, "//input[@id='otp']/ancestor::form//button[not(contains(text(), 'Resend'))]").click()

        # Wait for portal to fully redirect to the dashboard after OTP
        log(label, "⏳ Waiting for dashboard to load after OTP...")
        if not wait_for_dashboard(driver):
            raise TimeoutException("Dashboard did not load after OTP")
        time.sleep(3)
        log(label, "✅ Dashboard loaded. Login complete.")
        return driver  # ✅ THIS SHOULD BE THE LAST LINE INSIDE THE try BLOCK

    except Exception as e:
        log(label, f"❌ Error during login: {e}")
        save_debug_screenshot(driver, f"login_error_{label}.png")
        driver.quit()
        return None


def relogin(driver, login_email, login_password, label="Worker-1", imap_email=None, imap_password=None):
    log(label, "🔁 Re-logging in after session timeout...")
    relogin_start_time = datetime.now()

    try:
        captcha_element = None
        otp_modal_found = False
        for attempt in range(10):
            log(label, f"Re-login CAPTCHA Attempt {attempt + 1}")
            try:
                # Open login modal only if not already open (avoids modal toggle flicker/blink)
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
                    time.sleep(0.5)

                username_input = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="userName"]'))
                )
                if not username_input.get_attribute("value"):
                    username_input.clear()
                    username_input.send_keys(login_email)

                password_input = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))
                )
                if not password_input.get_attribute("value"):
                    password_input.clear()
                    password_input.send_keys(login_password)

                wait_for_loader(driver, timeout=5)
                captcha_xpath = "//img[@alt='Loaded Image']"
                captcha_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, captcha_xpath))
                )

                if attempt > 0:
                    refresh_search_captcha(driver, captcha_element)
                    wait_for_loader(driver, timeout=5)

                captcha_bytes = get_captcha_bytes(driver, captcha_element)
                captcha_text = keep_black_text_only(captcha_bytes, label=label)

                if not captcha_text or len(captcha_text) != 6:
                    save_captcha_sample(captcha_bytes, captcha_text, is_success=False, context="relogin_ocr_fail", label=label)
                    refresh_search_captcha(driver, captcha_element)
                    time.sleep(0.5)
                    continue

                captcha_input = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="loginCaptchaUserInput1"]'))
                )
                captcha_input.clear()
                captcha_input.send_keys(captcha_text)

                driver.find_element(By.XPATH, "//button[text()='Log In']").click()

                # Fast check for Alert, OTP, or invalid captcha
                relogin_verdict = None
                poll_start = time.time()
                poll_end = poll_start + 4.5
                while time.time() < poll_end:
                    try:
                        alert = driver.switch_to.alert
                        alert_text = alert.text
                        alert.accept()
                        if any(w in alert_text.lower() for w in ["multiple", "session"]):
                            time.sleep(1.0)
                            relogin_verdict = "session_cleared"
                            break
                        elif any(w in alert_text.lower() for w in ["captcha", "invalid"]):
                            relogin_verdict = "captcha_invalid"
                            break
                        else:
                            time.sleep(1.0)
                            break
                    except Exception:
                        pass

                    try:
                        otp_el = driver.find_element(By.XPATH, '//*[@id="otp"]')
                        if otp_el.is_displayed():
                            relogin_verdict = "success"
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
                                    relogin_verdict = "captcha_invalid"
                                    break
                            if relogin_verdict:
                                break
                        except Exception:
                            pass

                    time.sleep(0.15)

                if relogin_verdict == "success":
                    otp_modal_found = True
                    save_captcha_sample(captcha_bytes, captcha_text, is_success=True, context="relogin", label=label)
                    break
                elif relogin_verdict == "session_cleared":
                    log(label, "🔄 Active session cleared by Kaveri during re-login. Re-submitting...")
                    time.sleep(1.0)
                    continue
                elif relogin_verdict == "captcha_invalid":
                    save_captcha_sample(captcha_bytes, captcha_text, is_success=False, context="relogin_invalid", label=label)
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
                        otp_modal_found = True
                        save_captcha_sample(captcha_bytes, captcha_text, is_success=True, context="relogin", label=label)
                        break
                    except Exception:
                        save_captcha_sample(captcha_bytes, captcha_text, is_success=False, context="relogin_timeout", label=label)
                        time.sleep(0.5)
                        continue

            except (TimeoutException, WebDriverException):
                time.sleep(1)
                continue

        if not otp_modal_found:
            raise TimeoutException("OTP modal did not open during re-login")

        # OTP fetching and entering
        time.sleep(10)  # Wait for OTP email, matching launch_and_login
        otp = fetch_otp(after_time=relogin_start_time, label=label, imap_email=imap_email, imap_password=imap_password)

        if not otp:
            log(label, "❌ OTP fetch failed during re-login.")
            save_debug_screenshot(driver, f"otp_error_relogin_{label}.png")
            driver.quit()
            return None

        otp_input = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="otp"]'))
        )
        otp_input.clear()
        # Retry sending OTP up to 3 times
        max_input_attempts = 3
        for attempt in range(max_input_attempts):
            try:
                log(label, f"OTP input attempt {attempt + 1}")
                otp_input.click()  # Ensure field is focused
                otp_input.clear()
                otp_input.send_keys(otp)

                # Verify OTP was entered
                entered_value = otp_input.get_attribute("value")
                if entered_value == otp:
                    log(label, "✅ OTP entered successfully.")
                    break
                else:
                    log(label, f"⚠️ OTP input failed: expected {otp}, found {entered_value}")
                    time.sleep(1)  # Brief wait before retry
                    continue
            except Exception as e:
                log(label, f"⚠️ OTP input attempt {attempt + 1} failed: {e}")
                time.sleep(1)  # Brief wait before retry
                continue
        else:
            log(label, f"❌ Failed to enter OTP after {max_input_attempts} attempts: expected {otp}, found {otp_input.get_attribute('value')}")
            log(label, f"Field status: enabled={otp_input.is_enabled()}, displayed={otp_input.is_displayed()}")
            save_debug_screenshot(driver, f"otp_input_error_relogin_{label}.png")
            driver.quit()
            return None

        time.sleep(2)
        log(label, "Entered OTP.")

        # Click OTP submit button
        time.sleep(2)
        driver.find_element(By.XPATH, "//input[@id='otp']/ancestor::form//button[not(contains(text(), 'Resend'))]").click()

        # Wait for portal to redirect to dashboard after OTP
        log(label, "⏳ Waiting for dashboard to load after OTP...")
        if not wait_for_dashboard(driver):
            raise TimeoutException("Dashboard did not load after re-login OTP")
        time.sleep(3)
        log(label, "✅ Dashboard loaded. Re-login complete.")

        return driver  # Return driver on success, matching launch_and_login

    except Exception as e:
        log(label, f"❌ Re-login failed: {e}")
        save_debug_screenshot(driver, f"relogin_error_{label}.png")
        return None


# ======================================================================================
# FORM & SEARCH NAVIGATION
# ======================================================================================
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


def ensure_property_boundary_selected(driver, label=""):
    """Ensures the Property/Survey boundary radio button is selected on the EC form."""
    try:
        wait_for_loader(driver)
        candidates = [
            "/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[4]/div[1]/div[2]/label",
            "//form/div[4]/div[1]/div[2]/label",
            "//label[contains(text(), 'Property') or contains(text(), 'Survey')]",
        ]
        for xpath in candidates:
            elems = driver.find_elements(By.XPATH, xpath)
            for el in elems:
                if el.is_displayed():
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                        driver.execute_script("arguments[0].click();", el)
                        wait_for_loader(driver)
                        time.sleep(0.5)
                        return True
                    except Exception:
                        pass
    except Exception as e:
        if label:
            log(label, f"⚠️ Error ensuring property boundary radio selected: {e}")
    return False


def get_property_number_type_options(driver):
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
        calendar_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, date_picker_xpaths[date_type]))
        )
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

        # Ensure datepicker backdrop overlay is dismissed
        time.sleep(0.5)
        try:
            WebDriverWait(driver, 5).until(
                EC.invisibility_of_element_located((By.CLASS_NAME, 'mat-datepicker-content'))
            )
        except Exception:
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                time.sleep(0.5)
            except Exception:
                pass
        wait_for_loader(driver)
        time.sleep(0.5)
        return True
    except Exception as e:
        log("System", f"Error selecting {date_type} date: {str(e)}")
        return False


def navigate_to_search_page(driver, label="Worker-1"):
    log(label, "🧭 Navigating to EC Search Citizen page...")
    wait_for_loader(driver)
    wait_for_page_ready(driver)

    # Check if already on the EC Search Citizen page
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
                        break
                if dashboard_clicked:
                    break
            except Exception:
                pass

        wait_for_loader(driver)
        time.sleep(1)

        # 2. Click document / application type image
        app_type_clicked = False
        app_type_selectors = [
            "/html/body/div[1]/app-root/div/div/app-kaveri-dashboard/div/div[2]/div/app-application-type/div/div/div[2]/div[2]/div[2]/img",
            "//app-application-type//div[2]//div[2]//img",
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
                        break
                if app_type_clicked:
                    break
            except Exception:
                pass

        wait_for_loader(driver)
        time.sleep(1)

        # 3. Dismiss prerequisite modals (Modal 1 and Modal 2)
        for dialog_step in range(1, 4):
            wait_for_loader(driver)
            time.sleep(1)
            dialog_btns = driver.find_elements(
                By.XPATH,
                "//app-perquisite//button | //mat-dialog-container//button"
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
                        break
                    except Exception:
                        pass
            if not btn_clicked:
                break
            wait_for_loader(driver)
            time.sleep(1)

    # 4. Wait for EC Search page elements to appear
    wait_for_loader(driver)
    WebDriverWait(driver, 25).until(
        EC.presence_of_element_located((By.XPATH, DISTRICT_XPATH))
    )

    # 5. Select District (Bengaluru Urban)
    district_elem = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, DISTRICT_XPATH))
    )
    current_district = get_selected_dropdown_text(driver, DISTRICT_XPATH)
    is_district_valid = current_district and any(
        k in current_district.lower() for k in ["bengaluru", "bangalore", "ಬೆಂಗಳೂರು", "basavanagudi", "ಬಸವನಗುಡಿ"]
    )
    if not is_district_valid:
        end_time = time.time() + 15
        while time.time() < end_time:
            opts = Select(district_elem).options
            if len(opts) > 1:
                break
            time.sleep(0.5)

        select_obj = Select(district_elem)
        target_val = None
        for opt in select_obj.options:
            t = opt.text.strip().lower()
            if any(k in t for k in ["bengaluru urban", "bangalore urban", "ಬೆಂಗಳೂರು ನಗರ", "basavanagudi", "ಬಸವನಗುಡಿ"]):
                target_val = opt.get_attribute("value")
                break
        if not target_val:
            if len(select_obj.options) > 3:
                target_val = select_obj.options[3].get_attribute("value")
            elif len(select_obj.options) > 1:
                target_val = select_obj.options[1].get_attribute("value")

        if target_val:
            select_dropdown_value(driver, DISTRICT_XPATH, target_val)

    wait_for_loader(driver)
    time.sleep(1)

    # 6. Select Taluka (Anekal)
    taluka_elem = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, TALUKA_XPATH))
    )
    current_taluka = get_selected_dropdown_text(driver, TALUKA_XPATH)
    is_taluka_valid = current_taluka and any(
        k in current_taluka.lower() for k in ["anekal", "ಆನೇಕಲ್"]
    )
    # Check if Hobli options have already loaded
    hobli_opts = get_dropdown_options(driver, HOBLI_XPATH, timeout=2)

    # If taluka is not Anekal, or if Hobli options are not yet populated, select Anekal to trigger Hobli fetch
    if not is_taluka_valid or not hobli_opts:
        end_time = time.time() + 15
        while time.time() < end_time:
            opts = Select(taluka_elem).options
            if len(opts) > 1:
                break
            time.sleep(0.5)

        select_obj = Select(taluka_elem)
        target_val = None
        for opt in select_obj.options:
            t = opt.text.strip().lower()
            if "anekal" in t or "ಆನೇಕಲ್" in t:
                target_val = opt.get_attribute("value")
                break
        if not target_val:
            if len(select_obj.options) > 1:
                target_val = select_obj.options[1].get_attribute("value")

        if target_val:
            select_dropdown_value(driver, TALUKA_XPATH, target_val)

    wait_for_loader(driver)
    time.sleep(1)

    # Ensure Hobli dropdown is populated before completing navigation
    end_hobli_wait = time.time() + 12
    while time.time() < end_hobli_wait:
        if get_dropdown_options(driver, HOBLI_XPATH, timeout=2):
            break
        time.sleep(1)

    # 7. Select Property/Survey boundary radio
    ensure_property_boundary_selected(driver, label=label)

    # 8. Set initial date range if empty (From 01-JAN-2020 to TODAY's date in India)
    try:
        d1 = driver.find_element(By.XPATH, '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[8]/div[1]//input').get_attribute("value")
        d2 = driver.find_element(By.XPATH, '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[8]/div[2]//input').get_attribute("value")
        now = datetime.now()
        curr_day = now.day
        curr_month = now.strftime("%b").upper()
        curr_year = now.year
        if not d1:
            select_date(driver, 1, "JAN", 2020, "first")
        if not d2:
            select_date(driver, curr_day, curr_month, curr_year, "second")
    except Exception:
        pass

    log(label, "✅ Reached EC Search page; District and Taluka ready.")


def get_dropdown_options(driver, xpath, timeout=15, dedup_by_name=True, label=None):
    """
    Waits for and returns all valid options from a <select> dropdown.
    If dedup_by_name=True, filters out duplicate option names (case-insensitive)
    keeping only the first unique occurrence.
    """
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
            dup_count = 0
            for option in Select(select_element).options:
                text = option.text.strip()
                value = option.get_attribute("value")
                if text and value not in (None, "", "0") and not any(
                    marker in text.lower() for marker in ("select", "choose", "--")
                ):
                    norm = text.lower()
                    if dedup_by_name:
                        if norm in seen_names:
                            dup_count += 1
                            continue
                        seen_names.add(norm)
                    options.append((value, text))
            if options:
                if dup_count > 0 and label:
                    log(label, f"🧹 Filtered out {dup_count} duplicate options in dropdown; {len(options)} unique options remain.")
                return options
        except Exception:
            pass
        time.sleep(0.4)
    return []


def get_selected_dropdown_text(driver, xpath):
    """Returns the visible text of the currently selected option in a dropdown."""
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
            selected_option = Select(select_elements[0]).first_selected_option
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

    # Wait up to 10s if the option with this value is still being populated by Angular
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

    # Ensure value assignment in DOM and dispatch input, change, and blur events for Angular
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
    time.sleep(1)


# ======================================================================================
# SCRAPING WORKER LOOP
# ======================================================================================
def extract_valid_records_from_table(
    table_html,
    survey_value,
    batch_idx,
    district_name="",
    taluka_name="",
    hobli_name="",
    village_name="",
):
    """
    Parses table_html and returns a valid DataFrame if actual transaction records exist.
    Primary parse is pd.read_html with automatic fallback to native BeautifulSoup
    if parser issues occur on different PCs. Handles merged cells (rowspan) and avoids false drops.
    """
    if not table_html:
        return None

    df = None
    # Primary: pd.read_html
    try:
        soup = BeautifulSoup(str(table_html), 'html.parser')
        target_table = soup.find('table') or soup
        dfs = pd.read_html(io.StringIO(str(target_table)))
        if dfs and not dfs[0].empty:
            df = dfs[0]
    except Exception as read_err:
        # Secondary fallback: native BeautifulSoup extraction (works on all PCs without lxml)
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
        except Exception as soup_err:
            log("System", f"⚠️ Table parse error: {soup_err} (pd.read_html error: {read_err})")
            return None

    if df is None or df.empty:
        return None

    try:
        # Flatten MultiIndex columns if present
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [' '.join(str(c) for c in col).strip() for col in df.columns.values]

        # Clean column names
        df.columns = [str(c).strip() for c in df.columns]

        # Discard only if table has 1 row AND that row contains placeholder text
        if len(df) <= 1:
            all_text = " ".join(df.astype(str).values.flatten()).lower()
            if any(marker in all_text for marker in ["no data", "no record", "record not found", "not available", "ಮಾಹಿತಿ ಲಭ್ಯವಿಲ್ಲ"]):
                return None

        # Check for valid Document Reference column (both Kannada and English variants)
        doc_col = next(
            (c for c in df.columns if any(k in str(c).lower() for k in [
                'ದಸ್ತಾವೇಜಿನ ಉಲ್ಲೇಖ', 'ದಸ್ತಾವೇಜು', 'document reference', 'doc ref', 'registration no', 'doc no'
            ])),
            None
        )

        # STRICT: A table without a Document Reference column is NEVER a transaction result table!
        if not doc_col:
            return None

        # Forward-fill document reference when multiple property rows belong to one deed (rowspan handling)
        df[doc_col] = df[doc_col].ffill()

        # Row-level validation on document reference number
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
        return df
    except Exception as parse_e:
        log("System", f"⚠️ Error processing table DataFrame: {parse_e}")
        return None


def load_progress_state(file_path):
    """Load hierarchical progress while accepting legacy integer progress files."""
    default_state = {"hobli_index": 0, "village_index": 0, "property_value": 1, "consecutive_empty": 0, "village_name": ""}
    if not os.path.exists(file_path):
        return default_state

    try:
        with open(file_path, "r", encoding="utf-8") as progress_file:
            raw_progress = progress_file.read().strip()
        if not raw_progress:
            return default_state

        parsed_progress = json.loads(raw_progress)
        if isinstance(parsed_progress, dict):
            return {
                "hobli_index": max(0, int(parsed_progress.get("hobli_index", 0))),
                "village_index": max(0, int(parsed_progress.get("village_index", 0))),
                "property_value": max(1, int(parsed_progress.get("property_value", 1))),
                "consecutive_empty": max(0, int(parsed_progress.get("consecutive_empty", 0))),
                "village_name": str(parsed_progress.get("village_name", "")),
            }

        return {**default_state, "property_value": max(1, int(parsed_progress))}
    except (ValueError, TypeError, json.JSONDecodeError):
        return default_state


def save_progress_state(file_path, hobli_index, village_index, property_value, consecutive_empty=0, village_name="", scraped_by=""):
    """Atomically save the worker's current hobli, village, property value, consecutive empty count, and worker label."""
    progress_state = {
        "hobli_index": max(0, int(hobli_index)),
        "village_index": max(0, int(village_index)),
        "property_value": max(1, int(property_value)),
        "consecutive_empty": max(0, int(consecutive_empty)),
        "village_name": str(village_name),
        "scraped_by": str(scraped_by),
    }
    temporary_path = f"{file_path}.tmp"
    with open(temporary_path, "w", encoding="utf-8") as progress_file:
        json.dump(progress_state, progress_file)
    os.replace(temporary_path, file_path)


# ======================================================================================
# VILLAGE SCRAPING SUMMARY TRACKING (JSON & EXCEL)
# ======================================================================================
def load_village_summary(json_path=VILLAGE_SUMMARY_JSON):
    """Loads current village scraping summary dictionary from JSON."""
    if os.path.exists(json_path):
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_village_summary(summary_data, json_path=VILLAGE_SUMMARY_JSON, xlsx_path=VILLAGE_SUMMARY_XLSX):
    """Saves village scraping summary to both JSON and Excel atomically."""
    tmp_json = f"{json_path}.tmp"
    try:
        with open(tmp_json, 'w', encoding='utf-8') as f:
            json.dump(summary_data, f, indent=2, ensure_ascii=False)
        os.replace(tmp_json, json_path)
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
            save_df_atomically_with_lock_wait(sdf, xlsx_path)
        except Exception:
            pass


def update_village_summary_record(village_name, hobli_name, taluka_name, district_name, searched_val=None, scraped_val=None, new_recs_count=0, is_completed=False, scraped_by=None, status=None):
    """Updates a single village's summary record with the latest searched or scraped property values."""
    if not village_name:
        return
    v_key = str(village_name).strip()
    summary = load_village_summary()
    record = summary.get(v_key, {
        'district': str(district_name or 'Basavanagudi'),
        'taluka': str(taluka_name or 'Anekal'),
        'hobli': str(hobli_name or 'Anekal Kasaba Hobli'),
        'village': v_key,
        'scraped_by': str(scraped_by or ''),
        'last_scraped_value': 0,
        'last_searched_value': 0,
        'total_records': 0,
        'status': 'IN_PROGRESS',
        'last_updated': ''
    })

    if scraped_by:
        record['scraped_by'] = str(scraped_by)
    if searched_val is not None:
        record['last_searched_value'] = max(int(record.get('last_searched_value', 0)), int(searched_val))
    if scraped_val is not None:
        record['last_scraped_value'] = max(int(record.get('last_scraped_value', 0)), int(scraped_val))
    if new_recs_count > 0:
        record['total_records'] = int(record.get('total_records', 0)) + int(new_recs_count)
    if status:
        record['status'] = str(status)
    elif is_completed:
        record['status'] = 'COMPLETED'
    if searched_val is not None and (is_completed or status == 'COMPLETED'):
        record['last_searched_value'] = max(int(record.get('last_searched_value', 0)), int(searched_val))

    record['last_updated'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    summary[v_key] = record
    save_village_summary(summary)


def sync_village_summary(excel_file=OUTPUT_FILE, json_path=VILLAGE_SUMMARY_JSON, xlsx_path=VILLAGE_SUMMARY_XLSX):
    """Synchronizes village scraping summary from both Excel and active worker progress files."""
    summary_data = load_village_summary(json_path)

    # 1. Update from primary Excel
    if os.path.exists(excel_file):
        try:
            df = pd.read_excel(excel_file)
            if 'Village' in df.columns:
                for village, grp in df.groupby('Village'):
                    v_str = str(village).strip()
                    hobli = grp['Hobli/Town'].iloc[0] if 'Hobli/Town' in grp.columns else ''
                    taluka = grp['Taluka'].iloc[0] if 'Taluka' in grp.columns else 'Anekal'
                    dist = grp['District'].iloc[0] if 'District' in grp.columns else 'Basavanagudi'
                    max_scraped = int(grp['Survey Value'].max()) if 'Survey Value' in grp.columns and not grp['Survey Value'].dropna().empty else 0
                    total_recs = len(grp)
                    rec = summary_data.get(v_str, {})
                    rec.update({
                        'district': str(dist),
                        'taluka': str(taluka),
                        'hobli': str(hobli),
                        'village': v_str,
                        'last_scraped_value': max(max_scraped, rec.get('last_scraped_value', 0)),
                        'last_searched_value': max(rec.get('last_searched_value', 0), max_scraped),
                        'total_records': total_recs,
                        'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                    if 'status' not in rec:
                        rec['status'] = 'IN_PROGRESS'
                    summary_data[v_str] = rec
        except Exception:
            pass

    # 2. Update from worker progress files
    for pf in glob.glob(os.path.join(BASE_DIR, 'igr_last_scraped_*.txt')):
        try:
            with open(pf, 'r', encoding='utf-8') as f:
                data = json.load(f)
            vname = data.get('village_name', '').strip()
            val = int(data.get('property_value', 1))
            w_by = data.get('scraped_by', '').strip()
            if vname:
                rec = summary_data.get(vname, {
                    'district': 'Basavanagudi',
                    'taluka': 'Anekal',
                    'hobli': 'Anekal Kasaba Hobli',
                    'village': vname,
                    'scraped_by': w_by,
                    'last_scraped_value': 0,
                    'last_searched_value': 0,
                    'total_records': 0,
                    'status': 'IN_PROGRESS',
                    'last_updated': ''
                })
                if w_by and not rec.get('scraped_by'):
                    rec['scraped_by'] = w_by
                rec['last_searched_value'] = max(rec.get('last_searched_value', 0), val)
                rec['last_updated'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                summary_data[vname] = rec
        except Exception:
            pass

    save_village_summary(summary_data, json_path, xlsx_path)
    return summary_data


def run_scraping_for_current_village(
    driver,
    data_queue,
    excel_file=OUTPUT_FILE,
    progress_file="igr_last_scraped_1.txt",
    label="Worker-1",
    login_email=None,
    login_password=None,
    imap_email=None,
    imap_password=None,
    hobli_index=0,
    village_index=0,
    hobli_value=None,
    village_value=None,
    shared_seen_docs=None,
    district_name="",
    taluka_name="",
    hobli_name="",
    village_name="",
    shared_worker_status=None,
):
    if not district_name:
        district_name = get_selected_dropdown_text(driver, DISTRICT_XPATH) or "Basavanagudi"
    if not taluka_name:
        taluka_name = get_selected_dropdown_text(driver, TALUKA_XPATH) or "Anekal"
    if not hobli_name:
        hobli_name = get_selected_dropdown_text(driver, HOBLI_XPATH)
    if not village_name:
        village_name = get_selected_dropdown_text(driver, VILLAGE_XPATH)

    wait_for_loader(driver)
    wait_for_page_ready(driver)
    ensure_property_boundary_selected(driver, label=label)
    all_property_types = get_property_number_type_options(driver)
    BATCH_SIZE = 10
    batches = [all_property_types[i:i + BATCH_SIZE] for i in range(0, len(all_property_types), BATCH_SIZE)] if all_property_types else []

    log(label, f"📋 Detected {len(all_property_types)} Property Number Types split into {len(batches)} batches.")

    summary = load_village_summary()
    v_rec = summary.get(str(village_name).strip(), {})
    total_village_records = int(v_rec.get("total_records", 0))

    progress_state = load_progress_state(progress_file)
    if (
        progress_state["hobli_index"] == hobli_index
        and (
            progress_state["village_index"] == village_index
            or (progress_state.get("village_name") and progress_state.get("village_name").strip().lower() == str(village_name).strip().lower())
        )
    ):
        start_value = progress_state["property_value"]
        consecutive_empty_values = progress_state.get("consecutive_empty", 0)
    else:
        if v_rec and v_rec.get("status") == "IN_PROGRESS" and v_rec.get("last_searched_value", 0) > 0:
            start_value = int(v_rec.get("last_searched_value", 0)) + 1
            consecutive_empty_values = 0
            log(label, f"🔄 Resuming village '{village_name}' from Survey #{start_value} (Existing records: {total_village_records})")
        else:
            start_value = 1
            consecutive_empty_values = 0
    update_worker_status(
        shared_worker_status, label,
        village=village_name,
        survey=f"#{start_value}",
        empty_streak=f"{consecutive_empty_values}/{MAX_CONSECUTIVE_EMPTY}",
        records=total_village_records,
        state="SCRAPING",
        detail=f"Searching #{start_value}"
    )

    for value in itertools.count(start=start_value):
        log(label, f"🔍 [{village_name}] Survey #{value} | Searching... (Empty streak: {consecutive_empty_values}/{MAX_CONSECUTIVE_EMPTY})")
        update_worker_status(
            shared_worker_status, label,
            village=village_name,
            survey=f"#{value}",
            empty_streak=f"{consecutive_empty_values}/{MAX_CONSECUTIVE_EMPTY}",
            records=total_village_records,
            state="SCRAPING",
            detail=f"Survey #{value}"
        )

        value_had_new_records = False
        value_had_portal_records = False
        value_completed = False

        for value_attempt in range(1, 4):
            try:
                if not batches:
                    all_property_types = get_property_number_type_options(driver)
                    batches = [all_property_types[i:i + BATCH_SIZE] for i in range(0, len(all_property_types), BATCH_SIZE)] if all_property_types else [[]]

                for batch_idx, batch in enumerate(batches, start=1):
                    wait_for_loader(driver)
                    close_result_modal(driver)
                    wait_for_page_ready(driver)

                    if batch and not is_batch_already_setup(driver, batch):
                        setup_property_type_batch(driver, batch)

                    # Fast atomic JS input injection with Angular event dispatching
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
                    """, str(value))

                    inputs = driver.find_elements(By.XPATH, "//form//div[4]//input[not(contains(@placeholder, 'DD/MM')) and not(@type='radio') and not(@type='checkbox')]")
                    if not inputs:
                        log(label, f"⚠️ Property inputs not found on page; re-setting up property batch rows...")
                        if batch:
                            setup_property_type_batch(driver, batch)
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
                        """, str(value))
                        inputs = driver.find_elements(By.XPATH, "//form//div[4]//input[not(contains(@placeholder, 'DD/MM')) and not(@type='radio') and not(@type='checkbox')]")

                    entered_values = [inp.get_attribute("value") for inp in inputs]
                    if not entered_values or any(entered_value != str(value) for entered_value in entered_values):
                        for inp in inputs:
                            try:
                                if inp.get_attribute("value") != str(value):
                                    driver.execute_script(
                                        "arguments[0].value = arguments[1];"
                                        "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));"
                                        "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                                        inp,
                                        str(value),
                                    )
                            except Exception:
                                pass
                        entered_values = [inp.get_attribute("value") for inp in inputs]
                        if not entered_values or any(entered_value != str(value) for entered_value in entered_values):
                            raise TimeoutException(
                                f"Property value {value} was not entered in every property field: {entered_values}"
                            )

                    # Ensure dates (from 1-JAN-2020 to current date)
                    dates_ok = driver.execute_script("""
                        const inps = document.querySelectorAll("div:nth-child(8) input");
                        return inps.length >= 2 && inps[0].value.trim() !== '' && inps[1].value.trim() !== '';
                    """)
                    if not dates_ok:
                        now = datetime.now()
                        curr_day = now.day
                        curr_month = now.strftime("%b").upper()
                        curr_year = now.year
                        try:
                            select_date(driver, 1, "JAN", 2020, "first")
                            select_date(driver, curr_day, curr_month, curr_year, "second")
                        except Exception:
                            pass

                    # CAPTCHA solving for search
                    max_attempts = 15
                    attempt = 0
                    captcha2_success = False
                    no_records_confirmed = False

                    while not captcha2_success and attempt < max_attempts:
                        attempt += 1

                        try:
                            wait_for_loader(driver)
                            captcha2_input = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[9]/div[2]/input'))
                            )
                            captcha2_input.click()

                            captcha2_element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[9]/div[1]/img'))
                            )

                            if attempt > 1:
                                refresh_search_captcha(driver, captcha2_element)
                                try:
                                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", captcha2_element)
                                except Exception:
                                    pass

                            # Ultra-fast in-memory CAPTCHA solve (zero disk I/O, DOM/canvas extraction)
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
                                log(label, f"⚠️ CAPTCHA OCR failed for value {value} (attempt {attempt}/{max_attempts}); refreshing CAPTCHA and retrying.")
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
                                entered_captcha = captcha2_input.get_attribute("value") or ""
                                if entered_captcha.strip().upper() != captcha2_text:
                                    log(label, f"⚠️ CAPTCHA input was not filled for value {value} (attempt {attempt}/{max_attempts}); retrying.")
                                    refresh_search_captcha(driver, captcha2_element)
                                    continue

                            # Click Search button
                            try:
                                search_btn = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Search' or contains(text(), 'Search')]"))
                                )
                            except Exception:
                                search_btn = driver.find_element(By.XPATH, '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[10]/button[1]')

                            try:
                                results_container = driver.find_element(By.ID, "PdfData")
                                previous_results = results_container.get_attribute("innerHTML") or ""
                                previous_table = results_container.find_elements(By.CSS_SELECTOR, "table")
                                previous_table_html = (
                                    previous_table[0].get_attribute("outerHTML")
                                    if previous_table
                                    else ""
                                )
                                if previous_table:
                                    driver.execute_script("arguments[0].remove();", previous_table[0])
                                driver.execute_script("arguments[0].innerHTML = '';", results_container)
                            except NoSuchElementException:
                                results_container = None
                                previous_results = ""
                                previous_table_html = ""

                            search_btn.click()
                            time.sleep(0.05)
                            wait_for_loader(driver, timeout=25)

                            # Check browser alert
                            is_alert_no_records = False
                            try:
                                alert = driver.switch_to.alert
                                alert_text = alert.text
                                alert.accept()
                                if any(w in alert_text.lower() for w in ["captcha", "invalid", "incorrect"]):
                                    log(label, f"⚠️ CAPTCHA rejected by browser alert for value {value}: {alert_text}")
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
                                log(label, f"⚠️ CAPTCHA rejected for value {value} (attempt {attempt}/{max_attempts}); refreshing and retrying.")
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
                                        log(label, f"⚠️ CAPTCHA inline error for value {value}: {error_element.text.strip()}")
                                        save_captcha_sample(captcha2_bytes, captcha2_text, is_success=False, context="search_inline_err", label=label)
                                        refresh_search_captcha(driver, captcha2_element)
                                        is_captcha_error = True
                                        break
                            except Exception:
                                pass

                            if is_captcha_error:
                                time.sleep(0.2)
                                continue

                            # If no captcha errors, search request was accepted
                            captcha2_success = True
                            save_captcha_sample(captcha2_bytes, captcha2_text, is_success=True, context="search", label=label)
                            break

                        except (TimeoutException, WebDriverException) as e:
                            log(label, f"⚠️ Search CAPTCHA attempt {attempt}/{max_attempts} failed for value {value}: {e}; retrying.")
                            try:
                                current_captcha = driver.find_element(
                                    By.XPATH,
                                    '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[9]/div[1]/img'
                                )
                                refresh_search_captcha(driver, current_captcha)
                            except Exception:
                                if attempt % 3 == 0:
                                    driver.refresh()
                                    wait_for_loader(driver)
                            time.sleep(0.2)
                            continue

                    if not captcha2_success:
                        raise TimeoutException(f"Search CAPTCHA not solved after {max_attempts} attempts for value {value}")

                    if no_records_confirmed:
                        # FAST-PATH: Kaveri officially confirmed 0 records for this search.
                        # Instantly skip table polling and BeautifulSoup (saves ~12s per empty survey)!
                        close_result_modal(driver)
                        try:
                            driver.execute_script("var el = document.getElementById('PdfData'); if (el) { el.innerHTML = ''; }")
                        except Exception:
                            pass
                        continue

                    # Wait for results table or loader to finish
                    wait_for_loader(driver, timeout=25)

                    # Extract table data and push to multiprocess queue
                    try:
                        table_html = None
                        # Dynamic fast polling up to 25 seconds for table rows to be rendered inside #PdfData
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

                        if not table_html:
                            try:
                                soup = BeautifulSoup(driver.page_source, 'html.parser')
                                for sel in ['#PdfData table', '#PdfData > table', 'div[id*="PdfData"] table']:
                                    t_found = soup.select_one(sel)
                                    if t_found and len(t_found.find_all('tr')) > 1:
                                        table_html = str(t_found)
                                        break
                            except Exception:
                                pass

                        valid_df = extract_valid_records_from_table(
                            table_html, value, batch_idx,
                            district_name=district_name,
                            taluka_name=taluka_name,
                            hobli_name=hobli_name,
                            village_name=village_name,
                        )
                        if valid_df is not None and not valid_df.empty:
                            value_had_portal_records = True
                            # Check if incoming records already exist in Excel
                            doc_col = next(
                                (c for c in valid_df.columns if any(k in str(c).lower() for k in [
                                    'ದಸ್ತಾವೇಜಿನ ಉಲ್ಲೇಖ', 'ದಸ್ತಾವೇಜು', 'document reference', 'doc ref', 'registration no', 'doc no'
                                ])),
                                None
                            )
                            if doc_col:
                                incoming_refs = [str(r).strip() for r in valid_df[doc_col].dropna() if str(r).strip()]
                                if shared_seen_docs is not None:
                                    new_refs = [r for r in incoming_refs if r not in shared_seen_docs]
                                else:
                                    existing_refs = set()
                                    if os.path.exists(excel_file):
                                        try:
                                            edf = pd.read_excel(excel_file)
                                            ecol = next(
                                                (c for c in edf.columns if any(k in str(c).lower() for k in [
                                                    'ದಸ್ತಾವೇಜಿನ ಉಲ್ಲೇಖ', 'ದಸ್ತಾವೇಜು', 'document reference', 'doc ref'
                                                ])),
                                                None
                                            )
                                            if ecol:
                                                existing_refs = set(edf[ecol].dropna().astype(str).str.strip())
                                        except Exception:
                                            pass
                                    new_refs = [r for r in incoming_refs if r not in existing_refs]

                                if not new_refs:
                                    pass
                                else:
                                    value_had_new_records = True
                                    if shared_seen_docs is not None:
                                        for r in new_refs:
                                            shared_seen_docs[r] = True
                                    new_rows_df = valid_df[valid_df[doc_col].astype(str).str.strip().isin(new_refs)]
                                    data_queue.put((new_rows_df, excel_file, value, batch_idx, label))
                                    total_village_records += len(new_rows_df)
                                    log(label, f"⭐ [{village_name}] Survey #{value} (Batch {batch_idx}) | Found {len(new_rows_df)} new unique records! (Village Total: {total_village_records} docs)")
                            else:
                                value_had_new_records = True
                                data_queue.put((valid_df, excel_file, value, batch_idx, label))
                                total_village_records += len(valid_df)
                                log(label, f"⭐ [{village_name}] Survey #{value} (Batch {batch_idx}) | Found {len(valid_df)} new records! (Village Total: {total_village_records} docs)")
                        else:
                            # Batch has no records; silently continue to next batch without console spam
                            pass

                    except Exception as e:
                        log(label, f"❌ Table extraction error: {e}")

                    close_result_modal(driver)
                    wait_for_page_ready(driver)
                    try:
                        driver.execute_script("var el = document.getElementById('PdfData'); if (el) { el.innerHTML = ''; }")
                    except Exception:
                        pass

                value_completed = True
                break

            except (TimeoutException, WebDriverException) as e:
                log(label, f"⚠️ Attempt {value_attempt}/3 failed for value {value}: {e}. Refreshing page/session...")
                try:
                    login_form_visible = bool(driver.find_elements(By.ID, "userName"))
                    if login_form_visible and login_email and login_password:
                        driver = relogin(
                            driver,
                            login_email,
                            login_password,
                            label=label,
                            imap_email=imap_email,
                            imap_password=imap_password,
                        )
                        if not driver:
                            raise WebDriverException("Re-login failed after session expiry")
                    else:
                        driver.refresh()
                        wait_for_loader(driver)
                    navigate_to_search_page(driver, label=label)
                    hobli_xpath = '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[1]/div[3]/select'
                    village_xpath = '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[1]/div[4]/select'
                    if hobli_value:
                        select_dropdown_value(driver, hobli_xpath, hobli_value)
                    if village_value:
                        select_dropdown_value(driver, village_xpath, village_value)
                    ensure_property_boundary_selected(driver, label=label)
                except Exception as recovery_err:
                    log(label, f"⚠️ Recovery attempt failed for {label}: {recovery_err}")
                if value_attempt == 3:
                    log(label, f"❌ Value {value} failed after 3 attempts. Advancing to next value.")
            except Exception as e:
                log(label, f"Error processing value {value}: {e}")
                try:
                    navigate_to_search_page(driver, label=label)
                except Exception:
                    driver.refresh()
                if value_attempt == 3:
                    break

        if value_completed and (value_had_new_records or value_had_portal_records):
            consecutive_empty_values = 0
            scraped_this_val = value
            status_msg = "Records saved" if value_had_new_records else "Existing records verified"
            log(label, f"✅ [{village_name}] Survey #{value} completed | {status_msg} (Empty streak reset to 0)")
        else:
            consecutive_empty_values += 1
            scraped_this_val = None
            log(label, f"⏩ [{village_name}] Survey #{value} | 0 records | Empty streak: {consecutive_empty_values}/{MAX_CONSECUTIVE_EMPTY}")

        save_progress_state(progress_file, hobli_index, village_index, value + 1, consecutive_empty_values, village_name=village_name, scraped_by=label)

        update_worker_status(
            shared_worker_status, label,
            village=village_name,
            survey=f"#{value}",
            empty_streak=f"{consecutive_empty_values}/{MAX_CONSECUTIVE_EMPTY}",
            records=total_village_records,
            state="SCRAPING",
            detail=f"Survey #{value} done"
        )

        if data_queue is not None:
            data_queue.put(('summary_update', district_name, taluka_name, hobli_name, village_name, value, scraped_this_val, False, label))

        if consecutive_empty_values >= MAX_CONSECUTIVE_EMPTY:
            log(label, f"🎉 [{village_name}] VILLAGE COMPLETED! Reached {MAX_CONSECUTIVE_EMPTY} consecutive empty numbers. Total records: {total_village_records}.")
            update_worker_status(
                shared_worker_status, label,
                village=village_name,
                survey=f"#{value}",
                empty_streak=f"{consecutive_empty_values}/{MAX_CONSECUTIVE_EMPTY}",
                records=total_village_records,
                state="VILLAGE_DONE",
                detail=f"Completed {village_name}"
            )
            if data_queue is not None:
                data_queue.put(('summary_update', district_name, taluka_name, hobli_name, village_name, value, scraped_this_val, True, label))
            return True

    return True


def run_hierarchy_scraping(
    driver,
    data_queue,
    excel_file=OUTPUT_FILE,
    progress_file="igr_last_scraped_1.txt",
    label="Worker-1",
    worker_index=0,
    total_workers=1,
    login_email=None,
    login_password=None,
    imap_email=None,
    imap_password=None,
    shared_seen_docs=None,
    shared_claimed_villages=None,
    shared_completed_villages=None,
    village_lock=None,
    shared_worker_status=None,
):
    hobli_xpath = '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[1]/div[3]/select'
    village_xpath = '/html/body/div[1]/app-root/div/div/app-ec-search-citizen/div/div[4]/div/div/form/div[1]/div[4]/select'

    hobli_options = get_dropdown_options(driver, HOBLI_XPATH, timeout=10)
    if not hobli_options:
        log(label, "⚠️ Hobli options not loaded; re-triggering Taluka selection...")
        try:
            taluka_elem = driver.find_element(By.XPATH, TALUKA_XPATH)
            select_obj = Select(taluka_elem)
            target_val = None
            for opt in select_obj.options:
                t = opt.text.strip().lower()
                if "anekal" in t or "ಆನೇಕಲ್" in t:
                    target_val = opt.get_attribute("value")
                    break
            if not target_val and len(select_obj.options) > 1:
                target_val = select_obj.options[1].get_attribute("value")
            if target_val:
                select_dropdown_value(driver, TALUKA_XPATH, target_val)
        except Exception:
            pass
        hobli_options = get_dropdown_options(driver, HOBLI_XPATH, timeout=15)

    if not hobli_options:
        log(label, "⛔ No hobli options found; worker stopping.")
        return

    district_name = get_selected_dropdown_text(driver, DISTRICT_XPATH) or "Basavanagudi"
    taluka_name = get_selected_dropdown_text(driver, TALUKA_XPATH) or "Anekal"
    log(label, f"📍 Location Hierarchy: District='{district_name}', Taluka='{taluka_name}'")

    progress_state = load_progress_state(progress_file)
    resume_index = min(progress_state["hobli_index"], len(hobli_options) - 1)

    target_opt_idx = TARGET_HOBLI_OPTION_INDEX
    if target_opt_idx is None and "/option[" in HOBLI_XPATH:
        try:
            m = re.search(r'option\[(\d+)\]', HOBLI_XPATH)
            if m:
                target_opt_idx = int(m.group(1))
        except Exception:
            pass

    if target_opt_idx is not None and target_opt_idx > 0:
        # In Kaveri dropdown, option[1] is placeholder '--Select Hobli--', option[2] is 1st actual hobli (index 0)
        target_idx = target_opt_idx - 2 if target_opt_idx >= 2 else 0
        if 0 <= target_idx < len(hobli_options):
            assigned_indices = [target_idx]
            target_name = hobli_options[target_idx][1]
            log(label, f"🎯 Targeted Hobli option[{target_opt_idx}]: '{target_name}' (value: {hobli_options[target_idx][0]}).")
        else:
            log(label, f"⚠️ option[{target_opt_idx}] exceeds available hoblis ({len(hobli_options)}); defaulting to all hoblis.")
            assigned_indices = list(range(resume_index, len(hobli_options)))
    else:
        # All browsers process hoblis; villages are distributed dynamically on demand
        assigned_indices = list(range(resume_index, len(hobli_options)))
        log(
            label,
            f"🧩 All browsers will work on {len(assigned_indices)} hoblis; villages are claimed dynamically."
        )

    for hobli_index in assigned_indices:
        hobli_value, hobli_name = hobli_options[hobli_index]
        select_dropdown_value(driver, HOBLI_XPATH, hobli_value)

        village_options = get_dropdown_options(driver, VILLAGE_XPATH, timeout=15, dedup_by_name=True, label=label)
        if not village_options:
            log(label, f"⚠️ Village options not loaded for hobli '{hobli_name}'; re-selecting hobli...")
            select_dropdown_value(driver, HOBLI_XPATH, hobli_value)
            village_options = get_dropdown_options(driver, VILLAGE_XPATH, timeout=15, dedup_by_name=True, label=label)
            if not village_options:
                log(label, f"ℹ️ No villages found for hobli '{hobli_name}'.")
                continue

        # ------------------------------------------------------------------------------
        # DYNAMIC TASK DISPATCH: Workers dynamically claim unclaimed / incomplete villages
        # ------------------------------------------------------------------------------
        while True:
            target_village = None
            summary = load_village_summary()

            if village_lock is not None and shared_claimed_villages is not None:
                with village_lock:
                    for v_idx, (v_val, v_name) in enumerate(village_options):
                        v_key = v_name.strip()
                        norm_key = v_key.lower()

                        # Skip if completed
                        v_rec = summary.get(v_key, {})
                        is_completed = (
                            v_rec.get("status") == "COMPLETED"
                            or (shared_completed_villages is not None and norm_key in shared_completed_villages)
                        )
                        if SKIP_COMPLETED_VILLAGES and is_completed:
                            if shared_completed_villages is not None:
                                shared_completed_villages[norm_key] = True
                            continue

                        # Skip if currently claimed by another active worker
                        if norm_key in shared_claimed_villages:
                            continue

                        # Claim this village for this worker
                        shared_claimed_villages[norm_key] = label
                        target_village = (v_idx, v_val, v_name)
                        break
            else:
                for v_idx, (v_val, v_name) in enumerate(village_options):
                    v_key = v_name.strip()
                    v_rec = summary.get(v_key, {})
                    if SKIP_COMPLETED_VILLAGES and v_rec.get("status") == "COMPLETED":
                        continue
                    target_village = (v_idx, v_val, v_name)
                    break

            if not target_village:
                log(label, f"🏁 All villages in hobli '{hobli_name}' are claimed or completed.")
                break

            village_index, village_value, village_name = target_village
            norm_vname = village_name.strip().lower()
            log(label, f"📌 [DYNAMIC CLAIM] Claimed village '{village_name}' ({village_index + 1}/{len(village_options)}) in hobli '{hobli_name}'.")

            update_worker_status(
                shared_worker_status, label,
                village=village_name,
                survey="#1",
                empty_streak="0/50",
                records=0,
                state="SCRAPING",
                detail=f"Claimed village {village_index + 1}/{len(village_options)}"
            )

            # Immediately register claimed village as IN_PROGRESS in summary tracker
            if data_queue is not None:
                data_queue.put(('summary_update', district_name, taluka_name, hobli_name, village_name, 0, None, False, label))

            # Navigate to clean search form and select hobli & village
            try:
                navigate_to_search_page(driver, label=label)
                select_dropdown_value(driver, HOBLI_XPATH, hobli_value)
                select_dropdown_value(driver, VILLAGE_XPATH, village_value)
                ensure_property_boundary_selected(driver, label=label)
            except Exception as nav_err:
                log(label, f"⚠️ Navigation for village '{village_name}' failed: {nav_err}. Refreshing page and retrying...")
                try:
                    driver.refresh()
                    wait_for_loader(driver)
                    navigate_to_search_page(driver, label=label)
                    select_dropdown_value(driver, HOBLI_XPATH, hobli_value)
                    select_dropdown_value(driver, VILLAGE_XPATH, village_value)
                    ensure_property_boundary_selected(driver, label=label)
                except Exception as fatal_v_err:
                    log(label, f"❌ Could not initialize village '{village_name}': {fatal_v_err}")
                    raise fatal_v_err

            log(label, f"🏘️ Starting scrape for village '{village_name}' in hobli '{hobli_name}'.")

            village_finished = False
            try:
                village_finished = run_scraping_for_current_village(
                    driver, data_queue,
                    excel_file=excel_file,
                    progress_file=progress_file,
                    label=label,
                    login_email=login_email,
                    login_password=login_password,
                    imap_email=imap_email,
                    imap_password=imap_password,
                    hobli_index=hobli_index,
                    village_index=village_index,
                    hobli_value=hobli_value,
                    village_value=village_value,
                    shared_seen_docs=shared_seen_docs,
                    district_name=district_name,
                    taluka_name=taluka_name,
                    hobli_name=hobli_name,
                    village_name=village_name,
                    shared_worker_status=shared_worker_status,
                )
            except Exception as v_err:
                log(label, f"⚠️ Error processing village '{village_name}': {v_err}")
                if data_queue is not None:
                    data_queue.put(('summary_update', district_name, taluka_name, hobli_name, village_name, 0, None, 'FAILED', label))
            finally:
                if village_lock is not None and shared_claimed_villages is not None:
                    with village_lock:
                        shared_claimed_villages.pop(norm_vname, None)
                        if village_finished and shared_completed_villages is not None:
                            shared_completed_villages[norm_vname] = True

            save_progress_state(progress_file, hobli_index, village_index + 1, 1, 0, village_name=village_name, scraped_by=label)

        log(label, f"✅ Completed all available villages for hobli '{hobli_name}'.")

    log(label, "✅ Completed village assignments across all hoblis.")


# ======================================================================================
# WORKER PROCESS ENTRYPOINT
# ======================================================================================
def worker_process_main(
    account, worker_index, total_workers, data_queue,
    start_delay_seconds=0, shared_seen_docs=None,
    shared_claimed_villages=None, shared_completed_villages=None, village_lock=None,
    shared_worker_status=None,
):
    """Entry point for each isolated worker OS process."""
    configure_console_utf8()
    enable_windows_ansi()

    label = account.get("label", f"Worker-{worker_index+1}")
    browser_type = account.get("browser_type", "chrome")
    excel_file = account.get("excel_file", OUTPUT_FILE)
    progress_file = account.get("progress_file", f"igr_last_scraped_{worker_index+1}.txt")
    login_email = account.get("login_email")
    login_password = account.get("login_password")
    imap_email = account.get("imap_email")
    imap_password = account.get("imap_app_password")

    if start_delay_seconds > 0:
        log(label, f"⏳ Stagger Delay: Waiting {start_delay_seconds}s before launch to prevent OTP collisions...")
        start_wait = time.time()
        while time.time() - start_wait < start_delay_seconds:
            remaining = int(start_delay_seconds - (time.time() - start_wait))
            update_worker_status(
                shared_worker_status, label,
                engine=browser_type.upper(),
                state="WAITING_DELAY",
                detail=f"Starts in {remaining}s",
            )
            time.sleep(min(5, max(1, remaining)))

    max_worker_crashes = 30
    crash_count = 0

    while crash_count < max_worker_crashes:
        driver = None
        try:
            update_worker_status(
                shared_worker_status, label,
                engine=browser_type.upper(),
                state="LOGGING_IN",
                detail=f"Launch browser (#{crash_count+1})" if crash_count > 0 else "Opening browser & OTP",
            )
            log(label, f"🚀 Launching {browser_type.upper()} browser and logging in (attempt {crash_count+1}/{max_worker_crashes})...")

            driver = launch_and_login(
                login_email=login_email,
                login_password=login_password,
                label=label,
                browser_type=browser_type,
                imap_email=imap_email,
                imap_password=imap_password,
            )
            if not driver:
                log(label, f"❌ Login attempt failed for {label}. Retrying in 15s...")
                update_worker_status(
                    shared_worker_status, label,
                    engine=browser_type.upper(),
                    state="RETRYING_LOGIN",
                    detail="Login retry in 15s",
                )
                crash_count += 1
                time.sleep(15)
                continue

            update_worker_status(
                shared_worker_status, label,
                engine=browser_type.upper(),
                state="NAVIGATING",
                detail="Opening EC Search",
            )
            with_page_reload_retry(navigate_to_search_page, driver, max_retries=3, label=label)

            run_hierarchy_scraping(
                driver, data_queue,
                excel_file=excel_file, progress_file=progress_file, label=label,
                worker_index=worker_index, total_workers=total_workers,
                login_email=login_email, login_password=login_password,
                imap_email=imap_email, imap_password=imap_password,
                shared_seen_docs=shared_seen_docs,
                shared_claimed_villages=shared_claimed_villages,
                shared_completed_villages=shared_completed_villages,
                village_lock=village_lock,
                shared_worker_status=shared_worker_status,
            )

            update_worker_status(
                shared_worker_status, label,
                engine=browser_type.upper(),
                state="COMPLETED_ALL",
                detail="All villages completed",
            )
            log(label, "🎉 All assigned villages completed successfully!")
            break

        except Exception as e:
            crash_count += 1
            log(label, f"❌ Worker process crashed with error: {e}")
            if driver:
                try:
                    save_debug_screenshot(driver, f"error_{label}.png")
                except Exception:
                    pass
                try:
                    log(label, f"🧹 Closing crashed {browser_type.upper()} browser session to release resources...")
                    driver.quit()
                except Exception:
                    pass
                driver = None

            if crash_count < max_worker_crashes:
                retry_wait = 15
                log(label, f"🔄 AUTO-FALLBACK: Re-launching fresh {browser_type.upper()} browser in {retry_wait}s (crash #{crash_count}/{max_worker_crashes})...")
                update_worker_status(
                    shared_worker_status, label,
                    engine=browser_type.upper(),
                    state="RECOVERING",
                    detail=f"Rebooting browser #{crash_count}",
                )
                time.sleep(retry_wait)
            else:
                log(label, f"⛔ Worker exceeded maximum crash retries ({max_worker_crashes}). Terminating.")
                update_worker_status(
                    shared_worker_status, label,
                    engine=browser_type.upper(),
                    state="CRASHED",
                    detail=str(e)[:22],
                )
                break
        finally:
            if crash_count >= max_worker_crashes:
                log(label, "Worker permanently finished.")


def save_df_atomically_with_lock_wait(df, excel_file, stop_event=None):
    """
    Saves DataFrame to excel_file atomically using a temporary file.
    If Microsoft Excel has the file open, politely waits and retries without crashing.
    """
    temp_excel = f"{excel_file}.tmp.xlsx"
    # Ensure no junk form columns exist in the DataFrame before writing
    junk_cols = [c for c in df.columns if any(k in str(c).lower() for k in ['enter one', 'unnamed', 'add more'])]
    if junk_cols:
        df = df.drop(columns=junk_cols)

    df.to_excel(temp_excel, index=False, engine='openpyxl')

    warned = False
    while True:
        try:
            os.replace(temp_excel, excel_file)
            if warned:
                log("ExcelSaver", f"✅ Excel file lock released; changes successfully saved to {os.path.basename(excel_file)}!")
            return True
        except (PermissionError, OSError):
            if stop_event and stop_event.is_set():
                time.sleep(1)
                try:
                    os.replace(temp_excel, excel_file)
                    return True
                except Exception:
                    return False
            if not warned:
                log("ExcelSaver", f"⚠️ [FILE LOCKED] '{os.path.basename(excel_file)}' is open in Microsoft Excel! Please close Excel so data can be saved. Waiting...")
                warned = True
            time.sleep(2)


def print_live_progress_dashboard(shared_seen_docs=None, shared_worker_status=None):
    """Prints a clean, colorized, aligned multi-process progress dashboard."""
    summary = load_village_summary()
    now_str = datetime.now().strftime("%H:%M:%S")
    total_summary_recs = sum(int(v.get("total_records", 0)) for v in summary.values()) if summary else 0
    total_docs = max(len(shared_seen_docs) if shared_seen_docs is not None else 0, total_summary_recs)
    completed_villages = [v for v in summary.values() if v.get("status") == "COMPLETED"]
    completed_count = len(completed_villages)
    total_villages = len(summary) if summary else 0

    border = "=" * 98
    sub_border = "-" * 98
    lines = [
        "",
        f"{COLOR_CYAN}{BOLD}{border}{RESET}",
        f"{COLOR_WHITE}{BOLD}📊 LIVE KAVERI SCRAPER PROGRESS MONITOR [{now_str}] | Saved Records: {total_summary_recs} | Completed: {completed_count}/{total_villages} Villages{RESET}",
        f"{COLOR_GRAY}{sub_border}{RESET}",
        f"{'Worker':<18} {'Engine':<8} {'Status':<14} {'Current Village':<24} {'Survey #':<10} {'Streak':<10} {'Records'}",
        f"{COLOR_GRAY}{sub_border}{RESET}",
    ]

    if shared_worker_status:
        for label, w_info in sorted(shared_worker_status.items()):
            color = WORKER_COLORS.get(label, COLOR_CYAN)
            engine = str(w_info.get("engine", "-"))[:7]
            raw_state = str(w_info.get("state", "UNKNOWN"))
            state = raw_state[:14]
            village = str(w_info.get("village", "-"))[:22]
            survey = str(w_info.get("survey", "-"))[:9]
            streak = str(w_info.get("empty_streak", "-"))[:9]
            
            # Sync with summary tracker so worker records always reflect real-time saved records in Excel
            v_key = village.strip()
            v_saved = int(summary.get(v_key, {}).get("total_records", 0)) if v_key and v_key != "-" else 0
            w_recs = int(w_info.get("records", 0))
            display_recs = max(v_saved, w_recs)
            recs_str = f"{display_recs} docs"

            if raw_state in ["WAITING_DELAY", "LOGGING_IN", "NAVIGATING", "INITIALIZING", "PENDING", "FAILED_LOGIN"]:
                detail = str(w_info.get("detail", raw_state))[:22]
                lines.append(f"{color}{BOLD}{label:<18}{RESET} {engine:<8} {COLOR_YELLOW}{state:<14}{RESET} {detail:<24} {'-':<10} {'-':<10} {recs_str}")
            else:
                st_color = COLOR_GREEN if raw_state == "SCRAPING" else COLOR_WHITE
                lines.append(f"{color}{BOLD}{label:<18}{RESET} {engine:<8} {st_color}{state:<14}{RESET} {village:<24} {survey:<10} {streak:<10} {recs_str}")
    else:
        lines.append("   (Worker telemetry initializing...)")

    lines.append(f"{COLOR_GRAY}{sub_border}{RESET}")
    if completed_villages:
        comp_summary = [f"{v.get('village')}({v.get('total_records', 0)} docs)" for v in completed_villages[-5:]]
        lines.append(f"{COLOR_GREEN}🏆 Completed Villages ({completed_count}): {', '.join(comp_summary)}{RESET}")
    lines.append(f"{COLOR_CYAN}{BOLD}{border}{RESET}")
    lines.append("")

    formatted_block = "\n".join(lines) + "\n"
    try:
        sys.stdout.write(formatted_block)
        sys.stdout.flush()
    except Exception:
        print(formatted_block, flush=True)


# ======================================================================================
# DEDICATED EXCEL SAVE WRITER PROCESS
# ======================================================================================
def excel_saver_process_main(data_queue, stop_event, shared_seen_docs=None, shared_worker_status=None):
    """
    Dedicated background process that receives scraped DataFrames from all workers
    and writes them safely to Excel with automatic deduplication.
    """
    configure_console_utf8()
    enable_windows_ansi()
    log("ExcelSaver", "Dedicated Excel Writer process started.")
    last_summary_sync = time.time()
    while True:
        try:
            job = data_queue.get(timeout=1)
        except queue.Empty:
            if stop_event.is_set():
                break
            if time.time() - last_summary_sync > 20:
                try:
                    sync_village_summary(OUTPUT_FILE, VILLAGE_SUMMARY_JSON, VILLAGE_SUMMARY_XLSX)
                    last_summary_sync = time.time()
                    print_live_progress_dashboard(shared_seen_docs, shared_worker_status)
                except Exception:
                    pass
            continue

        if job is None:
            break

        # Handle lightweight summary updates
        if isinstance(job, tuple) and len(job) >= 7 and job[0] == 'summary_update':
            w_label = None
            if len(job) >= 9:
                _, dist, taluka, hobli, village, val, sc_val, is_comp, w_label = job[:9]
            elif len(job) == 8:
                _, dist, taluka, hobli, village, val, sc_val, is_comp = job
            else:
                _, dist, taluka, hobli, village, val, is_comp = job
                sc_val = None
            try:
                status_val = is_comp if isinstance(is_comp, str) else ('COMPLETED' if is_comp else None)
                is_completed_bool = bool(is_comp is True or is_comp == 'COMPLETED')
                update_village_summary_record(village, hobli, taluka, dist, searched_val=val, scraped_val=sc_val, is_completed=is_completed_bool, scraped_by=w_label, status=status_val)
                if is_comp:
                    try:
                        sync_village_summary(OUTPUT_FILE, VILLAGE_SUMMARY_JSON, VILLAGE_SUMMARY_XLSX)
                        print_live_progress_dashboard(shared_seen_docs, shared_worker_status)
                    except Exception:
                        pass
            except Exception:
                pass
            continue

        df, excel_file, value, batch_idx, label = job
        try:
            doc_col = next(
                (c for c in df.columns if any(k in str(c).lower() for k in [
                    'ದಸ್ತಾವೇಜಿನ ಉಲ್ಲೇಖ', 'ದಸ್ತಾವೇಜು', 'document reference', 'doc ref', 'registration no', 'doc no'
                ])),
                None
            )

            # Extra guard: discard if incoming df has no valid doc_col
            if not doc_col:
                log("ExcelSaver", f"⚠️ [{label}] Discarded batch without document reference column.")
                continue

            if os.path.exists(excel_file):
                workbook = load_workbook(excel_file)
                worksheet = workbook.worksheets[0]
                first_row = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
                workbook_is_empty = worksheet.max_row <= 1 and not any(first_row)

                if workbook_is_empty:
                    save_df_atomically_with_lock_wait(df, excel_file, stop_event=stop_event)
                    if shared_seen_docs is not None and doc_col:
                        for d_ref in df[doc_col].dropna().astype(str).str.strip():
                            if d_ref:
                                shared_seen_docs[d_ref] = True
                    v_name = df['Village'].iloc[0] if 'Village' in df.columns else ''
                    log("ExcelSaver", f"✅ [{label}] Initialized {os.path.basename(excel_file)} with {len(df)} records for '{v_name}' (Survey #{value})")
                    print_live_progress_dashboard(shared_seen_docs, shared_worker_status)
                    continue

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
                except Exception as read_err:
                    log("ExcelSaver", f"⚠️ Error reading {excel_file} for dedup: {read_err}")
                    new_rows = df

                if not new_rows.empty:
                    combined_df = pd.concat([existing_df, new_rows], ignore_index=True)
                    junk_cols = [c for c in combined_df.columns if any(k in str(c).lower() for k in ['enter one', 'unnamed', 'add more'])]
                    if junk_cols:
                        combined_df = combined_df.drop(columns=junk_cols)
                    save_df_atomically_with_lock_wait(combined_df, excel_file, stop_event=stop_event)
                    if shared_seen_docs is not None and doc_col:
                        for d_ref in new_rows[doc_col].dropna().astype(str).str.strip():
                            if d_ref:
                                shared_seen_docs[d_ref] = True
                    v_name = new_rows['Village'].iloc[0] if 'Village' in new_rows.columns else ''
                    h_name = new_rows['Hobli/Town'].iloc[0] if 'Hobli/Town' in new_rows.columns else ''
                    t_name = new_rows['Taluka'].iloc[0] if 'Taluka' in new_rows.columns else ''
                    d_name = new_rows['District'].iloc[0] if 'District' in new_rows.columns else ''
                    update_village_summary_record(v_name, h_name, t_name, d_name, searched_val=value, scraped_val=value, new_recs_count=len(new_rows), scraped_by=label)
                    total_docs = len(shared_seen_docs) if shared_seen_docs is not None else len(combined_df)
                    log("ExcelSaver", f"💾 [{label}] Saved {len(new_rows)} new unique records for '{v_name}' (Survey #{value}) | Total in Excel: {total_docs}")
                    print_live_progress_dashboard(shared_seen_docs, shared_worker_status)
                else:
                    log("ExcelSaver", f"ℹ️ [{label}] All {len(df)} records for Survey #{value} were duplicates; skipped.")
            else:
                save_df_atomically_with_lock_wait(df, excel_file, stop_event=stop_event)
                if shared_seen_docs is not None and doc_col:
                    for d_ref in df[doc_col].dropna().astype(str).str.strip():
                        if d_ref:
                            shared_seen_docs[d_ref] = True
                v_name = df['Village'].iloc[0] if 'Village' in df.columns else ''
                h_name = df['Hobli/Town'].iloc[0] if 'Hobli/Town' in df.columns else ''
                t_name = df['Taluka'].iloc[0] if 'Taluka' in df.columns else ''
                d_name = df['District'].iloc[0] if 'District' in df.columns else ''
                update_village_summary_record(v_name, h_name, t_name, d_name, searched_val=value, scraped_val=value, new_recs_count=len(df), scraped_by=label)
                total_docs = len(shared_seen_docs) if shared_seen_docs is not None else len(df)
                log("ExcelSaver", f"💾 [{label}] Created {os.path.basename(excel_file)} and saved {len(df)} records for '{v_name}' (Survey #{value}) | Total in Excel: {total_docs}")
                print_live_progress_dashboard(shared_seen_docs, shared_worker_status)

        except Exception as e:
            log("ExcelSaver", f"❌ Error saving data to Excel: {e}")

    log("ExcelSaver", "Excel Writer process finished.")


def check_system_dependencies():
    """Validates critical runtime dependencies and informs the user."""
    missing = []
    warnings = []
    try:
        import pandas as pd
    except ImportError:
        missing.append("pandas")
    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")
    try:
        import bs4
    except ImportError:
        missing.append("beautifulsoup4")
    try:
        import cv2
    except ImportError:
        missing.append("opencv-python")
    try:
        import PIL
    except ImportError:
        missing.append("pillow")
    try:
        import selenium
    except ImportError:
        missing.append("selenium")

    try:
        import lxml
    except ImportError:
        try:
            import html5lib
        except ImportError:
            warnings.append("Neither 'lxml' nor 'html5lib' is installed. BeautifulSoup fallback will be used for table parsing.")

    if missing:
        print("=" * 70)
        print("❌ PRE-FLIGHT ERROR: Missing required Python packages:")
        for pkg in missing:
            print(f"   - {pkg}")
        print("   Please run: pip install " + " ".join(missing))
        print("=" * 70)
        sys.exit(1)

    if warnings:
        for w in warnings:
            print(f"⚠️ [PRE-FLIGHT WARNING] {w}")


# ======================================================================================
# MAIN LAUNCHER
# ======================================================================================
if __name__ == "__main__":
    # Windows multiprocessing safety
    mp.freeze_support()

    check_system_dependencies()

    if not os.path.exists(TESSERACT_PATH):
        print("=" * 70)
        print(f"❌ PRE-FLIGHT ERROR: Tesseract OCR was not found at:")
        print(f"   {TESSERACT_PATH}")
        print("   Please install Tesseract OCR or verify its installation path.")
        print("=" * 70)
        sys.exit(1)

    if not os.path.exists(OUTPUT_FILE):
        print(f"📄 Output workbook will be created when the first result arrives: {OUTPUT_FILE}")

    print("=" * 70)
    print("🚀 STARTING KAVERI IGR MULTI-BROWSER MULTIPROCESSING SCRAPER")
    print(f"🎯 Target Hobli Option: option[{TARGET_HOBLI_OPTION_INDEX}]" if TARGET_HOBLI_OPTION_INDEX else "🎯 Target Hobli Option: ALL HOBLIS")
    print(f"👥 Total Workers Configured: {len(ACCOUNTS)}")
    for i, acc in enumerate(ACCOUNTS):
        print(f"   - Worker {i+1}: {acc['label']} (Engine: {acc.get('browser_type', 'chrome').upper()}, Account: {acc['login_email']})")
    print("=" * 70)

    # Shared manager dict to track all seen document reference numbers across all processes
    manager = mp.Manager()
    shared_seen_docs = manager.dict()
    shared_worker_status = manager.dict()

    for acc in ACCOUNTS:
        w_lbl = acc.get("label")
        shared_worker_status[w_lbl] = {
            "engine": acc.get("browser_type", "chrome").upper(),
            "state": "PENDING",
            "village": "-",
            "survey": "-",
            "empty_streak": "-",
            "records": 0,
            "detail": "Worker initialized",
        }

    # Preload existing document reference numbers from Excel if it exists
    if os.path.exists(OUTPUT_FILE):
        try:
            print(f"📖 Preloading existing document references from {OUTPUT_FILE} for instant deduplication...", flush=True)
            existing_df = pd.read_excel(OUTPUT_FILE)
            doc_col = next((c for c in existing_df.columns if any(k in str(c).lower() for k in [
                'ದಸ್ತಾವೇಜಿನ ಉಲ್ಲೇಖ', 'ದಸ್ತಾವೇಜು', 'document reference', 'doc ref', 'registration no', 'doc no'
            ])), None)
            if doc_col:
                need_save = False
                # Sanitize and drop accidental form columns/rows if any exist
                junk_cols = [c for c in existing_df.columns if any(k in str(c).lower() for k in ['enter one', 'unnamed', 'add more'])]
                if junk_cols or existing_df[doc_col].isna().any():
                    existing_df = existing_df[existing_df[doc_col].notna() & (existing_df[doc_col].astype(str).str.strip() != '')].copy()
                    if junk_cols:
                        existing_df = existing_df.drop(columns=junk_cols)
                    need_save = True

                # Standardize legacy District entries from 'Bengaluru Urban' to 'Basavanagudi'
                if 'District' in existing_df.columns and (existing_df['District'] == 'Bengaluru Urban').any():
                    existing_df['District'] = existing_df['District'].replace('Bengaluru Urban', 'Basavanagudi')
                    need_save = True

                if need_save:
                    try:
                        save_df_atomically_with_lock_wait(existing_df, OUTPUT_FILE)
                        print(f"🧹 Standardized District to 'Basavanagudi' and cleaned {os.path.basename(OUTPUT_FILE)}.", flush=True)
                    except Exception as clean_err:
                        print(f"⚠️ Could not auto-clean Excel file: {clean_err}", flush=True)

                loaded_count = 0
                for doc_ref in existing_df[doc_col].dropna().astype(str).str.strip():
                    if doc_ref and doc_ref not in shared_seen_docs:
                        shared_seen_docs[doc_ref] = True
                        loaded_count += 1
                print(f"✅ Preloaded {loaded_count} existing unique document references into shared cache.", flush=True)
        except Exception as e:
            print(f"⚠️ Could not preload existing records from {OUTPUT_FILE}: {e}", flush=True)

    # Initialize / synchronize village scraping summary tracking files
    try:
        sync_village_summary(OUTPUT_FILE, VILLAGE_SUMMARY_JSON, VILLAGE_SUMMARY_XLSX)
        print(f"📊 Village scraping summary synchronized: {os.path.basename(VILLAGE_SUMMARY_XLSX)} and {os.path.basename(VILLAGE_SUMMARY_JSON)}", flush=True)
    except Exception as sum_err:
        print(f"⚠️ Could not sync village summary: {sum_err}", flush=True)

    # Multiprocessing primitives
    data_queue = mp.Queue()
    stop_event = mp.Event()

    # Start Excel saver process
    saver_proc = mp.Process(
        target=excel_saver_process_main,
        args=(data_queue, stop_event, shared_seen_docs, shared_worker_status),
        name="ExcelSaverProcess"
    )
    saver_proc.start()

    # Dynamic village task coordination primitives
    shared_claimed_villages = manager.dict()
    shared_completed_villages = manager.dict()
    village_lock = manager.Lock()

    worker_processes = []
    total_workers = len(ACCOUNTS)

    try:
        # Launch worker processes with staggered delay
        for i, acc in enumerate(ACCOUNTS):
            stagger = i * STAGGER_DELAY_SECONDS
            p = mp.Process(
                target=worker_process_main,
                args=(
                    acc, i, total_workers, data_queue, stagger,
                    shared_seen_docs, shared_claimed_villages, shared_completed_villages, village_lock,
                    shared_worker_status
                ),
                name=f"WorkerProcess-{acc.get('label', i+1)}"
            )
            worker_processes.append(p)
            p.start()

        # Wait for all workers to complete
        for p in worker_processes:
            p.join()

    except KeyboardInterrupt:
        print("\n🛑 Keyboard interrupt received! Stopping workers...")
        for p in worker_processes:
            if p.is_alive():
                p.terminate()
            if p.pid:
                subprocess.run(
                    ["taskkill", "/PID", str(p.pid), "/T", "/F"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                )
        for p in worker_processes:
            p.join(timeout=10)

    finally:
        print("Draining data queue and terminating saver process...")
        stop_event.set()
        data_queue.put(None)
        saver_proc.join(timeout=60)
        if saver_proc.is_alive():
            print("⚠️ Excel saver did not finish within 60 seconds; terminating it.")
            saver_proc.terminate()
        print("🎉 All scraper processes have completed successfully.")
