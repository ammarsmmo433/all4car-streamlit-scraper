import os
# Generated from the proven V9 notebook scraper engine.
from pathlib import Path

# =========================
# إعدادات قابلة للتعديل
# =========================
BASE_URL = "https://all4car.com.ua"
APP_DIR = Path(os.environ.get("ALL4CAR_DATA_DIR", str(Path.cwd() / "all4car_batch_scraper")))
OUTPUT_DIR = APP_DIR / "outputs"
DATA_DIR = APP_DIR / "data"
LOG_DIR = APP_DIR / "logs"
SCREENSHOT_DIR = LOG_DIR / "screenshots"
DEBUG_HTML_DIR = LOG_DIR / "debug_pages"

MAX_RETRIES = 3
FAST_MODE = os.environ.get("ALL4CAR_FAST_MODE", "1").strip().lower() in ("1", "true", "yes", "on")
PAGE_LOAD_TIMEOUT = 75 if FAST_MODE else 180
ELEMENT_WAIT = 45 if FAST_MODE else 120
BASE_RETRY_DELAY = 2 if FAST_MODE else 5
MAX_RETRY_DELAY = 20 if FAST_MODE else 60
SITE_CHECK_INTERVAL = 45 if FAST_MODE else 120
MIN_REQUEST_DELAY = 0.15 if FAST_MODE else 0.8
MAX_REQUEST_DELAY = 4.0 if FAST_MODE else 12.0
HEADLESS = os.environ.get("ALL4CAR_HEADLESS", "0").strip().lower() in ("1", "true", "yes", "on")
SCREENSHOT_ON_ERROR = True
SAVE_DEBUG_HTML = True

# المحددات مأخوذة من النوتبوك الأصلي
SELECTORS = {
    "vin_input": "input.search-form__input",
    "main_popup": "ul._4uWvJ1pkaiA-",
    "main_popup_links": "a",
    "sub_container_item": "ul._9ikbUAgVfYQ- > li.sXbh6y72f90-",
    "sub_item": "li.sXbh6y72f90-",
    "sub_anchor": ":scope > a.OZsh-nsSg-8-",
    "sub_heading_fallback": "h2.shxLyX8OMsU-",
    "show_more": "button._710xV-kIMAg-",
    "part_row": "li.K-x4ukCMJlg-",
    "part_name": "div._0hmgoA1dDwA- > span",
    "part_description": "div.EzYOnHuko-o-",
    "part_number": "div.TMOQvY9f-1w- span",
}

for p in (OUTPUT_DIR, DATA_DIR, LOG_DIR, SCREENSHOT_DIR, DEBUG_HTML_DIR):
    p.mkdir(parents=True, exist_ok=True)

print("Workspace:", APP_DIR)


import re
import time
import json
import queue
import random
import sqlite3
import logging
import threading
import traceback
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import requests
from openpyxl import Workbook, load_workbook

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, WebDriverException, NoSuchElementException,
    StaleElementReferenceException, InvalidSessionIdException
)

# =========================================================
# Logging
# =========================================================
log_file = LOG_DIR / f"scraper_{datetime.now():%Y-%m-%d}.log"
logger = logging.getLogger("all4car")
logger.setLevel(logging.INFO)
if not logger.handlers:
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
    logger.addHandler(fh)

# =========================================================
# Helpers
# =========================================================
def now_iso():
    return datetime.now().isoformat(timespec="seconds")

def normalize_text(text):
    text = unicodedata.normalize("NFKC", text or "")
    text = re.sub(r"\s+", " ", text).strip().casefold()
    return text

def safe_filename(text):
    text = re.sub(r'[<>:"/\\|?*]+', '_', str(text)).strip().strip('.')
    return re.sub(r"\s+", "_", text)[:120] or "unnamed"

def dedupe_keep_order(values):
    seen, out = set(), []
    for value in values:
        v = value.strip()
        if not v:
            continue
        key = v.casefold()
        if key not in seen:
            seen.add(key)
            out.append(v)
    return out

def backoff_delay(attempt):
    base = min(MAX_RETRY_DELAY, BASE_RETRY_DELAY * (2 ** max(0, attempt - 1)))
    return min(MAX_RETRY_DELAY, base + random.uniform(0, max(1, base * 0.35)))

# =========================================================
# SQLite checkpoint
# =========================================================
class CheckpointDB:
    def __init__(self, path):
        self.path = str(path)
        self.lock = threading.RLock()
        self._init_db()

    def connect(self):
        con = sqlite3.connect(self.path, timeout=30)
        con.row_factory = sqlite3.Row
        con.execute("PRAGMA journal_mode=WAL")
        con.execute("PRAGMA synchronous=NORMAL")
        return con

    def _init_db(self):
        with self.lock, self.connect() as con:
            con.executescript("""
            CREATE TABLE IF NOT EXISTS batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                status TEXT NOT NULL,
                categories_json TEXT NOT NULL,
                output_dir TEXT NOT NULL,
                force_rescrape INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS batch_vins (
                batch_id INTEGER NOT NULL,
                vin TEXT NOT NULL,
                position INTEGER NOT NULL,
                status TEXT NOT NULL DEFAULT 'PENDING',
                last_error TEXT,
                updated_at TEXT NOT NULL,
                PRIMARY KEY(batch_id, vin)
            );
            CREATE TABLE IF NOT EXISTS categories (
                batch_id INTEGER NOT NULL,
                vin TEXT NOT NULL,
                category_name TEXT NOT NULL,
                category_url TEXT,
                status TEXT NOT NULL DEFAULT 'PENDING',
                last_error TEXT,
                updated_at TEXT NOT NULL,
                PRIMARY KEY(batch_id, vin, category_name)
            );
            CREATE TABLE IF NOT EXISTS sub_links (
                batch_id INTEGER NOT NULL,
                vin TEXT NOT NULL,
                category_name TEXT NOT NULL,
                sub_name TEXT NOT NULL,
                sub_url TEXT NOT NULL,
                position INTEGER NOT NULL,
                status TEXT NOT NULL DEFAULT 'PENDING',
                retry_count INTEGER NOT NULL DEFAULT 0,
                last_error TEXT,
                updated_at TEXT NOT NULL,
                PRIMARY KEY(batch_id, vin, category_name, sub_url)
            );
            CREATE TABLE IF NOT EXISTS completed_global (
                vin TEXT NOT NULL,
                category_name TEXT NOT NULL,
                completed_at TEXT NOT NULL,
                PRIMARY KEY(vin, category_name)
            );
            CREATE TABLE IF NOT EXISTS errors (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id INTEGER,
                vin TEXT,
                category_name TEXT,
                sub_url TEXT,
                error_type TEXT,
                message TEXT,
                created_at TEXT NOT NULL
            );
            """)

    def create_batch(self, vins, categories, output_dir, force_rescrape=False):
        ts = now_iso()
        with self.lock, self.connect() as con:
            cur = con.execute(
                "INSERT INTO batches(created_at,updated_at,status,categories_json,output_dir,force_rescrape) VALUES(?,?,?,?,?,?)",
                (ts, ts, "PENDING", json.dumps(categories, ensure_ascii=False), str(output_dir), int(force_rescrape))
            )
            batch_id = cur.lastrowid
            con.executemany(
                "INSERT INTO batch_vins(batch_id,vin,position,status,updated_at) VALUES(?,?,?,?,?)",
                [(batch_id, vin, i+1, "PENDING", ts) for i, vin in enumerate(vins)]
            )
        return batch_id

    def latest_unfinished_batch(self):
        with self.connect() as con:
            return con.execute(
                "SELECT * FROM batches WHERE status NOT IN ('COMPLETED','STOPPED') ORDER BY id DESC LIMIT 1"
            ).fetchone()

    def batch(self, batch_id):
        with self.connect() as con:
            return con.execute("SELECT * FROM batches WHERE id=?", (batch_id,)).fetchone()

    def vins(self, batch_id):
        with self.connect() as con:
            return con.execute("SELECT * FROM batch_vins WHERE batch_id=? ORDER BY position", (batch_id,)).fetchall()

    def set_batch_status(self, batch_id, status):
        with self.lock, self.connect() as con:
            con.execute("UPDATE batches SET status=?,updated_at=? WHERE id=?", (status, now_iso(), batch_id))

    def set_vin_status(self, batch_id, vin, status, error=None):
        with self.lock, self.connect() as con:
            con.execute("UPDATE batch_vins SET status=?,last_error=?,updated_at=? WHERE batch_id=? AND vin=?",
                        (status, error, now_iso(), batch_id, vin))

    def upsert_category(self, batch_id, vin, name, url=None, status="PENDING", error=None):
        with self.lock, self.connect() as con:
            con.execute("""
            INSERT INTO categories(batch_id,vin,category_name,category_url,status,last_error,updated_at)
            VALUES(?,?,?,?,?,?,?)
            ON CONFLICT(batch_id,vin,category_name) DO UPDATE SET
                category_url=COALESCE(excluded.category_url,categories.category_url),
                status=excluded.status,last_error=excluded.last_error,updated_at=excluded.updated_at
            """, (batch_id, vin, name, url, status, error, now_iso()))

    def category_status(self, batch_id, vin, name):
        with self.connect() as con:
            row = con.execute("SELECT status FROM categories WHERE batch_id=? AND vin=? AND category_name=?",
                              (batch_id, vin, name)).fetchone()
            return row[0] if row else None

    def globally_completed(self, vin, category):
        with self.connect() as con:
            return con.execute("SELECT 1 FROM completed_global WHERE vin=? AND category_name=?",
                               (vin, category)).fetchone() is not None

    def mark_global_completed(self, vin, category):
        with self.lock, self.connect() as con:
            con.execute("INSERT OR REPLACE INTO completed_global(vin,category_name,completed_at) VALUES(?,?,?)",
                        (vin, category, now_iso()))

    def register_sub_links(self, batch_id, vin, category, links):
        with self.lock, self.connect() as con:
            for pos, (name, url) in enumerate(links, 1):
                con.execute("""
                INSERT OR IGNORE INTO sub_links(batch_id,vin,category_name,sub_name,sub_url,position,status,updated_at)
                VALUES(?,?,?,?,?,?,?,?)
                """, (batch_id, vin, category, name, url, pos, "PENDING", now_iso()))

    def pending_sub_links(self, batch_id, vin, category):
        with self.connect() as con:
            return con.execute("""
                SELECT * FROM sub_links
                WHERE batch_id=? AND vin=? AND category_name=? AND status!='COMPLETED'
                ORDER BY position
            """, (batch_id, vin, category)).fetchall()

    def set_sub_status(self, batch_id, vin, category, url, status, error=None, retry_count=None):
        with self.lock, self.connect() as con:
            if retry_count is None:
                con.execute("""UPDATE sub_links SET status=?,last_error=?,updated_at=?
                               WHERE batch_id=? AND vin=? AND category_name=? AND sub_url=?""",
                            (status, error, now_iso(), batch_id, vin, category, url))
            else:
                con.execute("""UPDATE sub_links SET status=?,last_error=?,retry_count=?,updated_at=?
                               WHERE batch_id=? AND vin=? AND category_name=? AND sub_url=?""",
                            (status, error, retry_count, now_iso(), batch_id, vin, category, url))

    def log_error(self, batch_id, vin, category, sub_url, error_type, message):
        with self.lock, self.connect() as con:
            con.execute("""INSERT INTO errors(batch_id,vin,category_name,sub_url,error_type,message,created_at)
                           VALUES(?,?,?,?,?,?,?)""",
                        (batch_id, vin, category, sub_url, error_type, str(message)[:4000], now_iso()))

DB = CheckpointDB(DATA_DIR / "scraper_state.db")

# =========================================================
# Atomic Excel writer
# =========================================================
class ExcelWriter:
    HEADERS = ["VIN", "Main Category", "Sub Category", "Name", "Description", "OEM", "Source URL", "Scraped At"]

    def __init__(self, root_dir):
        self.root_dir = Path(root_dir)
        self.lock = threading.RLock()

    def path_for(self, vin, category):
        vin_dir = self.root_dir / safe_filename(vin)
        vin_dir.mkdir(parents=True, exist_ok=True)
        return vin_dir / f"{safe_filename(vin)}_{safe_filename(category)}.xlsx"

    def _load_keys(self, path):
        keys = set()
        if not path.exists():
            return keys
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            for row in wb.active.iter_rows(min_row=2, values_only=True):
                if len(row) >= 7:
                    keys.add((str(row[0] or ''), str(row[1] or ''), str(row[2] or ''),
                              str(row[5] or ''), str(row[6] or '')))
        finally:
            wb.close()
        return keys

    @staticmethod
    def _cleanup_tmp(tmp):
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass

    def reset_category_file(self, vin, category, retries=10, delay=2.0):
        path = self.path_for(vin, category)
        tmp = path.with_suffix(".tmp.xlsx")
        with self.lock:
            self._cleanup_tmp(tmp)
            if not path.exists():
                return
            last = None
            for attempt in range(1, retries + 1):
                try:
                    path.unlink()
                    return
                except (PermissionError, OSError) as e:
                    last = e
                    if isinstance(e, OSError) and not isinstance(e, PermissionError) and getattr(e, "winerror", None) != 32:
                        raise
                    if attempt < retries:
                        time.sleep(delay)
            raise ExcelFileLocked(
                f"ملف Excel مفتوح أو محجوز: {path}. أغلقه في Excel ثم أعد المحاولة."
            ) from last

    def append_subpage_atomic(self, vin, category, sub_name, sub_url, rows,
                              replace_retries=10, replace_delay=2.0):
        path = self.path_for(vin, category)
        tmp = path.with_suffix(".tmp.xlsx")
        with self.lock:
            self._cleanup_tmp(tmp)

            if path.exists():
                existing = self._load_keys(path)
                wb = load_workbook(path)
                ws = wb.active
            else:
                existing = set()
                wb = Workbook()
                ws = wb.active
                ws.title = "Parts"
                ws.append(self.HEADERS)

            added = 0
            ts = now_iso()
            for item in rows:
                key = (vin, category, sub_name, item.get("oem", ""), sub_url)
                if key in existing:
                    continue
                ws.append([vin, category, sub_name, item.get("name", ""),
                           item.get("description", ""), item.get("oem", ""),
                           sub_url, ts])
                existing.add(key)
                added += 1

            # Save and CLOSE before replacing destination.
            try:
                wb.save(tmp)
            finally:
                wb.close()

            last = None
            for attempt in range(1, replace_retries + 1):
                try:
                    os.replace(tmp, path)
                    return added, path
                except (PermissionError, OSError) as e:
                    last = e
                    if isinstance(e, OSError) and not isinstance(e, PermissionError) and getattr(e, "winerror", None) != 32:
                        self._cleanup_tmp(tmp)
                        raise
                    if attempt < replace_retries:
                        time.sleep(replace_delay)

            self._cleanup_tmp(tmp)
            raise ExcelFileLocked(
                f"تعذر تحديث {path.name} بعد {replace_retries} محاولات لأن الملف مفتوح أو محجوز. "
                "أغلقه في Excel ثم استخدم Resume Previous Batch."
            ) from last

# =========================================================
# Browser + scraper
# =========================================================
class ExcelFileLocked(Exception): pass
class SiteUnavailable(Exception): pass
class VinNotFound(Exception): pass
class CategoryNotFound(Exception): pass
class StopRequested(Exception): pass

class All4CarScraper:
    def __init__(self, db, output_dir, event_queue=None):
        self.db = db
        self.writer = ExcelWriter(output_dir)
        self.q = event_queue or queue.Queue()
        self.driver = None
        self.wait = None
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set()
        self.current_delay = MIN_REQUEST_DELAY
        self.driver_path = None

    def emit(self, kind="log", **payload):
        payload["kind"] = kind
        self.q.put(payload)
        msg = payload.get("message")
        if msg: logger.info(msg)

    def controlled_sleep(self, seconds):
        end = time.time() + seconds
        while time.time() < end:
            if self.stop_event.is_set(): raise StopRequested()
            self.pause_event.wait(0.25)
            time.sleep(min(0.25, max(0, end-time.time())))

    def checkpoint_gate(self):
        if self.stop_event.is_set(): raise StopRequested()
        while not self.pause_event.is_set():
            if self.stop_event.is_set(): raise StopRequested()
            time.sleep(0.25)

    def create_driver(self):
        """Create Chrome/Chromium driver for Windows local use or Linux hosting."""
        options = Options()
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-gpu")
        if HEADLESS:
            options.add_argument("--headless=new")

        # Docker/Render installs Chromium in these common locations.
        linux_chrome = None
        for candidate in ("/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome"):
            if os.path.exists(candidate):
                linux_chrome = candidate
                break
        if linux_chrome:
            options.binary_location = linux_chrome

        chromedriver = "/usr/bin/chromedriver"
        if os.path.exists(chromedriver):
            driver = webdriver.Chrome(service=Service(chromedriver), options=options)
        else:
            # On Windows/macOS Selenium Manager resolves the compatible driver.
            driver = webdriver.Chrome(options=options)

        driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
        self.driver = driver
        self.wait = WebDriverWait(driver, ELEMENT_WAIT)
        return driver

    def close_driver(self):
        try:
            if self.driver: self.driver.quit()
        except Exception:
            pass
        self.driver = self.wait = None

    def recover_driver(self):
        self.emit(message="إعادة تشغيل Chrome مع الحفاظ على الـ checkpoint...")
        self.close_driver()
        self.controlled_sleep(3)
        self.create_driver()

    def internet_available(self):
        try:
            requests.get("https://www.google.com/generate_204", timeout=10)
            return True
        except Exception:
            return False

    def safe_get(self, url, batch_id=None, vin=None, category=None, sub_url=None):
        last = None
        for attempt in range(1, MAX_RETRIES + 1):
            self.checkpoint_gate()
            try:
                t0 = time.perf_counter()
                self.emit(message=f"فتح: {url}")
                self.driver.get(url)
                # Fast mode uses only a tiny settling delay; explicit element waits do the real synchronization.
                self.controlled_sleep(self.current_delay)
                self.emit(message=f"V8 Fast: فتح الصفحة خلال {time.perf_counter()-t0:.2f}s")
                return True
            except (InvalidSessionIdException, WebDriverException) as e:
                last = e
                if attempt < MAX_RETRIES:
                    self.recover_driver()
            except Exception as e:
                last = e
            delay = backoff_delay(attempt)
            self.current_delay = min(MAX_REQUEST_DELAY, max(self.current_delay, delay / 5))
            self.emit(message=f"فشل فتح الصفحة - محاولة {attempt}/{MAX_RETRIES}. انتظار {delay:.1f}s")
            if attempt < MAX_RETRIES:
                self.controlled_sleep(delay)
        self.capture_debug(vin, category, "safe_get_failed")
        if batch_id:
            self.db.log_error(batch_id, vin, category, sub_url, "PAGE_LOAD_ERROR", last)
        raise SiteUnavailable(str(last))

    def safe_wait(self, selector, by=By.CSS_SELECTOR, timeout=None):
        timeout = timeout or ELEMENT_WAIT
        last = None
        for attempt in range(1, MAX_RETRIES + 1):
            self.checkpoint_gate()
            try:
                return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((by, selector)))
            except Exception as e:
                last = e
                if attempt < MAX_RETRIES:
                    self.controlled_sleep(backoff_delay(attempt))
        raise TimeoutException(f"Element not found: {selector}; {last}")

    def safe_click(self, element):
        last = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
                self.controlled_sleep(0.08 if FAST_MODE else 0.4)
                self.driver.execute_script("arguments[0].click();", element)
                return True
            except Exception as e:
                last = e
                if attempt < MAX_RETRIES: self.controlled_sleep(backoff_delay(attempt))
        raise WebDriverException(f"Click failed: {last}")

    def capture_debug(self, vin=None, category=None, label="error"):
        if not self.driver: return
        stem = safe_filename(f"{vin or 'NO_VIN'}_{category or 'NO_CATEGORY'}_{label}_{datetime.now():%Y%m%d_%H%M%S}")
        try:
            if SCREENSHOT_ON_ERROR:
                self.driver.save_screenshot(str(SCREENSHOT_DIR / f"{stem}.png"))
            if SAVE_DEBUG_HTML:
                (DEBUG_HTML_DIR / f"{stem}.html").write_text(self.driver.page_source, encoding="utf-8")
        except Exception:
            logger.exception("Failed to save debug artifact")

    def is_site_healthy(self):
        """فحص صحة الموقع بدون الاعتماد على CSS الديناميكي.
        صفحات /cats وواجهات SPA قد لا تحتوي vin_input أو selectors القديمة رغم أن الموقع سليم.
        """
        if not self.driver:
            return False
        try:
            src = (self.driver.page_source or "").casefold()
            title = (self.driver.title or "").casefold()
            bad_markers = [
                "too many requests", "429 too many", "captcha", "cloudflare",
                "access denied", "service unavailable", "temporarily unavailable",
                "bad gateway", "gateway timeout"
            ]
            if any(x in src or x in title for x in bad_markers):
                return False
            parsed = urlparse(self.driver.current_url)
            if "all4car.com.ua" not in (parsed.netloc or ""):
                return False
            try:
                body = self.driver.find_element(By.TAG_NAME, "body")
                body_text = (body.text or "").strip()
            except Exception:
                body_text = ""
            # SPA سليمة إذا كانت على نطاق All4Car ولها DOM/نص فعلي، حتى لو تغيرت classes.
            return len(src) > 800 and (len(body_text) > 20 or "catalog" in src or "categories" in src)
        except Exception:
            return False

    def wait_until_site_returns(self, batch_id, vin=None, category=None):
        self.db.set_batch_status(batch_id, "PAUSED")
        self.emit(kind="site", status="WAITING", message="الموقع غير متاح. تم إيقاف الـ Batch مؤقتًا مع حفظ نقطة التوقف.")
        while not self.stop_event.is_set():
            self.pause_event.wait()
            try:
                if self.driver is None:
                    self.create_driver()
                self.driver.get(BASE_URL)
                # لا نعتمد على طول HTML؛ ننتظر حقل VIN نفسه
                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, SELECTORS["vin_input"]))
                )
                if self.is_site_healthy():
                    self.current_delay = max(MIN_REQUEST_DELAY, self.current_delay * 0.9)
                    self.db.set_batch_status(batch_id, "IN_PROGRESS")
                    self.emit(kind="site", status="ONLINE", message="عاد الموقع للعمل. متابعة التنفيذ من الـ checkpoint.")
                    return
            except Exception:
                pass
            self.emit(kind="site", status="WAITING", message=f"الموقع ما زال غير متاح. فحص جديد بعد {SITE_CHECK_INTERVAL} ثانية.")
            self.controlled_sleep(SITE_CHECK_INTERVAL)
        raise StopRequested()

    def _wait_for_home(self, batch_id=None, vin=None):
        """فتح الصفحة الرئيسية وانتظار حقل VIN بدل الانتظار الأعمى."""
        self.safe_get(BASE_URL, batch_id, vin)
        try:
            return WebDriverWait(self.driver, 35).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, SELECTORS["vin_input"]))
            )
        except TimeoutException as e:
            self.capture_debug(vin, None, "home_vin_input_missing")
            raise SiteUnavailable("All4Car opened but VIN search input did not become ready") from e

    def _category_candidates_from_popup(self):
        """إرجاع عناصر popup حتى لو لم يكن لها href. هذه هي طريقة النسخة القديمة التي كانت تنجح بالضغط."""
        try:
            popup = WebDriverWait(self.driver, 35).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, SELECTORS["main_popup"]))
            )
            return popup.find_elements(By.CSS_SELECTOR, "a, button, [role='button'], li")
        except TimeoutException:
            return []

    def _find_category_element_by_text(self, category_name):
        """
        V2 Web: wait for current All4Car VIN results and find a Main Category
        specifically among #/schemas links. This mirrors the successful desktop flow.
        """
        target = normalize_text(category_name)

        # VIN results are rendered asynchronously; wait for real /schemas links,
        # not merely for the input/page to exist.
        try:
            WebDriverWait(self.driver, 60).until(
                lambda d: len(d.find_elements(
                    By.XPATH,
                    "//a[contains(@href,'#/schemas?') or contains(@href,'/schemas?')]"
                )) > 0
            )
        except TimeoutException:
            pass

        candidates = self.driver.find_elements(
            By.XPATH,
            "//a[contains(@href,'#/schemas?') or contains(@href,'/schemas?')]"
        )
        self.emit(message=f"V3 Web: عدد روابط Main Category الظاهرة = {len(candidates)}")

        # exact visible-text match first
        for el in candidates:
            try:
                txt = normalize_text(el.text)
                if txt == target:
                    return (el.text or category_name).strip(), el
            except Exception:
                continue

        # then unique partial match
        partial = []
        for el in candidates:
            try:
                txt = normalize_text(el.text)
                if target and target in txt:
                    partial.append(el)
            except Exception:
                continue
        if len(partial) == 1:
            el = partial[0]
            return (el.text or category_name).strip(), el

        # retain the older popup/DOM fallback as a last resort
        for el in self._category_candidates_from_popup():
            try:
                txt = normalize_text(el.text)
                if txt == target:
                    return (el.text or category_name).strip(), el
            except Exception:
                continue

        raise CategoryNotFound(
            f"Main Category not found for requested name: {category_name}"
        )

    def open_category_by_name(self, batch_id, vin, requested):
        """
        يبحث عن VIN ثم يفتح Main Category بالاسم.
        لا يحتاج href: إذا كان href موجودًا يستخدمه، وإلا يضغط العنصر كما في النسخة القديمة.
        يعاد تنفيذ البحث لكل Category، لذلك يعمل أيضًا عند وجود أكثر من مجموعة لنفس VIN.
        """
        inp = self._wait_for_home(batch_id, vin)
        inp.clear()
        inp.send_keys(vin)
        inp.send_keys(Keys.ENTER)
        # Wait for category links instead of always sleeping two seconds.
        try:
            WebDriverWait(self.driver, 12 if FAST_MODE else 30, poll_frequency=0.15 if FAST_MODE else 0.5).until(
                lambda d: len(d.find_elements(By.XPATH, "//a[contains(@href,'#/schemas?') or contains(@href,'/schemas?')]")) > 0
            )
        except TimeoutException:
            pass
        self.emit(message=f"البحث عن المجموعة بالاسم: {requested}")

        try:
            actual_name, el = self._find_category_element_by_text(requested)
        except CategoryNotFound as e:
            self.capture_debug(vin, requested, "category_not_found")
            # إذا حقل VIN ما زال موجودًا فالموقع نفسه سليم؛ المشكلة VIN/category وليست outage.
            if self.is_site_healthy():
                raise
            raise SiteUnavailable("Category lookup failed while site/page was unhealthy") from e

        href = None
        try:
            href = el.get_attribute("href")
        except Exception:
            pass

        old_url = self.driver.current_url
        self.emit(message=f"تم العثور على المجموعة: {actual_name} | {'href موجود' if href else 'سيتم الضغط عليها مباشرة'}")

        try:
            if href and href.lower().startswith(("http://", "https://")):
                self.safe_get(href, batch_id, vin, requested)
            else:
                self.safe_click(el)

            # النجاح لا يعتمد فقط على تغير URL؛ بعض المواقع SPA.
            WebDriverWait(self.driver, 20 if FAST_MODE else 45, poll_frequency=0.15 if FAST_MODE else 0.5).until(
                lambda d: (
                    d.current_url != old_url
                    or len(d.find_elements(By.CSS_SELECTOR, SELECTORS["sub_item"])) > 0
                    or len(d.find_elements(By.CSS_SELECTOR, SELECTORS["sub_heading_fallback"])) > 0
                    or len(d.find_elements(By.CSS_SELECTOR, SELECTORS["show_more"])) > 0
                )
            )
        except TimeoutException as e:
            self.capture_debug(vin, requested, "category_click_no_result")
            raise SiteUnavailable(f"Found category '{actual_name}' but click/navigation did not open its sub-items") from e

        final_url = self.driver.current_url
        self.emit(message=f"تم فتح المجموعة بنجاح: {actual_name} | URL: {final_url}")
        return actual_name, final_url

    # أبقينا هاتين الدالتين للتوافق مع أي استدعاء قديم، لكن التنفيذ الجديد يستخدم open_category_by_name.
    def search_vin_and_get_categories(self, batch_id, vin):
        inp = self._wait_for_home(batch_id, vin)
        inp.clear(); inp.send_keys(vin); inp.send_keys(Keys.ENTER)
        candidates = self._category_candidates_from_popup()
        categories = []
        for el in candidates:
            try:
                name = (el.text or "").strip()
                href = el.get_attribute("href")
                if name:
                    categories.append((name, href or ""))
            except StaleElementReferenceException:
                continue
        if not categories:
            self.capture_debug(vin, None, "vin_no_categories")
            if self.is_site_healthy():
                raise VinNotFound(f"VIN not found or no categories appeared: {vin}")
            raise SiteUnavailable("VIN result page is unhealthy")
        return categories

    def find_category(self, available, requested):
        target = normalize_text(requested)
        exact = [(n,u) for n,u in available if normalize_text(n) == target]
        if exact:
            return exact[0]
        partial = [(n,u) for n,u in available if target in normalize_text(n) or normalize_text(n) in target]
        if len(partial) == 1:
            return partial[0]
        raise CategoryNotFound(requested)

    def _generic_sub_link_candidates(self, vin=None):
        """
        V5: اكتشاف روابط العناصر الفرعية الحقيقية. All4Car يستخدم #/parts? مع groupId للصفحة النهائية.
        V3 كان يقبل أي رابط /cats/ تقريبًا، ولذلك جمع مئات روابط التنقل العامة.
        هنا نقبل فقط route ينتقل من قائمة schemas إلى schema مفردة أو رابط يحمل schema id.
        """
        current = self.driver.current_url
        current_norm = current.rstrip("/")
        out, seen = [], set()

        try:
            anchors = self.driver.find_elements(By.XPATH, "//a[@href]")
        except Exception:
            anchors = []

        for a in anchors:
            try:
                href = (a.get_attribute("href") or "").strip()
                if not href or href.startswith(("javascript:", "mailto:", "tel:")):
                    continue

                p = urlparse(href)
                if "all4car.com.ua" not in (p.netloc or ""):
                    continue
                if href.rstrip("/") == current_norm:
                    continue

                low = href.casefold()
                fragment = (p.fragment or "").casefold()

                # الرابط الحقيقي للعنصر/المخطط يجب أن يكون أعمق من صفحة قائمة schemas.
                # أمثلة مقبولة: #/parts?...&groupId=... (المسار الفعلي في All4Car)، مع إبقاء دعم schema القديم.
                is_single_schema_route = (
                    fragment.startswith("/parts?") or
                    fragment.startswith("parts?") or
                    fragment.startswith("/schema?") or
                    fragment.startswith("schema?") or
                    "groupid=" in low or
                    "schemaid=" in low or
                    "illustrationid=" in low or
                    "imageid=" in low
                )

                # استبعاد صفحة المجموعة نفسها (#/schemas?) وروابط التنقل العامة.
                is_schema_list = (
                    fragment.startswith("/schemas?") or
                    fragment.startswith("schemas?")
                )

                if is_schema_list or not is_single_schema_route:
                    continue

                # حافظ على سياق VIN إن كان موجودًا في الرابط.
                if vin and "article=" in low and vin.casefold() not in low:
                    continue

                txt = (
                    a.text
                    or a.get_attribute("title")
                    or a.get_attribute("aria-label")
                    or ""
                ).strip()

                if href not in seen:
                    seen.add(href)
                    out.append((txt or "Catalog diagram", href))

            except (StaleElementReferenceException, WebDriverException):
                continue

        return out

    def expand_all_items(self, vin=None):
        """
        Web V4:
        حمّل جميع بطاقات Sub Group قبل جمع الروابط.

        لا نعتمد على class قديم لزر Show more.
        نراقب العدد الحقيقي:
            a[data-test-id='parts-link']

        وفي كل دورة:
        1) نمرر إلى آخر بطاقة/أسفل الصفحة لتحفيز lazy loading.
        2) نبحث عن أي زر ظاهر يحمل Show more / Load more / More / عرض المزيد.
        3) نضغطه إن وجد.
        4) ننتظر زيادة عدد parts-link.
        5) لا نتوقف إلا بعد عدة دورات مستقرة بدون زر وبدون زيادة.
        """
        self.emit(message="V4 Web: بدء توسيع جميع Sub Groups قبل جمع الروابط...")

        selector = "a[data-test-id='parts-link']"
        last_count = -1
        stable_rounds = 0
        clicks = 0

        for round_no in range(1, 61):
            self.checkpoint_gate()

            cards = self.driver.find_elements(By.CSS_SELECTOR, selector)
            before = len(cards)

            # Scroll to the actual last card first; this is important for React/lazy lists.
            if cards:
                try:
                    self.driver.execute_script(
                        "arguments[0].scrollIntoView({block:'end'});", cards[-1]
                    )
                except Exception:
                    pass

            try:
                self.driver.execute_script(
                    "window.scrollTo(0, document.body.scrollHeight);"
                )
            except Exception:
                pass

            self.controlled_sleep(0.18 if FAST_MODE else 0.6)

            # Do not depend on generated CSS classes. Search visible buttons by meaning.
            clicked = False
            candidates = self.driver.find_elements(
                By.XPATH,
                "//button | //*[@role='button']"
            )
            for btn in candidates:
                try:
                    if not btn.is_displayed() or not btn.is_enabled():
                        continue

                    txt = normalize_text(
                        (btn.text or "")
                        + " "
                        + (btn.get_attribute("aria-label") or "")
                        + " "
                        + (btn.get_attribute("title") or "")
                    )

                    show_more_words = (
                        "show more",
                        "load more",
                        "more",
                        "show all",
                        "load all",
                        "عرض المزيد",
                        "إظهار المزيد",
                    )
                    if any(word in txt for word in show_more_words):
                        self.driver.execute_script(
                            "arguments[0].scrollIntoView({block:'center'});", btn
                        )
                        self.controlled_sleep(0.05 if FAST_MODE else 0.2)
                        self.driver.execute_script("arguments[0].click();", btn)
                        clicks += 1
                        clicked = True
                        self.emit(
                            message=f"V4 Web: ضغط زر التوسيع #{clicks} | البطاقات قبل الضغط={before}"
                        )
                        break
                except (StaleElementReferenceException, WebDriverException):
                    continue

            # Wait briefly for either lazy-load or button expansion to add cards.
            deadline = time.time() + (1.5 if FAST_MODE else 4.0)
            after = before
            while time.time() < deadline:
                self.checkpoint_gate()
                after = len(self.driver.find_elements(By.CSS_SELECTOR, selector))
                if after > before:
                    break
                time.sleep(0.10 if FAST_MODE else 0.25)

            self.emit(
                message=f"V4 Web: دورة التوسيع {round_no} | Sub Groups: {before} -> {after}"
            )

            if after > before:
                stable_rounds = 0
            elif clicked:
                # Button was clicked but rendering may need another round.
                stable_rounds = 0
            else:
                stable_rounds += 1

            # Three full stable rounds protects against delayed/lazy rendering.
            if stable_rounds >= (2 if FAST_MODE else 3):
                break

            last_count = after

        final_count = len(self.driver.find_elements(By.CSS_SELECTOR, selector))
        try:
            self.driver.execute_script("window.scrollTo(0, 0);")
        except Exception:
            pass

        self.emit(
            message=f"V4 Web: انتهى التوسيع | العدد النهائي للمجموعات={final_count} | ضغطات التوسيع={clicks}"
        )

    def collect_sub_links(self, vin=None):
        """
        V9 - يعتمد على البنية الفعلية الحالية للموقع:
        كل بطاقة Sub Group لها رابط رئيسي واحد فقط:
            a[data-test-id='parts-link']
        الروابط الموجودة داخل البطاقة والتي تحتوي partNameId هي روابط قطع/فلاتر
        وليست Sub Groups، لذلك لا نجمعها إطلاقًا.
        """
        self.emit(message="V4 Web: جمع جميع الروابط بعد اكتمال التوسيع من data-test-id=parts-link...")

        try:
            WebDriverWait(self.driver, 6 if FAST_MODE else 15, poll_frequency=0.15 if FAST_MODE else 0.5).until(
                lambda d: len(d.find_elements(
                    By.CSS_SELECTOR, "a[data-test-id='parts-link']"
                )) > 0
            )
        except TimeoutException:
            pass

        anchors = self.driver.find_elements(
            By.CSS_SELECTOR, "a[data-test-id='parts-link']"
        )
        self.emit(message=f"V9: عدد بطاقات Sub Group الظاهرة = {len(anchors)}")

        out = []
        seen_group_ids = set()

        for a in anchors:
            try:
                href = (a.get_attribute("href") or "").strip()
                if not href or "groupId=" not in href:
                    continue

                # رابط البطاقة الرئيسي يجب ألا يحوي partNameId.
                # حتى لو تغير DOM مستقبلًا، ننظفه احتياطيًا.
                href = re.sub(r"([&?])partNameId=[^&#]*", "", href, flags=re.I).rstrip("&?")

                m = re.search(r"(?:[?&])groupId=([^&#]+)", href, flags=re.I)
                if not m:
                    continue
                group_id = m.group(1)

                if group_id in seen_group_ids:
                    continue

                # الاسم الحقيقي للبطاقة: h2 > span، ثم alt للصورة كـ fallback.
                name = ""
                try:
                    name = a.find_element(By.CSS_SELECTOR, "h2 span").text.strip()
                except Exception:
                    pass
                if not name:
                    try:
                        name = (a.find_element(By.CSS_SELECTOR, "img").get_attribute("alt") or "").strip()
                    except Exception:
                        pass
                if not name:
                    name = f"Parts group {len(out)+1}"

                seen_group_ids.add(group_id)
                out.append((name, href))

            except (StaleElementReferenceException, WebDriverException):
                continue

        self.emit(message=f"V9: المجموعات الفرعية الفريدة = {len(out)}")
        if out:
            for i, (name, url) in enumerate(out[:5], 1):
                self.emit(message=f"V9 sample {i}: {name} | groupId={re.search(r'groupId=([^&#]+)', url).group(1)[:35]}...")

        return out

    def extract_rows(self):
        """
        V9 - استخراج جدول القطع من DOM الحالي الذي زوده المستخدم.

        الصف:
            li[data-part-expand]

        داخل كل صف:
            [data-test-id='expansion']

        رقم OEM:
            span السابق مباشرة لزر data-test-id='copy-button'

        الاسم والوصف:
            نصل إلى الحاوية التي تحتوي OEM ثم نقرأ أبناء div المباشرين:
            أول div = Part Name
            آخر div يحوي copy-button = OEM block
            ما بينهما = Description (إن وجد)

        بهذه الطريقة لا نعتمد على أسماء CSS العشوائية المتغيرة.
        """
        rows = self.driver.find_elements(By.CSS_SELECTOR, "li[data-part-expand]")
        self.emit(message=f"V9: عدد صفوف القطع في الجدول = {len(rows)}")

        data = []
        seen = set()

        for row in rows:
            try:
                expansion = row.find_element(
                    By.CSS_SELECTOR, "[data-test-id='expansion']"
                )

                copy_btn = expansion.find_element(
                    By.CSS_SELECTOR, "button[data-test-id='copy-button']"
                )

                oem_block = copy_btn.find_element(By.XPATH, "..")
                oem = ""
                try:
                    oem = oem_block.find_element(
                        By.XPATH, "./span[1]"
                    ).text.strip()
                except Exception:
                    pass

                # parent of OEM block is the semantic text container.
                text_container = oem_block.find_element(By.XPATH, "..")
                children = text_container.find_elements(By.XPATH, "./div")

                name = ""
                desc = ""

                semantic_children = []
                for child in children:
                    try:
                        # OEM block itself is not name/description.
                        if child.find_elements(
                            By.CSS_SELECTOR, "button[data-test-id='copy-button']"
                        ):
                            continue
                        txt = child.text.strip()
                        if txt:
                            semantic_children.append(txt)
                    except Exception:
                        continue

                if semantic_children:
                    name = semantic_children[0]
                    if len(semantic_children) > 1:
                        desc = " | ".join(semantic_children[1:])

                # fallback للاسم: من data-part-expand إذا فشل النص، لكن لا نخمن الوصف.
                if not oem:
                    # OEM يمكن استعادته من أول جزء من data-part-expand في بعض الصفوف،
                    # لكن نستخدمه فقط إذا لم نجد span الحقيقي.
                    raw_key = (row.get_attribute("data-part-expand") or "").strip()
                    if "__" in raw_key:
                        candidate = raw_key.split("__", 1)[0].strip()
                        if candidate:
                            oem = candidate

                if not name and not desc and not oem:
                    continue

                # المفتاح الصحيح للصف هو OEM + الاسم + الوصف.
                # لا نكرر نفس القطعة داخل نفس صفحة الرسم.
                key = (oem.casefold(), name.casefold(), desc.casefold())
                if key in seen:
                    continue
                seen.add(key)

                data.append({
                    "name": name,
                    "description": desc,
                    "oem": oem,
                })

            except (StaleElementReferenceException, NoSuchElementException):
                continue
            except Exception:
                continue

        self.emit(message=f"V9: تم استخراج {len(data)} صف قطعة من الصفحة.")
        return data

    def process_sub_page(self, batch_id, vin, category, sub_row):
        sub_name = sub_row["sub_name"]
        original_db_url = sub_row["sub_url"]
        sub_url = re.sub(r"([&?])partNameId=[^&#]*", "", original_db_url, flags=re.I).rstrip("&?")

        sub_started = time.perf_counter()
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                self.db.set_sub_status(
                    batch_id, vin, category, original_db_url,
                    "IN_PROGRESS", retry_count=attempt - 1
                )

                self.safe_get(sub_url, batch_id, vin, category, original_db_url)

                # النسخة الأصلية كانت تنتظر ثانيتين فقط ثم تقرأ الصفوف.
                # هنا نضيف انتظارًا قصيرًا ذكيًا بحد أقصى 12 ثانية، وليس 120 ثانية.
                try:
                    WebDriverWait(self.driver, 5 if FAST_MODE else 12, poll_frequency=0.12 if FAST_MODE else 0.5).until(
                        lambda d: len(d.find_elements(By.CSS_SELECTOR, "li[data-part-expand]")) > 0
                    )
                except TimeoutException:
                    pass

                self.controlled_sleep(0.12 if FAST_MODE else 0.5)
                data = self.extract_rows()

                # إذا الصفحة فتحت لكن React لم يرسم الجدول بعد، محاولة قصيرة ثانية.
                if not data:
                    self.controlled_sleep(0.7 if FAST_MODE else 2)
                    data = self.extract_rows()

                added, path = self.writer.append_subpage_atomic(
                    vin, category, sub_name, sub_url, data
                )

                self.db.set_sub_status(
                    batch_id, vin, category, original_db_url,
                    "COMPLETED", retry_count=attempt - 1
                )
                self.current_delay = max(MIN_REQUEST_DELAY, self.current_delay * 0.92)
                self.emit(
                    kind="progress",
                    message=f"{vin} | {category} | {sub_name}: استخراج {len(data)} / حفظ {added} صف",
                    items=added,
                    file=str(path)
                )
                self.emit(message=f"V8 Fast: زمن Sub Group = {time.perf_counter()-sub_started:.2f}s")
                return

            except SiteUnavailable:
                raise
            except PermissionError as e:
                self.db.set_sub_status(
                    batch_id, vin, category, original_db_url, "FAILED", str(e), attempt
                )
                self.db.log_error(
                    batch_id, vin, category, original_db_url, "SAVE_ERROR", e
                )
                raise
            except Exception as e:
                self.db.set_sub_status(
                    batch_id, vin, category, original_db_url, "RETRYING", str(e), attempt
                )
                self.db.log_error(
                    batch_id, vin, category, original_db_url, "EXTRACTION_ERROR", e
                )
                self.capture_debug(vin, category, f"sub_retry_{attempt}")
                if attempt >= MAX_RETRIES:
                    self.db.set_sub_status(
                        batch_id, vin, category, original_db_url, "FAILED", str(e), attempt
                    )
                    self.emit(message=f"فشل الرابط بعد {MAX_RETRIES} محاولات: {sub_url}")
                    return
                self.controlled_sleep(backoff_delay(attempt))

    def run_batch(self, batch_id):
        batch = self.db.batch(batch_id)
        if not batch: raise ValueError("Batch not found")
        categories = json.loads(batch["categories_json"])
        force = bool(batch["force_rescrape"])
        vins = self.db.vins(batch_id)
        self.db.set_batch_status(batch_id, "IN_PROGRESS")
        try:
            if self.driver is None: self.create_driver()
            total = len(vins)
            for vrow in vins:
                self.checkpoint_gate()
                vin, pos = vrow["vin"], vrow["position"]
                if vrow["status"] == "COMPLETED" and not force:
                    continue
                self.emit(kind="vin", vin=vin, position=pos, total=total, message=f"VIN {pos}/{total}: {vin}")
                self.db.set_vin_status(batch_id, vin, "IN_PROGRESS")
                try:
                    for category in categories:
                        self.checkpoint_gate()
                        if not force and self.db.globally_completed(vin, category):
                            self.db.upsert_category(batch_id, vin, category, status="ALREADY_COMPLETED")
                            self.emit(message=f"تجاوز مكتمل سابقًا: {vin} / {category}")
                            continue
                        if not force and self.db.category_status(batch_id, vin, category) == "COMPLETED":
                            continue

                        # V5: Force Re-scrape starts with a genuinely fresh workbook.
                        if force:
                            self.writer.reset_category_file(vin, category)
                            self.emit(message=f"Force Re-scrape: ملف جديد لـ {vin} / {category}")

                        while True:
                            try:
                                actual_name, main_url = self.open_category_by_name(batch_id, vin, category)
                                self.db.upsert_category(batch_id, vin, category, main_url, "IN_PROGRESS")
                                # بعد الضغط على المجموعة، نوسع العناصر مباشرة. لا نشترط selector واحد قبل التوسيع.
                                self.expand_all_items(vin)
                                sub_links = self.collect_sub_links(vin)
                                if not sub_links:
                                    self.capture_debug(vin, category, "no_sub_links_after_category_open")
                                    # صفر هنا يعني أن بنية عناصر المجموعة لم تُكتشف، وليس أن الموقع متوقف.
                                    # لا نرجع للصفحة الرئيسية ولا نكرر البحث عن المجموعة بلا داعٍ.
                                    if not self.is_site_healthy():
                                        raise SiteUnavailable("Category page is genuinely unhealthy/unavailable")
                                    self.db.upsert_category(batch_id, vin, category, main_url, "FAILED",
                                                            "Category opened, but no sub links were detected")
                                    self.emit(message=f"المجموعة فُتحت لكن لم تُكتشف روابط فرعية: {category}. تم حفظ Debug بدون إعادة البحث المتكرر.")
                                    sub_links = []
                                    break
                                break
                            except CategoryNotFound:
                                self.db.upsert_category(batch_id, vin, category, status="CATEGORY_NOT_FOUND", error="Category not found by visible text")
                                self.emit(message=f"المجموعة غير موجودة بالاسم: {vin} / {category}")
                                sub_links = []
                                main_url = None
                                break
                            except SiteUnavailable:
                                self.wait_until_site_returns(batch_id, vin, category)
                        if not sub_links:
                            continue
                        self.db.register_sub_links(batch_id, vin, category, sub_links)
                        pending = self.db.pending_sub_links(batch_id, vin, category)
                        total_sub = len(sub_links)
                        for srow in pending:
                            self.checkpoint_gate()
                            self.emit(kind="sub", vin=vin, category=category, position=srow["position"], total=total_sub,
                                      message=f"{vin} | {category} | Sub {srow['position']}/{total_sub}")
                            while True:
                                try:
                                    self.process_sub_page(batch_id, vin, category, srow)
                                    break
                                except SiteUnavailable:
                                    self.wait_until_site_returns(batch_id, vin, category)
                        # category complete only if no pending/failed remain
                        remaining = self.db.pending_sub_links(batch_id, vin, category)
                        failed = [r for r in remaining if r["status"] == "FAILED"]
                        if failed:
                            self.db.upsert_category(batch_id, vin, category, main_url, "FAILED", f"{len(failed)} sub links failed")
                        else:
                            self.db.upsert_category(batch_id, vin, category, main_url, "COMPLETED")
                            self.db.mark_global_completed(vin, category)
                    self.db.set_vin_status(batch_id, vin, "COMPLETED")
                except VinNotFound as e:
                    self.db.set_vin_status(batch_id, vin, "VIN_NOT_FOUND", str(e))
                    self.db.log_error(batch_id, vin, None, None, "VIN_NOT_FOUND", e)
                    self.emit(message=f"VIN غير موجود: {vin}")
                except PermissionError as e:
                    self.db.set_vin_status(batch_id, vin, "FAILED", str(e))
                    self.emit(message=f"توقف VIN بسبب مشكلة حفظ Excel: {e}")
                except Exception as e:
                    self.db.set_vin_status(batch_id, vin, "FAILED", str(e))
                    self.db.log_error(batch_id, vin, None, None, "UNEXPECTED", traceback.format_exc())
                    self.capture_debug(vin, None, "vin_error")
                    self.emit(message=f"خطأ غير متوقع في VIN {vin}: {e}")
            self.db.set_batch_status(batch_id, "COMPLETED")
            self.emit(kind="done", message=f"اكتملت Batch #{batch_id}")
        except StopRequested:
            self.db.set_batch_status(batch_id, "STOPPED")
            self.emit(kind="done", message=f"تم إيقاف Batch #{batch_id} بأمان")
        except Exception as e:
            self.db.set_batch_status(batch_id, "PAUSED")
            self.emit(kind="fatal", message=f"خطأ رئيسي: {e}")
            logger.exception("Fatal batch error")
        finally:
            self.close_driver()

    def pause(self):
        self.pause_event.clear()
        self.emit(kind="site", status="PAUSED", message="تم طلب Pause؛ سيتم التوقف قبل الخطوة التالية.")

    def resume(self):
        self.pause_event.set()
        self.emit(kind="site", status="WORKING", message="تم استئناف التنفيذ.")

    def stop(self):
        self.stop_event.set()
        self.pause_event.set()
        self.emit(message="تم طلب Stop Safely...")

print("تم تحميل محرك All4Car Batch Scraper بنجاح.")
