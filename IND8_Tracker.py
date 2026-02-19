import os
import sys
import json
import sqlite3
import datetime
import csv
import re
import ssl
import shutil
import tempfile
import base64
import threading
import urllib.request
import urllib.parse
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk

# ----------------- APP META / VERSIONING -----------------

APP_NAME = "IND8 Tracker"
APP_VERSION = "2.0.0"

# ----------------- CONFIG / PATHS -----------------

CONFIG_FILE_NAME = "ind8_config.json"

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def load_config():
    base_dir = get_base_dir()
    config_path = os.path.join(base_dir, CONFIG_FILE_NAME)
    default_config = {
        "cloud_sync": False,
        "fclm_cookie": "",
        "fclm_warehouse_id": "IND8",
    }
    if not os.path.exists(config_path):
        save_config(default_config)
        return default_config
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for k, v in default_config.items():
            if k not in data:
                data[k] = v
        return data
    except Exception:
        return default_config

def save_config(config):
    base_dir = get_base_dir()
    config_path = os.path.join(base_dir, CONFIG_FILE_NAME)
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
    except Exception:
        pass

def get_db_path(config):
    cloud_sync = config.get("cloud_sync", False)
    if cloud_sync:
        one_drive = os.getenv("OneDrive")
        if one_drive:
            cloud_dir = os.path.join(one_drive, "IND8Tracker")
            os.makedirs(cloud_dir, exist_ok=True)
            return os.path.join(cloud_dir, "indirect_tracking.db")
    local_appdata = os.getenv("LOCALAPPDATA")
    if local_appdata:
        app_dir = os.path.join(local_appdata, "IND8Tracker")
        os.makedirs(app_dir, exist_ok=True)
        return os.path.join(app_dir, "indirect_tracking.db")
    return os.path.join(get_base_dir(), "indirect_tracking.db")

CONFIG = load_config()
DB_FILE = get_db_path(CONFIG)

WARNING_THRESHOLD_HOURS = 5
INDIRECT_LIMIT_HOURS = 6

# ----------------- FCLM CONFIGURATION -----------------

FCLM_BASE_URL = "https://fclm-portal.amazon.com"
FCLM_SHIFT_START_HOUR = 18  # 6 PM
FCLM_SHIFT_END_HOUR = 6     # 6 AM next day

# Process IDs for function rollup reports
FCLM_PROCESS_IDS = {
    "C-Returns Support": "1003058",
    "C-Returns Processed": "1003026",
    "V-Returns": "1003059",
    "WHD Grading": "1002979",
    "WHD Grading Support": "1003060",
}

# Restricted paths that can cause MPV violations
FCLM_RESTRICTED_PATHS = [
    "Vreturns WaterSpider",
    "C-Returns_EndofLine",
    "Water Spider",
    "WHD Waterspider",
    "WHD Water Spider",
    "Team_Mech_Wspider",
]

# Map FCLM path titles to IND8 area and role
FCLM_PATH_MAP = {
    "C-Returns_EndofLine":   {"area": "CRET", "role": "Water Spider", "type": "INDIRECT"},
    "Vreturns WaterSpider":  {"area": "VRET", "role": "Water Spider", "type": "INDIRECT"},
    "Water Spider":          {"area": "CRET", "role": "Water Spider", "type": "INDIRECT"},
    "WHD Waterspider":       {"area": "CRET", "role": "Water Spider", "type": "INDIRECT"},
    "WHD Water Spider":      {"area": "CRET", "role": "Water Spider", "type": "INDIRECT"},
    "Team_Mech_Wspider":     {"area": "CRET", "role": "Water Spider", "type": "INDIRECT"},
    "C-Returns Support":     {"area": "CRET", "role": "N/A",          "type": "DIRECT"},
    "C-Returns Processed":   {"area": "CRET", "role": "N/A",          "type": "DIRECT"},
    "V-Returns":             {"area": "VRET", "role": "N/A",          "type": "DIRECT"},
    "WHD Grading":           {"area": "CRET", "role": "Down Stack",   "type": "DIRECT"},
    "WHD Grading Support":   {"area": "CRET", "role": "Down Stack",   "type": "DIRECT"},
}

# Short display names for restricted paths
FCLM_PATH_SHORT_NAMES = {
    "C-Returns_EndofLine":   "CREOL",
    "Vreturns WaterSpider":  "VRWS",
    "Water Spider":          "CRSDCNTF",
    "WHD Waterspider":       "WHDWTSP",
    "WHD Water Spider":      "WHDWTSP",
    "Team_Mech_Wspider":     "TMWSP",
}

# MPV max time on restricted path (4 hours 30 minutes)
MPV_MAX_TIME_MINUTES = 270

# Work code to restricted path mapping (for MPV checking)
WORK_CODE_TO_PATH = {
    "CREOL":    "C-Returns_EndofLine",
    "EOL":      "C-Returns_EndofLine",
    "VRWS":     "Vreturns WaterSpider",
    "VRETWS":   "Vreturns WaterSpider",
    "CRSDCNTF": "Water Spider",
    "WHDWTSP":  "WHD Waterspider",
    "TMWSP":    "Team_Mech_Wspider",
}

# ----------------- DB SETUP -----------------

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS associates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            badge_id TEXT UNIQUE NOT NULL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            associate_id INTEGER NOT NULL,
            start_time TEXT NOT NULL,
            end_time TEXT,
            work_type TEXT NOT NULL,
            area TEXT NOT NULL,
            role TEXT NOT NULL,
            FOREIGN KEY (associate_id) REFERENCES associates(id)
        )
    """)
    conn.commit()
    conn.close()

def get_conn():
    return sqlite3.connect(DB_FILE)

def get_or_create_associate(badge_id, name=None):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, name FROM associates WHERE badge_id = ?", (badge_id,))
    row = c.fetchone()
    if row:
        conn.close()
        return row[0]
    if not name:
        name = badge_id
    c.execute("INSERT INTO associates (name, badge_id) VALUES (?, ?)", (name, badge_id))
    conn.commit()
    assoc_id = c.lastrowid
    conn.close()
    return assoc_id

def end_active_session(associate_id):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id FROM sessions WHERE associate_id = ? AND end_time IS NULL", (associate_id,))
    row = c.fetchone()
    if row:
        now = datetime.datetime.now().isoformat()
        c.execute("UPDATE sessions SET end_time = ? WHERE id = ?", (now, row[0]))
        conn.commit()
    conn.close()

def start_session(associate_id, work_type, area, role):
    end_active_session(associate_id)
    conn = get_conn()
    c = conn.cursor()
    now = datetime.datetime.now().isoformat()
    c.execute("""
        INSERT INTO sessions (associate_id, start_time, work_type, area, role)
        VALUES (?, ?, ?, ?, ?)
    """, (associate_id, now, work_type, area, role))
    conn.commit()
    conn.close()

def get_today_sessions(associate_id):
    conn = get_conn()
    c = conn.cursor()
    today = datetime.date.today().isoformat()
    c.execute("""
        SELECT id, start_time, end_time, work_type, area, role
        FROM sessions
        WHERE associate_id = ?
          AND date(start_time) = ?
        ORDER BY start_time
    """, (associate_id, today))
    rows = c.fetchall()
    conn.close()
    return rows

def compute_indirect_hours_today(associate_id):
    sessions = get_today_sessions(associate_id)
    total_seconds = 0
    now = datetime.datetime.now()
    for _, start, end, work_type, _, _ in sessions:
        if work_type != "INDIRECT":
            continue
        start_dt = datetime.datetime.fromisoformat(start)
        end_dt = datetime.datetime.fromisoformat(end) if end else now
        total_seconds += (end_dt - start_dt).total_seconds()
    return total_seconds / 3600.0

def get_all_associates():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, name, badge_id FROM associates ORDER BY name")
    rows = c.fetchall()
    conn.close()
    return rows

# ----------------- FCLM CLIENT -----------------

class FclmClient:
    """Client for fetching employee data from FCLM Portal."""

    def __init__(self, cookie="", warehouse_id="IND8",
                 on_cookie_refreshed=None):
        self.cookie = self._sanitize_cookie(cookie)
        self.warehouse_id = warehouse_id
        self._ssl_ctx = ssl.create_default_context()
        # Callback invoked with the new cookie string after a silent refresh
        self._on_cookie_refreshed = on_cookie_refreshed

    @staticmethod
    def _sanitize_cookie(cookie):
        """Strip non-ASCII chars (e.g. ellipsis from browser truncation)."""
        return "".join(c for c in cookie if ord(c) < 128)

    def is_connected(self):
        return bool(self.cookie.strip())

    # ---------- Shift date range ----------

    def get_shift_date_range(self):
        """Calculate night shift date range (6PM - 6AM)."""
        now = datetime.datetime.now()
        hour = now.hour

        if hour >= FCLM_SHIFT_START_HOUR:
            # Evening: started today, ends tomorrow
            start = now
            end = now + datetime.timedelta(days=1)
        elif hour < FCLM_SHIFT_END_HOUR:
            # Early morning: started yesterday, ends today
            start = now - datetime.timedelta(days=1)
            end = now
        else:
            # Daytime: check previous night shift
            start = now - datetime.timedelta(days=1)
            end = now

        return {
            "start_date": start.strftime("%Y/%m/%d"),
            "end_date": end.strftime("%Y/%m/%d"),
            "start_hour": FCLM_SHIFT_START_HOUR,
            "end_hour": FCLM_SHIFT_END_HOUR,
        }

    # ---------- HTTP helpers ----------

    @staticmethod
    def _is_login_page(html):
        """Return True if the HTML looks like a Midway login redirect."""
        lower = html.lower()
        return ("midway" in lower or "sign in" in lower
                or ("/login" in lower and "ganttChart" not in html))

    def _raw_get(self, url):
        """Single authenticated GET (no retry)."""
        req = urllib.request.Request(url)
        req.add_header("Cookie", self.cookie)
        req.add_header("User-Agent", "IND8Tracker/2.0")
        resp = urllib.request.urlopen(req, context=self._ssl_ctx, timeout=20)
        return resp.read().decode("utf-8", errors="replace")

    def _request(self, url):
        """Authenticated GET with silent cookie auto-refresh on expiry.

        If the first attempt returns a Midway login page, re-reads
        cookies from the browser and retries once.
        """
        html = self._raw_get(url)
        if not self._is_login_page(html):
            return html

        # Cookie expired – try to silently grab fresh ones
        new_cookie, _browser = BrowserCookieReader.auto_detect()
        if not new_cookie:
            return html  # no fresh cookie available, return the login page

        self.cookie = self._sanitize_cookie(new_cookie)
        if self._on_cookie_refreshed:
            try:
                self._on_cookie_refreshed(self.cookie)
            except Exception:
                pass  # don't let callback errors break the request

        return self._raw_get(url)

    def test_connection(self):
        """Test FCLM connectivity. Returns (success, message)."""
        try:
            url = f"{FCLM_BASE_URL}/employee/timeDetails?warehouseId={self.warehouse_id}"
            html = self._request(url)
            if self._is_login_page(html):
                return False, "Cookie expired or invalid - FCLM is asking to log in."
            if "ganttChart" in html or "Time Details" in html:
                return True, "Connected to FCLM successfully."
            # Show a snippet of what we got back for debugging
            snippet = html[:300].strip()
            return False, f"Unexpected response (cookie may be wrong format).\n\nFirst 300 chars:\n{snippet}"
        except Exception as e:
            return False, f"Connection failed: {e}"

    # ---------- Employee time details ----------

    def fetch_employee_time_details(self, employee_id):
        """Fetch and parse time details for an employee."""
        shift = self.get_shift_date_range()
        params = urllib.parse.urlencode({
            "employeeId": employee_id,
            "warehouseId": self.warehouse_id,
            "startDateDay": shift["end_date"],
            "maxIntradayDays": "1",
            "spanType": "Intraday",
            "startDateIntraday": shift["start_date"],
            "startHourIntraday": str(shift["start_hour"]),
            "startMinuteIntraday": "0",
            "endDateIntraday": shift["end_date"],
            "endHourIntraday": str(shift["end_hour"]),
            "endMinuteIntraday": "0",
        })
        url = f"{FCLM_BASE_URL}/employee/timeDetails?{params}"
        html = self._request(url)
        return self._parse_time_details(html, employee_id)

    def _parse_time_details(self, html, employee_id):
        """Parse FCLM time details HTML (gantt chart table)."""
        result = {
            "employee_id": employee_id,
            "sessions": [],
            "current_activity": None,
            "is_clocked_in": False,
            "hours_on_task": 0.0,
            "total_scheduled_hours": 0.0,
        }

        # Find the gantt chart table
        table_m = re.search(
            r'<table[^>]*class="[^"]*ganttChart[^"]*"[^>]*>(.*?)</table>',
            html, re.DOTALL | re.IGNORECASE,
        )
        if not table_m:
            return result

        table_html = table_m.group(1)

        # Track function-seg totals for correction
        func_seg_totals = {}

        # Find all <tr> with a class attribute
        for row_m in re.finditer(r'<tr[^>]*class="([^"]*)"[^>]*>(.*?)</tr>', table_html, re.DOTALL):
            row_class = row_m.group(1)
            row_html = row_m.group(2)

            if "totSummary" in row_class or "<th" in row_html:
                continue

            cells = re.findall(r'<td[^>]*>(.*?)</td>', row_html, re.DOTALL)
            if len(cells) < 4:
                continue

            title, start, end, duration = "", "", "", ""

            if "job-seg" in row_class:
                if len(cells) >= 5:
                    title = self._strip_html(cells[1])
                    start = self._strip_html(cells[2])
                    end = self._strip_html(cells[3])
                    duration = self._strip_html(cells[4])
            elif "function-seg" in row_class or "clock-seg" in row_class:
                raw_title = self._strip_html(cells[0])
                if "\u2666" in raw_title:  # diamond separator
                    title = raw_title.split("\u2666")[-1].strip()
                else:
                    title = raw_title
                start = self._strip_html(cells[1])
                end = self._strip_html(cells[2])
                duration = self._strip_html(cells[3])
            else:
                # Generic fallback
                if len(cells) >= 5:
                    first_text = self._strip_html(cells[0])
                    if len(first_text) <= 2:
                        title = self._strip_html(cells[1])
                        start = self._strip_html(cells[2])
                        end = self._strip_html(cells[3])
                        duration = self._strip_html(cells[4])
                    else:
                        title = first_text
                        start = self._strip_html(cells[1])
                        end = self._strip_html(cells[2])
                        duration = self._strip_html(cells[3])
                else:
                    title = self._strip_html(cells[0])
                    start = self._strip_html(cells[1])
                    end = self._strip_html(cells[2])
                    duration = self._strip_html(cells[3])

            if not title or "OffClock" in title or "OnClock" in title:
                continue

            dur_mins = self._parse_fclm_duration(duration)

            # Track function-seg totals (aggregate per path)
            if "function-seg" in row_class:
                func_seg_totals[title] = func_seg_totals.get(title, 0) + dur_mins
                continue  # Don't add function-seg as sessions (would double-count)

            session = {
                "title": title,
                "start": start,
                "end": end,
                "duration": duration,
                "duration_minutes": dur_mins,
                "row_type": "job" if "job-seg" in row_class else "other",
            }
            result["sessions"].append(session)

            if not end or end.strip() == "":
                result["current_activity"] = session
                result["is_clocked_in"] = True

        # Correction: if function-seg total > sum of job-seg for a path,
        # add a synthetic session for the difference
        job_seg_totals = {}
        for s in result["sessions"]:
            job_seg_totals[s["title"]] = job_seg_totals.get(s["title"], 0) + s["duration_minutes"]

        for path, func_mins in func_seg_totals.items():
            job_mins = job_seg_totals.get(path, 0)
            if func_mins > job_mins:
                diff = func_mins - job_mins
                result["sessions"].append({
                    "title": path,
                    "start": "",
                    "end": "",
                    "duration": f"{int(diff)}:{int((diff % 1) * 60):02d}",
                    "duration_minutes": diff,
                    "row_type": "correction",
                })

        # Extract "Hours on Task"
        hours_m = re.search(r"Hours on Task:\s*([\d.]+)\s*/\s*([\d.]+)", html)
        if hours_m:
            result["hours_on_task"] = float(hours_m.group(1))
            result["total_scheduled_hours"] = float(hours_m.group(2))

        return result

    # ---------- Function rollup (path AAs) ----------

    def fetch_all_path_aas(self):
        """Fetch all associates on restricted paths from FCLM function rollup."""
        shift = self.get_shift_date_range()
        all_aas = {p: [] for p in FCLM_RESTRICTED_PATHS}

        errors = []
        for process_name, process_id in FCLM_PROCESS_IDS.items():
            try:
                params = urllib.parse.urlencode({
                    "reportFormat": "HTML",
                    "warehouseId": self.warehouse_id,
                    "processId": process_id,
                    "maxIntradayDays": "1",
                    "spanType": "Intraday",
                    "startDateIntraday": shift["start_date"],
                    "startHourIntraday": str(shift["start_hour"]),
                    "startMinuteIntraday": "0",
                    "endDateIntraday": shift["end_date"],
                    "endHourIntraday": str(shift["end_hour"]),
                    "endMinuteIntraday": "0",
                })
                url = f"{FCLM_BASE_URL}/reports/functionRollup?{params}"
                html = self._request(url)
                if self._is_login_page(html):
                    errors.append(f"{process_name}: Cookie expired (login page returned)")
                    continue
                self._parse_function_rollup(html, process_name, all_aas)
            except Exception as e:
                errors.append(f"{process_name}: {e}")

        # Sort each path by hours descending
        for path in all_aas:
            all_aas[path].sort(key=lambda a: a["hours"], reverse=True)

        # Attach errors so the UI can report them
        all_aas["_errors"] = errors
        return all_aas

    def _parse_function_rollup(self, html, process_name, result):
        """Parse function rollup HTML and populate result dict."""
        # Split HTML into table blocks
        tables = re.findall(r'<table[^>]*>(.*?)</table>', html, re.DOTALL | re.IGNORECASE)

        for table_html in tables:
            # Determine which restricted path this table belongs to
            table_path = None
            for path in FCLM_RESTRICTED_PATHS:
                if path in table_html:
                    table_path = path
                    break
            if not table_path:
                continue

            # Special handling: Water Spider from WHD process -> WHD Waterspider
            if table_path == "Water Spider" and "WHD" in process_name:
                table_path = "WHD Waterspider"

            if table_path not in result:
                result[table_path] = []

            seen_badges = {aa["badge_id"] for aa in result[table_path]}

            # Find data rows
            rows = re.findall(r'<tr[^>]*>(.*?)</tr>', table_html, re.DOTALL)
            for row_html in rows:
                cells = re.findall(r'<td[^>]*>(.*?)</td>', row_html, re.DOTALL)
                if len(cells) < 5:
                    continue
                first = self._strip_html(cells[0])
                if first not in ("AMZN", "TEMP"):
                    continue

                badge_id = self._strip_html(cells[1])
                if not badge_id or not badge_id.isdigit():
                    continue
                if badge_id in seen_badges:
                    continue

                name = self._strip_html(cells[2])

                # Find total hours (last numeric cell)
                hours = 0.0
                for i in range(len(cells) - 1, 2, -1):
                    text = self._strip_html(cells[i])
                    try:
                        hours = float(text)
                        break
                    except ValueError:
                        continue

                # If still 0, sum all numeric cells in the middle
                if hours == 0:
                    for i in range(4, len(cells) - 1):
                        text = self._strip_html(cells[i])
                        try:
                            hours += float(text)
                        except ValueError:
                            pass

                result[table_path].append({
                    "badge_id": badge_id,
                    "name": name,
                    "hours": hours,
                    "minutes": hours * 60,
                })
                seen_badges.add(badge_id)

    # ---------- Helpers ----------

    @staticmethod
    def _strip_html(html_str):
        """Remove HTML tags, preferring link text."""
        link = re.search(r'<a[^>]*>(.*?)</a>', html_str, re.DOTALL)
        if link:
            return link.group(1).strip()
        return re.sub(r'<[^>]+>', '', html_str).strip()

    @staticmethod
    def _parse_fclm_duration(duration):
        """Parse FCLM duration (MM:SS format like '210:35') to minutes."""
        if not duration:
            return 0.0
        parts = duration.split(":")
        if len(parts) >= 2:
            try:
                mins = int(parts[0])
                secs = int(parts[1])
                return mins + secs / 60.0
            except ValueError:
                return 0.0
        return 0.0


# ----------------- BROWSER COOKIE READER -----------------

class BrowserCookieReader:
    """Read FCLM session cookies directly from browser cookie databases.

    Supports Edge, Chrome (Windows via DPAPI + AES-256-GCM) and Firefox
    (all platforms, plain SQLite).  The user just needs to be logged into
    FCLM in their browser – no DevTools required.
    """

    FCLM_DOMAINS = [".amazon.com", "fclm-portal.amazon.com"]

    # --- public API ---

    @classmethod
    def auto_detect(cls):
        """Try installed browsers and return (cookie_str, browser_name).

        Returns (None, error_message) when no cookies are found.
        """
        browsers = []
        if sys.platform == "win32":
            browsers = [
                ("Edge", cls._read_edge),
                ("Chrome", cls._read_chrome),
                ("Firefox", cls._read_firefox),
            ]
        else:
            # Linux / macOS – only Firefox uses plain SQLite
            browsers = [
                ("Firefox", cls._read_firefox),
            ]

        errors = []
        for name, method in browsers:
            try:
                cookie = method()
                if cookie:
                    return cookie, name
            except Exception as exc:
                errors.append(f"{name}: {exc}")

        if errors:
            return None, "No FCLM cookies found.\n" + "\n".join(errors)
        return None, "No supported browser found with FCLM cookies."

    # --- Firefox ---

    @classmethod
    def _read_firefox(cls):
        if sys.platform == "win32":
            profiles_dir = os.path.join(
                os.getenv("APPDATA", ""), "Mozilla", "Firefox", "Profiles"
            )
        elif sys.platform == "darwin":
            profiles_dir = os.path.expanduser(
                "~/Library/Application Support/Firefox/Profiles"
            )
        else:
            profiles_dir = os.path.expanduser("~/.mozilla/firefox")

        if not os.path.isdir(profiles_dir):
            return None

        # Pick the first profile that has a cookies.sqlite
        cookie_db = None
        for entry in sorted(os.listdir(profiles_dir)):
            candidate = os.path.join(profiles_dir, entry, "cookies.sqlite")
            if os.path.isfile(candidate):
                cookie_db = candidate
                break
        if not cookie_db:
            return None

        return cls._read_sqlite_cookies(
            cookie_db,
            table="moz_cookies",
            host_col="host",
            name_col="name",
            value_col="value",
        )

    # --- Chrome / Edge (Windows) ---

    @classmethod
    def _read_chrome(cls):
        return cls._read_chromium("Google", "Chrome")

    @classmethod
    def _read_edge(cls):
        return cls._read_chromium("Microsoft", "Edge")

    @classmethod
    def _read_chromium(cls, vendor, browser):
        if sys.platform != "win32":
            return None

        local_appdata = os.getenv("LOCALAPPDATA", "")
        user_data = os.path.join(local_appdata, vendor, browser, "User Data")
        if not os.path.isdir(user_data):
            return None

        # Obtain the AES key from Local State
        key = cls._get_chromium_key(user_data)
        if key is None:
            return None

        # Try Default profile first, then Profile 1, Profile 2 …
        for profile in ["Default", "Profile 1", "Profile 2"]:
            db_path = os.path.join(user_data, profile, "Network", "Cookies")
            if not os.path.isfile(db_path):
                continue
            cookie = cls._read_chromium_cookies(db_path, key)
            if cookie:
                return cookie
        return None

    @classmethod
    def _get_chromium_key(cls, user_data_dir):
        local_state_path = os.path.join(user_data_dir, "Local State")
        if not os.path.isfile(local_state_path):
            return None
        try:
            with open(local_state_path, "r", encoding="utf-8") as fh:
                state = json.load(fh)
            encrypted_key = base64.b64decode(
                state["os_crypt"]["encrypted_key"]
            )
            # Strip the "DPAPI" prefix (5 bytes)
            encrypted_key = encrypted_key[5:]
            return cls._dpapi_decrypt(encrypted_key)
        except Exception:
            return None

    @classmethod
    def _read_chromium_cookies(cls, db_path, key):
        tmp = tempfile.mktemp(suffix=".sqlite")
        shutil.copy2(db_path, tmp)
        try:
            conn = sqlite3.connect(tmp)
            cur = conn.cursor()
            cookies = []
            for domain in cls.FCLM_DOMAINS:
                cur.execute(
                    "SELECT name, encrypted_value FROM cookies "
                    "WHERE host_key = ? OR host_key LIKE ?",
                    (domain, f"%{domain}"),
                )
                for name, enc_val in cur.fetchall():
                    val = cls._decrypt_chromium_value(enc_val, key)
                    if val:
                        cookies.append(f"{name}={val}")
            conn.close()
            return "; ".join(cookies) if cookies else None
        finally:
            try:
                os.unlink(tmp)
            except OSError:
                pass

    @classmethod
    def _decrypt_chromium_value(cls, encrypted, key):
        if not encrypted:
            return None
        # v10 / v11 → AES-256-GCM
        if encrypted[:3] in (b"v10", b"v11"):
            try:
                from cryptography.hazmat.primitives.ciphers.aead import AESGCM
                nonce = encrypted[3:15]
                ciphertext = encrypted[15:]
                return AESGCM(key).decrypt(nonce, ciphertext, None).decode("utf-8")
            except Exception:
                return None
        # Older DPAPI-only cookies
        raw = cls._dpapi_decrypt(encrypted)
        return raw.decode("utf-8", errors="replace") if raw else None

    # --- DPAPI (Windows) ---

    @staticmethod
    def _dpapi_decrypt(data):
        if sys.platform != "win32":
            return None
        import ctypes
        import ctypes.wintypes

        class _BLOB(ctypes.Structure):
            _fields_ = [
                ("cbData", ctypes.wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char)),
            ]

        blob_in = _BLOB(len(data), ctypes.create_string_buffer(data, len(data)))
        blob_out = _BLOB()
        if ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(blob_in), None, None, None, None, 0,
            ctypes.byref(blob_out),
        ):
            result = ctypes.string_at(blob_out.pbData, blob_out.cbData)
            ctypes.windll.kernel32.LocalFree(blob_out.pbData)
            return result
        return None

    # --- shared SQLite reader (Firefox) ---

    @classmethod
    def _read_sqlite_cookies(cls, db_path, *, table, host_col, name_col, value_col):
        tmp = tempfile.mktemp(suffix=".sqlite")
        shutil.copy2(db_path, tmp)
        try:
            conn = sqlite3.connect(tmp)
            cur = conn.cursor()
            cookies = []
            for domain in cls.FCLM_DOMAINS:
                cur.execute(
                    f"SELECT {name_col}, {value_col} FROM {table} "
                    f"WHERE {host_col} = ? OR {host_col} LIKE ?",
                    (domain, f"%{domain}"),
                )
                for name, value in cur.fetchall():
                    if value:
                        cookies.append(f"{name}={value}")
            conn.close()
            return "; ".join(cookies) if cookies else None
        finally:
            try:
                os.unlink(tmp)
            except OSError:
                pass


# ----------------- MPV CHECKING -----------------

def classify_fclm_session(title):
    """Classify an FCLM session title as restricted or not.
    Returns the restricted path name if it matches, else None."""
    if not title:
        return None
    normalized = title.lower().replace(" ", "").replace("_", "")

    for path in FCLM_RESTRICTED_PATHS:
        path_norm = path.lower().replace(" ", "").replace("_", "")
        if path_norm in normalized or normalized in path_norm:
            return path
        if path in title:
            return path

    # Additional keyword matching
    if "waterspider" in normalized:
        if "whd" in normalized:
            return "WHD Waterspider"
        if "vreturn" in normalized:
            return "Vreturns WaterSpider"
        return "Water Spider"
    if "endofline" in normalized or "eol" in normalized:
        return "C-Returns_EndofLine"
    if "teammech" in normalized or "tmwsp" in normalized:
        return "Team_Mech_Wspider"

    return None


def compute_fclm_path_times(sessions):
    """Calculate total minutes per restricted path from FCLM sessions."""
    path_times = {}
    for s in sessions:
        rp = classify_fclm_session(s["title"])
        if rp:
            path_times[rp] = path_times.get(rp, 0) + s.get("duration_minutes", 0)
    return path_times


def check_mpv_risk(sessions, target_work_code=None):
    """Check for MPV risk given FCLM sessions and a target work code.
    Returns a dict with risk info."""
    result = {
        "has_risk": False,
        "reason": None,
        "details": None,
        "worked_paths": [],
        "target_path": None,
        "path_times": {},
        "remaining_minutes": None,
        "current_minutes": None,
    }

    path_times = compute_fclm_path_times(sessions)
    result["path_times"] = path_times
    result["worked_paths"] = list(path_times.keys())

    # Determine target restricted path from work code
    target_path = None
    if target_work_code:
        upper = target_work_code.upper().replace(" ", "").replace("_", "")
        for code, path in WORK_CODE_TO_PATH.items():
            if upper == code or upper.startswith(code) or code.startswith(upper):
                target_path = path
                break
    result["target_path"] = target_path

    if not target_path:
        return result

    # Rule 1: Path switch
    for wp in result["worked_paths"]:
        if wp != target_path:
            result["has_risk"] = True
            result["reason"] = "PATH_SWITCH"
            t = path_times[wp]
            result["details"] = (
                f"Already worked {wp} ({_fmt_mins(t)}). "
                f"Cannot switch to {target_path}."
            )
            return result

    # Rule 2: Time exceeded
    target_time = path_times.get(target_path, 0)
    if target_time >= MPV_MAX_TIME_MINUTES:
        result["has_risk"] = True
        result["reason"] = "TIME_EXCEEDED"
        result["details"] = (
            f"Already {_fmt_mins(target_time)} on {target_path}. "
            f"Max allowed is {_fmt_mins(MPV_MAX_TIME_MINUTES)}."
        )
        return result

    # Same path, under limit
    if target_time > 0:
        result["remaining_minutes"] = MPV_MAX_TIME_MINUTES - target_time
        result["current_minutes"] = target_time

    return result


def _fmt_mins(mins):
    h = int(mins) // 60
    m = int(mins) % 60
    if h > 0 and m > 0:
        return f"{h}h {m}m"
    if h > 0:
        return f"{h}h"
    return f"{m}m"


# ----------------- SPLASH SCREEN -----------------

def show_splash(root):
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(bg="#1a1a1a")
    try:
        img = Image.open("splash.png")
        splash_img = ImageTk.PhotoImage(img)
        label = tk.Label(splash, image=splash_img, bg="#1a1a1a")
        label.image = splash_img
        label.pack()
    except Exception:
        label = tk.Label(
            splash,
            text=f"{APP_NAME}\nv{APP_VERSION}",
            bg="#1a1a1a",
            fg="#ff9900",
            font=("Segoe UI", 20, "bold"),
            padx=40,
            pady=30
        )
        label.pack()
    splash.update_idletasks()
    w = splash.winfo_width()
    h = splash.winfo_height()
    x = (splash.winfo_screenwidth() - w) // 2
    y = (splash.winfo_screenheight() - h) // 2
    splash.geometry(f"{w}x{h}+{x}+{y}")
    splash.after(2000, splash.destroy)
    return splash

# ----------------- APP -----------------

class App:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("1100x700")
        self.root.configure(bg="#1a1a1a")

        self.current_associate_id = None

        self.badge_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.area_var = tk.StringVar(value="CRET")
        self.role_var = tk.StringVar(value="Water Spider")
        self.status_var = tk.StringVar(value="No associate selected")
        self.indirect_hours_var = tk.StringVar(value="0.00")
        self.cloud_sync_var = tk.BooleanVar(value=self.config.get("cloud_sync", False))

        # FCLM client
        self.fclm = FclmClient(
            cookie=self.config.get("fclm_cookie", ""),
            warehouse_id=self.config.get("fclm_warehouse_id", "IND8"),
            on_cookie_refreshed=self._on_cookie_refreshed,
        )
        # Cached FCLM data
        self._fclm_employee_data = None
        self._fclm_path_aas = None

        self._setup_style()
        self._build_layout()

    # ---------- STYLE ----------

    def _setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.TFrame", background="#1a1a1a")
        style.configure("DarkInner.TFrame", background="#222222")
        style.configure("Dark.TLabel", background="#1a1a1a", foreground="#e6e6e6", font=("Segoe UI", 11))
        style.configure("DarkBanner.TLabel", background="#222222", foreground="#ff9900",
                        font=("Segoe UI", 22, "bold"), padding=20)
        style.configure("DarkNav.TButton", background="#333333", foreground="#e6e6e6",
                        padding=8, font=("Segoe UI", 11, "bold"))
        style.map("DarkNav.TButton", background=[("active", "#444444")])
        style.configure("Dark.Treeview",
                        background="#222222",
                        fieldbackground="#222222",
                        foreground="#e6e6e6",
                        rowheight=24)
        style.configure("Dark.Treeview.Heading",
                        background="#333333",
                        foreground="#e6e6e6",
                        font=("Segoe UI", 10, "bold"))

    # ---------- LAYOUT ----------

    def _build_layout(self):
        # Banner
        banner_frame = tk.Frame(self.root, bg="#222222", height=80)
        banner_frame.pack(side="top", fill="x")

        try:
            logo_img = Image.open("ind8logo.png").resize((60, 60))
            self.logo = ImageTk.PhotoImage(logo_img)
            logo_label = tk.Label(banner_frame, image=self.logo, bg="#222222")
            logo_label.pack(side="left", padx=20)
        except Exception:
            pass

        title_label = ttk.Label(banner_frame, text="IND8", style="DarkBanner.TLabel")
        title_label.pack(side="left")

        # FCLM status indicator in banner
        self.fclm_status_label = tk.Label(
            banner_frame,
            text="FCLM: Disconnected",
            bg="#222222",
            fg="#e74c3c",
            font=("Segoe UI", 9, "bold"),
            padx=10,
        )
        self.fclm_status_label.pack(side="right", padx=(0, 10))

        version_label = tk.Label(
            banner_frame,
            text=f"v{APP_VERSION}",
            bg="#222222",
            fg="#cccccc",
            font=("Segoe UI", 10, "bold")
        )
        version_label.pack(side="right", padx=10)

        # Main area
        main_frame = tk.Frame(self.root, bg="#1a1a1a")
        main_frame.pack(fill="both", expand=True)

        # Navigation sidebar
        nav = tk.Frame(main_frame, bg="#111111", width=200)
        nav.pack(side="left", fill="y")

        self._add_nav_button(nav, "Home", lambda: None)

        # FCLM section header
        fclm_header = tk.Label(nav, text="--- FCLM ---", bg="#111111", fg="#ff9900",
                               font=("Segoe UI", 9, "bold"))
        fclm_header.pack(fill="x", pady=(8, 2), padx=8)
        self._add_nav_button(nav, "FCLM Lookup", self.fclm_lookup_badge)
        self._add_nav_button(nav, "FCLM Sync Paths", self.fclm_sync_paths)
        self._add_nav_button(nav, "FCLM Dashboard", self.fclm_open_dashboard)
        self._add_nav_button(nav, "FCLM Settings", self.fclm_open_settings)

        # Manual section header
        manual_header = tk.Label(nav, text="--- Manual ---", bg="#111111", fg="#888888",
                                 font=("Segoe UI", 9, "bold"))
        manual_header.pack(fill="x", pady=(8, 2), padx=8)
        self._add_nav_button(nav, "Start Direct", self.start_direct)
        self._add_nav_button(nav, "Start Indirect", self.start_indirect)
        self._add_nav_button(nav, "End Current", self.end_current)
        self._add_nav_button(nav, "Dashboard", self.open_dashboard)

        # Tools section
        tools_header = tk.Label(nav, text="--- Tools ---", bg="#111111", fg="#888888",
                                font=("Segoe UI", 9, "bold"))
        tools_header.pack(fill="x", pady=(8, 2), padx=8)
        self._add_nav_button(nav, "Export to CSV", self.export_to_excel)
        self._add_nav_button(nav, "Shift Report", self.export_shift_report)
        self._add_nav_button(nav, "Legend", self.open_legend)

        cloud_frame = tk.Frame(nav, bg="#111111")
        cloud_frame.pack(fill="x", pady=10, padx=8)
        cloud_cb = tk.Checkbutton(
            cloud_frame,
            text="Cloud sync (OneDrive)",
            variable=self.cloud_sync_var,
            bg="#111111",
            fg="#e6e6e6",
            selectcolor="#111111",
            activebackground="#111111",
            activeforeground="#e6e6e6",
            command=self.toggle_cloud_sync
        )
        cloud_cb.pack(anchor="w")

        # Content area
        content = tk.Frame(main_frame, bg="#1a1a1a")
        content.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        # Input row
        top = tk.Frame(content, bg="#1a1a1a")
        top.pack(fill="x", pady=(0, 10))

        tk.Label(top, text="Badge ID:", bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w")
        self.badge_entry = tk.Entry(top, textvariable=self.badge_var, width=18,
                                    bg="#333333", fg="#e6e6e6", insertbackground="white",
                                    relief="flat")
        self.badge_entry.grid(row=0, column=1, sticky="w", padx=(5, 20))
        self.badge_entry.bind("<Return>", self.scan_badge)

        tk.Label(top, text="Name (optional):", bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=0, column=2, sticky="w")
        self.name_entry = tk.Entry(top, textvariable=self.name_var, width=20,
                                   bg="#333333", fg="#e6e6e6", insertbackground="white",
                                   relief="flat")
        self.name_entry.grid(row=0, column=3, sticky="w", padx=(5, 20))

        tk.Label(top, text="Area:", bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=(8, 0))
        area_cb = ttk.Combobox(top, textvariable=self.area_var,
                               values=["CRET", "VRET", "TRANSHIP", "ILS"],
                               width=15, state="readonly")
        area_cb.grid(row=1, column=1, sticky="w", pady=(8, 0))

        tk.Label(top, text="Indirect Role:", bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=1, column=2, sticky="w", pady=(8, 0))
        role_cb = ttk.Combobox(top, textvariable=self.role_var,
                               values=["Water Spider", "Down Stack", "Unloads", "N/A"],
                               width=18, state="readonly")
        role_cb.grid(row=1, column=3, sticky="w", pady=(8, 0))

        load_btn = ttk.Button(top, text="Load Associate", style="DarkNav.TButton",
                              command=self.load_associate)
        load_btn.grid(row=0, column=4, padx=(10, 0))

        # Session table
        mid = tk.Frame(content, bg="#1a1a1a")
        mid.pack(fill="both", expand=True)

        columns = ("start", "end", "type", "area", "role", "duration")
        self.tree = ttk.Treeview(mid, columns=columns, show="headings",
                                 style="Dark.Treeview")
        for col in columns:
            self.tree.heading(col, text=col.capitalize(), anchor="w")
            self.tree.column(col, anchor="w", width=100)
        self.tree.pack(fill="both", expand=True)

        # Status bar
        bottom = tk.Frame(content, bg="#1a1a1a")
        bottom.pack(fill="x", pady=(10, 0))

        tk.Label(bottom, text="Indirect hours today:",
                 bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w")
        tk.Label(bottom, textvariable=self.indirect_hours_var,
                 bg="#1a1a1a", fg="#ffcc00",
                 font=("Segoe UI", 11, "bold")).grid(row=0, column=1, sticky="w", padx=(5, 20))

        tk.Label(bottom, text="Status:",
                 bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=0, column=2, sticky="w")
        self.status_label = tk.Label(bottom, textvariable=self.status_var,
                                     bg="#1a1a1a", fg="#00cc66",
                                     font=("Segoe UI", 10, "bold"))
        self.status_label.grid(row=0, column=3, sticky="w", padx=(5, 0))

        db_label = tk.Label(
            bottom,
            text=f"DB: {DB_FILE}",
            bg="#1a1a1a",
            fg="#777777",
            font=("Segoe UI", 8),
            anchor="w",
            justify="left",
            wraplength=600
        )
        db_label.grid(row=1, column=0, columnspan=4, sticky="w", pady=(5, 0))

        # Update FCLM status indicator
        self._update_fclm_status_label()

    # ---------- DIRECT / INDIRECT HELPERS ----------

    def has_direct_today(self, associate_id):
        sessions = get_today_sessions(associate_id)
        return any(work_type == "DIRECT" for _, _, _, work_type, _, _ in sessions)

    def has_indirect_today(self, associate_id):
        sessions = get_today_sessions(associate_id)
        return any(work_type == "INDIRECT" for _, _, _, work_type, _, _ in sessions)

    def get_indirect_roles_today(self, associate_id):
        sessions = get_today_sessions(associate_id)
        roles = set()
        for _, _, _, work_type, _, role in sessions:
            if work_type == "INDIRECT":
                roles.add(role)
        return roles

    # ---------- NAV BUTTON ----------

    def _add_nav_button(self, parent, text, command):
        frame = tk.Frame(parent, bg="#111111")
        frame.pack(fill="x", pady=4, padx=8)

        btn = tk.Label(frame, text=text, bg="#333333", fg="#e6e6e6",
                       font=("Segoe UI", 10, "bold"),
                       padx=12, pady=6)
        btn.pack(fill="x")
        btn.bind("<Button-1>", lambda e: command())
        btn.bind("<Enter>", lambda e: btn.config(bg="#444444"))
        btn.bind("<Leave>", lambda e: btn.config(bg="#333333"))

    # ---------- CLOUD SYNC ----------

    def toggle_cloud_sync(self):
        new_value = self.cloud_sync_var.get()
        self.config["cloud_sync"] = new_value
        save_config(self.config)
        messagebox.showinfo(
            "Cloud Sync",
            "Cloud sync setting saved.\n\n"
            "Please close and reopen the app for the change to take effect."
        )

    # ---------- LEGEND PANEL ----------

    def open_legend(self):
        legend = tk.Toplevel(self.root)
        legend.title("Dashboard Legend")
        legend.configure(bg="#1a1a1a")
        legend.geometry("350x330")

        items = [
            ("DIRECT only", "#003366"),
            ("INDIRECT only", "#665511"),
            ("DIRECT + INDIRECT", "#CC6600"),
            ("No work today", "#444444"),
            ("Warning (5+ hrs indirect)", "#665511"),
            ("Violation (6+ hrs indirect)", "#661111"),
            ("MPV restricted path", "#663399"),
        ]

        tk.Label(
            legend,
            text="Dashboard Color Legend",
            bg="#1a1a1a",
            fg="#ff9900",
            font=("Segoe UI", 14, "bold"),
            pady=10
        ).pack()

        for label, color in items:
            row = tk.Frame(legend, bg="#1a1a1a")
            row.pack(fill="x", pady=4, padx=10)

            swatch = tk.Label(row, bg=color, width=4, height=2)
            swatch.pack(side="left", padx=(0, 10))

            tk.Label(
                row,
                text=label,
                bg="#1a1a1a",
                fg="#e6e6e6",
                font=("Segoe UI", 11)
            ).pack(side="left")

    # ---------- BADGE / ASSOCIATE ----------

    def scan_badge(self, event=None):
        badge = self.badge_var.get().strip()
        if not badge:
            return
        # If FCLM is connected, do FCLM lookup instead of just local
        if self.fclm.is_connected():
            self._fclm_lookup_employee(badge)
            return

        assoc_id = get_or_create_associate(badge, self.name_var.get().strip() or None)
        self.current_associate_id = assoc_id
        self.refresh_view()
        self._update_direct_indirect_status(assoc_id)

    def load_associate(self):
        badge = self.badge_var.get().strip()
        if not badge:
            messagebox.showerror("Error", "Enter a badge ID.")
            return
        # If FCLM is connected, do FCLM lookup
        if self.fclm.is_connected():
            self._fclm_lookup_employee(badge)
            return

        assoc_id = get_or_create_associate(badge, self.name_var.get().strip() or None)
        self.current_associate_id = assoc_id
        self.refresh_view()
        self._update_direct_indirect_status(assoc_id)

    def _update_direct_indirect_status(self, assoc_id):
        did_direct = self.has_direct_today(assoc_id)
        did_indirect = self.has_indirect_today(assoc_id)
        roles = self.get_indirect_roles_today(assoc_id)
        roles_text = ", ".join(sorted(roles)) if roles else "None"

        if did_direct and did_indirect:
            self.status_var.set(f"DIRECT + INDIRECT ({roles_text})")
            self.status_label.config(fg="#ffcc00")
        elif did_direct:
            self.status_var.set("DIRECT work today")
            self.status_label.config(fg="#00ccff")
        elif did_indirect:
            self.status_var.set(f"INDIRECT work today ({roles_text})")
            self.status_label.config(fg="#ffcc00")
        else:
            self.status_var.set("No DIRECT or INDIRECT work today")
            self.status_label.config(fg="#00cc66")

    # ---------- SESSION CONTROLS ----------

    def start_direct(self):
        if not self.current_associate_id:
            messagebox.showerror("Error", "Load an associate first.")
            return
        start_session(self.current_associate_id, "DIRECT", self.area_var.get(), "N/A")
        self.refresh_view()
        self._update_direct_indirect_status(self.current_associate_id)

    def start_indirect(self):
        if not self.current_associate_id:
            messagebox.showerror("Error", "Load an associate first.")
            return
        start_session(self.current_associate_id, "INDIRECT", self.area_var.get(), self.role_var.get())
        self.refresh_view()
        self._update_direct_indirect_status(self.current_associate_id)

    def end_current(self):
        if not self.current_associate_id:
            messagebox.showerror("Error", "Load an associate first.")
            return
        end_active_session(self.current_associate_id)
        self.refresh_view()
        self._update_direct_indirect_status(self.current_associate_id)

    # ---------- DASHBOARD (LOCAL DATA) ----------

    def open_dashboard(self):
        dash = tk.Toplevel(self.root)
        dash.title("Dashboard - All Associates (Local DB)")
        dash.configure(bg="#1a1a1a")
        dash.geometry("800x500")

        columns = ("name", "badge", "indirect_hours", "indirect_roles", "status")
        tree = ttk.Treeview(dash, columns=columns, show="headings",
                            style="Dark.Treeview")
        for col in columns:
            tree.heading(col, text=col.replace("_", " ").title(), anchor="w")
            tree.column(col, anchor="w", width=150)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        associates = get_all_associates()
        rows = []
        for assoc_id, name, badge in associates:
            hours = compute_indirect_hours_today(assoc_id)
            status = self.get_status_text(hours)
            did_direct = self.has_direct_today(assoc_id)
            did_indirect = self.has_indirect_today(assoc_id)
            roles = self.get_indirect_roles_today(assoc_id)
            roles_text = ", ".join(sorted(roles)) if roles else "None"
            rows.append((hours, name, badge, roles_text, status, did_direct, did_indirect, assoc_id))

        rows.sort(reverse=True, key=lambda x: x[0])

        for hours, name, badge, roles_text, status, did_direct, did_indirect, assoc_id in rows:
            if hours >= INDIRECT_LIMIT_HOURS:
                tag = "violation"
            elif hours >= WARNING_THRESHOLD_HOURS:
                tag = "warning"
            elif did_direct and did_indirect:
                tag = "both"
            elif did_direct:
                tag = "direct"
            elif did_indirect:
                tag = "indirect"
            else:
                tag = "none"

            tree.insert(
                "",
                "end",
                values=(name, badge, f"{hours:.2f}", roles_text, status),
                tags=(tag,)
            )

        tree.tag_configure("direct", background="#003366")
        tree.tag_configure("indirect", background="#665511")
        tree.tag_configure("both", background="#CC6600")
        tree.tag_configure("none", background="#444444")
        tree.tag_configure("warning", background="#665511")
        tree.tag_configure("violation", background="#661111")

    # ---------- EXPORTS ----------

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save as"
        )
        if not file_path:
            return

        # If we have FCLM path data, export that; otherwise local
        if self._fclm_path_aas:
            self._export_fclm_csv(file_path)
        else:
            self._export_local_csv(file_path)

    def _export_local_csv(self, file_path):
        associates = get_all_associates()
        with open(file_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Name", "Badge ID", "Indirect Hours Today", "Indirect Roles", "Status"])
            for assoc_id, name, badge in associates:
                hours = compute_indirect_hours_today(assoc_id)
                status = self.get_status_text(hours)
                roles = self.get_indirect_roles_today(assoc_id)
                roles_text = ", ".join(sorted(roles)) if roles else "None"
                writer.writerow([name, badge, f"{hours:.2f}", roles_text, status])
        messagebox.showinfo("Export", "Local data exported successfully.")

    def _export_fclm_csv(self, file_path):
        with open(file_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Path", "Name", "Badge ID", "Hours", "Status"])
            for path in FCLM_RESTRICTED_PATHS:
                aas = self._fclm_path_aas.get(path, [])
                short = FCLM_PATH_SHORT_NAMES.get(path, path)
                for aa in aas:
                    hrs = aa["hours"]
                    mins = hrs * 60
                    if mins >= MPV_MAX_TIME_MINUTES:
                        status = "VIOLATION"
                    elif mins >= MPV_MAX_TIME_MINUTES - 60:
                        status = "Near limit"
                    else:
                        status = "OK"
                    writer.writerow([short, aa["name"], aa["badge_id"], f"{hrs:.2f}", status])
        messagebox.showinfo("Export", "FCLM path data exported successfully.")

    def export_shift_report(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save shift report as"
        )
        if not file_path:
            return

        associates = get_all_associates()
        today = datetime.date.today().isoformat()

        with open(file_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Date", "Name", "Badge ID", "Total Indirect Hours", "Indirect Roles", "Status"])
            for assoc_id, name, badge in associates:
                hours = compute_indirect_hours_today(assoc_id)
                status = self.get_status_text(hours)
                roles = self.get_indirect_roles_today(assoc_id)
                roles_text = ", ".join(sorted(roles)) if roles else "None"
                writer.writerow([today, name, badge, f"{hours:.2f}", roles_text, status])

        messagebox.showinfo("Shift Report", "Shift report exported successfully.")

    # ---------- VIEW / STATUS ----------

    def refresh_view(self):
        if not self.current_associate_id:
            return

        for row in self.tree.get_children():
            self.tree.delete(row)

        sessions = get_today_sessions(self.current_associate_id)
        now = datetime.datetime.now()
        active_path = "No active session"

        for _, start, end, work_type, area, role in sessions:
            start_dt = datetime.datetime.fromisoformat(start)
            end_dt = datetime.datetime.fromisoformat(end) if end else now
            duration_hours = (end_dt - start_dt).total_seconds() / 3600.0
            self.tree.insert("", "end", values=(
                start_dt.strftime("%H:%M"),
                end_dt.strftime("%H:%M") if end else "ACTIVE",
                work_type,
                area,
                role,
                f"{duration_hours:.2f}"
            ))
            if end is None:
                active_path = f"{work_type} > {area} > {role}"

        hours = compute_indirect_hours_today(self.current_associate_id)
        self.indirect_hours_var.set(f"{hours:.2f}")

        status = self.get_status_text(hours)
        self.status_var.set(f"{status} | Path: {active_path}")
        self.apply_status_color(hours)

        if WARNING_THRESHOLD_HOURS <= hours < INDIRECT_LIMIT_HOURS:
            messagebox.showwarning(
                "Warning",
                f"Associate is at {hours:.2f} indirect hours (approaching 6h violation)."
            )
        elif hours >= INDIRECT_LIMIT_HOURS:
            messagebox.showerror(
                "Violation",
                f"Associate is at {hours:.2f} indirect hours (VIOLATION threshold reached)."
            )

    def get_status_text(self, hours):
        if hours >= INDIRECT_LIMIT_HOURS:
            return "VIOLATION (MPV risk)"
        elif hours >= WARNING_THRESHOLD_HOURS:
            return "Near violation"
        else:
            return "OK"

    def apply_status_color(self, hours):
        if hours >= INDIRECT_LIMIT_HOURS:
            color = "#ff3333"
        elif hours >= WARNING_THRESHOLD_HOURS:
            color = "#ffcc00"
        else:
            color = "#00cc66"
        self.status_label.config(fg=color)

    # ============================================================
    #                   FCLM INTEGRATION
    # ============================================================

    def _on_cookie_refreshed(self, new_cookie):
        """Called by FclmClient when it silently refreshes an expired cookie."""
        self.config["fclm_cookie"] = new_cookie
        save_config(self.config)
        # Update the banner on the main thread (this may be called from a bg thread)
        try:
            self.root.after(0, self._update_fclm_status_label)
        except Exception:
            pass

    def _update_fclm_status_label(self):
        if self.fclm.is_connected():
            self.fclm_status_label.config(text="FCLM: Connected", fg="#2ecc71")
        else:
            self.fclm_status_label.config(text="FCLM: Not configured", fg="#e74c3c")

    # ---------- FCLM SETTINGS ----------

    def fclm_open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("FCLM Settings")
        win.configure(bg="#1a1a1a")
        win.geometry("550x560")
        win.resizable(False, False)

        tk.Label(win, text="FCLM Connection Settings", bg="#1a1a1a", fg="#ff9900",
                 font=("Segoe UI", 14, "bold")).pack(pady=(15, 5))

        # --- Auto-detect section (primary) ---
        auto_frame = tk.Frame(win, bg="#222222", highlightbackground="#ff9900",
                              highlightthickness=1)
        auto_frame.pack(fill="x", padx=20, pady=(10, 5))

        tk.Label(auto_frame, text="Automatic Setup", bg="#222222", fg="#ff9900",
                 font=("Segoe UI", 12, "bold")).pack(pady=(10, 2))
        tk.Label(auto_frame,
                 text="Log into fclm-portal.amazon.com in your browser,\n"
                      "then click the button below.",
                 bg="#222222", fg="#cccccc", font=("Segoe UI", 10),
                 justify="center").pack(pady=(0, 8))

        # Status
        status_var = tk.StringVar(value="")
        status_label = tk.Label(win, textvariable=status_var, bg="#1a1a1a",
                                fg="#aaaaaa", font=("Segoe UI", 10),
                                wraplength=500, justify="left")
        status_label.pack(pady=5)

        # Warehouse ID
        wh_frame = tk.Frame(win, bg="#1a1a1a")
        wh_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(wh_frame, text="Warehouse ID:", bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).pack(side="left")
        wh_var = tk.StringVar(value=self.config.get("fclm_warehouse_id", "IND8"))
        wh_entry = tk.Entry(wh_frame, textvariable=wh_var, width=10,
                            bg="#333333", fg="#e6e6e6", insertbackground="white", relief="flat")
        wh_entry.pack(side="left", padx=(10, 0))

        def grab_from_browser():
            status_var.set("Searching for FCLM cookies in your browser...")
            status_label.config(fg="#aaaaaa")
            win.update()

            cookie, info = BrowserCookieReader.auto_detect()
            if cookie:
                # Fill the cookie field and auto-save
                cookie_text.delete("1.0", "end")
                cookie_text.insert("1.0", cookie)
                wh = wh_var.get().strip() or "IND8"
                self.config["fclm_cookie"] = cookie
                self.config["fclm_warehouse_id"] = wh
                save_config(self.config)
                self.fclm = FclmClient(cookie=cookie, warehouse_id=wh,
                                       on_cookie_refreshed=self._on_cookie_refreshed)
                self._update_fclm_status_label()

                # Test the connection
                status_var.set(f"Found cookies from {info}. Testing connection...")
                status_label.config(fg="#aaaaaa")
                win.update()
                ok, msg = self.fclm.test_connection()
                if ok:
                    status_var.set(f"Connected via {info}!")
                    status_label.config(fg="#2ecc71")
                else:
                    status_var.set(
                        f"Cookies found in {info} but connection failed:\n{msg}\n\n"
                        "Try logging into FCLM in your browser and clicking again."
                    )
                    status_label.config(fg="#e74c3c")
            else:
                # info contains the error message
                status_var.set(
                    f"{info}\n\n"
                    "Make sure you are logged into fclm-portal.amazon.com\n"
                    "in Edge, Chrome, or Firefox, then try again."
                )
                status_label.config(fg="#e74c3c")

        grab_btn = tk.Button(auto_frame, text="Grab Cookies from Browser",
                             bg="#ff9900", fg="#1a1a1a",
                             font=("Segoe UI", 12, "bold"), relief="flat",
                             padx=20, pady=8, cursor="hand2",
                             command=grab_from_browser)
        grab_btn.pack(pady=(0, 12))

        # --- Manual paste section (secondary / fallback) ---
        tk.Label(win, text="Or paste cookie manually:", bg="#1a1a1a", fg="#888888",
                 font=("Segoe UI", 9)).pack(anchor="w", padx=20, pady=(10, 2))
        cookie_text = tk.Text(win, height=4, bg="#333333", fg="#e6e6e6",
                              insertbackground="white", relief="flat",
                              font=("Consolas", 9), wrap="word")
        cookie_text.pack(fill="x", padx=20)
        cookie_text.insert("1.0", self.config.get("fclm_cookie", ""))

        # Buttons
        btn_frame = tk.Frame(win, bg="#1a1a1a")
        btn_frame.pack(pady=10)

        def test_connection():
            cookie = cookie_text.get("1.0", "end").strip()
            wh = wh_var.get().strip()
            if not cookie:
                status_var.set("Enter a cookie first.")
                status_label.config(fg="#e74c3c")
                return
            status_var.set("Testing connection...")
            status_label.config(fg="#aaaaaa")
            win.update()

            client = FclmClient(cookie=cookie, warehouse_id=wh)
            ok, msg = client.test_connection()
            status_var.set(msg)
            status_label.config(fg="#2ecc71" if ok else "#e74c3c")

        def save_settings():
            cookie = cookie_text.get("1.0", "end").strip()
            wh = wh_var.get().strip() or "IND8"
            self.config["fclm_cookie"] = cookie
            self.config["fclm_warehouse_id"] = wh
            save_config(self.config)
            self.fclm = FclmClient(cookie=cookie, warehouse_id=wh,
                                   on_cookie_refreshed=self._on_cookie_refreshed)
            self._update_fclm_status_label()
            status_var.set("Settings saved!")
            status_label.config(fg="#2ecc71")

        def clear_cookie():
            cookie_text.delete("1.0", "end")
            self.config["fclm_cookie"] = ""
            save_config(self.config)
            self.fclm = FclmClient(cookie="", warehouse_id=wh_var.get().strip() or "IND8",
                                   on_cookie_refreshed=self._on_cookie_refreshed)
            self._update_fclm_status_label()
            status_var.set("Cookie cleared.")
            status_label.config(fg="#aaaaaa")

        test_btn = tk.Button(btn_frame, text="Test Connection", bg="#3498db", fg="white",
                             font=("Segoe UI", 10, "bold"), relief="flat", padx=12, pady=4,
                             command=test_connection)
        test_btn.pack(side="left", padx=5)

        save_btn = tk.Button(btn_frame, text="Save", bg="#27ae60", fg="white",
                             font=("Segoe UI", 10, "bold"), relief="flat", padx=12, pady=4,
                             command=save_settings)
        save_btn.pack(side="left", padx=5)

        clear_btn = tk.Button(btn_frame, text="Clear", bg="#e74c3c", fg="white",
                              font=("Segoe UI", 10, "bold"), relief="flat", padx=12, pady=4,
                              command=clear_cookie)
        clear_btn.pack(side="left", padx=5)

    # ---------- FCLM BADGE LOOKUP ----------

    def fclm_lookup_badge(self):
        """Prompt for badge ID and do FCLM lookup."""
        badge = self.badge_var.get().strip()
        if not badge:
            messagebox.showerror("Error", "Enter a badge ID first.")
            return
        if not self.fclm.is_connected():
            messagebox.showerror("Error",
                                 "FCLM not connected.\n\nGo to FCLM Settings to configure your cookie.")
            return
        self._fclm_lookup_employee(badge)

    def _fclm_lookup_employee(self, badge_id):
        """Fetch employee time details from FCLM in a background thread."""
        self.status_var.set(f"Looking up {badge_id} in FCLM...")
        self.status_label.config(fg="#3498db")

        # Also create/load the local associate
        assoc_id = get_or_create_associate(badge_id, self.name_var.get().strip() or None)
        self.current_associate_id = assoc_id

        def _fetch():
            try:
                data = self.fclm.fetch_employee_time_details(badge_id)
                self.root.after(0, lambda: self._display_fclm_employee(data, assoc_id))
            except Exception as e:
                self.root.after(0, lambda: self._fclm_lookup_error(badge_id, str(e)))

        thread = threading.Thread(target=_fetch, daemon=True)
        thread.start()

    def _display_fclm_employee(self, data, assoc_id):
        """Display FCLM employee data in the main view."""
        self._fclm_employee_data = data

        # Clear and populate the tree with FCLM sessions
        for row in self.tree.get_children():
            self.tree.delete(row)

        total_indirect_mins = 0.0
        restricted_paths_worked = set()

        # Consolidate sessions by classified path so segmented gantt
        # entries (e.g. 9 segments of "Liquidations Pick") become one row.
        consolidated = {}  # key -> {work_type, area, role, label, mins, first_start, last_end}
        display_order = []

        for s in data["sessions"]:
            title = s["title"]
            dur_mins = s.get("duration_minutes", 0)

            # Classify the session
            rp = classify_fclm_session(title)
            mapping = FCLM_PATH_MAP.get(rp or title, FCLM_PATH_MAP.get(title, None))

            if rp:
                restricted_paths_worked.add(rp)
                work_type = "INDIRECT"
                area = mapping["area"] if mapping else "CRET"
                role = mapping["role"] if mapping else "Water Spider"
                total_indirect_mins += dur_mins
                group_key = rp  # Group all segments of same restricted path
                label = role
            elif mapping:
                work_type = mapping["type"]
                area = mapping["area"]
                role = mapping["role"]
                if work_type == "INDIRECT":
                    total_indirect_mins += dur_mins
                group_key = title  # Group by exact title for mapped paths
                label = role if work_type == "INDIRECT" else title[:20]
            else:
                work_type = "DIRECT"
                area = "--"
                role = "--"
                group_key = title  # Group by exact title for direct work
                label = title[:20]

            if group_key in consolidated:
                entry = consolidated[group_key]
                entry["mins"] += dur_mins
                # Track the last end time (empty = still active)
                s_end = s.get("end", "") or ""
                if not s_end:
                    entry["last_end"] = ""  # ACTIVE
                elif entry["last_end"] != "":
                    entry["last_end"] = s_end
            else:
                consolidated[group_key] = {
                    "work_type": work_type,
                    "area": area,
                    "label": label,
                    "mins": dur_mins,
                    "first_start": s.get("start", "--"),
                    "last_end": s.get("end", "--") or "",
                }
                display_order.append(group_key)

        for key in display_order:
            entry = consolidated[key]
            dur_h = entry["mins"] / 60.0
            self.tree.insert("", "end", values=(
                entry["first_start"],
                entry["last_end"] or "ACTIVE",
                entry["work_type"],
                entry["area"],
                entry["label"],
                f"{dur_h:.2f}",
            ))

        # Update hours and status
        indirect_hours = total_indirect_mins / 60.0
        self.indirect_hours_var.set(f"{indirect_hours:.2f}")

        # Build status text
        paths_text = ", ".join(FCLM_PATH_SHORT_NAMES.get(p, p) for p in restricted_paths_worked)
        current = data.get("current_activity")
        current_text = current["title"] if current else "None"

        if indirect_hours >= INDIRECT_LIMIT_HOURS:
            self.status_var.set(f"VIOLATION | Paths: {paths_text or 'None'} | Current: {current_text}")
            self.status_label.config(fg="#ff3333")
        elif indirect_hours >= WARNING_THRESHOLD_HOURS:
            self.status_var.set(f"Near violation | Paths: {paths_text or 'None'} | Current: {current_text}")
            self.status_label.config(fg="#ffcc00")
        else:
            self.status_var.set(f"OK | Paths: {paths_text or 'None'} | Current: {current_text}")
            self.status_label.config(fg="#2ecc71")

        # Show hours on task if available
        if data.get("hours_on_task"):
            hot = data["hours_on_task"]
            tsh = data["total_scheduled_hours"]
            self.status_var.set(self.status_var.get() + f" | HoT: {hot:.1f}/{tsh:.1f}")

        total_sessions = len(data["sessions"])
        if total_sessions == 0:
            self.status_var.set(f"No FCLM sessions found for badge {data['employee_id']}")
            self.status_label.config(fg="#aaaaaa")

    def _fclm_lookup_error(self, badge_id, error_msg):
        """Handle FCLM lookup failure."""
        self.status_var.set(f"FCLM lookup failed for {badge_id}: {error_msg}")
        self.status_label.config(fg="#e74c3c")
        # Fall back to local data
        if self.current_associate_id:
            self.refresh_view()

    # ---------- FCLM SYNC ALL PATHS ----------

    def fclm_sync_paths(self):
        """Fetch all associates on restricted paths from FCLM."""
        if not self.fclm.is_connected():
            messagebox.showerror("Error",
                                 "FCLM not connected.\n\nGo to FCLM Settings to configure your cookie.")
            return

        self.status_var.set("Syncing restricted path data from FCLM...")
        self.status_label.config(fg="#3498db")

        def _fetch():
            try:
                data = self.fclm.fetch_all_path_aas()
                self.root.after(0, lambda: self._on_path_sync_done(data))
            except Exception as e:
                self.root.after(0, lambda: self._on_path_sync_error(str(e)))

        thread = threading.Thread(target=_fetch, daemon=True)
        thread.start()

    def _on_path_sync_done(self, data):
        errors = data.pop("_errors", [])
        self._fclm_path_aas = data
        total = sum(len(aas) for aas in data.values())
        paths_with_data = sum(1 for aas in data.values() if aas)

        if errors and total == 0:
            error_detail = "\n".join(errors)
            self.status_var.set(f"FCLM sync: 0 AAs found ({len(errors)} errors)")
            self.status_label.config(fg="#e74c3c")
            messagebox.showwarning("FCLM Sync Issues",
                                   f"No AA data returned.\n\nErrors:\n{error_detail}")
            return

        msg = f"FCLM sync complete: {total} AAs on {paths_with_data} restricted paths"
        if errors:
            msg += f" ({len(errors)} process errors)"
        self.status_var.set(msg)
        self.status_label.config(fg="#2ecc71")

        # Auto-open FCLM dashboard
        self.fclm_open_dashboard()

    def _on_path_sync_error(self, error_msg):
        self.status_var.set(f"FCLM sync failed: {error_msg}")
        self.status_label.config(fg="#e74c3c")
        messagebox.showerror("FCLM Sync Error", f"Failed to sync from FCLM:\n\n{error_msg}")

    # ---------- FCLM DASHBOARD ----------

    def fclm_open_dashboard(self):
        """Open a dashboard showing all associates on restricted paths (from FCLM)."""
        if not self._fclm_path_aas:
            if not self.fclm.is_connected():
                messagebox.showerror("Error",
                                     "FCLM not connected.\n\nGo to FCLM Settings to configure your cookie.")
                return
            # Auto-sync first
            self.fclm_sync_paths()
            return

        dash = tk.Toplevel(self.root)
        dash.title("FCLM Dashboard - Restricted Path Associates")
        dash.configure(bg="#1a1a1a")
        dash.geometry("900x600")

        # Title
        tk.Label(dash, text="Restricted Path Associates (from FCLM)", bg="#1a1a1a",
                 fg="#ff9900", font=("Segoe UI", 14, "bold")).pack(pady=(10, 5))

        # Timestamp
        tk.Label(dash, text=f"Synced: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                 bg="#1a1a1a", fg="#888888", font=("Segoe UI", 9)).pack()

        # Create notebook with tabs for each path
        notebook = ttk.Notebook(dash)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Summary tab
        summary_frame = tk.Frame(notebook, bg="#1a1a1a")
        notebook.add(summary_frame, text="Summary")

        summary_cols = ("path", "count", "over_limit")
        summary_tree = ttk.Treeview(summary_frame, columns=summary_cols, show="headings",
                                    style="Dark.Treeview")
        summary_tree.heading("path", text="Restricted Path", anchor="w")
        summary_tree.heading("count", text="Associates", anchor="w")
        summary_tree.heading("over_limit", text="Over 4.5h Limit", anchor="w")
        summary_tree.column("path", width=250)
        summary_tree.column("count", width=100)
        summary_tree.column("over_limit", width=150)
        summary_tree.pack(fill="both", expand=True, padx=5, pady=5)

        for path in FCLM_RESTRICTED_PATHS:
            aas = self._fclm_path_aas.get(path, [])
            if not aas:
                continue
            short = FCLM_PATH_SHORT_NAMES.get(path, path)
            over = sum(1 for a in aas if a["minutes"] >= MPV_MAX_TIME_MINUTES)
            tag = "violation" if over > 0 else "ok"
            summary_tree.insert("", "end", values=(short, len(aas), over), tags=(tag,))

        summary_tree.tag_configure("violation", background="#661111")
        summary_tree.tag_configure("ok", background="#222222")

        # Per-path tabs
        for path in FCLM_RESTRICTED_PATHS:
            aas = self._fclm_path_aas.get(path, [])
            if not aas:
                continue

            short = FCLM_PATH_SHORT_NAMES.get(path, path)
            frame = tk.Frame(notebook, bg="#1a1a1a")
            notebook.add(frame, text=f"{short} ({len(aas)})")

            cols = ("name", "badge", "hours", "status")
            tree = ttk.Treeview(frame, columns=cols, show="headings",
                                style="Dark.Treeview")
            tree.heading("name", text="Name", anchor="w")
            tree.heading("badge", text="Badge ID", anchor="w")
            tree.heading("hours", text="Hours", anchor="w")
            tree.heading("status", text="Status", anchor="w")
            tree.column("name", width=200)
            tree.column("badge", width=120)
            tree.column("hours", width=100)
            tree.column("status", width=150)
            tree.pack(fill="both", expand=True, padx=5, pady=5)

            for aa in aas:
                hrs = aa["hours"]
                mins = aa["minutes"]
                if mins >= MPV_MAX_TIME_MINUTES:
                    status = "OVER LIMIT"
                    tag = "violation"
                elif mins >= MPV_MAX_TIME_MINUTES - 60:
                    remaining = MPV_MAX_TIME_MINUTES - mins
                    status = f"{_fmt_mins(remaining)} remaining"
                    tag = "warning"
                else:
                    remaining = MPV_MAX_TIME_MINUTES - mins
                    status = f"{_fmt_mins(remaining)} remaining"
                    tag = "ok"

                tree.insert("", "end", values=(
                    aa["name"], aa["badge_id"], f"{hrs:.2f}", status
                ), tags=(tag,))

            tree.tag_configure("violation", background="#661111")
            tree.tag_configure("warning", background="#665511")
            tree.tag_configure("ok", background="#222222")

            # Lookup on double-click
            def _on_double_click(event, t=tree):
                sel = t.selection()
                if sel:
                    vals = t.item(sel[0], "values")
                    if vals and len(vals) >= 2:
                        self.badge_var.set(vals[1])  # badge ID
                        self._fclm_lookup_employee(vals[1])

            tree.bind("<Double-1>", _on_double_click)

        # Refresh button
        btn_frame = tk.Frame(dash, bg="#1a1a1a")
        btn_frame.pack(pady=(0, 10))

        def _refresh():
            dash.destroy()
            self.fclm_sync_paths()

        tk.Button(btn_frame, text="Refresh from FCLM", bg="#3498db", fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", padx=12, pady=4,
                  command=_refresh).pack(side="left", padx=5)

        def _export():
            fp = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                title="Export FCLM path data"
            )
            if fp:
                self._export_fclm_csv(fp)

        tk.Button(btn_frame, text="Export to CSV", bg="#27ae60", fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", padx=12, pady=4,
                  command=_export).pack(side="left", padx=5)


# ----------------- MAIN -----------------

def main():
    init_db()
    root = tk.Tk()
    root.withdraw()
    show_splash(root)
    root.after(2000, lambda: root.deiconify())
    app = App(root, CONFIG)
    root.mainloop()

if __name__ == "__main__":
    main()
