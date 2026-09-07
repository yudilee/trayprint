"""
TrayPrint — Flask API Server

Provides a local REST API for printing, job queue management, printer discovery,
and hub synchronization. Used by the TrayPrint tray application and Windows service.

Logging is configured via log_utils with rotating file handler support.
"""

import json
import os
import sys
import uuid
import time
import subprocess
import platform as _platform
import tempfile
from datetime import datetime
from collections import OrderedDict
import threading

# ── Helper: suppress console window for subprocesses on Windows ──
if sys.platform == 'win32':
    import subprocess as _sp
    _BASE_SPAWN = _sp.CREATE_NO_WINDOW if hasattr(_sp, 'CREATE_NO_WINDOW') else 0
else:
    _BASE_SPAWN = 0

def _run_hidden(args, **kwargs):
    """Run a subprocess without showing a console window (Windows only)."""
    kwargs.setdefault('creationflags', 0)
    kwargs['creationflags'] |= _BASE_SPAWN
    return subprocess.run(args, **kwargs)

start_time = time.time()

import printer
from path_utils import get_root_dir, get_data_dir
from logger import get_logger, get_log_path
from log_utils import setup_logging as setup_log_utils, get_log_path as get_log_utils_path

# Initialize logging using log_utils (with config-based rotation settings)
_config_for_logging = None
try:
    _config_path = os.path.join(get_root_dir(), "config.json")
    if os.path.exists(_config_path):
        with open(_config_path, "r") as f:
            _config_for_logging = json.load(f)
except Exception:
    pass

log = setup_log_utils(
    name="trayprint",
    config=_config_for_logging,
)

# Try to import psutil for memory/disk monitoring (optional)
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False
    log.info("psutil not available — memory/disk monitoring will use fallback methods")

# ─────────────────────────────────────────────
#  Shared Config Loading (merged from both locations)
# ─────────────────────────────────────────────

def _load_config():
    """Load config.json merging from install dir first, then user data dir.

    When installed via MSI (frozen):
      - Bundled config is at get_root_dir()         → C:\\Program Files\\PrintHub\\TrayPrint\\config.json
      - User settings are saved to get_data_dir()  → %LOCALAPPDATA%\\TrayPrint\\config.json

    This ensures that user settings saved via the Settings UI take precedence
    and override the installation-wide defaults.
    """
    config = {}
    # Install dir as fallback (bundled defaults)
    install_config = os.path.join(get_root_dir(), 'config.json')
    if os.path.exists(install_config):
        try:
            with open(install_config, 'r') as f:
                config.update(json.load(f))
        except Exception as e:
            log.debug("Could not load install config from %s: %s", install_config, e)
    # User data dir takes priority (where Settings UI saves)
    user_config = os.path.join(get_data_dir(), 'config.json')
    if os.path.exists(user_config):
        try:
            with open(user_config, 'r') as f:
                config.update(json.load(f))
        except Exception as e:
            log.debug("Could not load user config from %s: %s", user_config, e)
    return config

# ─────────────────────────────────────────────
#  Printer Config Merge Support
# ─────────────────────────────────────────────

def load_printer_configs():
    """Load per-printer saved configs from config.json (merged from both locations)."""
    try:
        data = _load_config()
        return data.get('printer_configs', {})
    except Exception as e:
        log.error("Error loading printer configs: %s", e)
    return {}

def merge_printer_config(printer_name, options):
    """
    Merge saved per-printer control fields into the given options dict.
    
    Strategy:
      1. Look up printer_name in saved printer_configs from config.json
      2. If found, use those values as base defaults
      3. Override with any profile / request-level options
      4. Profile / request options take precedence
    """
    if not printer_name:
        return options

    printer_configs = load_printer_configs()
    saved = printer_configs.get(printer_name)
    if not saved:
        return options  # No saved config for this printer

    # Saved config keys that map to printer control fields
    control_keys = {'tray_source', 'color_mode', 'print_quality',
                    'scaling_percentage', 'media_type', 'collate', 'reverse_order'}

    # Start with saved config as base (only control fields)
    base = {k: v for k, v in saved.items() if k in control_keys}

    # Override with incoming options (profile / request values take precedence)
    merged = {**base, **options}
    log.info("Merged printer config for '%s': base=%s, override=%s → merged=%s",
             printer_name, base, options, merged)
    return merged

APP_VERSION = "3.0.2"


# ─────────────────────────────────────────────
#  Job Queue (in-memory, thread-safe)
# ─────────────────────────────────────────────

import sqlite3

class JobQueue:
    """Thread-safe persistent print job tracker."""

    def __init__(self, db_name='jobs.db', max_history=50):
        self._lock = threading.Lock()
        self.db_path = os.path.join(get_data_dir(), db_name)
        self._max = max_history
        self._init_db()

    def _init_db(self):
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute('''
                    CREATE TABLE IF NOT EXISTS jobs (
                        id TEXT PRIMARY KEY,
                        printer TEXT,
                        type TEXT,
                        options TEXT,
                        status TEXT,
                        error TEXT,
                        created_at TEXT,
                        completed_at TEXT,
                        data_preview TEXT,
                        _raw_data TEXT
                    )
                ''')
                conn.commit()

    def create(self, printer_name, job_type, options=None, data_preview='', job_id=None):
        if not job_id:
            job_id = str(uuid.uuid4())[:8]
        
        now = datetime.now().isoformat()
        job = {
            'id': job_id,
            'printer': printer_name,
            'type': job_type,
            'options': options or {},
            'status': 'pending',
            'error': None,
            'created_at': now,
            'completed_at': None,
            'data_preview': data_preview[:80] if data_preview else '',
        }
        
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute('''
                    INSERT INTO jobs (id, printer, type, options, status, error, created_at, completed_at, data_preview)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (job_id, printer_name, job_type, json.dumps(options or {}), 'pending', None, now, None, job['data_preview']))
                
                conn.execute('''
                    DELETE FROM jobs WHERE id NOT IN (
                        SELECT id FROM jobs ORDER BY created_at DESC LIMIT ?
                    )
                ''', (self._max,))
                conn.commit()
        return job

    def complete(self, job_id, success, error_msg=''):
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                status = 'success' if success else 'failed'
                error = error_msg if not success else None
                now = datetime.now().isoformat()
                conn.execute('''
                    UPDATE jobs SET status = ?, error = ?, completed_at = ? WHERE id = ?
                ''', (status, error, now, job_id))
                conn.commit()

    def _row_to_dict(self, row):
        if not row: return None
        return {
            'id': row[0],
            'printer': row[1],
            'type': row[2],
            'options': json.loads(row[3]) if row[3] else {},
            'status': row[4],
            'error': row[5],
            'created_at': row[6],
            'completed_at': row[7],
            'data_preview': row[8]
        }

    def get(self, job_id):
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                cur = conn.execute('SELECT * FROM jobs WHERE id = ?', (job_id,))
                return self._row_to_dict(cur.fetchone())

    def list_recent(self, limit=50):
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                cur = conn.execute('SELECT * FROM jobs ORDER BY created_at ASC LIMIT ?', (limit,))
                return [self._row_to_dict(row) for row in cur.fetchall()]

    def get_job_data(self, job_id):
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                cur = conn.execute('SELECT _raw_data FROM jobs WHERE id = ?', (job_id,))
                row = cur.fetchone()
                return row[0] if row else None

    def store_job_data(self, job_id, data):
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute('UPDATE jobs SET _raw_data = ? WHERE id = ?', (data, job_id))
                conn.commit()


# ─────────────────────────────────────────────
#  Offline Job Store
#  Buffers completed jobs locally when the hub is unreachable,
#  and replays them once connectivity is restored.
# ─────────────────────────────────────────────

class OfflineJobStore:
    """Thread-safe SQLite store for jobs that could not be reported to the hub."""

    def __init__(self, db_name='offline_jobs.db'):
        self.db_path = os.path.join(get_data_dir(), db_name)
        self._lock = threading.Lock()
        self._init_db()

    def _init_db(self):
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute("""
                    CREATE TABLE IF NOT EXISTS offline_jobs (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        job_id TEXT NOT NULL,
                        printer TEXT NOT NULL,
                        job_type TEXT NOT NULL,
                        status TEXT NOT NULL,
                        error TEXT,
                        options TEXT,
                        created_at TEXT NOT NULL,
                        completed_at TEXT NOT NULL,
                        stored_at TEXT NOT NULL DEFAULT (datetime('now')),
                        retry_count INTEGER NOT NULL DEFAULT 0
                    )
                """)
                conn.commit()

    def store(self, job):
        """Store a completed job for later sync to hub."""
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute(
                    """INSERT INTO offline_jobs (job_id, printer, job_type, status, error, options, created_at, completed_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        job.get('id', job.get('job_id', '')),
                        job.get('printer', ''),
                        job.get('type', ''),
                        job.get('status', ''),
                        job.get('error'),
                        json.dumps(job.get('options', {})),
                        job.get('created_at', ''),
                        job.get('completed_at', ''),
                    )
                )
                conn.commit()

    def list_pending(self):
        """Return all unsynced jobs ordered by stored_at."""
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                rows = conn.execute(
                    "SELECT * FROM offline_jobs ORDER BY stored_at ASC"
                ).fetchall()
                return [dict(r) for r in rows]

    def mark_synced(self, job_id):
        """Remove a job from the offline store after successful sync."""
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute("DELETE FROM offline_jobs WHERE job_id = ?", (job_id,))
                conn.commit()

    def increment_retry(self, job_id):
        """Increment retry count for a failed sync attempt."""
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute(
                    "UPDATE offline_jobs SET retry_count = retry_count + 1 WHERE job_id = ?",
                    (job_id,)
                )
                conn.commit()

    def count(self):
        """Return the number of pending offline jobs."""
        with self._lock:
            with sqlite3.connect(self.db_path) as conn:
                row = conn.execute("SELECT COUNT(*) as cnt FROM offline_jobs").fetchone()
                return row[0] if row else 0


# Global job queue
_job_queue = JobQueue()
_offline_store = OfflineJobStore()
_notification_callback = None

# ── Queue refresh request (for WebSocket-triggered immediate poll) ──
_queue_refresh_request = threading.Event()

# ── Printing state (for tray indicator) ──
_is_printing = False
_printing_lock = threading.Lock()

# ── Cancel flags: job_id → threading.Event ──
_cancel_flags = {}       # type: dict[str, threading.Event]
_cancel_flags_lock = threading.Lock()

def is_printing():
    """Returns True if a job is currently being processed/printed."""
    with _printing_lock:
        return _is_printing

def _set_printing(state: bool):
    """Set the printing flag (thread-safe)."""
    with _printing_lock:
        global _is_printing
        _is_printing = state

def request_queue_refresh():
    """Signal the sync loop to poll the queue immediately on its next iteration."""
    _queue_refresh_request.set()


def get_queue_status():
    """
    Returns a list of active/pending/recent jobs suitable for the queue dialog.
    Each entry: {job_id, document_name, status, printer, created_at, type, error}
    """
    jobs = _job_queue.list_recent(50)
    pending = [j for j in jobs if j['status'] == 'pending']
    completed = [j for j in jobs if j['status'] != 'pending']
    result = list(reversed(pending)) + list(reversed(completed))
    return result

def cancel_job(job_id):
    """
    Cancel a print job.
    - Pending: remove from queue via a cancel flag
    - Printing: signal the running thread to abort
    Returns True if cancel was initiated, False if job not found.
    """
    job = _job_queue.get(job_id)
    if not job:
        return False

    with _cancel_flags_lock:
        if job_id not in _cancel_flags:
            _cancel_flags[job_id] = threading.Event()
        _cancel_flags[job_id].set()

    log.info("Cancel flag set for job %s (status=%s)", job_id, job['status'])
    return True

def _check_cancelled(job_id):
    """Check if a job has been cancelled. Returns True if cancelled."""
    with _cancel_flags_lock:
        flag = _cancel_flags.get(job_id)
        if flag and flag.is_set():
            return True
    return False

def _cleanup_cancel_flag(job_id):
    """Remove cancel flag after job completes."""
    with _cancel_flags_lock:
        _cancel_flags.pop(job_id, None)


_allowed_origins = ["http://127.0.0.1:*", "http://localhost:*"]


# ─────────────────────────────────────────────
#  Profiles (synced from hub or local config)
# ─────────────────────────────────────────────

_profiles = {}

def load_profiles_from_config():
    """Load profiles from config.json (merged from both locations)."""
    global _profiles
    try:
        data = _load_config()
        _profiles = data.get('profiles', {})
        log.info("Loaded %d queue(s) from config", len(_profiles))
    except Exception as e:
        log.error("Error loading profiles: %s", e)

def save_profiles_to_config(profiles):
    """Persist profiles back to config.json."""
    global _profiles
    _profiles = profiles
    config_path = os.path.join(get_data_dir(), 'config.json')
    try:
        with open(config_path, 'r') as f:
            data = json.load(f)
        data['profiles'] = profiles
        with open(config_path, 'w') as f:
            json.dump(data, f, indent=2)
        log.info("Saved %d queue(s) to config", len(profiles))
    except Exception as e:
        log.error("Error saving profiles: %s", e)

def get_profiles():
    return _profiles


# ─────────────────────────────────────────────
#  Hub Response Validation
# ─────────────────────────────────────────────

def _check_hub_response(json_data, context=""):
    """Validate Hub API response. Returns parsed data dict or None on failure."""
    if not isinstance(json_data, dict):
        log.error("Hub %s: invalid response type: %s", context, type(json_data).__name__)
        return None
    if not json_data.get('success', False):
        error_msg = json_data.get('error', {}).get('message', json_data.get('error', 'Unknown error'))
        log.error("Hub %s: API error — %s", context, error_msg)
        return None
    return json_data.get('data', {})


# ─────────────────────────────────────────────
#  Hub Sync & Spooler (background threads)
# ─────────────────────────────────────────────
import queue as _queue
_internal_print_queue = _queue.Queue()
_hub_last_status = "Disconnected"
_hub_sync_running = False
_hub_sync_stop_event = threading.Event()
_cached_printer_count = 0
_cached_printer_count_lock = threading.Lock()

# Hot-reloadable sync config — updated by reload_config() so the sync loop
# picks up new interval/retry values without restarting.
_sync_config = {
    "interval": 60,
    "max_retries": 3,
    "retry_delay": 60,
}
_sync_config_lock = threading.Lock()

def get_hub_status():
    return _hub_last_status

def get_cached_printer_count():
    """Return the cached printer count, falling back to direct enumeration if not yet seeded."""
    global _cached_printer_count
    with _cached_printer_count_lock:
        if _cached_printer_count == 0:
            try:
                printers_list = printer.get_printers()
                _cached_printer_count = len(printers_list)
                log.info("get_cached_printer_count: fallback enumeration found %d printer(s)", _cached_printer_count)
            except Exception as e:
                log.debug("get_cached_printer_count: fallback failed: %s", e)
        return _cached_printer_count

def start_hub_sync(hub_url, agent_key, interval, max_retries=3, retry_delay=60):
    """Periodically pull profiles and print queue from the central hub.
    
    If a previous sync loop is running, it will be signalled to stop
    before starting a new one. This allows hot-reloading the hub URL
    or agent key without restarting the entire application.
    """
    global _hub_sync_running
    # Signal the old sync loop to stop (if running)
    _hub_sync_stop_event.set()
    # Give it a moment to notice the stop signal
    import time as _time
    _time.sleep(0.5)
    # Clear the stop event for the new loop
    _hub_sync_stop_event.clear()
    import requests
    import base64

    hub_url = hub_url.rstrip('/')

    # Seed the hot-reloadable sync config
    with _sync_config_lock:
        _sync_config["interval"] = interval
        _sync_config["max_retries"] = max_retries
        _sync_config["retry_delay"] = retry_delay

    def _get_sync_interval():
        """Read the current sync interval from the hot-reloadable config."""
        with _sync_config_lock:
            return _sync_config.get("interval", interval)

    def _get_max_retries():
        """Read the current max retries from the hot-reloadable config."""
        with _sync_config_lock:
            return _sync_config.get("max_retries", max_retries)

    def _get_retry_delay():
        """Read the current retry delay from the hot-reloadable config."""
        with _sync_config_lock:
            return _sync_config.get("retry_delay", retry_delay)

    def sync_loop():
        # Send initial heartbeat + status report immediately so the hub marks
        # this agent as online without waiting for the first interval cycle.
        try:
            log.info("Sync loop started (hub_url=%s) — sending initial heartbeat", hub_url)
            report_status_to_hub(hub_url, agent_key)
            headers_init = {'Authorization': f'Bearer {agent_key}'}
            hb_url = f'{hub_url}/api/print-hub/heartbeat'
            log.info("Sending heartbeat to %s", hb_url)
            resp = requests.post(hb_url, headers=headers_init, timeout=5)
            if resp and resp.status_code == 200:
                log.info("Initial heartbeat sent — agent marked online on hub")
            else:
                code = resp.status_code if resp else 'no response'
                log.warning("Initial heartbeat returned HTTP %s", code)
        except Exception as e:
            log.warning("Initial heartbeat/status failed: %s", e)

        current_interval = _get_sync_interval()
        profile_counter = current_interval
        status_counter = 0  # Reset to 0 since we just sent initial status
        backoff = 1
        max_backoff = 30  # Cap at 30s so heartbeats stay within hub's 2-min online threshold

        while True:
            jobs_found = False
            try:
                headers = {'Authorization': f'Bearer {agent_key}'}

                # Report status (printers) every sync interval
                if status_counter >= interval:
                    report_status_to_hub(hub_url, agent_key)
                    status_counter = 0

                # Heartbeat — lightweight keepalive
                try:
                    resp = requests.post(f'{hub_url}/api/print-hub/heartbeat', headers=headers, timeout=5)
                    if resp and resp.status_code == 200:
                        log.debug("Heartbeat sent")
                except Exception:
                    pass  # heartbeat failures are non-critical

                # Fast polling for queue
                resp_queue = requests.get(f'{hub_url}/api/print-hub/queue', headers=headers, timeout=5)
                jobs_found = False
                if resp_queue.status_code == 200:
                    json_data = resp_queue.json()
                    data = _check_hub_response(json_data, "queue")
                    if data is not None:
                        global _hub_last_status
                        _hub_last_status = "Connected"
                        jobs = data.get('jobs', [])
                        for j in jobs:
                            log.info("Pulled job %s from hub.", j['job_id'])
                            _internal_print_queue.put(j)
                        jobs_found = len(jobs) > 0

                        # Flush any offline-buffered jobs now that we're connected
                        try:
                            _flush_offline_jobs(hub_url, agent_key)
                        except Exception as offline_e:
                            log.debug("Offline job flush error: %s", offline_e)

                # Slower polling for profiles and CORS
                if profile_counter >= interval:
                    resp_prof = requests.get(f'{hub_url}/api/print-hub/profiles', headers=headers, timeout=10)
                    if resp_prof.status_code == 200:
                        json_data = resp_prof.json()
                        data = _check_hub_response(json_data, "getProfiles")
                        if data is not None:
                            _hub_last_status = "Connected"
                            # Hub returns profiles array → convert to dict keyed by name
                            new_profiles_list = data.get('profiles', [])
                            new_profiles = {}
                            for p in new_profiles_list:
                                name = p.get('name', '')
                                if name:
                                    new_profiles[name] = p
                            log.info("Synced %d profile(s) from hub", len(new_profiles))
                            save_profiles_to_config(new_profiles)

                    try:
                        resp_cors = requests.get(f'{hub_url}/api/print-hub/cors-origins', headers=headers, timeout=10)
                        if resp_cors.status_code == 200:
                            json_data = resp_cors.json()
                            data = _check_hub_response(json_data, "cors-origins")
                            if data is not None:
                                origins = data.get('allowed_origins')
                                if origins is not None:
                                    global _allowed_origins
                                    _allowed_origins = origins
                    except Exception as ce:
                        log.debug("Hub CORS sync failed: %s", ce)

                    profile_counter = 0

            except Exception as e:
                log.debug("Hub sync failed (hub may be offline): %s", e)
                _hub_last_status = "Offline"

            # Check for queue refresh request (e.g., from WebSocket event)
            if _queue_refresh_request.is_set():
                _queue_refresh_request.clear()
                log.debug("Queue refresh requested, resetting backoff")
                backoff = 1

            # Re-read sync interval from hot-reloadable config on each iteration
            current_interval = _get_sync_interval()

            # Exponential backoff
            if jobs_found:
                backoff = 1
            else:
                backoff = min(backoff * 2, max_backoff)

            # Check if we should stop (e.g., on sync restart with new hub URL)
            if _hub_sync_stop_event.is_set():
                log.info("Sync loop stopping (restart requested)")
                _hub_sync_running = False
                return

            time.sleep(backoff)
            profile_counter += backoff
            status_counter += backoff

    def spooler_loop():
        retry_queue = _queue.Queue()

        def _process_job(job_data):
            """Process a single print job with retries. Returns True on success."""
            hub_job = job_data
            job_id = hub_job['job_id']
            printer_name = hub_job['printer']
            job_type = hub_job['type']
            options = hub_job['options'] or {}
            b64_data = hub_job.get('document_base64')
            current_max_retries = _get_max_retries()
            current_retry_delay = _get_retry_delay()

            # Skip jobs pending approval
            approval_status = hub_job.get('approval_status', 'auto_approved')
            if approval_status == 'pending':
                log.info("Job %s skipped — pending approval", job_id)
                return False

            # Resolve profile from job options if available
            profile_name = options.get('profile') or options.get('queue')
            if profile_name and profile_name in _profiles:
                profile = _profiles[profile_name]
                # Profile printer takes precedence if no explicit printer was set
                if not printer_name or printer_name == 'default':
                    printer_name = profile.get('printer', printer_name)
                # Merge profile fields as base (metadata keys excluded), request options override
                meta_keys = {'id', 'name', 'printer', 'description', 'print_agent_id'}
                profile_opts = {k: v for k, v in profile.items() if k not in meta_keys and v is not None}
                merged = {**profile_opts, **options}
                options = merged
                log.info("Hub job %s: resolved profile '%s' → printer=%s, %d option(s)",
                         job_id, profile_name, printer_name, len(options))

            # Merge per-printer saved config (local defaults) — profile/request values override
            if printer_name:
                options = merge_printer_config(printer_name, options)

            if not b64_data:
                return False

            raw_data = base64.b64decode(b64_data)

            # Skip if already processed (non-pending)
            existing = _job_queue.get(job_id)
            if existing and existing['status'] != 'pending':
                return True  # already handled

            # Check if cancelled before we even start
            if _check_cancelled(job_id):
                log.info("Job %s cancelled before starting", job_id)
                _job_queue.complete(job_id, False, 'Cancelled')
                _cleanup_cancel_flag(job_id)
                return False

            job = _job_queue.create(printer_name, job_type, options, '(Pulled from Hub)', job_id=job_id)

            success = False
            error_msg = ''

            # Set printing flag
            _set_printing(True)

            try:
                for attempt in range(current_max_retries):
                    # Check cancel flag before each attempt
                    if _check_cancelled(job_id):
                        log.info("Job %s cancelled (attempt %d)", job_id, attempt + 1)
                        error_msg = 'Cancelled'
                        success = False
                        break

                    try:
                        if job_type == 'pdf':
                            success, error_msg = printer.print_pdf(printer_name, b64_data, options)
                        else:
                            success, error_msg = printer.print_raw(printer_name, raw_data, options)

                        if success:
                            break
                    except Exception as e:
                        success = False
                        error_msg = str(e)

                    if not success and attempt < current_max_retries - 1:
                        # Check cancel flag during retry delay (poll every 1s)
                        log.warning("Print failed. Retrying in %ds... (%d/%d)", current_retry_delay, attempt + 1, current_max_retries)
                        for _ in range(current_retry_delay):
                            if _check_cancelled(job_id):
                                log.info("Job %s cancelled during retry wait", job_id)
                                error_msg = 'Cancelled'
                                success = False
                                break
                            time.sleep(1)
                        if error_msg == 'Cancelled':
                            break
            finally:
                _set_printing(False)

            _job_queue.complete(job_id, success, error_msg)
            _cleanup_cancel_flag(job_id)

            if success:
                if _notification_callback:
                    _notification_callback("Hub Print Success", f"Printed on {printer_name}")
            elif error_msg == 'Cancelled':
                # Don't retry cancelled jobs
                if _notification_callback:
                    _notification_callback("Hub Print Cancelled", f"Cancelled job on {printer_name}")
            else:
                # All retries exhausted — move to retry queue for later
                retry_queue.put(job_data)
                if _notification_callback:
                    _notification_callback("Hub Print Failed", f"Failed on {printer_name}: {error_msg}")

            report_job_to_hub(hub_url, agent_key, _job_queue.get(job_id))
            return success

        while True:
            # Process retry queue items first (non-blocking)
            retry_items = []
            while not retry_queue.empty():
                try:
                    retry_items.append(retry_queue.get_nowait())
                except _queue.Empty:
                    break

            for job_data in retry_items:
                log.info("Retrying previously failed job %s", job_data.get('job_id', 'unknown'))
                _process_job(job_data)

            # Process new jobs from hub (non-blocking with timeout)
            try:
                hub_job = _internal_print_queue.get(timeout=1)
                _process_job(hub_job)
                _internal_print_queue.task_done()
            except _queue.Empty:
                time.sleep(1)

    global _hub_sync_running
    _hub_sync_running = True
    threading.Thread(target=sync_loop, daemon=True).start()
    threading.Thread(target=spooler_loop, daemon=True).start()
    log.info("Hub sync & spooler started → %s (every %ds)", hub_url, interval)

def report_status_to_hub(hub_url, agent_key):
    """Report local status (printers + capabilities + hardware telemetry) to the central hub."""
    import requests
    if not hub_url:
        return
    try:
        headers = {'Authorization': f'Bearer {agent_key}',
                   'Content-Type': 'application/json'}
        # Fetch current printers from the OS
        printers_list = printer.get_printers()
        global _cached_printer_count
        _cached_printer_count = len(printers_list)

        # Build capabilities & hardware status for each printer
        capabilities_dict = {}
        hardware_status_dict = {}
        try:
            import capabilities as caps_mod
            for p in printers_list:
                name = p.get('name', '')
                if name:
                    try:
                        caps = caps_mod.discover_capabilities(name)
                        if 'error' not in caps:
                            capabilities_dict[name] = caps
                    except Exception:
                        pass
                    try:
                        if hasattr(printer, 'get_printer_hardware_status'):
                            hardware_status_dict[name] = printer.get_printer_hardware_status(name)
                    except Exception:
                        pass
        except Exception:
            pass

        payload = {
            'printers': [p['name'] for p in printers_list],
            'capabilities': capabilities_dict,
            'hardware_status': hardware_status_dict,
        }
        
        # Post to telemetry endpoint (syncs printers, capabilities and hardware health)
        resp = requests.post(f'{hub_url}/api/print-hub/telemetry', json=payload, headers=headers, timeout=10)
        global _hub_last_status
        if resp.status_code == 200:
            json_data = resp.json()
            data = _check_hub_response(json_data, "telemetry report")
            if data is not None:
                _hub_last_status = "Connected"
                log.info("Reported %d printers (+ capabilities & telemetry) to hub", len(printers_list))
            else:
                _hub_last_status = "Offline (API error)"
        else:
            # Fallback to legacy status endpoint if telemetry returned 404
            if resp.status_code == 404:
                resp = requests.post(f'{hub_url}/api/print-hub/status', json=payload, headers=headers, timeout=10)
            if resp.status_code == 200:
                _hub_last_status = "Connected"
            else:
                _hub_last_status = f"Offline ({resp.status_code})"
                log.warning("Hub rejected status report (HTTP %d): %s", resp.status_code, resp.text)
    except Exception as e:
        _hub_last_status = "Offline"
        log.debug("Failed to report status to hub: %s", e)

def report_job_to_hub(hub_url, agent_key, job):
    """Report a completed job back to the central hub (fire-and-forget).
    
    Falls back to the offline job store if the hub is unreachable.
    Offline jobs are automatically replayed when connectivity is restored.
    """
    import requests

    def _report():
        try:
            headers = {'Authorization': f'Bearer {agent_key}',
                       'Content-Type': 'application/json'}
            payload = {
                'job_id': job['id'],
                'printer': job['printer'],
                'type': job['type'],
                'status': job['status'],
                'error': job.get('error'),
                'options': job.get('options', {}),
                'created_at': job['created_at'],
                'completed_at': job['completed_at'],
            }
            resp = requests.post(f'{hub_url}/api/print-hub/jobs', json=payload, headers=headers, timeout=10)
            if resp.status_code == 200:
                json_data = resp.json()
                _check_hub_response(json_data, "job report")
            else:
                # Non-200 response — buffer offline
                log.warning("Hub returned %d for job report, buffering offline", resp.status_code)
                _offline_store.store(job)
        except requests.ConnectionError:
            log.info("Hub unreachable — buffering job %s in offline store", job.get('id', '?'))
            _offline_store.store(job)
        except Exception as e:
            log.debug("Failed to report job to hub: %s — buffering offline", e)
            _offline_store.store(job)

    threading.Thread(target=_report, daemon=True).start()


def _flush_offline_jobs(hub_url, agent_key):
    """Replay buffered offline jobs to the hub. Called from sync loop when connected."""
    import requests
    pending = _offline_store.list_pending()
    if not pending:
        return

    log.info("Flushing %d offline job(s) to hub", len(pending))
    headers = {'Authorization': f'Bearer {agent_key}',
               'Content-Type': 'application/json'}

    for job in pending:
        try:
            payload = {
                'job_id': job['job_id'],
                'printer': job['printer'],
                'type': job['job_type'],
                'status': job['status'],
                'error': job.get('error'),
                'options': json.loads(job.get('options', '{}')),
                'created_at': job['created_at'],
                'completed_at': job['completed_at'],
            }
            resp = requests.post(f'{hub_url}/api/print-hub/jobs', json=payload, headers=headers, timeout=10)
            if resp.status_code == 200:
                _offline_store.mark_synced(job['job_id'])
                log.info("Offline job %s synced to hub", job['job_id'])
            else:
                _offline_store.increment_retry(job['job_id'])
                log.warning("Failed to sync offline job %s (HTTP %d)", job['job_id'], resp.status_code)
        except Exception as e:
            _offline_store.increment_retry(job['job_id'])
            log.debug("Failed to sync offline job %s: %s", job['job_id'], e)
            break  # Stop trying on connection error — will retry next cycle


# ─────────────────────────────────────────────
#  Hub Connection Test
# ─────────────────────────────────────────────

def test_hub_connection():
    """Test if the configured hub is reachable. Returns True/False."""
    import requests
    # Use get_data_dir() to match where the Settings UI saves config.json,
    # and also fall back to get_root_dir() for the bundled config.
    config_locations = [
        os.path.join(get_data_dir(), 'config.json'),
        os.path.join(get_root_dir(), 'config.json'),
    ]
    try:
        data = {}
        for config_path in config_locations:
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    data.update(json.load(f))
        hub_url = data.get('hub_url', '').rstrip('/')
        agent_key = data.get('agent_key', '')
        if not hub_url or not agent_key:
            return False
        headers = {'Authorization': f'Bearer {agent_key}'}
        resp = requests.get(f'{hub_url}/api/print-hub/heartbeat', headers=headers, timeout=5)
        return resp.status_code == 200
    except Exception as e:
        log.debug("test_hub_connection failed: %s", e)
        return False


# ─────────────────────────────────────────────
#  Watchdog — Spooler / CUPS Monitor
# ─────────────────────────────────────────────

_watchdog_thread = None
_watchdog_status = {
    "running": False,
    "last_check": None,
    "restart_attempts": 0,
    "memory_usage": "N/A",
    "disk_space": "N/A",
    "websocket_status": "N/A",
}
_watchdog_status_lock = threading.Lock()

# Watchdog ring buffer (last 50 check results)
_watchdog_log = []
_watchdog_log_lock = threading.Lock()
WATCHDOG_LOG_MAX = 50

# Watchdog checks configuration (loaded from config.json)
_watchdog_checks_config = {
    "spooler": True,
    "memory": True,
    "disk": True,
    "websocket": True,
}
_watchdog_checks_lock = threading.Lock()


def _add_watchdog_log_entry(message):
    """Add an entry to the watchdog ring buffer."""
    with _watchdog_log_lock:
        _watchdog_log.append({
            "timestamp": datetime.now().isoformat(),
            "message": message,
        })
        # Trim to max size
        while len(_watchdog_log) > WATCHDOG_LOG_MAX:
            _watchdog_log.pop(0)


def get_watchdog_log():
    """Return the watchdog log ring buffer."""
    with _watchdog_log_lock:
        return list(_watchdog_log)


def load_watchdog_checks_config():
    """Load watchdog_checks from config.json, falling back to defaults."""
    try:
        data = _load_config()
        checks = data.get('watchdog_checks', {})
        if checks:
            with _watchdog_checks_lock:
                for key in _watchdog_checks_config:
                    if key in checks:
                        _watchdog_checks_config[key] = bool(checks[key])
            log.info("Loaded watchdog checks config: %s", _watchdog_checks_config)
    except Exception as e:
        log.error("Error loading watchdog checks config: %s", e)


def is_watchdog_check_enabled(check_name):
    """Check if a specific watchdog check is enabled."""
    with _watchdog_checks_lock:
        return _watchdog_checks_config.get(check_name, True)

class WatchdogThread(threading.Thread):
    """Periodically checks spooler health, memory usage, disk space, and WebSocket."""

    def __init__(self, check_interval=60):
        super().__init__(daemon=True)
        self.check_interval = check_interval
        self._stop_event = threading.Event()
        self._restart_attempts = 0
        self._consecutive_failures = 0

    def run(self):
        log.info("Watchdog started (interval=%ds)", self.check_interval)
        with _watchdog_status_lock:
            _watchdog_status["running"] = True
            _watchdog_status["last_check"] = None
            _watchdog_status["restart_attempts"] = 0
            _watchdog_status["memory_usage"] = "N/A"
            _watchdog_status["disk_space"] = "N/A"
            _watchdog_status["websocket_status"] = "N/A"

        _add_watchdog_log_entry("Watchdog started (interval=%ds)" % self.check_interval)

        while not self._stop_event.is_set():
            self._perform_check()
            self._stop_event.wait(self.check_interval)

        with _watchdog_status_lock:
            _watchdog_status["running"] = False
        _add_watchdog_log_entry("Watchdog stopped")
        log.info("Watchdog stopped")

    def stop(self):
        self._stop_event.set()

    def _check_windows_spooler(self):
        """Check if spoolsv.exe is running on Windows."""
        try:
            import psutil
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and proc.info['name'].lower() == 'spoolsv.exe':
                    return True
            return False
        except ImportError:
            # Fallback using tasklist
            try:
                result = _run_hidden(
                    ['tasklist', '/FI', 'IMAGENAME eq spoolsv.exe'],
                    capture_output=True, text=True, timeout=10
                )
                return 'spoolsv.exe' in result.stdout
            except Exception:
                return True  # Assume running if we can't check

    def _check_linux_cups(self):
        """Check if CUPS scheduler is running on Linux/macOS."""
        try:
            result = subprocess.run(
                ['lpstat', '-r'],
                capture_output=True, text=True, timeout=10
            )
            return 'scheduler is running' in result.stdout
        except Exception:
            return True  # Assume running if we can't check

    def _check_macos_cups(self):
        """Check if CUPS is running on macOS via cupsctl or lpstat."""
        try:
            result = subprocess.run(
                ['cupsctl'],
                capture_output=True, text=True, timeout=10
            )
            return result.returncode == 0
        except FileNotFoundError:
            try:
                result = subprocess.run(
                    ['lpstat', '-r'],
                    capture_output=True, text=True, timeout=10
                )
                return 'scheduler is running' in result.stdout
            except Exception:
                return True
        except Exception:
            return True

    def _check_memory_usage(self):
        """Check process memory usage via psutil or /proc/self/status."""
        try:
            if HAS_PSUTIL:
                proc = psutil.Process(os.getpid())
                mem_info = proc.memory_info()
                rss_mb = mem_info.rss / (1024 * 1024)
                vms_mb = mem_info.vms / (1024 * 1024)
                mem_percent = proc.memory_percent()
                result = f"RSS={rss_mb:.1f}MB, VMS={vms_mb:.1f}MB, {mem_percent:.1f}%"
                with _watchdog_status_lock:
                    _watchdog_status["memory_usage"] = result
                return result
            else:
                # Fallback: read /proc/self/status on Linux
                if sys.platform == 'linux':
                    with open('/proc/self/status', 'r') as f:
                        for line in f:
                            if line.startswith('VmRSS:'):
                                rss = line.split()[1]
                                result = f"RSS={rss}kB"
                                with _watchdog_status_lock:
                                    _watchdog_status["memory_usage"] = result
                                return result
                raise RuntimeError("No memory monitoring available")
        except Exception as e:
            log.debug("Watchdog memory check failed: %s", e)
            with _watchdog_status_lock:
                _watchdog_status["memory_usage"] = "Error: %s" % str(e)
            return "Error: %s" % str(e)

    def _check_disk_space(self):
        """Check that temp directory has sufficient free space (>100MB)."""
        try:
            temp_dir = tempfile.gettempdir()
            if HAS_PSUTIL:
                usage = psutil.disk_usage(temp_dir)
                free_mb = usage.free / (1024 * 1024)
                total_mb = usage.total / (1024 * 1024)
                status = "OK" if free_mb > 100 else "LOW"
                result = f"{status}: {free_mb:.0f}MB free of {total_mb:.0f}MB ({temp_dir})"
            else:
                # Fallback: use os.statvfs on Unix
                if sys.platform != 'win32':
                    st = os.statvfs(temp_dir)
                    free_mb = (st.f_bavail * st.f_frsize) / (1024 * 1024)
                else:
                    # Windows fallback: use ctypes
                    import ctypes
                    free_bytes = ctypes.c_ulonglong(0)
                    ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                        ctypes.c_wchar_p(temp_dir), None, None, ctypes.pointer(free_bytes)
                    )
                    free_mb = free_bytes.value / (1024 * 1024)
                status = "OK" if free_mb > 100 else "LOW"
                result = f"{status}: {free_mb:.0f}MB free ({temp_dir})"

            with _watchdog_status_lock:
                _watchdog_status["disk_space"] = result
            return result
        except Exception as e:
            log.debug("Watchdog disk check failed: %s", e)
            result = "Error: %s" % str(e)
            with _watchdog_status_lock:
                _watchdog_status["disk_space"] = result
            return result

    def _check_websocket_health(self):
        """Check if WebSocket client is connected and responsive."""
        try:
            # Import websocket client module to check its status
            import websocket_client as ws_mod
            # The ws_client is stored in app.py's TrayApp, but we can check
            # via the module-level reference if available
            if hasattr(ws_mod, '_ws_client_instance') and ws_mod._ws_client_instance:
                client = ws_mod._ws_client_instance
                if client.is_connected:
                    result = "Connected"
                elif client.is_running:
                    result = "Reconnecting"
                else:
                    result = "Disconnected"
            else:
                result = "Not configured"
            with _watchdog_status_lock:
                _watchdog_status["websocket_status"] = result
            return result
        except Exception as e:
            log.debug("Watchdog WebSocket check failed: %s", e)
            result = "Error: %s" % str(e)
            with _watchdog_status_lock:
                _watchdog_status["websocket_status"] = result
            return result

    def _perform_check(self):
        """Run all enabled watchdog checks."""
        check_time = datetime.now().isoformat()
        results = []

        # 1. Spooler health check
        if is_watchdog_check_enabled("spooler"):
            try:
                if _platform.system() == 'Windows':
                    alive = self._check_windows_spooler()
                elif _platform.system() == 'Darwin':
                    alive = self._check_macos_cups()
                else:
                    alive = self._check_linux_cups()

                if not alive:
                    self._consecutive_failures += 1
                    msg = "Spooler check FAILED (%d/3)" % self._consecutive_failures
                    results.append(msg)
                    log.warning("Watchdog: %s", msg)
                    if self._consecutive_failures >= 3:
                        log.warning("Watchdog: spooler appears dead, attempting restart...")
                        self._attempt_restart()
                        self._consecutive_failures = 0
                else:
                    if self._consecutive_failures > 0:
                        msg = "Spooler recovered after %d failures" % self._consecutive_failures
                        results.append(msg)
                        log.info("Watchdog: %s", msg)
                    self._consecutive_failures = 0
                    results.append("Spooler OK")
            except Exception as e:
                results.append("Spooler error: %s" % str(e))
                log.error("Watchdog spooler check error: %s", e)
        else:
            results.append("Spooler check disabled")

        # 2. Memory usage check
        if is_watchdog_check_enabled("memory"):
            try:
                mem_result = self._check_memory_usage()
                results.append("Memory: %s" % mem_result)
            except Exception as e:
                results.append("Memory error: %s" % str(e))
        else:
            results.append("Memory check disabled")

        # 3. Disk space check
        if is_watchdog_check_enabled("disk"):
            try:
                disk_result = self._check_disk_space()
                results.append("Disk: %s" % disk_result)
            except Exception as e:
                results.append("Disk error: %s" % str(e))
        else:
            results.append("Disk check disabled")

        # 4. WebSocket health check
        if is_watchdog_check_enabled("websocket"):
            try:
                ws_result = self._check_websocket_health()
                results.append("WebSocket: %s" % ws_result)
            except Exception as e:
                results.append("WebSocket error: %s" % str(e))
        else:
            results.append("WebSocket check disabled")

        # Update last check time
        with _watchdog_status_lock:
            _watchdog_status["last_check"] = check_time

        # Add combined result to log
        summary = "; ".join(results)
        _add_watchdog_log_entry("Check at %s: %s" % (check_time, summary))

    def _attempt_restart(self):
        """Try to restart the spooler service."""
        self._restart_attempts += 1
        with _watchdog_status_lock:
            _watchdog_status["restart_attempts"] = self._restart_attempts
        _add_watchdog_log_entry("Attempting spooler restart (#%d)" % self._restart_attempts)
        try:
            if _platform.system() == 'Windows':
                result = subprocess.run(
                    ['net', 'start', 'spoolsv'],
                    capture_output=True, text=True, timeout=30
                )
                if result.returncode == 0:
                    log.info("Watchdog: spoolsv.exe restarted successfully")
                    _add_watchdog_log_entry("spoolsv.exe restarted successfully")
                else:
                    log.warning("Watchdog: failed to restart spoolsv.exe: %s", result.stderr)
                    _add_watchdog_log_entry("Failed to restart spoolsv.exe: %s" % result.stderr)
            elif _platform.system() == 'Darwin':
                try:
                    result = subprocess.run(
                        ['sudo', 'launchctl', 'kickstart', '-k', 'system/org.cups.cupsd'],
                        capture_output=True, text=True, timeout=30
                    )
                except FileNotFoundError:
                    result = subprocess.run(
                        ['sudo', 'cupsctl', '--reset'],
                        capture_output=True, text=True, timeout=30
                    )
                if result.returncode == 0:
                    log.info("Watchdog: macOS CUPS restarted successfully")
                    _add_watchdog_log_entry("macOS CUPS restarted successfully")
                else:
                    log.warning("Watchdog: failed to restart macOS CUPS: %s", result.stderr)
                    _add_watchdog_log_entry("Failed to restart macOS CUPS: %s" % result.stderr)
            else:
                try:
                    result = subprocess.run(
                        ['sudo', 'systemctl', 'start', 'cups'],
                        capture_output=True, text=True, timeout=30
                    )
                except FileNotFoundError:
                    result = subprocess.run(
                        ['sudo', 'service', 'cups', 'start'],
                        capture_output=True, text=True, timeout=30
                    )
                if result.returncode == 0:
                    log.info("Watchdog: CUPS restarted successfully")
                    _add_watchdog_log_entry("CUPS restarted successfully")
                else:
                    log.warning("Watchdog: failed to restart CUPS: %s", result.stderr)
                    _add_watchdog_log_entry("Failed to restart CUPS: %s" % result.stderr)
        except Exception as e:
            log.error("Watchdog restart failed: %s", e)
            _add_watchdog_log_entry("Restart failed: %s" % str(e))


# ─────────────────────────────────────────────
#  reload_config — Apply settings live
# ─────────────────────────────────────────────

_hub_url = ""
_agent_key = ""

def reload_config():
    """Re-read config.json and update runtime state without restarting."""
    global _hub_url, _agent_key, _allowed_origins
    try:
        data = _load_config()
        _hub_url = data.get('hub_url', '')
        _agent_key = data.get('agent_key', '')
        if 'allowed_origins' in data:
            _allowed_origins = data.get('allowed_origins', ["http://127.0.0.1:*", "http://localhost:*"])
        load_profiles_from_config()
        load_watchdog_checks_config()

        # Update hot-reloadable sync config so new interval/retry values take effect immediately
        with _sync_config_lock:
            new_interval = data.get('sync_interval_seconds', _sync_config.get("interval", 60))
            new_max_retries = data.get('max_retries', _sync_config.get("max_retries", 3))
            new_retry_delay = data.get('retry_delay_seconds', _sync_config.get("retry_delay", 60))
            _sync_config["interval"] = new_interval
            _sync_config["max_retries"] = new_max_retries
            _sync_config["retry_delay"] = new_retry_delay
            log.info("Sync config hot-reloaded: interval=%ds, max_retries=%d, retry_delay=%ds",
                     new_interval, new_max_retries, new_retry_delay)

        # Signal the sync loop to pick up new interval on next iteration
        request_queue_refresh()

        log.info("Config reloaded: hub_url=%s, profiles=%d", _hub_url, len(_profiles))
        return True
    except Exception as e:
        log.error("Error reloading config: %s", e)
    return False

def get_watchdog_status():
    """Return the current watchdog status dict including log entries."""
    with _watchdog_status_lock:
        status = dict(_watchdog_status)
    status['log'] = get_watchdog_log()
    status['checks_config'] = dict(_watchdog_checks_config)
    return status


# ─────────────────────────────────────────────
#  Flask App Factory
# ─────────────────────────────────────────────

def get_resource_path(relative_path):
    import sys
    import os
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath(os.path.dirname(__file__)), relative_path)

def create_app():
    from flask import Flask, request, jsonify, render_template
    app = Flask(__name__, template_folder=get_resource_path('templates'))

    # Load settings (merged from both user data dir and install dir)
    config_data = {"port": 49211, "allowed_origins": ["*"]}
    try:
        config_data.update(_load_config())
    except Exception as e:
        log.error("Error loading config.json: %s", e)

    global _allowed_origins
    _allowed_origins = config_data.get('allowed_origins', ["http://127.0.0.1:*", "http://localhost:*"])

    # Load watchdog checks config
    load_watchdog_checks_config()

    @app.before_request
    def handle_options():
        if request.method == 'OPTIONS':
            resp = app.response_class()
            resp.status_code = 204
            return resp

    @app.after_request
    def add_cors_headers(response):
        origin = request.headers.get('Origin')
        if not origin:
            return response

        import re
        allowed = False
        for allowed_origin in _allowed_origins:
            if allowed_origin == '*':
                allowed = True
                break
            pattern = "^" + re.escape(allowed_origin).replace("\\*", ".*") + "$"
            if re.match(pattern, origin):
                allowed = True
                break

        if allowed:
            response.headers['Access-Control-Allow-Origin'] = origin
            response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization, X-API-Key, X-Agent-Key'
            response.headers['Access-Control-Allow-Methods'] = 'GET, PUT, POST, DELETE, OPTIONS'

        return response

    # Load local profiles
    load_profiles_from_config()

    # Store hub config for job-reporting (sync is started from app.py)
    hub_url = config_data.get('hub_url', '')
    agent_key = config_data.get('agent_key', '')
    app.config['HUB_URL'] = hub_url
    app.config['AGENT_KEY'] = agent_key

    # ── Routes ──

    @app.route('/status', methods=['GET'])
    def status():
        pending = len([j for j in _job_queue.list_recent(50) if j['status'] == 'pending'])
        uptime = int(time.time() - start_time)
        return jsonify({
            "status": "running",
            "version": APP_VERSION,
            "uptime_seconds": uptime,
            "pending_jobs": pending
        }), 200

    @app.route('/printers', methods=['GET'])
    def list_printers():
        printers_list = printer.get_printers()
        return jsonify({"printers": printers_list}), 200

    @app.route('/profiles', methods=['GET'])
    @app.route('/queues', methods=['GET'])
    def list_profiles():
        return jsonify({"profiles": get_profiles(), "queues": get_profiles()}), 200

    @app.route('/print', methods=['POST'])
    def handle_print():
        data = request.get_json()
        if not data:
            return jsonify({"error": "No JSON payload provided."}), 400

        printer_name = data.get('printer')
        raw_data = data.get('data')
        job_type = data.get('type', 'raw')
        options = data.get('options', {})
        queue_name = data.get('queue') or data.get('profile')

        # Resolve profile if specified
        if queue_name and queue_name in _profiles:
            profile = _profiles[queue_name]
            if not printer_name:
                printer_name = profile.get('printer', '')
            # Profile fields are at top level; filter out metadata keys,
            # use profile fields as base, request-level options override
            meta_keys = {'id', 'name', 'printer', 'description', 'print_agent_id'}
            profile_opts = {k: v for k, v in profile.items() if k not in meta_keys and v is not None}
            merged = {**profile_opts, **options}
            options = merged
            log.info("Using queue '%s' → printer=%s, options=%s", queue_name, printer_name, options)

        # Merge per-printer saved config (local defaults) — profile/request values override
        if printer_name:
            options = merge_printer_config(printer_name, options)

        if not printer_name or not raw_data:
            return jsonify({"error": "Missing 'printer' or 'data' in payload."}), 400

        # Create job
        preview = raw_data[:80] if job_type == 'raw' else '(PDF binary)'
        job = _job_queue.create(printer_name, job_type, options, preview)
        _job_queue.store_job_data(job['id'], raw_data)

        # Execute with printing flag
        _set_printing(True)
        try:
            if job_type == 'pdf':
                success, error_msg = printer.print_pdf(printer_name, raw_data, options)
            else:
                success, error_msg = printer.print_raw(printer_name, raw_data, options)
        finally:
            _set_printing(False)

        _job_queue.complete(job['id'], success, error_msg)
        if _notification_callback:
            if success:
                _notification_callback("Local Print Success", f"Printed on {printer_name}")
            else:
                _notification_callback("Local Print Failed", f"Failed on {printer_name}: {error_msg}")

        # Report to hub if configured
        if app.config['HUB_URL']:
            updated_job = _job_queue.get(job['id'])
            report_job_to_hub(app.config['HUB_URL'], app.config['AGENT_KEY'], updated_job)

        if success:
            return jsonify({"status": "success", "job_id": job['id'],
                            "message": f"Sent to {printer_name}"}), 200
        else:
            return jsonify({"status": "error", "job_id": job['id'],
                            "error": error_msg}), 500

    @app.route('/jobs', methods=['GET'])
    def list_jobs():
        limit = request.args.get('limit', 50, type=int)
        jobs = _job_queue.list_recent(limit)
        # Strip internal _raw_data from response
        clean = [{k: v for k, v in j.items() if not k.startswith('_')} for j in jobs]
        return jsonify({"jobs": clean}), 200

    @app.route('/jobs/<job_id>', methods=['GET'])
    def get_job(job_id):
        job = _job_queue.get(job_id)
        if not job:
            return jsonify({"error": "Job not found"}), 404
        clean = {k: v for k, v in job.items() if not k.startswith('_')}
        return jsonify(clean), 200

    @app.route('/jobs/<job_id>/retry', methods=['POST'])
    def retry_job(job_id):
        job = _job_queue.get(job_id)
        if not job:
            return jsonify({"error": "Job not found"}), 404

        raw_data = _job_queue.get_job_data(job_id)
        if not raw_data:
            return jsonify({"error": "Job data expired, cannot retry"}), 410

        # Re-execute
        if job['type'] == 'pdf':
            success, error_msg = printer.print_pdf(job['printer'], raw_data, job.get('options'))
        else:
            success, error_msg = printer.print_raw(job['printer'], raw_data, job.get('options'))

        _job_queue.complete(job_id, success, error_msg)
        if _notification_callback:
            if success:
                _notification_callback("Retry Print Success", f"Printed on {printer_name}")
            else:
                _notification_callback("Retry Print Failed", f"Failed on {printer_name}: {error_msg}")

        if success:
            return jsonify({"status": "success", "message": "Retry successful"}), 200
        else:
            return jsonify({"status": "error", "error": error_msg}), 500

    @app.route('/jobs/<job_id>/cancel', methods=['POST'])
    def cancel_job_route(job_id):
        """Cancel a pending or in-progress print job."""
        job = _job_queue.get(job_id)
        if not job:
            return jsonify({"error": "Job not found"}), 404

        if job['status'] not in ('pending',):
            return jsonify({"error": f"Cannot cancel job with status '{job['status']}'"}), 400

        cancelled = cancel_job(job_id)
        if cancelled:
            # If it's pending, mark it as cancelled immediately
            if job['status'] == 'pending':
                _job_queue.complete(job_id, False, 'Cancelled')
                _cleanup_cancel_flag(job_id)
                if _notification_callback:
                    _notification_callback("Job Cancelled", f"Cancelled job {job_id}")

            return jsonify({"status": "cancelled", "job_id": job_id}), 200
        else:
            return jsonify({"error": "Failed to cancel job"}), 500

    @app.route('/queue-status', methods=['GET'])
    def queue_status_route():
        """Return the current queue status for the tray UI."""
        jobs = get_queue_status()
        return jsonify({
            "jobs": jobs,
            "is_printing": is_printing()
        }), 200

    # ── Capabilities ──

    @app.route('/api/capabilities', methods=['GET'])
    def capabilities_route():
        """Return capabilities for a specific printer or all printers."""
        import capabilities as caps_mod
        printer_name = request.args.get('printer')
        if printer_name:
            caps = caps_mod.discover_capabilities(printer_name)
            return jsonify({"capabilities": caps}), 200
        else:
            all_caps = caps_mod.discover_all_printers_capabilities()
            return jsonify({"capabilities": all_caps}), 200

    @app.route('/api/capabilities/all', methods=['GET'])
    def all_capabilities_route():
        """Return capabilities for all printers on this system."""
        import capabilities as caps_mod
        all_caps = caps_mod.discover_all_printers_capabilities()
        return jsonify({"capabilities": all_caps}), 200

    # ── Diagnostics ──

    @app.route('/api/diagnostics', methods=['GET'])
    def diagnostics():
        """Return full diagnostics data for the TrayPrint agent."""
        cfg = _load_config()
        hub_url = cfg.get('hub_url', '')
        printers_list = printer.get_printers()
        queue = _job_queue.list_recent(50)
        running_uptime = time.time() - start_time
        hours = int(running_uptime // 3600)
        minutes = int((running_uptime % 3600) // 60)
        seconds = int(running_uptime % 60)
        uptime_str = f"{hours}h {minutes}m {seconds}s"

        wd_status = get_watchdog_status()

        # System info
        import socket
        system_info = {
            "os_version": _platform.platform(),
            "python_version": sys.version.split()[0],
            "architecture": _platform.machine(),
            "machine": _platform.machine(),
            "processor": _platform.processor(),
            "hostname": socket.gethostname(),
        }

        # Try to get PySide6 and PyMuPDF versions
        pyside6_version = ""
        pymupdf_version = ""
        try:
            from PySide6 import QtCore
            pyside6_version = QtCore.__version__
        except Exception:
            pass
        try:
            import fitz
            pymupdf_version = fitz.version
        except Exception:
            pass
        system_info["pyside6_version"] = pyside6_version
        system_info["pymupdf_version"] = pymupdf_version

        # WebSocket status
        websocket_status = {"connected": False, "reconnect_count": 0}
        try:
            import websocket_client as ws_mod
            if hasattr(ws_mod, '_ws_instance') and ws_mod._ws_instance:
                ws = ws_mod._ws_instance
                websocket_status["connected"] = ws.connected if hasattr(ws, 'connected') else False
                websocket_status["reconnect_count"] = ws.reconnect_count if hasattr(ws, 'reconnect_count') else 0
                websocket_status["url"] = ws.url if hasattr(ws, 'url') else ""
        except Exception:
            pass

        # Proxy config
        proxy_config = {
            "http_proxy": os.environ.get('HTTP_PROXY', '') or os.environ.get('http_proxy', ''),
            "https_proxy": os.environ.get('HTTPS_PROXY', '') or os.environ.get('https_proxy', ''),
            "no_proxy": os.environ.get('NO_PROXY', '') or os.environ.get('no_proxy', ''),
        }

        # Printer capabilities (all printers)
        printer_capabilities = {}
        try:
            import capabilities as caps_mod
            printer_capabilities = caps_mod.discover_all_printers_capabilities()
        except Exception:
            pass

        # Default printer
        default_printer = ""
        try:
            default_printer = printer.get_default_printer()
        except Exception:
            pass

        # Jobs (full list, cleaned)
        jobs = []
        try:
            raw_jobs = _job_queue.list_recent(50)
            jobs = [{k: v for k, v in j.items() if not k.startswith('_')} for j in raw_jobs]
        except Exception:
            pass

        # Logs (last 50 lines) — try log_utils path first, fall back to logger path
        logs = []
        try:
            log_path = get_log_utils_path(name="trayprint")
            if not log_path or not os.path.exists(log_path):
                log_path = get_log_path()
            if log_path and os.path.exists(log_path):
                with open(log_path, 'r', encoding='utf-8', errors='replace') as f:
                    all_lines = f.readlines()
                    logs = [line.rstrip('\n') for line in all_lines[-50:]]
        except Exception:
            pass

        # Hub reachable and last sync time
        hub_reachable = test_hub_connection()
        last_sync_time = ""
        try:
            if hasattr(_hub_sync_thread, '_last_sync_time'):
                last_sync_time = _hub_sync_thread._last_sync_time
        except Exception:
            pass

        return jsonify({
            "app_version": APP_VERSION,
            "python_version": sys.version,
            "platform": sys.platform,
            "run_mode": "packaged" if getattr(sys, 'frozen', False) else "dev",
            "hub_url": hub_url,
            "hub_connected": hub_reachable,
            "hub_reachable": hub_reachable,
            "last_sync_time": last_sync_time,
            "printer_count": len(printers_list),
            "printers": [p['name'] for p in printers_list],
            "queue_length": len(queue),
            "uptime": uptime_str,
            "log_file": get_log_path(),
            "os_details": _platform.platform(),
            "system_info": system_info,
            "websocket_status": websocket_status,
            "proxy_config": proxy_config,
            "printer_capabilities": printer_capabilities,
            "default_printer": default_printer,
            "jobs": jobs,
            "logs": logs,
            "watchdog": wd_status,
        }), 200

    # ── Watchdog Status & Log ──

    @app.route('/api/watchdog/status', methods=['GET'])
    def watchdog_status_route():
        return jsonify(get_watchdog_status()), 200

    @app.route('/api/watchdog/log', methods=['GET'])
    def watchdog_log_route():
        """Return the watchdog check history (ring buffer)."""
        return jsonify({"log": get_watchdog_log()}), 200

    return app


def run_server(port):
    global _watchdog_thread
    app = create_app()
    import logging as stdlib_logging
    werkzeug_log = stdlib_logging.getLogger('werkzeug')
    werkzeug_log.setLevel(stdlib_logging.ERROR)
    log.info("Starting API server on 127.0.0.1:%d", port)

    # Perform an initial printer check to populate the cache
    try:
        # Pre-populate printer list for the local UI even without hub
        printers_list = printer.get_printers()
        global _cached_printer_count
        with _cached_printer_count_lock:
            _cached_printer_count = len(printers_list)
        log.info("Initial printer check: found %d printer(s)", _cached_printer_count)
    except Exception:
        pass

    # Load watchdog checks config before starting watchdog
    load_watchdog_checks_config()

    # Start watchdog thread
    _watchdog_thread = WatchdogThread(check_interval=60)
    _watchdog_thread.start()

    # Register shutdown handler to stop watchdog on exit
    try:
        from flask import request
        @app.teardown_appcontext
        def shutdown_watchdog(exception=None):
            if _watchdog_thread and _watchdog_thread.is_alive():
                _watchdog_thread.stop()
                _watchdog_thread.join(timeout=3)
                log.info("Watchdog stopped during server shutdown")
    except Exception:
        pass

    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)

    # Ensure watchdog stops after server exits
    if _watchdog_thread and _watchdog_thread.is_alive():
        _watchdog_thread.stop()


if __name__ == '__main__':
    run_server(49211)
