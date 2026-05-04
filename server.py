import json
import os
import uuid
import time
from datetime import datetime
from collections import OrderedDict
import threading

START_TIME = time.time()

from flask import Flask, request, jsonify, render_template
import printer
from path_utils import get_root_dir
from logger import get_logger

log = get_logger()

# ─────────────────────────────────────────────
#  Printer Config Merge Support
# ─────────────────────────────────────────────

def load_printer_configs():
    """Load per-printer saved configs from config.json."""
    config_path = os.path.join(get_root_dir(), 'config.json')
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                data = json.load(f)
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

APP_VERSION = "3.0.0"


# ─────────────────────────────────────────────
#  Job Queue (in-memory, thread-safe)
# ─────────────────────────────────────────────

import sqlite3

class JobQueue:
    """Thread-safe persistent print job tracker."""

    def __init__(self, db_name='jobs.db', max_history=50):
        self._lock = threading.Lock()
        self.db_path = os.path.join(get_root_dir(), db_name)
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


# Global job queue
_job_queue = JobQueue()
_notification_callback = None

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
    """Load profiles from config.json."""
    global _profiles
    config_path = os.path.join(get_root_dir(), 'config.json')
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                data = json.load(f)
                _profiles = data.get('profiles', {})
                log.info("Loaded %d queue(s) from config", len(_profiles))
    except Exception as e:
        log.error("Error loading profiles: %s", e)

def save_profiles_to_config(profiles):
    """Persist profiles back to config.json."""
    global _profiles
    _profiles = profiles
    config_path = os.path.join(get_root_dir(), 'config.json')
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
_cached_printer_count = 0

def get_hub_status():
    return _hub_last_status

def get_cached_printer_count():
    return _cached_printer_count

def start_hub_sync(hub_url, agent_key, interval, max_retries=3, retry_delay=60):
    """Periodically pull profiles and print queue from the central hub."""
    import requests
    import base64

    def sync_loop():
        profile_counter = interval  # Trigger immediately
        status_counter = interval   # Trigger immediately
        backoff = 1
        max_backoff = 60

        while True:
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

            # Exponential backoff
            if jobs_found:
                backoff = 1
            else:
                backoff = min(backoff * 2, max_backoff)

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
                for attempt in range(max_retries):
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

                    if not success and attempt < max_retries - 1:
                        # Check cancel flag during retry delay (poll every 1s)
                        log.warning("Print failed. Retrying in %ds... (%d/%d)", retry_delay, attempt + 1, max_retries)
                        for _ in range(retry_delay):
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

    threading.Thread(target=sync_loop, daemon=True).start()
    threading.Thread(target=spooler_loop, daemon=True).start()
    log.info("Hub sync & spooler started → %s (every %ds)", hub_url, interval)

def report_status_to_hub(hub_url, agent_key):
    """Report local status (printers) to the central hub."""
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

        payload = {
            'printers': [p['name'] for p in printers_list]
        }
        resp = requests.post(f'{hub_url}/api/print-hub/status', json=payload, headers=headers, timeout=10)
        global _hub_last_status
        if resp.status_code == 200:
            json_data = resp.json()
            data = _check_hub_response(json_data, "status report")
            if data is not None:
                _hub_last_status = "Connected"
                log.info("Reported %d printers to hub", len(printers_list))
            else:
                _hub_last_status = "Offline (API error)"
        else:
            _hub_last_status = f"Offline ({resp.status_code})"
            log.warning("Hub rejected status report (HTTP %d): %s", resp.status_code, resp.text)
    except Exception as e:
        _hub_last_status = "Offline"
        log.debug("Failed to report status to hub: %s", e)

def report_job_to_hub(hub_url, agent_key, job):
    """Report a completed job back to the central hub (fire-and-forget)."""
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
        except Exception as e:
            log.debug("Failed to report job to hub: %s", e)

    threading.Thread(target=_report, daemon=True).start()


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
    app = Flask(__name__, template_folder=get_resource_path('templates'))

    # Load settings
    config_path = os.path.join(get_root_dir(), 'config.json')
    config_data = {"port": 49211, "allowed_origins": ["*"]}
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                config_data.update(json.load(f))
    except Exception as e:
        log.error("Error loading config.json: %s", e)

    global _allowed_origins
    _allowed_origins = config_data.get('allowed_origins', ["http://127.0.0.1:*", "http://localhost:*"])

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

    # Start hub sync if configured
    hub_url = config_data.get('hub_url', '')
    agent_key = config_data.get('agent_key', '')
    if hub_url:
        interval = config_data.get('sync_interval_seconds', 60)
        max_retries = config_data.get('max_retries', 3)
        retry_delay = config_data.get('retry_delay_seconds', 60)
        start_hub_sync(hub_url, agent_key, interval, max_retries, retry_delay)

    # Store hub config for job-reporting
    app.config['HUB_URL'] = hub_url
    app.config['AGENT_KEY'] = agent_key

    # ── Routes ──

    @app.route('/status', methods=['GET'])
    def status():
        pending = len([j for j in _job_queue.list_recent(50) if j['status'] == 'pending'])
        uptime = int(time.time() - START_TIME)
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

    return app


def run_server(port):
    app = create_app()
    import logging as stdlib_logging
    werkzeug_log = stdlib_logging.getLogger('werkzeug')
    werkzeug_log.setLevel(stdlib_logging.ERROR)
    log.info("Starting API server on 127.0.0.1:%d", port)

    # Perform an initial printer check to populate the cache
    try:
        # Pre-populate printer list for the local UI even without hub
        printer.get_printers()
    except Exception:
        pass

    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)


if __name__ == '__main__':
    run_server(49211)
