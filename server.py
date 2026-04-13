import json
import os
import uuid
import threading
from datetime import datetime
from collections import OrderedDict

from flask import Flask, request, jsonify
from flask_cors import CORS
import printer
from path_utils import get_root_dir
from logger import get_logger

log = get_logger()


# ─────────────────────────────────────────────
#  Job Queue (in-memory, thread-safe)
# ─────────────────────────────────────────────

class JobQueue:
    """Thread-safe in-memory print job tracker."""

    def __init__(self, max_history=50):
        self._lock = threading.Lock()
        self._jobs = OrderedDict()
        self._max = max_history

    def create(self, printer_name, job_type, options=None, data_preview='', job_id=None):
        if not job_id:
            job_id = str(uuid.uuid4())[:8]
        job = {
            'id': job_id,
            'printer': printer_name,
            'type': job_type,
            'options': options or {},
            'status': 'pending',
            'error': None,
            'created_at': datetime.now().isoformat(),
            'completed_at': None,
            'data_preview': data_preview[:80] if data_preview else '',
        }
        with self._lock:
            self._jobs[job_id] = job
            # Trim old jobs
            while len(self._jobs) > self._max:
                self._jobs.popitem(last=False)
        return job

    def complete(self, job_id, success, error_msg=''):
        with self._lock:
            job = self._jobs.get(job_id)
            if job:
                job['status'] = 'success' if success else 'failed'
                job['error'] = error_msg if not success else None
                job['completed_at'] = datetime.now().isoformat()

    def get(self, job_id):
        with self._lock:
            return self._jobs.get(job_id)

    def list_recent(self, limit=50):
        with self._lock:
            items = list(self._jobs.values())
        return items[-limit:]

    def get_job_data(self, job_id):
        """Returns the raw data for retry. We store it transiently."""
        with self._lock:
            return self._jobs.get(job_id, {}).get('_raw_data')

    def store_job_data(self, job_id, data):
        with self._lock:
            job = self._jobs.get(job_id)
            if job:
                job['_raw_data'] = data


# Global job queue
_job_queue = JobQueue()


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
                log.info("Loaded %d profile(s) from config", len(_profiles))
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
        log.info("Saved %d profile(s) to config", len(profiles))
    except Exception as e:
        log.error("Error saving profiles: %s", e)

def get_profiles():
    return _profiles


# ─────────────────────────────────────────────
#  Hub Sync & Spooler (background threads)
# ─────────────────────────────────────────────
import queue
_internal_print_queue = queue.Queue()
_hub_last_status = "Disconnected"
_cached_printer_count = 0

def get_hub_status():
    return _hub_last_status

def get_cached_printer_count():
    return _cached_printer_count

def start_hub_sync(hub_url, agent_key, interval, max_retries=3, retry_delay=60):
    """Periodically pull profiles and print queue from the central hub."""
    import time
    import requests
    import base64

    def sync_loop():
        profile_counter = interval # Trigger immediately
        status_counter = interval # Trigger immediately
        while True:
            try:
                headers = {'Authorization': f'Bearer {agent_key}'}
                
                # Report status (printers) every sync interval
                if status_counter >= interval:
                    report_status_to_hub(hub_url, agent_key)
                    status_counter = 0

                # Fast polling for queue
                resp_queue = requests.get(f'{hub_url}/api/print-hub/queue', headers=headers, timeout=5)
                if resp_queue.status_code == 200:
                    global _hub_last_status
                    _hub_last_status = "Connected"
                    jobs = resp_queue.json().get('jobs', [])
                    for j in jobs:
                        log.info("Pulled job %s from hub.", j['job_id'])
                        _internal_print_queue.put(j)
                
                # Slower polling for profiles
                if profile_counter >= interval:
                    resp_prof = requests.get(f'{hub_url}/api/print-hub/profiles', headers=headers, timeout=10)
                    if resp_prof.status_code == 200:
                        _hub_last_status = "Connected"
                        new_profiles = resp_prof.json().get('profiles', {})
                        save_profiles_to_config(new_profiles)
                    profile_counter = 0

            except Exception as e:
                log.debug("Hub sync failed (hub may be offline): %s", e)
                _hub_last_status = "Offline"
            
            # Fast polling for queue every 5s
            time.sleep(5)
            profile_counter += 5
            status_counter += 5

    def spooler_loop():
        while True:
            hub_job = _internal_print_queue.get()
            job_id = hub_job['job_id']
            printer_name = hub_job['printer']
            job_type = hub_job['type']
            options = hub_job['options'] or {}
            b64_data = hub_job.get('document_base64')

            if not b64_data:
                _internal_print_queue.task_done()
                continue
                
            raw_data = base64.b64decode(b64_data)
            
            # Create local job record
            if job_id in _job_queue._jobs:
                _internal_print_queue.task_done()
                continue

            job = _job_queue.create(printer_name, job_type, options, '(Pulled from Hub)', job_id=job_id)
            
            success = False
            error_msg = ''
            
            for attempt in range(max_retries + 1):
                try:
                    if job_type == 'pdf':
                        # PASS BASE64 STRING DIRECTLY (Fixing double-decoding bug)
                        success, error_msg = printer.print_pdf(printer_name, b64_data, options)
                    else:
                        # For raw, we use the decoded bytes
                        success, error_msg = printer.print_raw(printer_name, raw_data, options)
                        
                    if success:
                        break
                except Exception as e:
                    success = False
                    error_msg = str(e)
                    
                if not success and attempt < max_retries:
                    log.warning("Print failed. Retrying in %ds... (%d/%d)", retry_delay, attempt+1, max_retries)
                    time.sleep(retry_delay)

            _job_queue.complete(job_id, success, error_msg)
            report_job_to_hub(hub_url, agent_key, _job_queue.get(job_id))
            _internal_print_queue.task_done()

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
            _hub_last_status = "Connected"
            log.info("Reported %d printers to hub", len(printers_list))
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
            requests.post(f'{hub_url}/api/print-hub/jobs', json=payload, headers=headers, timeout=10)
        except Exception as e:
            log.debug("Failed to report job to hub: %s", e)

    threading.Thread(target=_report, daemon=True).start()


# ─────────────────────────────────────────────
#  Flask App Factory
# ─────────────────────────────────────────────

def create_app():
    app = Flask(__name__)

    # Load settings
    config_path = os.path.join(get_root_dir(), 'config.json')
    config_data = {"port": 49211, "allowed_origins": ["*"]}
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                config_data.update(json.load(f))
    except Exception as e:
        log.error("Error loading config.json: %s", e)

    allowed_origins = config_data.get('allowed_origins', ['*'])
    CORS(app, origins=allowed_origins)

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
        return jsonify({"status": "running", "version": "2.0.0"}), 200

    @app.route('/printers', methods=['GET'])
    def list_printers():
        printers_list = printer.get_printers()
        return jsonify({"printers": printers_list}), 200

    @app.route('/profiles', methods=['GET'])
    def list_profiles():
        return jsonify({"profiles": get_profiles()}), 200

    @app.route('/print', methods=['POST'])
    def handle_print():
        data = request.get_json()
        if not data:
            return jsonify({"error": "No JSON payload provided."}), 400

        printer_name = data.get('printer')
        raw_data = data.get('data')
        job_type = data.get('type', 'raw')
        options = data.get('options', {})
        profile_name = data.get('profile')

        # Resolve profile if specified
        if profile_name and profile_name in _profiles:
            profile = _profiles[profile_name]
            if not printer_name:
                printer_name = profile.get('printer', '')
            # Merge profile options (profile is base, request overrides)
            merged = {**profile.get('options', {}), **options}
            options = merged
            log.info("Using profile '%s' → printer=%s, options=%s", profile_name, printer_name, options)

        if not printer_name or not raw_data:
            return jsonify({"error": "Missing 'printer' or 'data' in payload."}), 400

        # Create job
        preview = raw_data[:80] if job_type == 'raw' else '(PDF binary)'
        job = _job_queue.create(printer_name, job_type, options, preview)
        _job_queue.store_job_data(job['id'], raw_data)

        # Execute
        if job_type == 'pdf':
            success, error_msg = printer.print_pdf(printer_name, raw_data, options)
        else:
            success, error_msg = printer.print_raw(printer_name, raw_data, options)

        _job_queue.complete(job['id'], success, error_msg)

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

        if success:
            return jsonify({"status": "success", "message": "Retry successful"}), 200
        else:
            return jsonify({"status": "error", "error": error_msg}), 500

    # ── Web Settings Dashboard ──

    @app.route('/settings', methods=['GET'])
    def settings_page():
        config = {}
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
        except:
            pass

        printers_list = printer.get_printers()
        jobs = _job_queue.list_recent(20)
        clean_jobs = [{k: v for k, v in j.items() if not k.startswith('_')} for j in jobs]

        hub_status = "Not configured"
        hub_color = "#8b8fa3"
        if config.get('hub_url'):
            try:
                import requests as req
                headers = {'Authorization': f'Bearer {config.get("agent_key", "")}'}
                resp = req.get(f'{config["hub_url"]}/api/print-hub/profiles', headers=headers, timeout=3)
                if resp.status_code == 200:
                    hub_status = f"Connected ({len(resp.json().get('profiles', {}))} profiles)"
                    hub_color = "#22c55e"
                elif resp.status_code == 401:
                    hub_status = "Auth failed - check Agent Key"
                    hub_color = "#ef4444"
                else:
                    hub_status = f"Error (HTTP {resp.status_code})"
                    hub_color = "#ef4444"
            except:
                hub_status = "Cannot reach hub"
                hub_color = "#f59e0b"

        return f'''<!DOCTYPE html>
<html><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Trayprint Settings</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root {{ --bg:#0f1117; --surface:#1a1d27; --border:#2a2e3f; --text:#e4e6ed; --muted:#8b8fa3; --primary:#6366f1; --success:#22c55e; --danger:#ef4444; --warning:#f59e0b; }}
* {{ margin:0;padding:0;box-sizing:border-box; }}
body {{ font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);padding:2rem;max-width:900px;margin:0 auto; }}
h1 {{ font-size:1.8rem;font-weight:700;background:linear-gradient(135deg,var(--primary),#a855f7);-webkit-background-clip:text;-webkit-text-fill-color:transparent;margin-bottom:.25rem; }}
.sub {{ color:var(--muted);font-size:.85rem;margin-bottom:2rem; }}
.card {{ background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:1.5rem;margin-bottom:1.5rem; }}
.card h2 {{ font-size:1rem;font-weight:600;margin-bottom:1rem;display:flex;align-items:center;gap:.5rem; }}
.grid {{ display:grid;grid-template-columns:1fr 1fr;gap:1rem; }}
label {{ display:block;font-size:.8rem;font-weight:500;color:var(--muted);margin-bottom:.3rem; }}
input,select {{ width:100%;padding:.55rem .75rem;background:var(--bg);border:1px solid var(--border);border-radius:6px;color:var(--text);font-size:.85rem;font-family:inherit; }}
input:focus {{ outline:none;border-color:var(--primary); }}
.btn {{ display:inline-flex;align-items:center;gap:.4rem;padding:.6rem 1.2rem;border-radius:6px;font-size:.85rem;font-weight:600;border:none;cursor:pointer;transition:all .15s;text-decoration:none; }}
.btn-primary {{ background:var(--primary);color:white; }}
.btn-primary:hover {{ background:#818cf8; }}
.btn-secondary {{ background:var(--border);color:var(--text); }}
.btn-secondary:hover {{ background:#353952; }}
.badge {{ display:inline-block;padding:.15rem .5rem;border-radius:100px;font-size:.7rem;font-weight:600; }}
.badge-ok {{ background:rgba(34,197,94,.15);color:var(--success); }}
.badge-err {{ background:rgba(239,68,68,.15);color:var(--danger); }}
.badge-warn {{ background:rgba(245,158,11,.15);color:var(--warning); }}
.badge-info {{ background:rgba(99,102,241,.15);color:var(--primary); }}
table {{ width:100%;border-collapse:collapse;font-size:.85rem; }}
th {{ text-align:left;padding:.6rem .75rem;color:var(--muted);font-weight:500;font-size:.75rem;text-transform:uppercase;border-bottom:1px solid var(--border); }}
td {{ padding:.6rem .75rem;border-bottom:1px solid var(--border); }}
tr:hover {{ background:rgba(99,102,241,.04); }}
.mono {{ font-family:monospace;font-size:.8rem;background:var(--bg);padding:.15rem .4rem;border-radius:3px; }}
.dot {{ display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:5px; }}
.dot-g {{ background:var(--success);box-shadow:0 0 6px var(--success); }}
.dot-r {{ background:var(--danger); }}
.msg {{ padding:.75rem 1rem;border-radius:6px;margin-bottom:1rem;font-size:.85rem;display:none; }}
.form-group {{ margin-bottom:1rem; }}
</style>
</head><body>

<h1>Trayprint Settings</h1>
<p class="sub">v2.0.0 &middot; Local Print Service running on port {config.get("port", 49211)}</p>

<div id="msg" class="msg"></div>

<!-- Hub Connection -->
<div class="card">
  <h2>Hub Connection <span class="badge {"badge-ok" if hub_color=="#22c55e" else "badge-err" if hub_color=="#ef4444" else "badge-warn"}" style="margin-left:auto;">{hub_status}</span></h2>
  <form id="settingsForm">
    <div class="grid">
      <div class="form-group">
        <label>Print Hub URL</label>
        <input type="text" name="hub_url" id="hub_url" value="{config.get("hub_url","")}" placeholder="http://192.168.1.100:8082">
      </div>
      <div class="form-group">
        <label>Agent Key</label>
        <input type="text" name="agent_key" id="agent_key" value="{config.get("agent_key","")}" placeholder="Paste from Hub > Agents page">
      </div>
    </div>
    <div class="grid">
      <div class="form-group">
        <label>Local Port</label>
        <input type="number" name="port" id="port" value="{config.get("port", 49211)}">
      </div>
      <div class="form-group">
        <label>Sync Interval (seconds)</label>
        <input type="number" name="sync_interval_seconds" id="sync_interval" value="{config.get("sync_interval_seconds", 300)}">
      </div>
      <div class="form-group">
        <label>Max Retries (Print Queue)</label>
        <input type="number" name="max_retries" id="max_retries" value="{config.get("max_retries", 3)}">
      </div>
      <div class="form-group">
        <label>Retry Delay (seconds)</label>
        <input type="number" name="retry_delay_seconds" id="retry_delay_seconds" value="{config.get("retry_delay_seconds", 60)}">
      </div>
    </div>
    <button type="submit" class="btn btn-primary">Save Settings</button>
    <button type="button" class="btn btn-secondary" onclick="testConnection()" id="testBtn">Test Connection</button>
  </form>
</div>

<!-- Printers -->
<div class="card">
  <h2>Local Printers ({len(printers_list)})</h2>
  <table>
    <thead><tr><th>Printer</th><th>Default</th><th>Status</th></tr></thead>
    <tbody>
    {"".join(f'<tr><td><strong>{p["name"]}</strong></td><td>{"<span class=badge badge-ok>DEFAULT</span>" if p["is_default"] else ""}</td><td><span class="badge badge-info">{p["status"]}</span></td></tr>' for p in printers_list) if printers_list else '<tr><td colspan=3 style="color:var(--muted)">No printers found</td></tr>'}
    </tbody>
  </table>
</div>

<!-- Recent Jobs -->
<div class="card">
  <h2>Recent Jobs ({len(clean_jobs)})</h2>
  <table>
    <thead><tr><th>ID</th><th>Printer</th><th>Type</th><th>Status</th><th>Time</th></tr></thead>
    <tbody>
    {"".join(f'<tr><td class="mono">{j["id"]}</td><td>{j["printer"]}</td><td><span class="badge badge-info">{j["type"].upper()}</span></td><td><span class="badge {"badge-ok" if j["status"]=="success" else "badge-err"}">{j["status"]}</span></td><td style="color:var(--muted)">{j["created_at"][11:19] if j.get("created_at") else ""}</td></tr>' for j in reversed(clean_jobs)) if clean_jobs else '<tr><td colspan=5 style="color:var(--muted)">No jobs yet</td></tr>'}
    </tbody>
  </table>
</div>

<script>
const msg = document.getElementById('msg');
function showMsg(text, ok) {{
  msg.style.display='block';
  msg.style.background=ok?'rgba(34,197,94,.1)':'rgba(239,68,68,.1)';
  msg.style.color=ok?'#22c55e':'#ef4444';
  msg.style.border='1px solid '+(ok?'rgba(34,197,94,.2)':'rgba(239,68,68,.2)');
  msg.textContent=text;
}}
document.getElementById('settingsForm').addEventListener('submit', function(e) {{
  e.preventDefault();
  fetch('/settings/save', {{
    method:'POST', headers:{{'Content-Type':'application/json'}},
    body: JSON.stringify({{
      hub_url: document.getElementById('hub_url').value.replace(/\\/+$/,''),
      agent_key: document.getElementById('agent_key').value,
      port: parseInt(document.getElementById('port').value) || 49211,
      sync_interval_seconds: parseInt(document.getElementById('sync_interval').value) || 300,
      max_retries: parseInt(document.getElementById('max_retries').value) || 3,
      retry_delay_seconds: parseInt(document.getElementById('retry_delay_seconds').value) || 60
    }})
  }}).then(r=>r.json()).then(d=>{{ showMsg(d.message || 'Saved!', d.status==='ok'); }}).catch(()=>showMsg('Error saving',false));
}});
function testConnection() {{
  const btn=document.getElementById('testBtn');
  btn.textContent='Testing...'; btn.disabled=true;
  const hub=document.getElementById('hub_url').value.replace(/\\/+$/,'');
  const key=document.getElementById('agent_key').value;
  fetch('/settings/test', {{
    method:'POST', headers:{{'Content-Type':'application/json'}},
    body: JSON.stringify({{ hub_url:hub, agent_key:key }})
  }}).then(r=>r.json()).then(d=>{{ showMsg(d.message, d.status==='ok'); btn.textContent='Test Connection'; btn.disabled=false; }}).catch(()=>{{ showMsg('Network error',false); btn.textContent='Test Connection'; btn.disabled=false; }});
}}
</script>
</body></html>'''

    @app.route('/settings/save', methods=['POST'])
    def settings_save():
        data = request.get_json()
        try:
            with open(config_path, 'r') as f:
                cfg = json.load(f)
            cfg.update(data)
            with open(config_path, 'w') as f:
                json.dump(cfg, f, indent=2)
            log.info("Settings saved via web UI")
            return jsonify({"status": "ok", "message": "Settings saved! Restart Trayprint to apply."})
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 500

    @app.route('/settings/test', methods=['POST'])
    def settings_test():
        data = request.get_json()
        hub_url = data.get('hub_url', '')
        agent_key = data.get('agent_key', '')
        if not hub_url or not agent_key:
            return jsonify({"status": "error", "message": "Hub URL and Agent Key are required."})
        try:
            import requests as req
            headers = {'Authorization': f'Bearer {agent_key}'}
            resp = req.get(f'{hub_url}/api/print-hub/profiles', headers=headers, timeout=5)
            if resp.status_code == 200:
                count = len(resp.json().get('profiles', {}))
                return jsonify({"status": "ok", "message": f"Connected! {count} profile(s) found."})
            elif resp.status_code == 401:
                return jsonify({"status": "error", "message": "Invalid Agent Key."})
            else:
                return jsonify({"status": "error", "message": f"Hub returned HTTP {resp.status_code}"})
        except Exception as e:
            return jsonify({"status": "error", "message": f"Cannot reach hub: {e}"})

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
    except:
        pass

    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)


if __name__ == '__main__':
    run_server(49211)
