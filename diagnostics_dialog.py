"""Enhanced Diagnostics dialog for TrayPrint — agent health, configuration, system info,
network, printers with capabilities, job history, logs, and export."""

import json
import os
import sys
import platform as _platform
from datetime import datetime

import requests
from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QFormLayout, QGroupBox,
    QMessageBox, QWidget, QCheckBox, QFileDialog, QTabWidget,
    QTreeWidget, QTreeWidgetItem, QHeaderView,
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont

from path_utils import get_root_dir
from logger import get_logger, get_log_path
from theme import get_stylesheet, get_palette

log = get_logger()

API_BASE = "http://127.0.0.1:49211"


class DiagnosticsDialog(QDialog):
    """Enhanced diagnostics dialog with multiple sections and auto-refresh."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("TrayPrint Diagnostics")
        self.setMinimumSize(800, 650)
        self.resize(900, 700)

        self._theme_id = 'dark'
        self._theme = get_palette(self._theme_id)

        # Auto-refresh state
        self._auto_refresh_enabled = False
        self._auto_refresh_timer = QTimer(self)
        self._auto_refresh_timer.timeout.connect(self._fetch_and_populate)
        self._auto_refresh_interval = 30000  # 30 seconds

        self._setup_ui()
        self._fetch_and_populate()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        # Title
        title = QLabel("TrayPrint Diagnostics")
        title.setStyleSheet(
            "font-size: 16px; font-weight: 700; color: %s; background: transparent;"
            % self._theme['text_primary']
        )
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Tab widget for sections
        self._tabs = QTabWidget()
        self._tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid %s;
                border-radius: 4px;
                background: %s;
            }
            QTabBar::tab {
                background: %s;
                color: %s;
                padding: 6px 14px;
                border: 1px solid %s;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: %s;
                color: %s;
            }
        """ % (
            self._theme['border'], self._theme['bg'],
            self._theme['tab_bg'], self._theme['tab_text'],
            self._theme['border'],
            self._theme['tab_active_bg'], self._theme['tab_active_text'],
        ))

        # ── Tab 1: Overview ──
        self._tab_overview = QWidget()
        self._setup_overview_tab()
        self._tabs.addTab(self._tab_overview, "Overview")

        # ── Tab 2: System Info ──
        self._tab_system = QWidget()
        self._setup_system_tab()
        self._tabs.addTab(self._tab_system, "System Info")

        # ── Tab 3: Network ──
        self._tab_network = QWidget()
        self._setup_network_tab()
        self._tabs.addTab(self._tab_network, "Network")

        # ── Tab 4: Printers ──
        self._tab_printers = QWidget()
        self._setup_printers_tab()
        self._tabs.addTab(self._tab_printers, "Printers")

        # ── Tab 5: Job History ──
        self._tab_jobs = QWidget()
        self._setup_jobs_tab()
        self._tabs.addTab(self._tab_jobs, "Job History")

        # ── Tab 6: Logs ──
        self._tab_logs = QWidget()
        self._setup_logs_tab()
        self._tabs.addTab(self._tab_logs, "Logs")

        # ── Tab 7: Watchdog ──
        self._tab_watchdog = QWidget()
        self._setup_watchdog_tab()
        self._tabs.addTab(self._tab_watchdog, "Watchdog")

        layout.addWidget(self._tabs, stretch=1)

        # ── Button row ──
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)

        self._chk_autorefresh = QCheckBox("Auto-refresh (30s)")
        self._chk_autorefresh.setStyleSheet(
            "color: %s; background: transparent;" % self._theme['text_primary']
        )
        self._chk_autorefresh.stateChanged.connect(self._on_autorefresh_toggle)

        self._btn_refresh = QPushButton("Refresh")
        self._btn_refresh.setMinimumHeight(32)
        self._btn_refresh.setToolTip("Re-fetch diagnostics from the local API server")
        self._btn_refresh.clicked.connect(self._fetch_and_populate)

        self._btn_export = QPushButton("Export JSON")
        self._btn_export.setMinimumHeight(32)
        self._btn_export.setToolTip("Export all diagnostic data as JSON file")
        self._btn_export.clicked.connect(self._export_json)

        self._btn_copy = QPushButton("Copy All")
        self._btn_copy.setMinimumHeight(32)
        self._btn_copy.setToolTip("Copy full diagnostics report to clipboard")
        self._btn_copy.clicked.connect(self._copy_to_clipboard)

        self._btn_close = QPushButton("Close")
        self._btn_close.setMinimumHeight(32)
        self._btn_close.setToolTip("Close diagnostics dialog")
        self._btn_close.clicked.connect(self.accept)

        btn_layout.addWidget(self._chk_autorefresh)
        btn_layout.addWidget(self._btn_refresh)
        btn_layout.addWidget(self._btn_export)
        btn_layout.addWidget(self._btn_copy)
        btn_layout.addStretch()
        btn_layout.addWidget(self._btn_close)

        layout.addLayout(btn_layout)

        # Apply shared theme stylesheet
        self.setStyleSheet(get_stylesheet(self._theme_id))

    def _make_group(self, title, parent_layout):
        """Create a styled QGroupBox."""
        gb = QGroupBox(title)
        gb.setStyleSheet("""
            QGroupBox {
                font-weight: 600;
                border: 1px solid %s;
                border-radius: 4px;
                margin-top: 8px;
                padding-top: 14px;
                color: %s;
                background: %s;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }
        """ % (self._theme['border'], self._theme['text_primary'], self._theme['bg_card']))
        gl = QFormLayout(gb)
        gl.setSpacing(4)
        gl.setContentsMargins(10, 16, 10, 8)
        parent_layout.addWidget(gb)
        return gl

    def _add_row(self, form_layout, label, value_widget):
        """Add a labeled row to a form layout."""
        lbl = QLabel(label)
        lbl.setStyleSheet("color: %s; font-weight: 500; background: transparent;" % self._theme['text_secondary'])
        form_layout.addRow(lbl, value_widget)

    def _make_label(self, text="", mono=False):
        """Create a styled QLabel."""
        lbl = QLabel(text)
        font_family = "Courier New" if mono else ""
        lbl.setStyleSheet(
            "color: %s; background: transparent; %s" % (
                self._theme['text_primary'],
                "font-family: %s; font-size: 12px;" % font_family if font_family else ""
            )
        )
        lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        return lbl

    def _make_text_edit(self, read_only=True):
        """Create a styled QTextEdit."""
        te = QTextEdit()
        te.setReadOnly(read_only)
        te.setFont(QFont("Courier New", 10))
        te.setStyleSheet(
            "QTextEdit {"
            "  background: %s;"
            "  color: %s;"
            "  border: 1px solid %s;"
            "  border-radius: 4px;"
            "  padding: 6px;"
            "  font-size: 12px;"
            "}" % (self._theme['bg'], self._theme['text_primary'], self._theme['border'])
        )
        return te

    # ── Tab Setup Methods ──

    def _setup_overview_tab(self):
        layout = QVBoxLayout(self._tab_overview)
        layout.setContentsMargins(8, 8, 8, 8)

        gl = self._make_group("Agent Overview", layout)
        self._lbl_app_version = self._make_label()
        self._lbl_python_version = self._make_label()
        self._lbl_platform = self._make_label()
        self._lbl_os_details = self._make_label()
        self._lbl_run_mode = self._make_label()
        self._lbl_uptime = self._make_label()
        self._lbl_config_path = self._make_label()
        self._lbl_config_exists = self._make_label()
        self._lbl_log_file = self._make_label()

        self._add_row(gl, "App Version:", self._lbl_app_version)
        self._add_row(gl, "Python:", self._lbl_python_version)
        self._add_row(gl, "Platform:", self._lbl_platform)
        self._add_row(gl, "OS Details:", self._lbl_os_details)
        self._add_row(gl, "Run Mode:", self._lbl_run_mode)
        self._add_row(gl, "Uptime:", self._lbl_uptime)
        self._add_row(gl, "Config Path:", self._lbl_config_path)
        self._add_row(gl, "Config Exists:", self._lbl_config_exists)
        self._add_row(gl, "Log File:", self._lbl_log_file)

        layout.addStretch()

    def _setup_system_tab(self):
        layout = QVBoxLayout(self._tab_system)
        layout.setContentsMargins(8, 8, 8, 8)

        gl = self._make_group("System Information", layout)
        self._lbl_os_version = self._make_label()
        self._lbl_python_ver = self._make_label()
        self._lbl_pyside6_ver = self._make_label()
        self._lbl_pymupdf_ver = self._make_label()
        self._lbl_architecture = self._make_label()
        self._lbl_machine = self._make_label()
        self._lbl_processor = self._make_label()
        self._lbl_hostname = self._make_label()

        self._add_row(gl, "OS Version:", self._lbl_os_version)
        self._add_row(gl, "Python Version:", self._lbl_python_ver)
        self._add_row(gl, "PySide6 Version:", self._lbl_pyside6_ver)
        self._add_row(gl, "PyMuPDF Version:", self._lbl_pymupdf_ver)
        self._add_row(gl, "Architecture:", self._lbl_architecture)
        self._add_row(gl, "Machine:", self._lbl_machine)
        self._add_row(gl, "Processor:", self._lbl_processor)
        self._add_row(gl, "Hostname:", self._lbl_hostname)

        layout.addStretch()

    def _setup_network_tab(self):
        layout = QVBoxLayout(self._tab_network)
        layout.setContentsMargins(8, 8, 8, 8)

        gl = self._make_group("Network & Hub", layout)
        self._lbl_hub_url = self._make_label()
        self._lbl_hub_reachable = self._make_label()
        self._lbl_last_sync = self._make_label()
        self._lbl_ws_status = self._make_label()
        self._lbl_proxy = self._make_label()

        self._add_row(gl, "Hub URL:", self._lbl_hub_url)
        self._add_row(gl, "Hub Reachable:", self._lbl_hub_reachable)
        self._add_row(gl, "Last Sync:", self._lbl_last_sync)
        self._add_row(gl, "WebSocket:", self._lbl_ws_status)
        self._add_row(gl, "Proxy:", self._lbl_proxy)

        layout.addStretch()

    def _setup_printers_tab(self):
        layout = QVBoxLayout(self._tab_printers)
        layout.setContentsMargins(8, 8, 8, 8)

        self._printer_tree = QTreeWidget()
        self._printer_tree.setHeaderLabels([
            "Printer", "Default", "Status", "Trays",
            "Media Sizes", "Resolutions", "Color Modes", "Duplex"
        ])
        self._printer_tree.setAlternatingRowColors(True)
        self._printer_tree.setRootIsDecorated(False)
        self._printer_tree.header().setStretchLastSection(True)
        self._printer_tree.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self._printer_tree.setStyleSheet("""
            QTreeWidget {
                background: %s;
                color: %s;
                border: 1px solid %s;
                border-radius: 4px;
                alternate-background-color: %s;
                font-size: 12px;
            }
            QHeaderView::section {
                background: %s;
                color: %s;
                padding: 4px;
                border: 1px solid %s;
                font-weight: 600;
            }
        """ % (
            self._theme['bg'], self._theme['text_primary'], self._theme['border'],
            self._theme['list_alt'],
            self._theme['table_header'], self._theme['table_header_text'], self._theme['table_grid'],
        ))
        layout.addWidget(self._printer_tree, stretch=1)

    def _setup_jobs_tab(self):
        layout = QVBoxLayout(self._tab_jobs)
        layout.setContentsMargins(8, 8, 8, 8)

        gl = self._make_group("Job Statistics", layout)
        self._lbl_total_jobs = self._make_label()
        self._lbl_success_jobs = self._make_label()
        self._lbl_failed_jobs = self._make_label()
        self._lbl_pending_jobs = self._make_label()
        self._add_row(gl, "Total Jobs:", self._lbl_total_jobs)
        self._add_row(gl, "Successful:", self._lbl_success_jobs)
        self._add_row(gl, "Failed:", self._lbl_failed_jobs)
        self._add_row(gl, "Pending:", self._lbl_pending_jobs)

        gl2 = self._make_group("Last 10 Jobs", layout)
        self._jobs_text = self._make_text_edit()
        self._jobs_text.setMaximumHeight(250)
        gl2.addRow(self._jobs_text)

    def _setup_logs_tab(self):
        layout = QVBoxLayout(self._tab_logs)
        layout.setContentsMargins(8, 8, 8, 8)

        gl = self._make_group("Application Log (last 50 lines)", layout)
        self._logs_text = self._make_text_edit()
        gl.addRow(self._logs_text)

    def _setup_watchdog_tab(self):
        layout = QVBoxLayout(self._tab_watchdog)
        layout.setContentsMargins(8, 8, 8, 8)

        gl = self._make_group("Watchdog Status", layout)
        self._lbl_wd_running = self._make_label()
        self._lbl_wd_last_check = self._make_label()
        self._lbl_wd_restarts = self._make_label()
        self._lbl_wd_memory = self._make_label()
        self._lbl_wd_disk = self._make_label()
        self._lbl_wd_ws = self._make_label()
        self._add_row(gl, "Running:", self._lbl_wd_running)
        self._add_row(gl, "Last Check:", self._lbl_wd_last_check)
        self._add_row(gl, "Restart Attempts:", self._lbl_wd_restarts)
        self._add_row(gl, "Memory Usage:", self._lbl_wd_memory)
        self._add_row(gl, "Disk Space:", self._lbl_wd_disk)
        self._add_row(gl, "WebSocket:", self._lbl_wd_ws)

        gl2 = self._make_group("Watchdog Log", layout)
        self._wd_log_text = self._make_text_edit()
        self._wd_log_text.setMaximumHeight(200)
        gl2.addRow(self._wd_log_text)

        layout.addStretch()

    # ── Data Fetching ──

    def _fetch_and_populate(self):
        """Fetch diagnostics from the server and populate all tabs."""
        self._btn_refresh.setEnabled(False)
        self._btn_refresh.setText("Fetching...")
        QApplication.processEvents()

        try:
            resp = requests.get(f"{API_BASE}/api/diagnostics", timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                self._populate_from_data(data)
            else:
                error_text = f"Error: Server returned HTTP {resp.status_code}\n\nResponse: {resp.text}"
                self._show_error(error_text)
        except requests.ConnectionError:
            self._show_error(
                "Error: Could not connect to TrayPrint API server.\n\n"
                f"Expected at: {API_BASE}/api/diagnostics\n\n"
                "Make sure TrayPrint is running."
            )
        except Exception as e:
            self._show_error(f"Error fetching diagnostics: {e}")
        finally:
            self._btn_refresh.setEnabled(True)
            self._btn_refresh.setText("Refresh")

    def _show_error(self, message):
        """Display error in all tabs. Converts message to string to prevent TypeError."""
        msg_str = str(message)
        # Overview
        for lbl in [self._lbl_app_version, self._lbl_python_version, self._lbl_platform,
                     self._lbl_os_details, self._lbl_run_mode, self._lbl_uptime,
                     self._lbl_config_path, self._lbl_config_exists, self._lbl_log_file]:
            lbl.setText(msg_str)
        # System
        for lbl in [self._lbl_os_version, self._lbl_python_ver, self._lbl_pyside6_ver,
                     self._lbl_pymupdf_ver, self._lbl_architecture, self._lbl_machine,
                     self._lbl_processor, self._lbl_hostname]:
            lbl.setText(msg_str)
        # Network
        for lbl in [self._lbl_hub_url, self._lbl_hub_reachable, self._lbl_last_sync,
                     self._lbl_ws_status, self._lbl_proxy]:
            lbl.setText(msg_str)
        # Printers
        self._printer_tree.clear()
        # Jobs
        for lbl in [self._lbl_total_jobs, self._lbl_success_jobs, self._lbl_failed_jobs, self._lbl_pending_jobs]:
            lbl.setText(msg_str)
        self._jobs_text.setText(msg_str)
        # Logs
        self._logs_text.setText(msg_str)
        # Watchdog
        for lbl in [self._lbl_wd_running, self._lbl_wd_last_check, self._lbl_wd_restarts,
                     self._lbl_wd_memory, self._lbl_wd_disk, self._lbl_wd_ws]:
            lbl.setText(msg_str)
        self._wd_log_text.setText(msg_str)

    def _populate_from_data(self, data):
        """Populate all tabs from diagnostics data."""
        self._raw_data = data  # Store for export

        # ── Overview Tab ──
        self._lbl_app_version.setText(data.get('app_version', 'N/A'))
        self._lbl_python_version.setText(data.get('python_version', 'N/A'))
        self._lbl_platform.setText(data.get('platform', 'N/A'))
        self._lbl_os_details.setText(data.get('os_details', 'N/A'))
        self._lbl_run_mode.setText(data.get('run_mode', 'N/A'))
        self._lbl_uptime.setText(data.get('uptime', 'N/A'))
        self._lbl_config_path.setText(data.get('config_path', 'N/A'))
        self._lbl_config_exists.setText(str(data.get('config_exists', False)))
        self._lbl_log_file.setText(data.get('log_file', 'N/A'))

        # ── System Info Tab ──
        sys_info = data.get('system_info', {})
        self._lbl_os_version.setText(sys_info.get('os_version', _platform.platform()))
        self._lbl_python_ver.setText(sys_info.get('python_version', sys.version.split()[0]))
        self._lbl_pyside6_ver.setText(sys_info.get('pyside6_version', 'N/A'))
        self._lbl_pymupdf_ver.setText(sys_info.get('pymupdf_version', 'N/A'))
        self._lbl_architecture.setText(sys_info.get('architecture', 'N/A'))
        self._lbl_machine.setText(sys_info.get('machine', 'N/A'))
        self._lbl_processor.setText(sys_info.get('processor', 'N/A'))
        self._lbl_hostname.setText(sys_info.get('hostname', 'N/A'))

        # ── Network Tab ──
        self._lbl_hub_url.setText(data.get('hub_url', 'N/A'))
        hub_reachable = data.get('hub_reachable', data.get('hub_connected', False))
        self._lbl_hub_reachable.setText(str(hub_reachable))
        self._lbl_last_sync.setText(data.get('last_sync_time', 'N/A'))
        # websocket_status is a dict from server: {"connected": bool, "reconnect_count": int}
        ws_status = data.get('websocket_status', {})
        if isinstance(ws_status, dict):
            connected = ws_status.get('connected', False)
            reconnects = ws_status.get('reconnect_count', 0)
            ws_str = f"Connected: {connected} | Reconnects: {reconnects}"
        else:
            ws_str = str(ws_status)
        self._lbl_ws_status.setText(ws_str)
        proxy = data.get('proxy_config', {})
        if proxy and isinstance(proxy, dict):
            proxy_str = json.dumps(proxy, indent=2)
        else:
            proxy_str = str(proxy) if proxy else 'None configured'
        self._lbl_proxy.setText(proxy_str)

        # ── Printers Tab ──
        self._populate_printers(data)

        # ── Job History Tab ──
        self._populate_jobs(data)

        # ── Logs Tab ──
        self._populate_logs(data)

        # ── Watchdog Tab ──
        self._populate_watchdog(data)

    def _populate_printers(self, data):
        """Populate the printers tree with capabilities."""
        self._printer_tree.clear()
        printers = data.get('printers', [])
        capabilities = data.get('printer_capabilities', {})
        default_printer = data.get('default_printer', '')

        if not printers:
            item = QTreeWidgetItem(["No printers found"])
            self._printer_tree.addTopLevelItem(item)
            return

        for p in printers:
            name = p if isinstance(p, str) else p.get('name', str(p))
            is_default = "Yes" if name == default_printer else ""
            caps = capabilities.get(name, {})
            trays = ", ".join(caps.get('trays', [])) if isinstance(caps, dict) else ""
            media = ", ".join(caps.get('media_sizes', [])) if isinstance(caps, dict) else ""
            resolutions = ", ".join(caps.get('resolutions', [])) if isinstance(caps, dict) else ""
            color_modes = ", ".join(caps.get('color_modes', [])) if isinstance(caps, dict) else ""
            duplex = ", ".join(caps.get('duplex', [])) if isinstance(caps, dict) else ""
            status = p.get('status', '') if isinstance(p, dict) else ""

            item = QTreeWidgetItem([name, is_default, status, trays, media, resolutions, color_modes, duplex])
            self._printer_tree.addTopLevelItem(item)

    def _populate_jobs(self, data):
        """Populate job history section."""
        jobs = data.get('jobs', [])
        total = len(jobs)
        success = sum(1 for j in jobs if j.get('status') == 'success')
        failed = sum(1 for j in jobs if j.get('status') == 'failed')
        pending = sum(1 for j in jobs if j.get('status') == 'pending')

        self._lbl_total_jobs.setText(str(total))
        self._lbl_success_jobs.setText(str(success))
        self._lbl_failed_jobs.setText(str(failed))
        self._lbl_pending_jobs.setText(str(pending))

        # Last 10 jobs
        last_10 = jobs[-10:] if len(jobs) > 10 else jobs
        lines = []
        lines.append(f"{'ID':<10} {'Printer':<20} {'Type':<8} {'Status':<10} {'Time':<20}")
        lines.append("-" * 70)
        for j in reversed(last_10):
            job_id = j.get('id', j.get('job_id', '?'))[:8]
            printer = j.get('printer', '?')[:18]
            jtype = j.get('type', '?')[:6]
            status = j.get('status', '?')[:8]
            created = j.get('created_at', '?')[:19]
            lines.append(f"{job_id:<10} {printer:<20} {jtype:<8} {status:<10} {created:<20}")
        self._jobs_text.setText("\n".join(lines))

    def _populate_logs(self, data):
        """Populate logs section."""
        logs = data.get('logs', '')
        if logs:
            # Server may return logs as a list of strings or a single string
            if isinstance(logs, list):
                logs_str = "\n".join(logs)
            else:
                logs_str = str(logs)
            self._logs_text.setText(logs_str)
        else:
            # Try reading from log file directly
            log_path = get_log_path()
            try:
                if os.path.exists(log_path):
                    with open(log_path, 'r', encoding='utf-8', errors='replace') as f:
                        all_lines = f.readlines()
                        last_50 = all_lines[-50:] if len(all_lines) > 50 else all_lines
                        self._logs_text.setText("".join(last_50))
                else:
                    self._logs_text.setText("Log file not found: %s" % log_path)
            except Exception as e:
                self._logs_text.setText("Error reading log: %s" % e)

    def _populate_watchdog(self, data):
        """Populate watchdog tab."""
        wd = data.get('watchdog', {})
        self._lbl_wd_running.setText(str(wd.get('running', False)))
        self._lbl_wd_last_check.setText(str(wd.get('last_check', 'N/A')))
        self._lbl_wd_restarts.setText(str(wd.get('restart_attempts', 0)))
        self._lbl_wd_memory.setText(str(wd.get('memory_usage', 'N/A')))
        self._lbl_wd_disk.setText(str(wd.get('disk_space', 'N/A')))
        self._lbl_wd_ws.setText(str(wd.get('websocket_status', 'N/A')))

        # Watchdog log
        wd_log = wd.get('log', [])
        if wd_log:
            lines = []
            for entry in wd_log:
                if isinstance(entry, dict):
                    ts = entry.get('timestamp', '')
                    msg = entry.get('message', str(entry))
                    lines.append(f"[{ts}] {msg}")
                else:
                    lines.append(str(entry))
            self._wd_log_text.setText("\n".join(lines))
        else:
            self._wd_log_text.setText("No watchdog log entries available.")

    # ── Actions ──

    def _on_autorefresh_toggle(self, state):
        """Toggle auto-refresh timer."""
        if state == Qt.Checked:
            self._auto_refresh_enabled = True
            self._auto_refresh_timer.start(self._auto_refresh_interval)
            log.info("Diagnostics auto-refresh enabled (30s interval)")
        else:
            self._auto_refresh_enabled = False
            self._auto_refresh_timer.stop()
            log.info("Diagnostics auto-refresh disabled")

    def _export_json(self):
        """Export all diagnostic data as a JSON file."""
        if not hasattr(self, '_raw_data') or not self._raw_data:
            QMessageBox.warning(self, "Export", "No diagnostic data available yet. Click Refresh first.")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"trayprint_diagnostics_{timestamp}.json"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export Diagnostics", default_name, "JSON Files (*.json)"
        )
        if not file_path:
            return

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self._raw_data, f, indent=2, default=str)
            QMessageBox.information(self, "Export", f"Diagnostics exported to:\n{file_path}")
            log.info("Diagnostics exported to %s", file_path)
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export: {e}")
            log.error("Diagnostics export failed: %s", e)

    def _copy_to_clipboard(self):
        """Copy the full diagnostics report to the system clipboard."""
        text = self._build_full_report()
        if text:
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            self._btn_copy.setText("Copied!")
            QTimer.singleShot(2000, lambda: self._btn_copy.setText("Copy All"))

    def _build_full_report(self):
        """Build a comprehensive text report from all tabs."""
        if not hasattr(self, '_raw_data') or not self._raw_data:
            return "No diagnostic data available."

        data = self._raw_data
        lines = []
        lines.append("=" * 60)
        lines.append("  TrayPrint Agent Diagnostics \u2014 Full Report")
        lines.append("=" * 60)
        lines.append("")

        # ── Overview ──
        lines.append("-" * 60)
        lines.append("  AGENT OVERVIEW")
        lines.append("-" * 60)
        lines.append(f"  App Version:     {data.get('app_version', 'N/A')}")
        lines.append(f"  Python Version:  {data.get('python_version', 'N/A')}")
        lines.append(f"  Platform:        {data.get('platform', 'N/A')}")
        lines.append(f"  OS Details:      {data.get('os_details', 'N/A')}")
        lines.append(f"  Run Mode:        {data.get('run_mode', 'N/A')}")
        lines.append(f"  Uptime:          {data.get('uptime', 'N/A')}")
        lines.append(f"  Config Path:     {data.get('config_path', 'N/A')}")
        lines.append(f"  Config Exists:   {data.get('config_exists', False)}")
        lines.append(f"  Log File:        {data.get('log_file', 'N/A')}")
        lines.append("")

        # ── System Info ──
        sys_info = data.get('system_info', {})
        lines.append("-" * 60)
        lines.append("  SYSTEM INFORMATION")
        lines.append("-" * 60)
        lines.append(f"  OS Version:      {sys_info.get('os_version', 'N/A')}")
        lines.append(f"  Python:          {sys_info.get('python_version', 'N/A')}")
        lines.append(f"  PySide6:         {sys_info.get('pyside6_version', 'N/A')}")
        lines.append(f"  PyMuPDF:         {sys_info.get('pymupdf_version', 'N/A')}")
        lines.append(f"  Architecture:    {sys_info.get('architecture', 'N/A')}")
        lines.append(f"  Machine:         {sys_info.get('machine', 'N/A')}")
        lines.append(f"  Processor:       {sys_info.get('processor', 'N/A')}")
        lines.append(f"  Hostname:        {sys_info.get('hostname', 'N/A')}")
        lines.append("")

        # ── Network ──
        lines.append("-" * 60)
        lines.append("  NETWORK")
        lines.append("-" * 60)
        lines.append(f"  Hub URL:         {data.get('hub_url', 'N/A')}")
        lines.append(f"  Hub Reachable:   {data.get('hub_reachable', data.get('hub_connected', False))}")
        lines.append(f"  Last Sync:       {data.get('last_sync_time', 'N/A')}")
        lines.append(f"  WebSocket:       {data.get('websocket_status', 'N/A')}")
        proxy = data.get('proxy_config', {})
        lines.append(f"  Proxy:           {json.dumps(proxy) if proxy else 'None'}")
        lines.append("")

        # ── Printers ──
        lines.append("-" * 60)
        lines.append("  PRINTERS")
        lines.append("-" * 60)
        printers = data.get('printers', [])
        lines.append(f"  Count:           {len(printers)}")
        default_printer = data.get('default_printer', '')
        caps = data.get('printer_capabilities', {})
        for p in printers:
            name = p if isinstance(p, str) else p.get('name', str(p))
            is_def = " (default)" if name == default_printer else ""
            lines.append(f"    - {name}{is_def}")
            pc = caps.get(name, {})
            if isinstance(pc, dict) and pc:
                for key in ('trays', 'media_sizes', 'resolutions', 'color_modes', 'duplex'):
                    vals = pc.get(key, [])
                    if vals:
                        lines.append(f"      {key}: {', '.join(vals)}")
        lines.append("")

        # ── Job History ──
        lines.append("-" * 60)
        lines.append("  JOB HISTORY")
        lines.append("-" * 60)
        jobs = data.get('jobs', [])
        total = len(jobs)
        success = sum(1 for j in jobs if j.get('status') == 'success')
        failed = sum(1 for j in jobs if j.get('status') == 'failed')
        pending = sum(1 for j in jobs if j.get('status') == 'pending')
        lines.append(f"  Total:           {total}")
        lines.append(f"  Successful:      {success}")
        lines.append(f"  Failed:          {failed}")
        lines.append(f"  Pending:         {pending}")
        lines.append("")
        lines.append(f"  {'ID':<10} {'Printer':<20} {'Type':<8} {'Status':<10} {'Time':<20}")
        lines.append(f"  {'-' * 68}")
        for j in reversed(jobs[-10:]):
            job_id = j.get('id', j.get('job_id', '?'))[:8]
            printer = j.get('printer', '?')[:18]
            jtype = j.get('type', '?')[:6]
            status = j.get('status', '?')[:8]
            created = j.get('created_at', '?')[:19]
            lines.append(f"  {job_id:<10} {printer:<20} {jtype:<8} {status:<10} {created:<20}")
        lines.append("")

        # ── Logs ──
        lines.append("-" * 60)
        lines.append("  LOGS (last 50 lines)")
        lines.append("-" * 60)
        logs = data.get('logs', '')
        if logs:
            lines.append(logs)
        else:
            log_path = get_log_path()
            try:
                if os.path.exists(log_path):
                    with open(log_path, 'r', encoding='utf-8', errors='replace') as f:
                        all_lines = f.readlines()
                        last_50 = all_lines[-50:] if len(all_lines) > 50 else all_lines
                        lines.append("".join(last_50))
                else:
                    lines.append(f"Log file not found: {log_path}")
            except Exception as e:
                lines.append(f"Error reading log: {e}")
        lines.append("")

        # ── Watchdog ──
        lines.append("-" * 60)
        lines.append("  WATCHDOG")
        lines.append("-" * 60)
        wd = data.get('watchdog', {})
        lines.append(f"  Running:         {wd.get('running', False)}")
        lines.append(f"  Last Check:      {wd.get('last_check', 'N/A')}")
        lines.append(f"  Restart Attempts: {wd.get('restart_attempts', 0)}")
        lines.append(f"  Memory Usage:    {wd.get('memory_usage', 'N/A')}")
        lines.append(f"  Disk Space:      {wd.get('disk_space', 'N/A')}")
        lines.append(f"  WebSocket:       {wd.get('websocket_status', 'N/A')}")
        lines.append("")
        lines.append("=" * 60)

        return "\n".join(lines)
