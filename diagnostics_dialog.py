"""Diagnostics dialog for TrayPrint — shows agent health and configuration."""

import json
import os
import sys
from datetime import datetime

import requests
from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QFormLayout, QGroupBox,
    QMessageBox, QWidget
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont

from path_utils import get_root_dir
from logger import get_logger
from theme import get_stylesheet, get_palette

log = get_logger()

API_BASE = "http://127.0.0.1:49211"


class DiagnosticsDialog(QDialog):
    """Display TrayPrint diagnostics in a read-only dialog."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("TrayPrint Diagnostics")
        self.setMinimumSize(600, 500)
        self.resize(600, 500)

        self._theme_id = 'dark'
        self._theme = get_palette(self._theme_id)

        self._setup_ui()
        self._fetch_and_populate()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # Title
        title = QLabel("TrayPrint Diagnostics")
        title.setStyleSheet(
            "font-size: 16px; font-weight: 700; color: %s; background: transparent;"
            % self._theme['text_primary']
        )
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Text area for diagnostics
        self._text_area = QTextEdit()
        self._text_area.setReadOnly(True)
        self._text_area.setFont(QFont("Courier New", 10))
        self._text_area.setStyleSheet(
            "QTextEdit {"
            "  background: %s;"
            "  color: %s;"
            "  border: 1px solid %s;"
            "  border-radius: 6px;"
            "  padding: 10px;"
            "  font-size: 13px;"
            "}" % (self._theme['bg'], self._theme['text_primary'], self._theme['border'])
        )
        layout.addWidget(self._text_area, stretch=1)

        # Button row
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self._btn_refresh = QPushButton("Refresh")
        self._btn_refresh.setMinimumHeight(32)
        self._btn_refresh.setToolTip("Re-fetch diagnostics from the local API server")
        self._btn_refresh.clicked.connect(self._fetch_and_populate)

        self._btn_copy = QPushButton("Copy All")
        self._btn_copy.setMinimumHeight(32)
        self._btn_copy.setToolTip("Copy full diagnostics report to clipboard")
        self._btn_copy.clicked.connect(self._copy_to_clipboard)

        self._btn_close = QPushButton("Close")
        self._btn_close.setMinimumHeight(32)
        self._btn_close.setToolTip("Close diagnostics dialog")
        self._btn_close.clicked.connect(self.accept)

        btn_layout.addWidget(self._btn_refresh)
        btn_layout.addWidget(self._btn_copy)
        btn_layout.addStretch()
        btn_layout.addWidget(self._btn_close)

        layout.addLayout(btn_layout)

        # Apply shared theme stylesheet
        self.setStyleSheet(get_stylesheet(self._theme_id))

    def _fetch_and_populate(self):
        """Fetch diagnostics from the server and populate the text area."""
        self._btn_refresh.setEnabled(False)
        self._btn_refresh.setText("Fetching...")
        QApplication.processEvents()

        try:
            resp = requests.get(f"{API_BASE}/api/diagnostics", timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                self._populate_from_data(data)
            else:
                self._text_area.setText(
                    f"Error: Server returned HTTP {resp.status_code}\n\n"
                    f"Response: {resp.text}"
                )
        except requests.ConnectionError:
            self._text_area.setText(
                "Error: Could not connect to TrayPrint API server.\n\n"
                f"Expected at: {API_BASE}/api/diagnostics\n\n"
                "Make sure TrayPrint is running."
            )
        except Exception as e:
            self._text_area.setText(f"Error fetching diagnostics: {e}")
        finally:
            self._btn_refresh.setEnabled(True)
            self._btn_refresh.setText("Refresh")

    def _populate_from_data(self, data):
        """Format diagnostics data into human-readable text."""
        lines = []
        lines.append("=" * 56)
        lines.append("  TrayPrint Agent Diagnostics")
        lines.append("=" * 56)
        lines.append("")

        lines.append(f"  App Version:     {data.get('app_version', 'N/A')}")
        lines.append(f"  Python Version:  {data.get('python_version', 'N/A')}")
        lines.append(f"  Platform:        {data.get('platform', 'N/A')}")
        lines.append(f"  OS Details:      {data.get('os_details', 'N/A')}")
        lines.append(f"  Run Mode:        {data.get('run_mode', 'N/A')}")
        lines.append(f"  Uptime:          {data.get('uptime', 'N/A')}")
        lines.append("")

        lines.append("-" * 56)
        lines.append("  Configuration")
        lines.append("-" * 56)
        lines.append(f"  Config Path:     {data.get('config_path', 'N/A')}")
        lines.append(f"  Config Exists:   {data.get('config_exists', False)}")
        lines.append(f"  Hub URL:         {data.get('hub_url', 'N/A')}")
        lines.append(f"  Hub Connected:   {data.get('hub_connected', False)}")
        lines.append(f"  Log File:        {data.get('log_file', 'N/A')}")
        lines.append("")

        lines.append("-" * 56)
        lines.append("  Printers")
        lines.append("-" * 56)
        printers = data.get('printers', [])
        lines.append(f"  Printer Count:   {data.get('printer_count', 0)}")
        if printers:
            for p in printers:
                lines.append(f"    - {p}")
        lines.append("")

        lines.append("-" * 56)
        lines.append("  Queue")
        lines.append("-" * 56)
        lines.append(f"  Queue Length:    {data.get('queue_length', 0)}")
        lines.append("")

        lines.append("-" * 56)
        lines.append("  Watchdog")
        lines.append("-" * 56)
        wd = data.get('watchdog', {})
        lines.append(f"  Running:         {wd.get('running', False)}")
        lines.append(f"  Last Check:      {wd.get('last_check', 'N/A')}")
        lines.append(f"  Restart Attempts: {wd.get('restart_attempts', 0)}")
        lines.append("")
        lines.append("=" * 56)

        self._text_area.setText("\n".join(lines))

    def _copy_to_clipboard(self):
        """Copy the full diagnostics report to the system clipboard."""
        text = self._text_area.toPlainText()
        if text:
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            self._btn_copy.setText("Copied!")
            QTimer.singleShot(2000, lambda: self._btn_copy.setText("Copy All"))
