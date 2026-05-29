"""
Print Job Queue Viewer Dialog for TrayPrint.

Shows active/pending/recent print jobs in a QTableWidget with
Cancel, Refresh, and auto-refresh via QTimer.
"""

import os
import json
from datetime import datetime

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QLabel, QMessageBox, QWidget,
    QAbstractItemView,
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QFont

from path_utils import get_data_dir
from theme import get_stylesheet, get_palette


STATUS_STYLES = {
    'pending':   ('\u23f3  Pending',   '#f0b34b'),   # ⏳ amber
    'printing':  ('\U0001f7e2  Printing', '#4caf88'), # 🟢 green
    'success':   ('\u2705  Success',   '#4caf88'),   # ✅ green
    'failed':    ('\u274c  Failed',    '#e66565'),   # ❌ red
    'cancelled': ('\u274c  Cancelled', '#707090'),   # ❌ grey
}


class PrintQueueDialog(QDialog):
    """Dialog showing current print queue with cancel/refresh controls."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Print Queue')
        self.setMinimumSize(800, 500)
        self.resize(800, 500)

        # Load settings theme dynamically
        self._theme_id = 'dark'
        config_path = os.path.join(get_data_dir(), 'config.json')
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    self._theme_id = config.get('theme', 'dark')
            except Exception:
                pass
        self._theme = get_palette(self._theme_id)

        self._refresh_timer = QTimer(self)
        self._refresh_timer.timeout.connect(self.refresh)
        self._refresh_timer.start(3000)  # auto-refresh every 3 seconds

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── Top bar: title + refresh button ──
        top_layout = QHBoxLayout()

        title_label = QLabel('Print Job Queue')
        title_label.setStyleSheet(
            'font-weight: 700; font-size: 15px; color: %s; background: transparent;'
            % self._theme['text_primary']
        )
        top_layout.addWidget(title_label)
        top_layout.addStretch()

        self.btn_refresh = QPushButton('Refresh')
        self.btn_refresh.setMinimumHeight(30)
        self.btn_refresh.setToolTip('Refresh the print queue (Ctrl+R)')
        self.btn_refresh.clicked.connect(self.refresh)
        top_layout.addWidget(self.btn_refresh)

        layout.addLayout(top_layout)

        # ── Table ──
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels([
            'Job ID', 'Document', 'Status', 'Printer', 'Time', 'Type', 'Actions'
        ])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # ── Empty state label (shown when no jobs) ──
        self.empty_label = QLabel('No active jobs')
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setStyleSheet(
            'color: %s; font-size: 14px; padding: 40px; background: transparent;'
            % self._theme['text_muted']
        )
        self.empty_label.setVisible(False)
        layout.addWidget(self.empty_label)

        # ── Cancel confirmation dialog tracking ──
        self._cancel_in_progress = False

        # Apply shared theme
        self.setStyleSheet(get_stylesheet(self._theme_id))

        # Load initial data
        self.refresh()

    def refresh(self):
        """Fetch queue status from server and update the table."""
        import server

        jobs = server.get_queue_status()
        is_printing = server.is_printing()

        # Update title with printing indicator
        title_suffix = '  \U0001f4e8 Printing...' if is_printing else ''
        title_widget = self.layout().itemAt(0)
        if title_widget and title_widget.layout():
            title_label = title_widget.layout().itemAt(0)
            if title_label and title_label.widget():
                base_text = 'Print Job Queue'
                title_label.widget().setText(base_text + title_suffix)

        self.table.setRowCount(0)

        if not jobs:
            self.table.setVisible(False)
            self.empty_label.setVisible(True)
            return

        self.table.setVisible(True)
        self.empty_label.setVisible(False)

        for j in jobs:
            row = self.table.rowCount()
            self.table.insertRow(row)

            # Job ID
            id_item = QTableWidgetItem(j.get('id', ''))
            id_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 0, id_item)

            # Document name (from data_preview or type)
            doc_name = j.get('data_preview', '') or j.get('type', '')
            if not doc_name:
                doc_name = j.get('document_name', 'Unknown')
            self.table.setItem(row, 1, QTableWidgetItem(doc_name))

            # Status with color
            status = j.get('status', 'unknown')
            # Map 'pending' → show as 'printing' if currently printing
            display_status = status
            if status == 'pending' and is_printing and row == 0:
                display_status = 'printing'

            status_info = STATUS_STYLES.get(display_status,
                                            (display_status.capitalize(), '#707090'))
            status_text, status_color = status_info

            status_item = QTableWidgetItem(status_text)
            status_item.setForeground(QColor(status_color))
            font = status_item.font()
            font.setBold(True)
            font.setPointSize(11)
            status_item.setFont(font)
            self.table.setItem(row, 2, status_item)

            # Printer
            self.table.setItem(row, 3, QTableWidgetItem(j.get('printer', '')))

            # Time
            created = j.get('created_at', '')
            time_str = created[11:19] if len(created) >= 19 else created
            self.table.setItem(row, 4, QTableWidgetItem(time_str))

            # Type
            self.table.setItem(row, 5, QTableWidgetItem(j.get('type', '')))

            # Cancel button (only for pending jobs, not already cancelled)
            actions_widget = QWidget()
            actions_layout = QHBoxLayout(actions_widget)
            actions_layout.setContentsMargins(4, 0, 4, 0)
            actions_layout.setSpacing(4)

            can_cancel = status in ('pending',) and display_status != 'cancelled'

            if can_cancel:
                btn_cancel = QPushButton('Cancel')
                btn_cancel.setMinimumHeight(26)
                btn_cancel.setStyleSheet(
                    "QPushButton {"
                    "  padding: 4px 10px;"
                    "  background: %s;"
                    "  border: 1px solid %s;"
                    "  border-radius: 4px;"
                    "  color: %s;"
                    "  font-size: 11px;"
                    "  font-weight: 500;"
                    "}"
                    "QPushButton:hover {"
                    "  background: %s;"
                    "}"
                    "QPushButton:disabled {"
                    "  background: %s;"
                    "  color: %s;"
                    "  border-color: %s;"
                    "}"
                ) % (
                    self._theme['bg'],       # bg for cancel (error-toned)
                    self._theme['error'],     # border
                    self._theme['error'],     # text
                    '#3a2020' if self._theme_id == 'dark' else '#f8d7da', # hover (slightly lighter red)
                    self._theme['bg'],        # disabled bg
                    self._theme['text_muted'],# disabled text
                    self._theme['border'],    # disabled border
                )
                btn_cancel.setCursor(Qt.PointingHandCursor)
                job_id = j.get('id', '')
                btn_cancel.clicked.connect(
                    lambda checked, jid=job_id: self._on_cancel(jid)
                )
                actions_layout.addWidget(btn_cancel)
            else:
                # Show a disabled placeholder
                placeholder = QLabel('\u2014')
                placeholder.setStyleSheet('color: %s; background: transparent;' % self._theme['text_muted'])
                placeholder.setAlignment(Qt.AlignCenter)
                actions_layout.addWidget(placeholder)

            actions_layout.addStretch()
            self.table.setCellWidget(row, 6, actions_widget)

    def _on_cancel(self, job_id):
        """Show confirmation and cancel the job."""
        import server

        if self._cancel_in_progress:
            return
        self._cancel_in_progress = True

        try:
            reply = QMessageBox.question(
                self, 'Cancel Print Job',
                f'Are you sure you want to cancel job {job_id}?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )

            if reply == QMessageBox.Yes:
                result = server.cancel_job(job_id)
                if result:
                    self.refresh()
                else:
                    QMessageBox.warning(
                        self, 'Cancel Failed',
                        f'Could not cancel job {job_id}. It may have already completed.'
                    )
        finally:
            self._cancel_in_progress = False

    def keyPressEvent(self, event):
        """Handle keyboard shortcuts."""
        if event.key() == Qt.Key_R and event.modifiers() & Qt.ControlModifier:
            self.refresh()
        else:
            super().keyPressEvent(event)

    def closeEvent(self, event):
        """Stop the refresh timer when dialog closes."""
        self._refresh_timer.stop()
        super().closeEvent(event)
