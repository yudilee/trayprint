"""
Print Job Queue Viewer Dialog for TrayPrint.

Shows active/pending/recent print jobs in a QTableWidget with
Cancel, Refresh, and auto-refresh via QTimer.
"""

from datetime import datetime

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QLabel, QMessageBox, QWidget,
    QAbstractItemView,
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QFont


STATUS_STYLES = {
    'pending':   ('\u23f3  Pending',   '#f59e0b'),   # ⏳
    'printing':  ('\U0001f7e2  Printing', '#22c55e'), # 🟢
    'success':   ('\u2705  Success',   '#22c55e'),   # ✅
    'failed':    ('\u274c  Failed',    '#ef4444'),   # ❌
    'cancelled': ('\u274c  Cancelled', '#6b7280'),   # ❌ grey
}


class PrintQueueDialog(QDialog):
    """Dialog showing current print queue with cancel/refresh controls."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Print Queue')
        self.resize(700, 420)
        self.setMinimumSize(500, 300)

        self._refresh_timer = QTimer(self)
        self._refresh_timer.timeout.connect(self.refresh)
        self._refresh_timer.start(3000)  # auto-refresh every 3 seconds

        layout = QVBoxLayout(self)

        # ── Top bar: refresh button + empty state hint ──
        top_layout = QHBoxLayout()

        title_label = QLabel('Print Job Queue')
        title_label.setStyleSheet('font-weight: 600; font-size: 14px; color: #1e293b;')
        top_layout.addWidget(title_label)
        top_layout.addStretch()

        self.btn_refresh = QPushButton('Refresh')
        self.btn_refresh.setStyleSheet("""
            QPushButton {
                padding: 6px 14px;
                background: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                color: #475569;
                font-weight: 500;
            }
            QPushButton:hover {
                background: #f1f5f9;
                border-color: #cbd5e1;
            }
        """)
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
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                background: #ffffff;
                gridline-color: #f1f5f9;
                color: #1e293b;
                selection-background-color: #eff6ff;
                selection-color: #1e293b;
                alternate-background-color: #f8fafc;
            }
            QHeaderView::section {
                background-color: #f8fafc;
                padding: 10px 8px;
                border: none;
                border-bottom: 2px solid #e2e8f0;
                border-right: 1px solid #f1f5f9;
                font-weight: 600;
                color: #475569;
                font-size: 10px;
                letter-spacing: 0.5px;
                text-transform: uppercase;
            }
            QTableWidget::item {
                padding: 6px 8px;
                border-bottom: 1px solid #f8fafc;
            }
        """)
        layout.addWidget(self.table)

        # ── Empty state label (shown when no jobs) ──
        self.empty_label = QLabel('No active jobs')
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setStyleSheet(
            'color: #94a3b8; font-size: 13px; padding: 30px;'
        )
        self.empty_label.setVisible(False)
        layout.addWidget(self.empty_label)

        # ── Cancel confirmation dialog tracking ──
        self._cancel_in_progress = False

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
                                            (display_status.capitalize(), '#64748b'))
            status_text, status_color = status_info

            status_item = QTableWidgetItem(status_text)
            status_item.setForeground(QColor(status_color))
            font = status_item.font()
            font.setBold(True)
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
                btn_cancel.setStyleSheet("""
                    QPushButton {
                        padding: 4px 10px;
                        background: #fef2f2;
                        border: 1px solid #fecaca;
                        border-radius: 4px;
                        color: #dc2626;
                        font-size: 11px;
                        font-weight: 500;
                    }
                    QPushButton:hover {
                        background: #fee2e2;
                    }
                    QPushButton:disabled {
                        background: #f1f5f9;
                        color: #94a3b8;
                        border-color: #e2e8f0;
                    }
                """)
                btn_cancel.setCursor(Qt.PointingHandCursor)
                job_id = j.get('id', '')
                btn_cancel.clicked.connect(
                    lambda checked, jid=job_id: self._on_cancel(jid)
                )
                actions_layout.addWidget(btn_cancel)
            else:
                # Show a disabled placeholder
                placeholder = QLabel('\u2014')
                placeholder.setStyleSheet('color: #cbd5e1;')
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

    def closeEvent(self, event):
        """Stop the refresh timer when dialog closes."""
        self._refresh_timer.stop()
        super().closeEvent(event)
