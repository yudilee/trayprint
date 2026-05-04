"""
Notification History for TrayPrint.

Stores recent desktop notifications in memory (max 50) and provides
a QDialog to browse them in reverse chronological order.
"""

from datetime import datetime
from collections import deque
from dataclasses import dataclass, field
from typing import Optional

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget,
    QListWidgetItem, QLabel, QWidget, QMessageBox,
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QFont, QIcon, QColor, QPixmap, QPainter


# ─────────────────────────────────────────────
#  Data Model
# ─────────────────────────────────────────────

@dataclass
class NotificationEntry:
    """A single notification record."""
    timestamp: datetime
    title: str
    message: str
    icon: int = 0          # QStyle.StandardPixmap or 0 for default
    type: str = 'info'     # 'info', 'warning', 'error', 'success'
    entry_id: int = field(default=0, compare=False)


# ─────────────────────────────────────────────
#  Notification History Store
# ─────────────────────────────────────────────

class NotificationHistory:
    """Thread-safe in-memory store of recent notifications (max 50)."""

    def __init__(self, maxlen: int = 50):
        self._maxlen = maxlen
        self._entries: deque[NotificationEntry] = deque(maxlen=maxlen)
        self._counter = 0

    def add(self, title: str, message: str, icon: int = 0,
            type: str = 'info') -> NotificationEntry:
        """Add a notification and return the entry."""
        self._counter += 1
        entry = NotificationEntry(
            timestamp=datetime.now(),
            title=title,
            message=message,
            icon=icon,
            type=type,
            entry_id=self._counter,
        )
        self._entries.appendleft(entry)  # newest first
        return entry

    def get_all(self) -> list[NotificationEntry]:
        """Return all entries (newest first)."""
        return list(self._entries)

    def clear(self) -> None:
        """Remove all entries."""
        self._entries.clear()

    def __len__(self) -> int:
        return len(self._entries)


# ─────────────────────────────────────────────
#  Notification History Dialog
# ─────────────────────────────────────────────

_ICON_MAP = {
    'info':    '\u2139\ufe0f',   # ℹ️
    'warning': '\u26a0\ufe0f',   # ⚠️
    'error':   '\u274c',         # ❌
    'success': '\u2705',         # ✅
}

_COLOR_MAP = {
    'info':    '#3b82f6',
    'warning': '#f59e0b',
    'error':   '#ef4444',
    'success': '#22c55e',
}


class NotificationHistoryDialog(QDialog):
    """Dialog showing recent notification history."""

    def __init__(self, history: NotificationHistory, parent=None):
        super().__init__(parent)
        self._history = history
        self.setWindowTitle('Notification History')
        self.resize(520, 400)
        self.setMinimumSize(380, 280)

        layout = QVBoxLayout(self)

        # Header
        header_layout = QHBoxLayout()
        header_label = QLabel(f'Notifications ({len(self._history)})')
        header_label.setStyleSheet('font-weight: 600; font-size: 14px; color: #1e293b;')
        header_layout.addWidget(header_label)
        header_layout.addStretch()

        self.btn_clear = QPushButton('Clear All')
        self.btn_clear.setStyleSheet("""
            QPushButton {
                padding: 6px 14px;
                background: #fef2f2;
                border: 1px solid #fecaca;
                border-radius: 6px;
                color: #dc2626;
                font-weight: 500;
            }
            QPushButton:hover {
                background: #fee2e2;
            }
        """)
        self.btn_clear.clicked.connect(self._on_clear)
        header_layout.addWidget(self.btn_clear)

        layout.addLayout(header_layout)

        # Notification list
        self.list_widget = QListWidget()
        self.list_widget.setAlternatingRowColors(True)
        self.list_widget.setStyleSheet("""
            QListWidget {
                border: 1px solid #e2e8f0;
                border-radius: 6px;
                background: #ffffff;
                outline: none;
            }
            QListWidget::item {
                padding: 10px 12px;
                border-bottom: 1px solid #f1f5f9;
            }
            QListWidget::item:hover {
                background: #f8fafc;
            }
            QListWidget::item:selected {
                background: #eff6ff;
                color: #1e293b;
            }
        """)
        layout.addWidget(self.list_widget)

        self._populate()

    def _populate(self):
        """Fill the list widget with current notifications."""
        self.list_widget.clear()
        for entry in self._history.get_all():
            emoji = _ICON_MAP.get(entry.type, '\u2139\ufe0f')
            time_str = entry.timestamp.strftime('%H:%M:%S')
            display = f"{emoji}  [{time_str}] {entry.title}"
            item = QListWidgetItem(display)
            item.setToolTip(entry.message)

            # Color coding by type
            color = _COLOR_MAP.get(entry.type, '#475569')
            item.setForeground(QColor(color))
            item.setData(Qt.UserRole, entry.entry_id)
            self.list_widget.addItem(item)

        # Update header count
        parent = self.layout().itemAt(0)
        if parent and parent.layout():
            header_label = parent.layout().itemAt(0)
            if header_label and header_label.widget():
                header_label.widget().setText(
                    f'Notifications ({len(self._history)})'
                )

    def _on_clear(self):
        """Clear all notifications after confirmation."""
        reply = QMessageBox.question(
            self, 'Clear Notifications',
            'Remove all notification history?',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self._history.clear()
            self._populate()

    def refresh(self):
        """Refresh the list (called from external timer or after new notification)."""
        self._populate()
