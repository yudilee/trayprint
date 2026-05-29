"""
Notification History for TrayPrint.

Stores recent desktop notifications in memory (max 50) and provides
a QDialog to browse them in reverse chronological order.
"""

import os
import json
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

from path_utils import get_data_dir
from theme import get_stylesheet, get_palette


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
    'info':    '#5ba3d9',
    'warning': '#f0b34b',
    'error':   '#e66565',
    'success': '#4caf88',
}


class NotificationHistoryDialog(QDialog):
    """Dialog showing recent notification history."""

    def __init__(self, history: NotificationHistory, parent=None):
        super().__init__(parent)
        self._history = history
        self.setWindowTitle('Notification History')
        self.setMinimumSize(500, 400)
        self.resize(520, 420)

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

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # Header
        header_layout = QHBoxLayout()
        header_label = QLabel(f'Notifications ({len(self._history)})')
        header_label.setStyleSheet(
            'font-weight: 700; font-size: 15px; color: %s; background: transparent;'
            % self._theme['text_primary']
        )
        header_layout.addWidget(header_label)
        header_layout.addStretch()

        self.btn_clear = QPushButton('Clear All')
        self.btn_clear.setMinimumHeight(30)
        self.btn_clear.setToolTip('Remove all notification history')
        self.btn_clear.setStyleSheet(
            "QPushButton {"
            "  padding: 6px 14px;"
            "  background: %s;"
            "  border: 1px solid %s;"
            "  border-radius: 6px;"
            "  color: %s;"
            "  font-weight: 500;"
            "}"
            "QPushButton:hover {"
            "  background: %s;"
            "}"
            % (
                self._theme['bg'],
                self._theme['error'],
                self._theme['error'],
                '#3a2020' if self._theme_id == 'dark' else '#f8d7da',
            )
        )
        self.btn_clear.clicked.connect(self._on_clear)
        header_layout.addWidget(self.btn_clear)

        layout.addLayout(header_layout)

        # Notification list
        self.list_widget = QListWidget()
        self.list_widget.setAlternatingRowColors(True)
        layout.addWidget(self.list_widget)

        # Apply shared theme stylesheet
        self.setStyleSheet(get_stylesheet(self._theme_id))

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
            color = _COLOR_MAP.get(entry.type, self._theme['text_secondary'])
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
