import threading
import os
import sys
import json
import subprocess
import webbrowser
from datetime import datetime

from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor, QAction
from PySide6.QtCore import Qt, QTimer, QCoreApplication

import server
import autostart
import ui_settings
from path_utils import get_root_dir
from logger import get_logger, get_log_path
from notification_history import NotificationHistory, NotificationHistoryDialog
from queue_dialog import PrintQueueDialog

log = get_logger()

# ─────────────────────────────────────────────
#  Config Validation
# ─────────────────────────────────────────────

def validate_config(config):
    """Validate runtime configuration and return a list of error messages."""
    errors = []
    if not config.get('hub_url'):
        errors.append("hub_url is not configured")
    if not config.get('agent_key'):
        errors.append("agent_key is not configured")
    if config.get('sync_interval_seconds', 60) < 10:
        errors.append("sync_interval_seconds should be >= 10")
    return errors


# ─────────────────────────────────────────────
#  Tray Icon Generation
# ─────────────────────────────────────────────

def create_tray_icon(printing=False):
    """Generates a printer icon for the system tray using QPainter.
    When printing=True, adds a small animated-style overlay dot."""
    pixmap = QPixmap(64, 64)
    pixmap.fill(QColor(0, 0, 0, 0))

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)

    # Body
    painter.setBrush(QColor(50, 150, 250))
    painter.setPen(QColor(50, 150, 250))
    painter.drawRoundedRect(8, 16, 48, 40, 4, 4)

    # Top Paper
    painter.setBrush(QColor(255, 255, 255))
    painter.setPen(QColor(255, 255, 255))
    painter.drawRect(16, 8, 32, 16)

    # Bottom Paper (Exit)
    painter.drawRect(16, 40, 32, 20)

    # Lines on paper
    painter.setPen(QColor(100, 100, 100))
    painter.drawLine(20, 45, 44, 45)
    painter.drawLine(20, 52, 44, 52)

    if printing:
        # Green indicator dot in bottom-right corner
        painter.setBrush(QColor(34, 197, 94))   # green-500
        painter.setPen(QColor(34, 197, 94))
        painter.drawEllipse(48, 48, 12, 12)
        # Inner white dot for contrast
        painter.setBrush(QColor(255, 255, 255))
        painter.setPen(QColor(255, 255, 255))
        painter.drawEllipse(51, 51, 6, 6)

    painter.end()
    return QIcon(pixmap)


# ─────────────────────────────────────────────
#  Config
# ─────────────────────────────────────────────

def get_config():
    config_path = os.path.join(get_root_dir(), 'config.json')
    config_data = {"port": 49211}
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                config_data.update(json.load(f))
    except Exception as e:
        log.error("Error loading config: %s", e)
    return config_data


# ─────────────────────────────────────────────
#  App Class
# ─────────────────────────────────────────────

class TrayApp:
    def __init__(self, port):
        self.port = port
        self.app = QApplication(sys.argv)
        self.app.setQuitOnLastWindowClosed(False)

        # Printing state tracking
        self._is_printing = False
        self._normal_icon = create_tray_icon(printing=False)
        self._printing_icon = create_tray_icon(printing=True)

        self.tray = QSystemTrayIcon(self._normal_icon, self.app)
        self.tray.setToolTip(f"Trayprint v{server.APP_VERSION} - Local Print Service (Port {port})")

        self.menu = QMenu()
        self.tray.setContextMenu(self.menu)

        # Connect menu aboutToShow to dynamic update
        self.menu.aboutToShow.connect(self.update_menu)

        # Notification history store
        self.notification_history = NotificationHistory(maxlen=50)

        # Track dialogs (prevent multiple instances)
        self._queue_dialog = None
        self._notification_dialog = None

        # Polling timer: check server printing state every 2 seconds
        self._state_timer = QTimer(self.app)
        self._state_timer.timeout.connect(self._poll_printing_state)
        self._state_timer.start(2000)

        self.tray.show()
        log.info("Tray icon initialized (PySide6) — API on port %d", port)

    # ── Printing state management ──

    def _poll_printing_state(self):
        """Poll server for printing state and update tray icon/tooltip."""
        try:
            currently_printing = server.is_printing()
        except Exception:
            currently_printing = False

        if currently_printing != self._is_printing:
            self._is_printing = currently_printing
            self._update_tray_for_printing_state()

    def _update_tray_for_printing_state(self):
        """Update tray icon and tooltip based on printing state."""
        if self._is_printing:
            self.tray.setIcon(self._printing_icon)
            base = self.tray.toolTip().split('\n')[0]
            self.tray.setToolTip(base + "\n🖨️ Printing...")
        else:
            self.tray.setIcon(self._normal_icon)
            base = self.tray.toolTip().split('\n')[0]
            self.tray.setToolTip(base)

    # ── Menu building ──

    def update_menu(self):
        """Rebuilds the menu items with fresh status info."""
        self.menu.clear()

        # Header Info
        header = self.menu.addAction(f"Trayprint v{server.APP_VERSION} - Port {self.port}")
        header.setEnabled(False)

        hub_status = server.get_hub_status()
        printers_count = server.get_cached_printer_count()

        # Add printing indicator to hub status line
        printing_indicator = ""
        try:
            if server.is_printing():
                printing_indicator = " 🖨️ Printing..."
        except Exception:
            pass

        hub_info = self.menu.addAction(f"Hub: {hub_status}{printing_indicator}")
        hub_info.setEnabled(False)

        printers_info = self.menu.addAction(f"Printers: {printers_count} found")
        printers_info.setEnabled(False)

        self.menu.addSeparator()

        # ── View Queue action ──
        view_queue_act = self.menu.addAction("View Queue...")
        view_queue_act.triggered.connect(self.open_queue_dialog)

        self.menu.addSeparator()

        # ── Show Notifications action ──
        notif_count = len(self.notification_history)
        notif_act = self.menu.addAction(f"Show Notifications{f' ({notif_count})' if notif_count else ''}...")
        notif_act.triggered.connect(self.open_notification_history)

        # ── Settings ──
        settings_act = self.menu.addAction("Settings")
        settings_act.triggered.connect(self.open_settings_window)

        # Recent Jobs Submenu
        jobs_menu = self.menu.addMenu("Recent Jobs")
        self.populate_jobs_menu(jobs_menu)

        self.menu.addSeparator()

        view_logs_act = self.menu.addAction("View Logs")
        view_logs_act.triggered.connect(self.view_logs)

        autostart_act = self.menu.addAction("Auto-start on Login")
        autostart_act.setCheckable(True)
        autostart_act.setChecked(autostart.is_autostart_enabled())
        autostart_act.triggered.connect(self.toggle_autostart)

        self.menu.addSeparator()

        restart_act = self.menu.addAction("Restart App")
        restart_act.triggered.connect(self.restart_app)

        exit_act = self.menu.addAction("Exit")
        exit_act.triggered.connect(self.quit_app)

    # ── Dialog openers ──

    def open_queue_dialog(self):
        """Open the print queue viewer dialog (singleton pattern)."""
        if self._queue_dialog is not None and self._queue_dialog.isVisible():
            self._queue_dialog.raise_()
            self._queue_dialog.activateWindow()
            return

        self._queue_dialog = PrintQueueDialog()
        self._queue_dialog.setAttribute(Qt.WA_DeleteOnClose, True)
        self._queue_dialog.destroyed.connect(lambda: setattr(self, '_queue_dialog', None))
        self._queue_dialog.show()

    def open_notification_history(self):
        """Open the notification history dialog (singleton pattern)."""
        if self._notification_dialog is not None and self._notification_dialog.isVisible():
            self._notification_dialog.raise_()
            self._notification_dialog.activateWindow()
            return

        self._notification_dialog = NotificationHistoryDialog(self.notification_history)
        self._notification_dialog.setAttribute(Qt.WA_DeleteOnClose, True)
        self._notification_dialog.destroyed.connect(lambda: setattr(self, '_notification_dialog', None))
        self._notification_dialog.show()

    # ── Existing methods ──

    def open_settings_window(self):
        ui_settings.show_settings()

    def populate_jobs_menu(self, menu):
        jobs = server._job_queue.list_recent(10)
        if not jobs:
            act = menu.addAction("No recent jobs")
            act.setEnabled(False)
            return

        for j in reversed(jobs):
            icon_char = '\u2713' if j['status'] == 'success' else '\u2717' if j['status'] == 'failed' else '\u2026'
            label = f"{icon_char} {j['printer']} ({j['type']}) {j['created_at'][11:19]}"
            act = menu.addAction(label)
            act.setEnabled(False)

    def view_logs(self):
        log_path = get_log_path()
        if sys.platform == 'win32':
            os.startfile(log_path)
        else:
            subprocess.Popen(['xdg-open', log_path])

    def toggle_autostart(self):
        if autostart.is_autostart_enabled():
            autostart.disable_autostart()
            self.show_notification("Trayprint", "Auto-start disabled")
        else:
            autostart.enable_autostart()
            self.show_notification("Trayprint", "Auto-start enabled")

    def restart_app(self):
        log.info("User requested restart")
        QCoreApplication.quit()
        # Give it a moment to cleanup
        os.execl(sys.executable, sys.executable, *sys.argv)

    def quit_app(self):
        log.info("User requested exit")
        QCoreApplication.quit()
        os._exit(0)

    def show_notification(self, title, message):
        """Shows a toast notification AND stores it in notification history."""
        # Store in history
        self.notification_history.add(title, message, type='info')

        # Show toast on the main GUI thread
        QTimer.singleShot(0, lambda: self.tray.showMessage(title, message))

    def run(self):
        # Bind server notifications to our tray
        server._notification_callback = self.show_notification

        # Run Flask server in background
        server_thread = threading.Thread(target=server.run_server, args=(self.port,))
        server_thread.daemon = True
        server_thread.start()

        return self.app.exec()


def setup_tray(port):
    app = TrayApp(port)
    sys.exit(app.run())


if __name__ == '__main__':
    config = get_config()

    # Validate configuration on startup
    config_errors = validate_config(config)
    if config_errors:
        for err in config_errors:
            log.warning("Config validation: %s", err)

    setup_tray(config.get('port', 49211))
