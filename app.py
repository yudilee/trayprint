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
import websocket_client
import autostart
import ui_settings
from path_utils import get_root_dir
from logger import get_logger, get_log_path
from notification_history import NotificationHistory, NotificationHistoryDialog
from queue_dialog import PrintQueueDialog
from diagnostics_dialog import DiagnosticsDialog
from updater import UpdateChecker

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

        # macOS-specific: high-DPI pixmap support
        if sys.platform == 'darwin':
            self.app.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
            log.info("macOS: enabled High-DPI pixmap support")

        # Printing state tracking
        self._is_printing = False
        self._normal_icon = create_tray_icon(printing=False)
        self._printing_icon = create_tray_icon(printing=True)

        self.tray = QSystemTrayIcon(self._normal_icon, self.app)
        # Build a rich tooltip with status information
        self._update_tooltip_text()

        self.menu = QMenu()
        self.tray.setContextMenu(self.menu)

        # Connect menu aboutToShow to dynamic update
        self.menu.aboutToShow.connect(self.update_menu)

        # Notification history store
        self.notification_history = NotificationHistory(maxlen=50)

        # Track dialogs (prevent multiple instances)
        self._queue_dialog = None
        self._notification_dialog = None
        self._update_dialog_shown = False

        # Polling timer: check server printing state every 2 seconds
        self._state_timer = QTimer(self.app)
        self._state_timer.timeout.connect(self._poll_printing_state)
        self._state_timer.start(2000)

        # ── Auto-Update Checker ──
        self._update_checker = None
        self._init_update_checker()

        # ── WebSocket Client (optional, for real-time updates) ──
        self._ws_client = None

        self.tray.show()
        log.info("Tray icon initialized (PySide6) — API on port %d", port)

    def _init_update_checker(self):
        """Initialize the UpdateChecker from config settings."""
        config = get_config()
        hub_url = config.get("hub_url", "")
        agent_key = config.get("agent_key", "")
        if not hub_url or not agent_key:
            log.debug("UpdateChecker: hub not configured, skipping")
            return

        self._update_checker = UpdateChecker(
            hub_url=hub_url,
            agent_key=agent_key,
            check_interval=86400,  # daily
            on_update_available=self._on_update_available,
        )
        self._update_checker.start()
        log.info("UpdateChecker started (hub=%s)", hub_url)

    def _on_update_available(self, version, release_notes):
        """Callback when an update is detected."""
        log.info("Update available: v%s", version)
        msg = f"Update v{version} available — open Settings to install"
        if release_notes:
            msg += f"\n\nRelease notes:\n{release_notes[:200]}"
        self.show_notification("Update Available", msg)
        self._update_dialog_shown = False

    # ── Config reload (apply settings without restart) ──

    def reload_config(self):
        """Re-read config.json and apply changes live without restarting."""
        log.info("Reloading config from disk...")
        try:
            config = get_config()
            old_hub_url = server._hub_url

            # Reload server-side config
            server.reload_config()

            new_hub_url = config.get("hub_url", "")
            agent_key = config.get("agent_key", "")
            interval = config.get("sync_interval_seconds", 60)

            # If hub URL changed, restart the hub sync with new settings
            if new_hub_url and new_hub_url != old_hub_url:
                log.info("Hub URL changed: %s → %s", old_hub_url, new_hub_url)
                if agent_key:
                    max_retries = config.get("max_retries", 3)
                    retry_delay = config.get("retry_delay_seconds", 60)
                    server.start_hub_sync(new_hub_url, agent_key, interval, max_retries, retry_delay)

            # Update tray tooltip
            self.tray.setToolTip(f"Trayprint v{server.APP_VERSION} - Local Print Service (Port {self.port})")

            log.info("Config reloaded successfully — settings applied live")
        except Exception as e:
            log.error("Failed to reload config: %s", e)

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
        else:
            self.tray.setIcon(self._normal_icon)
        self._update_tooltip_text()

    def _update_tooltip_text(self):
        """Build a rich tooltip showing connection status, printer count, and queue length."""
        hub_status = server.get_hub_status()
        printer_count = server.get_cached_printer_count()
        try:
            queue_info = server.get_queue_status()
            queue_depth = queue_info.get('total_queued', 0)
            processing_count = queue_info.get('processing', 0)
        except Exception:
            queue_depth = 0
            processing_count = 0

        printing_indicator = ""
        try:
            if server.is_printing():
                printing_indicator = "\n🖨️ Printing..."
        except Exception:
            pass

        self.tray.setToolTip(
            f"TrayPrint v{server.APP_VERSION} — Port {self.port}\n"
            f"Hub: {hub_status}\n"
            f"Printers: {printer_count}\n"
            f"Queue: {queue_depth} pending ({processing_count} processing)"
            f"{printing_indicator}"
        )

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

        # ── Diagnostics action ──
        diagnostics_act = self.menu.addAction("Diagnostics...")
        diagnostics_act.triggered.connect(self.open_diagnostics)

        # ── Show Notifications action ──
        notif_count = len(self.notification_history)
        notif_act = self.menu.addAction(f"Show Notifications{f' ({notif_count})' if notif_count else ''}...")
        notif_act.triggered.connect(self.open_notification_history)

        # ── Check for Updates ──
        if self._update_checker:
            if self._update_checker.update_ready:
                check_upd_act = self.menu.addAction(
                    f"\u2b06 Update v{self._update_checker.latest_version} Available"
                )
            else:
                check_upd_act = self.menu.addAction("Check for Updates...")
            check_upd_act.triggered.connect(self.check_for_updates)

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

    def open_diagnostics(self):
        """Open the diagnostics dialog."""
        dlg = DiagnosticsDialog()
        dlg.exec()

    def open_settings_window(self):
        """Open the settings dialog with apply callback."""
        dlg = ui_settings.SettingsWindow()
        dlg.apply_callback = self.reload_config
        # Pass the update checker reference for the Software Updates section
        if hasattr(self, '_update_checker') and self._update_checker:
            dlg.update_checker = self._update_checker
        dlg.exec()

    def check_for_updates(self):
        """Trigger an immediate update check."""
        if not self._update_checker:
            self.show_notification("Update Check", "Update checker not configured (hub not set up)")
            return

        self.show_notification("Update Check", "Checking for updates...")
        # Run check in a thread to avoid blocking the UI
        def _do_check():
            try:
                self._update_checker.check_for_updates()
                if self._update_checker.update_ready:
                    self.show_notification(
                        "Update Available",
                        f"v{self._update_checker.latest_version} is available — open Settings to install"
                    )
                else:
                    self.show_notification("Update Check", "You're up to date!")
            except Exception as e:
                self.show_notification("Update Check Failed", str(e))
        threading.Thread(target=_do_check, daemon=True).start()

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
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', log_path])
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
        # Stop WebSocket client if running
        if self._ws_client:
            try:
                self._ws_client.stop()
            except Exception as e:
                log.warning("Error stopping WebSocket client: %s", e)
        QCoreApplication.quit()
        os._exit(0)

    def closeEvent(self, event):
        """Override the window close event to minimize to tray instead of quitting."""
        log.debug("Close event intercepted — minimizing to tray")
        event.ignore()
        self._hide_window()

    def _hide_window(self):
        """Hide any visible dialogs to the tray instead of closing."""
        if self._queue_dialog and self._queue_dialog.isVisible():
            self._queue_dialog.hide()
        if self._notification_dialog and self._notification_dialog.isVisible():
            self._notification_dialog.hide()
        # Show a notification to inform the user the app is still running
        self.tray.showMessage(
            "TrayPrint",
            "Application minimized to tray. Double-click the tray icon to restore.",
            QSystemTrayIcon.Information,
            3000
        )

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

        # Start hub sync loop immediately (don't wait for Flask factory)
        # This ensures status auto-connects on startup rather than staying "Disconnected"
        config = get_config()
        hub_url = config.get("hub_url", "")
        agent_key = config.get("agent_key", "")
        if hub_url and agent_key:
            interval = config.get("sync_interval_seconds", 60)
            max_retries = config.get("max_retries", 3)
            retry_delay = config.get("retry_delay_seconds", 60)
            server.start_hub_sync(hub_url, agent_key, interval, max_retries, retry_delay)
            log.info("Hub sync loop started from app startup (hub=%s)", hub_url)

        # Start optional WebSocket client for real-time queue updates
        self._on_ws_event(config)

        return self.app.exec()

    def _on_ws_event(self, config):
        """Start the optional WebSocket client and define event callback."""
        def _handle_ws_event(event_data):
            """Callback invoked when a WebSocket event is received from the hub."""
            event = event_data.get('event', '')
            channel = event_data.get('channel', '')
            data = event_data.get('data', {})
            log.debug("WebSocket event: %s on %s", event, channel)

            # Queue updated — request immediate queue poll
            if event == 'queue.updated':
                log.info("Queue update received via WebSocket, requesting refresh")
                try:
                    server.request_queue_refresh()
                except Exception as e:
                    log.warning("Queue refresh request error: %s", e)

            # Job status changes — show a notification
            elif event == 'job.status.updated':
                job_data = data if isinstance(data, dict) else {}
                job_id = job_data.get('job_id', '?')
                status = job_data.get('status', '?')
                self.show_notification("Job Update", f"Job {job_id}: {status}")

        # Start WebSocket client in background
        self._ws_client = websocket_client.start_websocket_client(
            config=config,
            on_event=_handle_ws_event,
        )
        if self._ws_client:
            log.info("WebSocket real-time client started")
        else:
            log.info("WebSocket client not started (will use polling fallback)")


def setup_tray(port, config=None):
    app = TrayApp(port)
    sys.exit(app.run())


if __name__ == '__main__':
    config = get_config()

    # Validate configuration on startup
    config_errors = validate_config(config)
    if config_errors:
        for err in config_errors:
            log.warning("Config validation: %s", err)

    setup_tray(config.get('port', 49211), config)
