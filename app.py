import threading
import os
import sys
import json
import subprocess
import webbrowser
from datetime import datetime

from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor, QAction
from PySide6.QtCore import QTimer, QCoreApplication

import server
import autostart
from path_utils import get_root_dir
from logger import get_logger, get_log_path

log = get_logger()

# ─────────────────────────────────────────────
#  Tray Icon Generation
# ─────────────────────────────────────────────

def create_tray_icon():
    """Generates a printer icon for the system tray using QPainter."""
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
        
        self.tray = QSystemTrayIcon(create_tray_icon(), self.app)
        self.tray.setToolTip(f"Trayprint - Local Print Service (Port {port})")
        
        self.menu = QMenu()
        self.tray.setContextMenu(self.menu)
        
        # Connect menu aboutToShow to dynamic update
        self.menu.aboutToShow.connect(self.update_menu)
        
        self.tray.show()
        log.info("Tray icon initialized (PySide6) — API on port %d", port)

    def update_menu(self):
        """Rebuilds the menu items with fresh status info."""
        self.menu.clear()
        
        # Header Info
        header = self.menu.addAction(f"Trayprint v3.0 - Port {self.port}")
        header.setEnabled(False)
        
        hub_status = server.get_hub_status()
        hub_info = self.menu.addAction(f"Hub: {hub_status}")
        hub_info.setEnabled(False)
        
        printers_count = server.get_cached_printer_count()
        printers_info = self.menu.addAction(f"Printers: {printers_count} found")
        printers_info.setEnabled(False)
        
        self.menu.addSeparator()
        
        # Actions
        settings_act = self.menu.addAction("Settings (Web Dashboard)")
        settings_act.triggered.connect(self.open_settings_browser)
        
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

    def open_settings_browser(self):
        webbrowser.open(f"http://127.0.0.1:{self.port}/settings")

    def populate_jobs_menu(self, menu):
        jobs = server._job_queue.list_recent(10)
        if not jobs:
            act = menu.addAction("No recent jobs")
            act.setEnabled(False)
            return

        for j in reversed(jobs):
            icon_char = '✓' if j['status'] == 'success' else '✗' if j['status'] == 'failed' else '…'
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
            self.tray.showMessage("Trayprint", "Auto-start disabled")
        else:
            autostart.enable_autostart()
            self.tray.showMessage("Trayprint", "Auto-start enabled")

    def restart_app(self):
        log.info("User requested restart")
        QCoreApplication.quit()
        # Give it a moment to cleanup
        os.execl(sys.executable, sys.executable, *sys.argv)

    def quit_app(self):
        log.info("User requested exit")
        QCoreApplication.quit()
        os._exit(0)

    def run(self):
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
    setup_tray(config.get('port', 49211))
