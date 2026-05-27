"""
TrayPrint — Main Application Entry Point

Supports command-line arguments and environment variables for unattended
/ silent deployment scenarios.

Command-line arguments:
    --config-path PATH     Specify config file location
    --hub-url URL          Set hub URL on first run
    --agent-key KEY        Set agent key on first run
    --silent               Run without showing UI (for service mode)
    --install-service      Install as Windows service (using nssm)
    --uninstall-service    Remove Windows service
    --version              Print version and exit

Environment variables:
    TRAYPRINT_HUB_URL       Override hub URL
    TRAYPRINT_AGENT_KEY     Override agent key
    TRAYPRINT_CONFIG_PATH   Override config file path
"""

import argparse
import threading
import os
import sys
import json
import subprocess
import webbrowser
from datetime import datetime

from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor, QAction
from PySide6.QtCore import Qt, QTimer, QCoreApplication, QSharedMemory

import server
import websocket_client
import autostart
import ui_settings
from path_utils import get_root_dir, get_data_dir
from logger import get_logger, get_log_path
from log_utils import setup_logging as setup_log_utils
from notification_history import NotificationHistory, NotificationHistoryDialog
from queue_dialog import PrintQueueDialog
from diagnostics_dialog import DiagnosticsDialog
from updater import UpdateChecker

# ── Version ──
APP_VERSION = getattr(server, "APP_VERSION", "3.0.0")

# ── CLI Argument Parser ──

def build_arg_parser():
    """Build and return the argument parser for CLI flags."""
    parser = argparse.ArgumentParser(
        description="TrayPrint — Local Print Agent for Print Hub",
        add_help=False,  # We handle --help ourselves to avoid conflicts
    )
    parser.add_argument(
        "--config-path",
        type=str,
        default=None,
        help="Specify config file location",
    )
    parser.add_argument(
        "--hub-url",
        type=str,
        default=None,
        help="Set hub URL on first run",
    )
    parser.add_argument(
        "--agent-key",
        type=str,
        default=None,
        help="Set agent key on first run",
    )
    parser.add_argument(
        "--silent",
        action="store_true",
        default=False,
        help="Run without showing UI (for service mode)",
    )
    parser.add_argument(
        "--install-service",
        action="store_true",
        default=False,
        help="Install as Windows service (using nssm)",
    )
    parser.add_argument(
        "--uninstall-service",
        action="store_true",
        default=False,
        help="Remove Windows service",
    )
    parser.add_argument(
        "--install-task",
        action="store_true",
        default=False,
        help="Register a Scheduled Task to auto-start TrayPrint at system boot (admin)",
    )
    parser.add_argument(
        "--uninstall-task",
        action="store_true",
        default=False,
        help="Remove the TrayPrint Scheduled Task",
    )
    parser.add_argument(
        "--version",
        action="store_true",
        default=False,
        help="Print version and exit",
    )
    parser.add_argument(
        "--help",
        action="store_true",
        default=False,
        help="Show this help message and exit",
    )
    return parser


def parse_cli_args():
    """
    Parse command-line arguments, filtering out PySide6/Qt arguments.
    Returns a namespace with parsed args.
    """
    parser = build_arg_parser()

    # Filter out Qt-specific arguments (e.g., -style, -platform)
    qt_args = {"-style", "-platform", "-stylesheet", "-qmljsdebugger",
               "-session", "-graphicssystem", "-native"}
    filtered_argv = [sys.argv[0]]
    i = 1
    while i < len(sys.argv):
        arg = sys.argv[i]
        if arg in qt_args:
            i += 2  # Skip the argument and its value
            continue
        if arg.startswith("--") and "=" not in arg:
            # Check if next arg is a value (not a flag)
            if i + 1 < len(sys.argv) and not sys.argv[i + 1].startswith("-"):
                # Check if this is a known argument that takes a value
                known_with_value = {"--config-path", "--hub-url", "--agent-key"}
                if arg in known_with_value:
                    filtered_argv.append(arg)
                    filtered_argv.append(sys.argv[i + 1])
                    i += 2
                    continue
        filtered_argv.append(arg)
        i += 1

    args, _ = parser.parse_known_args(filtered_argv)
    return args


# ── Environment Variable Helpers ──

def apply_env_overrides(config):
    """
    Apply environment variable overrides to the config dict.
    Environment variables take precedence over config.json values.

    Supported variables:
        TRAYPRINT_HUB_URL       → config['hub_url']
        TRAYPRINT_AGENT_KEY     → config['agent_key']
        TRAYPRINT_CONFIG_PATH   → (used to locate config file)
    """
    env_hub = os.environ.get("TRAYPRINT_HUB_URL")
    env_key = os.environ.get("TRAYPRINT_AGENT_KEY")
    env_config = os.environ.get("TRAYPRINT_CONFIG_PATH")

    if env_hub:
        config["hub_url"] = env_hub
    if env_key:
        config["agent_key"] = env_key
    if env_config:
        config["_config_path_override"] = env_config

    return config


def apply_cli_overrides(config, args):
    """
    Apply CLI argument overrides to the config dict.
    CLI arguments take precedence over both config.json and env vars.
    """
    if args.hub_url:
        config["hub_url"] = args.hub_url
    if args.agent_key:
        config["agent_key"] = args.agent_key
    if args.config_path:
        config["_config_path_override"] = args.config_path

    return config


# ── Service Installation Helpers ──

def install_windows_service():
    """Install TrayPrint as a Windows service using nssm."""
    if sys.platform != "win32":
        print("[ERROR] Service installation is only supported on Windows.")
        return False

    nssm_path = _find_nssm()
    if not nssm_path:
        print("[ERROR] nssm.exe not found. Download from https://nssm.cc")
        print("        Place nssm.exe in the application directory or in PATH.")
        return False

    service_name = "TrayPrint"
    app_path = os.path.abspath(sys.argv[0])
    config_path = os.path.join(get_root_dir(), "config.json")

    try:
        # Install the service
        subprocess.run(
            [
                nssm_path, "install", service_name,
                "Application", app_path,
                "AppParameters", "--silent",
                "AppDirectory", get_root_dir(),
                "DisplayName", "TrayPrint Print Agent Service",
                "Description",
                "Local print agent service for Print Hub. "
                "Provides a REST API for printing.",
                "Start", "SERVICE_AUTO_START",
                "AppStdout", os.path.join(get_data_dir(), "logs", "service.log"),
                "AppStderr", os.path.join(get_data_dir(), "logs", "service.err"),
            ],
            check=True,
        )
        print(f"[SUCCESS] Service '{service_name}' installed.")
        print(f"         Start with: nssm start {service_name}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Failed to install service: {e}")
        return False
    except FileNotFoundError:
        print("[ERROR] nssm.exe not found in PATH.")
        return False


def uninstall_windows_service():
    """Remove the TrayPrint Windows service."""
    if sys.platform != "win32":
        print("[ERROR] Service uninstallation is only supported on Windows.")
        return False

    nssm_path = _find_nssm()
    if not nssm_path:
        print("[ERROR] nssm.exe not found.")
        return False

    service_name = "TrayPrint"

    try:
        subprocess.run(
            [nssm_path, "remove", service_name, "confirm"],
            check=True,
        )
        print(f"[SUCCESS] Service '{service_name}' removed.")
        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Failed to remove service: {e}")
        return False


def _find_nssm():
    """Locate nssm.exe in PATH or application directory."""
    # Check PATH
    try:
        result = subprocess.run(
            ["where", "nssm.exe"] if sys.platform == "win32" else ["which", "nssm"],
            capture_output=True, text=True, timeout=5,
        )
        if result.returncode == 0:
            path = result.stdout.strip().split("\n")[0]
            if os.path.isfile(path):
                return path
    except Exception:
        pass

    # Check application directory
    local_path = os.path.join(get_root_dir(), "nssm.exe")
    if os.path.isfile(local_path):
        return local_path

    return None


# ── Scheduled Task Helpers ──


def _find_powershell():
    """Locate PowerShell executable on Windows."""
    if sys.platform != "win32":
        return None
    # Common PowerShell paths
    candidates = [
        r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe",
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    # Fallback: try PATH
    try:
        result = subprocess.run(
            ["where", "powershell.exe"],
            capture_output=True, text=True, timeout=5,
        )
        if result.returncode == 0:
            path = result.stdout.strip().split("\n")[0]
            if os.path.isfile(path):
                return path
    except Exception:
        pass
    return None


def _get_script_path(script_name):
    """Resolve the full path to a PowerShell script in the installer directory."""
    # When running from PyInstaller bundle, scripts are alongside the exe
    exe_dir = os.path.dirname(sys.executable)
    candidate = os.path.join(exe_dir, "installer", script_name)
    if os.path.isfile(candidate):
        return candidate
    # When running from source
    script_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "installer")
    candidate = os.path.join(script_dir, script_name)
    if os.path.isfile(candidate):
        return candidate
    return None


def install_scheduled_task():
    """Register a Scheduled Task to auto-start TrayPrint at system boot.

    Delegates to Register-TrayPrintTask.ps1 which creates a task running
    as NT AUTHORITY\SYSTEM with AtStartup trigger and restart-on-failure.
    """
    if sys.platform != "win32":
        print("[ERROR] Scheduled Task installation is only supported on Windows.")
        return False

    powershell = _find_powershell()
    if not powershell:
        print("[ERROR] PowerShell not found. Cannot register Scheduled Task.")
        return False

    script_path = _get_script_path("Register-TrayPrintTask.ps1")
    if not script_path:
        print("[ERROR] Register-TrayPrintTask.ps1 not found.")
        print("       Ensure the script is in the 'installer' directory.")
        return False

    # Determine the trayprint.exe path
    exe_path = sys.executable
    install_dir = os.path.dirname(exe_path)

    try:
        log.info("Registering Scheduled Task via %s", script_path)
        result = subprocess.run(
            [
                powershell,
                "-ExecutionPolicy", "Bypass",
                "-File", script_path,
                "-InstallPath", install_dir,
            ],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode == 0:
            print("[SUCCESS] Scheduled Task 'TrayPrintAgent' registered.")
            print("         TrayPrint will auto-start at system boot as SYSTEM.")
            if result.stdout.strip():
                print(result.stdout.strip())
            return True
        else:
            print(f"[ERROR] Failed to register Scheduled Task (exit code {result.returncode}).")
            if result.stderr.strip():
                print(f"       {result.stderr.strip()}")
            return False
    except subprocess.TimeoutExpired:
        print("[ERROR] PowerShell script timed out after 30 seconds.")
        return False
    except FileNotFoundError as e:
        print(f"[ERROR] PowerShell executable not found: {e}")
        return False
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
        return False


def uninstall_scheduled_task():
    """Remove the TrayPrint Scheduled Task.

    Delegates to Unregister-TrayPrintTask.ps1 which removes the task
    from Windows Task Scheduler.
    """
    if sys.platform != "win32":
        print("[ERROR] Scheduled Task uninstallation is only supported on Windows.")
        return False

    powershell = _find_powershell()
    if not powershell:
        print("[ERROR] PowerShell not found. Cannot unregister Scheduled Task.")
        return False

    script_path = _get_script_path("Unregister-TrayPrintTask.ps1")
    if not script_path:
        print("[ERROR] Unregister-TrayPrintTask.ps1 not found.")
        print("       Ensure the script is in the 'installer' directory.")
        return False

    try:
        log.info("Unregistering Scheduled Task via %s", script_path)
        result = subprocess.run(
            [
                powershell,
                "-ExecutionPolicy", "Bypass",
                "-File", script_path,
            ],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode == 0:
            print("[SUCCESS] Scheduled Task 'TrayPrintAgent' removed.")
            if result.stdout.strip():
                print(result.stdout.strip())
            return True
        else:
            print(f"[ERROR] Failed to unregister Scheduled Task (exit code {result.returncode}).")
            if result.stderr.strip():
                print(f"       {result.stderr.strip()}")
            return False
    except subprocess.TimeoutExpired:
        print("[ERROR] PowerShell script timed out after 30 seconds.")
        return False
    except FileNotFoundError as e:
        print(f"[ERROR] PowerShell executable not found: {e}")
        return False
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
        return False


# ── Logger Setup ──

# Initialize logging using log_utils (with config-based rotation settings)
_config_for_logging = None
try:
    config_path = os.path.join(get_root_dir(), "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r") as f:
            _config_for_logging = json.load(f)
except Exception:
    pass

log = setup_log_utils(
    name="trayprint",
    config=_config_for_logging,
)

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

def get_config(config_path_override=None):
    """
    Load configuration from config.json, with optional path override.

    Parameters
    ----------
    config_path_override : str or None
        If provided, use this path instead of the default config.json.

    Returns
    -------
    dict
        Configuration dictionary.
    """
    if config_path_override:
        config_path = config_path_override
    else:
        # Check for env var override
        env_config = os.environ.get("TRAYPRINT_CONFIG_PATH")
        if env_config:
            config_path = env_config
        else:
            config_path = os.path.join(get_root_dir(), 'config.json')

    config_data = {"port": 49211}
    loaded_from = None

    # When frozen (PyInstaller), bundled files are in sys._MEIPASS
    if getattr(sys, 'frozen', False):
        try:
            bundled_config = os.path.join(sys._MEIPASS, 'config.json')
            if os.path.exists(bundled_config):
                with open(bundled_config, 'r') as f:
                    config_data.update(json.load(f))
                loaded_from = bundled_config
        except Exception as e:
            log.debug("Could not load bundled config: %s", e)

        # When frozen, also look for existing user config in the source directory
        # (parent of dist/) and in the standard data directory, since users may
        # have previously configured the app from source
        try:
            import platform as _pf
            # Look in sibling directory (e.g., ../config.json relative to dist/)
            sibling_config = os.path.join(os.path.dirname(get_root_dir()), 'config.json')
            if os.path.exists(sibling_config) and sibling_config != config_path:
                with open(sibling_config, 'r') as f:
                    config_data.update(json.load(f))
                loaded_from = sibling_config
                log.info("Found existing config at %s", sibling_config)
            # Also check the user data directory
            data_dir_config = os.path.join(get_data_dir(), 'config.json')
            if os.path.exists(data_dir_config) and data_dir_config != config_path:
                with open(data_dir_config, 'r') as f:
                    config_data.update(json.load(f))
                loaded_from = data_dir_config
                log.info("Found existing config at %s", data_dir_config)
        except Exception as e:
            log.debug("Could not load fallback config: %s", e)

    # Load (or overlay) user config from the standard location
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r') as f:
                # User config overlays bundled defaults
                user_config = json.load(f)
                config_data.update(user_config)
            loaded_from = config_path
        else:
            # If no user config exists yet and we found existing config elsewhere,
            # copy it to the standard location so settings dialog can save there
            if loaded_from and loaded_from != config_path:
                try:
                    os.makedirs(os.path.dirname(config_path), exist_ok=True)
                    import shutil
                    shutil.copy2(loaded_from, config_path)
                    log.info("Copied existing config from %s to %s", loaded_from, config_path)
                except Exception as copy_err:
                    log.debug("Could not copy config: %s", copy_err)
            elif not loaded_from:
                log.warning("Config file not found: %s", config_path)
    except Exception as e:
        log.error("Error loading user config from %s: %s", config_path, e)

    # Apply environment variable overrides
    config_data = apply_env_overrides(config_data)

    return config_data


# ─────────────────────────────────────────────
#  App Class
# ─────────────────────────────────────────────

# ── Single-Instance Guard ──
INSTANCE_LOCK_KEY = "TrayPrint_SingleInstance_v3"


def _check_instance_lock():
    """Try to acquire a lock; return False if another instance is running.
    
    Uses a PID file as primary mechanism (more reliable across crashes),
    with QSharedMemory as secondary guard.
    """
    # ── PID file lock (primary) ──
    pid_path = os.path.join(get_data_dir(), "trayprint.pid")
    try:
        if os.path.exists(pid_path):
            with open(pid_path, 'r') as f:
                old_pid = int(f.read().strip())
            # Check if the process with this PID is still alive
            if os.path.exists(f'/proc/{old_pid}'):
                log.warning("Another TrayPrint instance is already running (PID %d exists)", old_pid)
                return False
            else:
                # Stale PID file — clean it up
                log.info("Removing stale PID file from previous instance (PID %d)", old_pid)
                os.remove(pid_path)
        # Write our PID
        os.makedirs(os.path.dirname(pid_path), exist_ok=True)
        with open(pid_path, 'w') as f:
            f.write(str(os.getpid()))
    except Exception as e:
        log.debug("PID file lock failed, falling back to shared memory: %s", e)
        # ── QSharedMemory (secondary fallback) ──
        shared_mem = QSharedMemory(INSTANCE_LOCK_KEY)
        if shared_mem.attach():
            log.warning("Another TrayPrint instance is already running (shared memory key exists)")
            shared_mem.detach()
            return False
        if not shared_mem.create(1):
            log.warning("Another TrayPrint instance is already running (create failed)")
            return False
        _check_instance_lock._shared_mem = shared_mem

    log.info("Single-instance lock acquired (PID %d)", os.getpid())
    return True


class TrayApp:
    def __init__(self, port):
        # ── Single-instance guard ──
        if not _check_instance_lock():
            log.critical("TrayPrint is already running — exiting.")
            sys.exit(1)

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
            # Failed jobs get a retry action
            if j['status'] == 'failed':
                retry_menu = QMenu("Retry", menu)
                retry_now = retry_menu.addAction("Retry Now")
                retry_now.triggered.connect(lambda checked, job_id=j['id']: self._retry_job(job_id))
                act.setMenu(retry_menu)
                # Also make the main item clickable for quick retry
                act.setToolTip(f"Failed: {j.get('error', 'Unknown error')}")
            else:
                act.setEnabled(False)

    def _retry_job(self, job_id):
        """Retry a failed print job by calling the server's retry endpoint."""
        import requests
        log.info("Retrying job %s from tray menu", job_id)
        self.show_notification("Retry", f"Retrying job {job_id}...")
        threading.Thread(
            target=self._do_retry,
            args=(job_id,),
            daemon=True,
        ).start()

    def _do_retry(self, job_id):
        """Execute the retry in a background thread."""
        try:
            resp = requests.post(
                f'http://127.0.0.1:{self.port}/jobs/{job_id}/retry',
                timeout=30,
            )
            if resp.status_code == 200:
                result = resp.json()
                if result.get('status') == 'success':
                    self.show_notification("Retry Success", f"Job {job_id} completed successfully")
                else:
                    self.show_notification("Retry Failed", f"Job {job_id}: {result.get('error', 'Unknown error')}")
            else:
                error_data = resp.json()
                self.show_notification("Retry Failed", f"Job {job_id}: HTTP {resp.status_code} - {error_data.get('error', '')}")
        except requests.exceptions.ConnectionError:
            self.show_notification("Retry Failed", f"Cannot connect to local server to retry job {job_id}")
        except Exception as e:
            log.error("Retry failed for job %s: %s", job_id, e)
            self.show_notification("Retry Error", f"Failed to retry job {job_id}: {e}")

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
            # Test hub connection immediately so UI doesn't show stale "Disconnected"
            if server.test_hub_connection():
                server._hub_last_status = "Connected"
                log.info("Hub reachable on startup (already connected)")
            else:
                server._hub_last_status = "Initializing..."
                log.info("Hub not reachable on startup — sync loop will retry")
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


# ─────────────────────────────────────────────
#  Crash Diagnostics Handler
#  Captures unhandled exceptions, writes a crash dump,
#  and attempts to upload it to the hub for diagnostics.
# ─────────────────────────────────────────────

CRASH_LOG_PATH = os.path.join(get_data_dir(), "crash_diagnostics.log")


def _crash_handler(exc_type, exc_value, exc_traceback):
    """Global exception hook — writes crash dump and attempts hub upload."""
    import traceback
    from datetime import datetime

    # Skip KeyboardInterrupt and SystemExit
    if issubclass(exc_type, (KeyboardInterrupt, SystemExit)):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    # Build crash report
    tb_lines = traceback.format_exception(exc_type, exc_value, exc_traceback)
    crash_report = {
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "version": APP_VERSION,
        "platform": sys.platform,
        "python_version": sys.version,
        "exception_type": exc_type.__name__,
        "exception_message": str(exc_value),
        "traceback": "".join(tb_lines),
    }

    # Write to local crash log
    try:
        with open(CRASH_LOG_PATH, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 60 + "\n")
            f.write(f"CRASH at {crash_report['timestamp']}\n")
            f.write(f"Version: {crash_report['version']}\n")
            f.write(f"Platform: {crash_report['platform']}\n")
            f.write(f"Python: {crash_report['python_version']}\n")
            f.write(f"Exception: {crash_report['exception_type']}: {crash_report['exception_message']}\n")
            f.write("Traceback:\n")
            f.write(crash_report['traceback'])
            f.write("\n")
    except Exception:
        pass  # Don't crash while handling a crash

    # Attempt to upload crash report to hub (fire-and-forget)
    try:
        config = get_config()
        hub_url = config.get("hub_url", "")
        agent_key = config.get("agent_key", "")
        if hub_url and agent_key:
            import requests
            headers = {
                'Authorization': f'Bearer {agent_key}',
                'Content-Type': 'application/json',
            }
            payload = {
                'event': 'crash.diagnostics',
                'agent_version': crash_report['version'],
                'platform': crash_report['platform'],
                'exception': crash_report['exception_type'],
                'message': crash_report['exception_message'],
                'traceback': crash_report['traceback'],
                'timestamp': crash_report['timestamp'],
            }
            threading.Thread(
                target=lambda: requests.post(
                    f'{hub_url}/api/print-hub/diagnostics/crash',
                    json=payload,
                    headers=headers,
                    timeout=5,
                ),
                daemon=True,
            ).start()
    except Exception:
        pass

    # Log to logger as well
    try:
        log.critical(
            "Unhandled %s: %s\n%s",
            crash_report['exception_type'],
            crash_report['exception_message'],
            crash_report['traceback'],
        )
    except Exception:
        pass

    # Call the original excepthook to preserve default behavior
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


def setup_tray(port, config=None):
    app = TrayApp(port)
    sys.exit(app.run())


def main():
    """Main entry point with CLI argument and environment variable support."""
    args = parse_cli_args()

    # ── Handle --version ──
    if args.version:
        print(f"TrayPrint v{APP_VERSION}")
        print(f"Python: {sys.version}")
        print(f"Platform: {sys.platform}")
        sys.exit(0)

    # ── Handle --help ──
    if args.help:
        parser = build_arg_parser()
        parser.print_help()
        print()
        print("Environment variables:")
        print("  TRAYPRINT_HUB_URL       Override hub URL")
        print("  TRAYPRINT_AGENT_KEY     Override agent key")
        print("  TRAYPRINT_CONFIG_PATH   Override config file path")
        print()
        print("Examples:")
        print("  python app.py --hub-url https://hub.example.com --agent-key mykey")
        print("  python app.py --silent")
        print("  python app.py --install-service")
        print("  python app.py --install-task")
        print("  python app.py --uninstall-task")
        print("  python app.py --version")
        sys.exit(0)

    # ── Handle --install-service ──
    if args.install_service:
        success = install_windows_service()
        sys.exit(0 if success else 1)

    # ── Handle --uninstall-service ──
    if args.uninstall_service:
        success = uninstall_windows_service()
        sys.exit(0 if success else 1)

    # ── Handle --install-task ──
    if args.install_task:
        success = install_scheduled_task()
        sys.exit(0 if success else 1)

    # ── Handle --uninstall-task ──
    if args.uninstall_task:
        success = uninstall_scheduled_task()
        sys.exit(0 if success else 1)

    # ── Load config with overrides ──
    config_path = args.config_path or os.environ.get("TRAYPRINT_CONFIG_PATH")
    config = get_config(config_path_override=config_path)

    # Apply CLI overrides (highest precedence)
    config = apply_cli_overrides(config, args)

    # Validate configuration on startup
    config_errors = validate_config(config)
    if config_errors:
        for err in config_errors:
            log.warning("Config validation: %s", err)

    # ── Handle --silent (no tray UI) ──
    if args.silent:
        log.info("Starting in silent mode (no tray UI)")
        port = config.get("port", 49211)

        # Start Flask server directly
        server_thread = threading.Thread(
            target=server.run_server,
            args=(port,),
            daemon=True,
        )
        server_thread.start()

        # Start hub sync if configured
        hub_url = config.get("hub_url", "")
        agent_key = config.get("agent_key", "")
        if hub_url and agent_key:
            interval = config.get("sync_interval_seconds", 60)
            max_retries = config.get("max_retries", 3)
            retry_delay = config.get("retry_delay_seconds", 60)
            server.start_hub_sync(hub_url, agent_key, interval, max_retries, retry_delay)

        log.info("Silent mode active — server running on port %d", port)

        # Keep the main thread alive
        try:
            threading.Event().wait()
        except KeyboardInterrupt:
            log.info("Shutting down (Ctrl+C)")
        sys.exit(0)

    # ── Normal mode (with tray UI) ──
    setup_tray(config.get('port', 49211), config)


# Install global crash handler before anything else
sys.excepthook = _crash_handler

if __name__ == '__main__':
    main()
