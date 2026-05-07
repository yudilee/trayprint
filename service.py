"""
TrayPrint — Windows Service Wrapper

Wraps the TrayPrint Flask server as a Windows service using pywin32.
Supports start, stop, restart, install, and remove commands.

Usage:
    python service.py install      # Install the Windows service
    python service.py remove       # Remove the Windows service
    python service.py start        # Start the service
    python service.py stop         # Stop the service
    python service.py restart      # Restart the service
    python service.py debug        # Run in foreground (for testing)

The service runs the Flask API server without the tray UI, suitable for
headless/server environments.

Logs are written to:
  - Windows Event Log (via pywin32)
  - logs/app.log in the application directory (via log_utils)
"""

import os
import sys
import threading
import time

# Optional pywin32 import — gracefully degrade if not available
try:
    import win32serviceutil
    import win32service
    import win32event
    import servicemanager
    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

# Ensure the project root is on sys.path
_project_root = os.path.abspath(os.path.dirname(__file__))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from path_utils import get_root_dir
from log_utils import setup_logging, get_log_path

# Configure logging early
log = setup_logging(name="trayprint-service")

# Import server module (will use the logging configured above)
import server

APP_VERSION = getattr(server, "APP_VERSION", "3.0.0")
DEFAULT_PORT = 49211


def _get_config():
    """Load configuration from config.json."""
    import json
    config_path = os.path.join(get_root_dir(), "config.json")
    try:
        if os.path.exists(config_path):
            with open(config_path, "r") as f:
                return json.load(f)
    except Exception as e:
        log.error("Failed to load config: %s", e)
    return {}


class TrayPrintService(win32serviceutil.ServiceFramework):
    """
    Windows Service implementation for TrayPrint Flask server.

    Runs the Flask API server in a background thread, handling start/stop
    lifecycle via Windows Service Control Manager.
    """

    # Windows Service metadata
    _svc_name_ = "TrayPrint"
    _svc_display_name_ = "TrayPrint Print Agent Service"
    _svc_description_ = (
        "Local print agent service for Print Hub. "
        "Provides a REST API for printing and communicates with the Print Hub server."
    )

    def __init__(self, args):
        super().__init__(args)
        self._stop_event = threading.Event()
        self._server_thread = None
        self._flask_app = None
        self._port = DEFAULT_PORT

        # Register event log source
        if HAS_PYWIN32:
            try:
                servicemanager.LogMsg(
                    servicemanager.EVENTLOG_INFORMATION_TYPE,
                    servicemanager.PYS_SERVICE_STARTING,
                    (self._svc_display_name_, ""),
                )
            except Exception:
                pass

    def SvcDoRun(self):
        """Main service entry point — called by the Service Control Manager."""
        try:
            if HAS_PYWIN32:
                servicemanager.LogMsg(
                    servicemanager.EVENTLOG_INFORMATION_TYPE,
                    servicemanager.PYS_SERVICE_STARTED,
                    (self._svc_display_name_, ""),
                )
            log.info("Service starting...")

            # Load config
            config = _get_config()
            self._port = config.get("port", DEFAULT_PORT)

            # Create the Flask app
            self._flask_app = server.create_app()

            # Run Flask in a background thread
            self._server_thread = threading.Thread(
                target=self._run_flask,
                daemon=True,
                name="FlaskServer",
            )
            self._server_thread.start()

            # Start hub sync if configured
            hub_url = config.get("hub_url", "")
            agent_key = config.get("agent_key", "")
            if hub_url and agent_key:
                interval = config.get("sync_interval_seconds", 60)
                max_retries = config.get("max_retries", 3)
                retry_delay = config.get("retry_delay_seconds", 60)
                server.start_hub_sync(
                    hub_url, agent_key, interval, max_retries, retry_delay
                )
                log.info("Hub sync started (hub=%s)", hub_url)

            log.info("Service started successfully on port %d", self._port)

            # Wait for stop signal
            self._stop_event.wait()

        except Exception as e:
            log.error("Service run error: %s", e)
            if HAS_PYWIN32:
                try:
                    servicemanager.LogMsg(
                        servicemanager.EVENTLOG_ERROR_TYPE,
                        servicemanager.PYS_SERVICE_STARTED,
                        (self._svc_display_name_, str(e)),
                    )
                except Exception:
                    pass
            raise

    def _run_flask(self):
        """Run the Flask server (runs in background thread)."""
        try:
            # Suppress Werkzeug access logs
            import logging as stdlib_logging
            werkzeug_log = stdlib_logging.getLogger("werkzeug")
            werkzeug_log.setLevel(stdlib_logging.ERROR)

            log.info("Starting Flask server on 127.0.0.1:%d", self._port)
            self._flask_app.run(
                host="127.0.0.1",
                port=self._port,
                debug=False,
                use_reloader=False,
            )
        except Exception as e:
            log.error("Flask server error: %s", e)

    def SvcStop(self):
        """Handle service stop request from SCM."""
        log.info("Service stop requested...")
        if HAS_PYWIN32:
            try:
                servicemanager.LogMsg(
                    servicemanager.EVENTLOG_INFORMATION_TYPE,
                    servicemanager.PYS_SERVICE_STOPPED,
                    (self._svc_display_name_, ""),
                )
            except Exception:
                pass

        # Signal the server thread to stop
        self._stop_event.set()

        # Notify SCM that we're stopping
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)

        # Give the server thread time to shut down
        if self._server_thread and self._server_thread.is_alive():
            self._server_thread.join(timeout=10)

        log.info("Service stopped")


def _run_debug():
    """Run the service in debug/foreground mode (for testing)."""
    print("=" * 60)
    print(f"  TrayPrint Service v{APP_VERSION} — Debug Mode")
    print("=" * 60)
    print()
    print("Running Flask server in foreground. Press Ctrl+C to stop.")
    print()

    config = _get_config()
    port = config.get("port", DEFAULT_PORT)

    app = server.create_app()

    # Start hub sync if configured
    hub_url = config.get("hub_url", "")
    agent_key = config.get("agent_key", "")
    if hub_url and agent_key:
        interval = config.get("sync_interval_seconds", 60)
        max_retries = config.get("max_retries", 3)
        retry_delay = config.get("retry_delay_seconds", 60)
        server.start_hub_sync(hub_url, agent_key, interval, max_retries, retry_delay)

    try:
        import logging as stdlib_logging
        werkzeug_log = stdlib_logging.getLogger("werkzeug")
        werkzeug_log.setLevel(stdlib_logging.INFO)

        app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)
    except KeyboardInterrupt:
        print("\nShutting down...")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


def main():
    """Entry point — parse command and dispatch."""
    if not HAS_PYWIN32:
        print("[WARNING] pywin32 not installed. Service commands require Windows with pywin32.")
        print("          Falling back to debug mode (foreground server).")
        print()

    if len(sys.argv) < 2:
        print("Usage: python service.py [command]")
        print()
        print("Commands:")
        print("  install      Install the Windows service")
        print("  remove       Remove the Windows service")
        print("  start        Start the service")
        print("  stop         Stop the service")
        print("  restart      Restart the service")
        print("  debug        Run in foreground (for testing)")
        print()
        print("Examples:")
        print("  python service.py install")
        print("  python service.py start")
        print("  python service.py debug")
        return

    command = sys.argv[1].lower()

    if command == "debug":
        _run_debug()
    elif HAS_PYWIN32:
        # Dispatch to pywin32 service framework
        try:
            if command == "install":
                # Before installing, ensure we handle the path correctly
                sys.argv[0] = os.path.abspath(__file__)
            win32serviceutil.HandleCommandLine(TrayPrintService)
        except Exception as e:
            print(f"[ERROR] Service command failed: {e}")
            sys.exit(1)
    else:
        print(f"[ERROR] pywin32 is required for '{command}' command.")
        print("        Install with: pip install pywin32")
        print("        Or use 'debug' mode for foreground testing.")
        sys.exit(1)


if __name__ == "__main__":
    main()
