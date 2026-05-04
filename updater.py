"""
Automatic update checker for TrayPrint.
Queries Print Hub for the latest agent version, downloads and installs updates.
"""
import os
import sys
import json
import time
import logging
import hashlib
import tempfile
import threading
from pathlib import Path

import requests

logger = logging.getLogger(__name__)


class UpdateChecker(threading.Thread):
    """Periodically checks for new TrayPrint versions from Print Hub."""

    def __init__(self, hub_url, agent_key, check_interval=86400,
                 on_update_available=None):
        super().__init__(daemon=True)
        self.hub_url = hub_url.rstrip('/')
        self.agent_key = agent_key
        self.check_interval = check_interval  # default: daily (86400s)
        self.on_update_available = on_update_available  # callback(version, notes)
        self._stop_event = threading.Event()
        self._latest_version = None
        self._download_url = None
        self._release_notes = None
        self._sha256 = None
        self._mandatory = False
        self._update_ready = False
        self._download_progress = 0.0

    # ── Public API ──

    def stop(self):
        """Signal the update checker thread to stop."""
        self._stop_event.set()

    @property
    def latest_version(self):
        return self._latest_version

    @property
    def download_url(self):
        return self._download_url

    @property
    def release_notes(self):
        return self._release_notes

    @property
    def update_ready(self):
        return self._update_ready

    @property
    def download_progress(self):
        return self._download_progress

    # ── Thread run loop ──

    def run(self):
        """Main loop — check for updates periodically."""
        logger.info("UpdateChecker started (interval=%ds)", self.check_interval)
        while not self._stop_event.is_set():
            try:
                self.check_for_updates()
            except Exception as e:
                logger.error("Update check failed: %s", e)
            self._stop_event.wait(self.check_interval)

    # ── Update checking ──

    def check_for_updates(self):
        """Query Print Hub for the latest agent version."""
        url = f"{self.hub_url}/api/v1/agents/version"
        try:
            resp = requests.get(url, headers={
                'X-API-Key': self.agent_key,
                'Accept': 'application/json'
            }, timeout=10)
            if resp.ok:
                data = resp.json().get('data', {})
                latest = data.get('latest_version', '')
                if latest and self._is_newer(latest):
                    self._latest_version = latest
                    self._download_url = data.get('download_url', '')
                    self._release_notes = data.get('release_notes', '')
                    self._sha256 = data.get('sha256', '')
                    self._mandatory = data.get('mandatory', False)
                    self._update_ready = True
                    logger.info("Update available: v%s (mandatory=%s)",
                                latest, self._mandatory)
                    if self.on_update_available:
                        self.on_update_available(latest, self._release_notes)
                else:
                    logger.debug("No newer version found (latest=%s)", latest)
            else:
                logger.debug("Version check returned HTTP %d", resp.status_code)
        except requests.RequestException:
            logger.debug("Version endpoint unreachable — hub may be offline")
        except Exception as e:
            logger.error("Unexpected error during update check: %s", e)

    def _is_newer(self, latest_version):
        """Compare version strings using packaging or fallback."""
        current = self._get_current_version()
        try:
            from packaging.version import parse
            return parse(latest_version) > parse(current)
        except ImportError:
            # Fallback: simple string comparison (only works for semver-like)
            logger.debug("packaging not available — using string comparison")
            return latest_version > current

    def _get_current_version(self):
        """Read the current version from config.json."""
        config_path = Path(__file__).parent / 'config.json'
        if config_path.exists():
            try:
                with open(config_path) as f:
                    config = json.load(f)
                    return config.get('version', '1.0.0')
            except (json.JSONDecodeError, OSError):
                pass
        return '1.0.0'

    # ── Download ──

    def download_update(self, progress_callback=None):
        """
        Download the new version installer to a temp directory.

        Args:
            progress_callback: callable(progress_float) called during download.

        Returns:
            Path to the downloaded file, or None on failure.
        """
        if not self._download_url:
            logger.error("No download URL available")
            return None

        try:
            # Determine file extension from URL
            url = self._download_url
            ext = os.path.splitext(url.split('?')[0])[1] or '.exe'
            tmp_dir = tempfile.mkdtemp(prefix='trayprint_update_')
            tmp_path = os.path.join(tmp_dir, f'trayprint_update{ext}')

            logger.info("Downloading update from %s → %s", url, tmp_path)

            resp = requests.get(url, stream=True, timeout=120)
            resp.raise_for_status()

            total = int(resp.headers.get('content-length', 0))
            downloaded = 0
            self._download_progress = 0.0

            with open(tmp_path, 'wb') as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    if self._stop_event.is_set():
                        logger.info("Download cancelled by stop event")
                        self._cleanup_temp(tmp_dir)
                        return None
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total > 0:
                        self._download_progress = downloaded / total
                    else:
                        self._download_progress = 0.0
                    if progress_callback:
                        progress_callback(self._download_progress)

            self._download_progress = 1.0
            if progress_callback:
                progress_callback(1.0)

            # ── Hash verification ──
            if self._sha256:
                logger.info("Verifying SHA-256 checksum...")
                file_hash = hashlib.sha256()
                with open(tmp_path, 'rb') as f:
                    for chunk in iter(lambda: f.read(65536), b''):
                        file_hash.update(chunk)
                if file_hash.hexdigest().lower() != self._sha256.lower():
                    logger.error("SHA-256 mismatch! Expected=%s, Got=%s",
                                 self._sha256, file_hash.hexdigest())
                    self._cleanup_temp(tmp_dir)
                    return None
                logger.info("SHA-256 checksum verified successfully")

            logger.info("Download complete: %s (%.1f MB)",
                        tmp_path, os.path.getsize(tmp_path) / (1024 * 1024))
            return tmp_path

        except requests.RequestException as e:
            logger.error("Download failed: %s", e)
            self._cleanup_temp(tmp_dir if 'tmp_dir' in locals() else None)
        except Exception as e:
            logger.error("Download error: %s", e)
            self._cleanup_temp(tmp_dir if 'tmp_dir' in locals() else None)
        return None

    def _cleanup_temp(self, tmp_dir):
        """Remove a temporary directory if it exists."""
        if tmp_dir and os.path.exists(tmp_dir):
            import shutil
            try:
                shutil.rmtree(tmp_dir)
            except Exception:
                pass

    # ── Install ──

    def install_update(self, update_path):
        """
        Apply the downloaded update.

        For packaged (PyInstaller) builds: the updater exits the current process
        and launches the new executable. For dev mode: just logs the availability.
        """
        if not update_path or not os.path.exists(update_path):
            logger.error("Update path does not exist: %s", update_path)
            return False

        is_frozen = getattr(sys, 'frozen', False)

        if is_frozen:
            # ── Packaged mode: replace executable ──
            try:
                current_exe = sys.executable
                backup_path = current_exe + '.bak'

                # Rename current exe to .bak
                if os.path.exists(backup_path):
                    os.remove(backup_path)
                os.rename(current_exe, backup_path)
                logger.info("Backed up %s → %s", current_exe, backup_path)

                # Move new exe in place
                import shutil
                shutil.move(update_path, current_exe)
                logger.info("Replaced %s with update", current_exe)

                # Make executable on Unix/macOS
                if sys.platform != 'win32':
                    os.chmod(current_exe, 0o755)

                # Restart the application
                logger.info("Restarting application with updated executable...")
                os.execv(current_exe, sys.argv)
                return True

            except Exception as e:
                logger.error("Failed to install update: %s", e)
                # Attempt to restore backup
                try:
                    if os.path.exists(backup_path):
                        import shutil
                        shutil.move(backup_path, current_exe)
                except Exception:
                    pass
                return False
        else:
            # ── Dev mode: just log ──
            version = self._latest_version or 'unknown'
            logger.info(
                "Update v%s available at %s — in dev mode, install manually",
                version, update_path
            )
            print(f"\n=== Update Available: v{version} ===")
            print(f"Downloaded to: {update_path}")
            print(f"Release notes: {self._release_notes or 'N/A'}")
            print("TrayPrint is running in development mode.")
            print("Please install the update manually or rebuild with build.py.\n")
            return True
