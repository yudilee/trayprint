"""
Tests for utility modules — autostart.py, path_utils.py, logger.py.
"""

import os
import sys

import pytest


# ===================================================================
# path_utils
# ===================================================================

class TestPathUtils:
    """Test path_utils.get_root_dir() behaviour."""

    def test_get_root_dir_not_frozen(self):
        """When not frozen, get_root_dir returns the script directory."""
        from path_utils import get_root_dir
        root = get_root_dir()
        # Should end with the project root (contains app.py)
        assert os.path.isdir(root)
        # The root should contain app.py
        assert "app.py" in os.listdir(root) or "tests" in os.listdir(root)

    def test_get_root_dir_frozen(self, mocker):
        """When frozen (PyInstaller), get_root_dir returns sys.executable directory."""
        mocker.patch("sys.frozen", True, create=True)
        mocker.patch("sys.executable", "/fake/path/trayprint.exe")
        from path_utils import get_root_dir
        assert get_root_dir() == "/fake/path"


# ===================================================================
# autostart
# ===================================================================

class TestAutostart:
    """Test autostart.enable_autostart and disable_autostart on each platform."""

    def test_enable_windows(self, mocker):
        """On Windows, enable_autostart writes to HKCU\...\Run."""
        mocker.patch("autostart.is_windows", return_value=True)
        mock_key = mocker.MagicMock()
        mock_OpenKey = mocker.patch("winreg.OpenKey", return_value=mock_key)
        mock_SetValueEx = mocker.patch("winreg.SetValueEx")

        from autostart import enable_autostart
        result = enable_autostart()

        assert result is True
        mock_OpenKey.assert_called_once()
        mock_SetValueEx.assert_called_once()

    def test_enable_windows_failure(self, mocker):
        """On Windows, if registry write fails, enable_autostart returns False."""
        mocker.patch("autostart.is_windows", return_value=True)
        mocker.patch("winreg.OpenKey", side_effect=PermissionError("Access denied"))

        from autostart import enable_autostart
        result = enable_autostart()
        assert result is False

    def test_enable_linux(self, mocker):
        """On Linux, enable_autostart creates an XDG .desktop file."""
        mocker.patch("autostart.is_windows", return_value=False)
        mocker.patch("autostart.is_macos", return_value=False)
        makedirs = mocker.patch("os.makedirs")
        open_func = mocker.patch("builtins.open", mocker.mock_open())

        # Ensure os.path.expanduser returns a temp path
        mocker.patch("os.path.expanduser", return_value="/tmp/.config/autostart")

        from autostart import enable_autostart
        result = enable_autostart()

        assert result is True
        makedirs.assert_called_once()
        open_func.assert_called_once()

    def test_enable_linux_failure(self, mocker):
        """On Linux, if file creation fails, enable_autostart returns False."""
        mocker.patch("autostart.is_windows", return_value=False)
        mocker.patch("autostart.is_macos", return_value=False)
        mocker.patch("os.makedirs", side_effect=OSError("Read-only fs"))

        from autostart import enable_autostart
        result = enable_autostart()
        assert result is False

    def test_disable_windows(self, mocker):
        """On Windows, disable_autostart removes the registry value."""
        mocker.patch("autostart.is_windows", return_value=True)
        mock_key = mocker.MagicMock()
        mocker.patch("winreg.OpenKey", return_value=mock_key)
        mock_DeleteValue = mocker.patch("winreg.DeleteValue")

        from autostart import disable_autostart
        result = disable_autostart()

        assert result is True
        mock_DeleteValue.assert_called_once()

    def test_disable_windows_failure(self, mocker):
        """On Windows, if DeleteValue fails, disable_autostart returns False."""
        mocker.patch("autostart.is_windows", return_value=True)
        mocker.patch("winreg.OpenKey", side_effect=PermissionError("Access denied"))

        from autostart import disable_autostart
        result = disable_autostart()
        assert result is False

    def test_disable_linux(self, mocker):
        """On Linux, disable_autostart removes the .desktop file."""
        mocker.patch("autostart.is_windows", return_value=False)
        mocker.patch("autostart.is_macos", return_value=False)
        mocker.patch("os.path.expanduser", return_value="/tmp/.config/autostart")
        mock_remove = mocker.patch("os.remove")

        from autostart import disable_autostart
        result = disable_autostart()
        assert result is True
        mock_remove.assert_called_once()

    def test_disable_linux_no_file(self, mocker):
        """If the .desktop file doesn't exist, disable_autostart still returns True."""
        mocker.patch("autostart.is_windows", return_value=False)
        mocker.patch("autostart.is_macos", return_value=False)
        mocker.patch("os.path.expanduser", return_value="/tmp/.config/autostart")
        mocker.patch("os.remove", side_effect=FileNotFoundError)

        from autostart import disable_autostart
        result = disable_autostart()
        assert result is True  # Should be idempotent — no file == already disabled

    def test_is_windows(self):
        """is_windows() returns True when sys.platform == 'win32'."""
        import autostart
        autostart.sys.platform = "win32"
        assert autostart.is_windows() is True

    def test_is_windows_false(self):
        """is_windows() returns False on other platforms."""
        import autostart
        for plat in ("linux", "darwin", "cygwin"):
            autostart.sys.platform = plat
            assert autostart.is_windows() is False

    def test_is_macos(self):
        """is_macos() returns True when sys.platform == 'darwin'."""
        import autostart
        autostart.sys.platform = "darwin"
        assert autostart.is_macos() is True

    def test_is_macos_false(self):
        """is_macos() returns False on other platforms."""
        import autostart
        autostart.sys.platform = "win32"
        assert autostart.is_macos() is False
