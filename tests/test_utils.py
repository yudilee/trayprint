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
    """Test path_utils.get_root_dir() and get_data_dir() behaviour."""

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

    def test_get_data_dir_not_frozen(self):
        """When not frozen, get_data_dir returns the project root."""
        from path_utils import get_data_dir
        data_dir = get_data_dir()
        assert os.path.isdir(data_dir)
        # In dev mode it returns the script directory
        assert "app.py" in os.listdir(data_dir) or "tests" in os.listdir(data_dir)

    def test_get_data_dir_frozen_windows(self, mocker):
        """When frozen on Windows, get_data_dir uses %LOCALAPPDATA%."""
        mocker.patch("sys.frozen", True, create=True)
        mocker.patch("platform.system", return_value="Windows")
        mocker.patch.dict(os.environ, {"LOCALAPPDATA": "C:\\Users\\Test\\AppData\\Local"})
        # Prevent actual directory creation
        mocker.patch("os.makedirs")

        from path_utils import get_data_dir
        result = get_data_dir()
        assert result == "C:\\Users\\Test\\AppData\\Local\\TrayPrint"

    def test_get_data_dir_frozen_windows_fallback(self, mocker):
        """When LOCALAPPDATA is missing, falls back to user home."""
        mocker.patch("sys.frozen", True, create=True)
        mocker.patch("platform.system", return_value="Windows")
        mocker.patch.dict(os.environ, {}, clear=True)
        mocker.patch("os.path.expanduser", return_value="C:\\Users\\Default")
        mocker.patch("os.makedirs")

        from path_utils import get_data_dir
        result = get_data_dir()
        assert result == "C:\\Users\\Default\\TrayPrint"

    def test_get_data_dir_frozen_macos(self, mocker):
        """When frozen on macOS, get_data_dir uses ~/Library/Application Support."""
        import os
        mocker.patch("sys.frozen", True, create=True)
        mocker.patch("platform.system", return_value="Darwin")
        mocker.patch("os.path.expanduser", return_value="/Users/testuser")
        mocker.patch("os.makedirs")

        from path_utils import get_data_dir
        result = get_data_dir()
        expected = os.path.join("/Users/testuser", "Library", "Application Support", "TrayPrint")
        assert result == expected

    def test_get_data_dir_frozen_linux(self, mocker):
        """When frozen on Linux, get_data_dir uses XDG_DATA_HOME."""
        import os
        mocker.patch("sys.frozen", True, create=True)
        mocker.patch("platform.system", return_value="Linux")
        mocker.patch.dict(os.environ, {"XDG_DATA_HOME": "/home/testuser/.local/share"})
        mocker.patch("os.makedirs")

        from path_utils import get_data_dir
        result = get_data_dir()
        expected = os.path.join("/home/testuser/.local/share", "TrayPrint")
        assert result == expected

    def test_get_data_dir_frozen_linux_xdg_fallback(self, mocker):
        """When frozen on Linux without XDG_DATA_HOME, falls back to ~/.local/share."""
        import os
        mocker.patch("sys.frozen", True, create=True)
        mocker.patch("platform.system", return_value="Linux")
        mocker.patch.dict(os.environ, {}, clear=True)
        mocker.patch("os.path.expanduser", return_value="/home/testuser")
        mocker.patch("os.makedirs")

        from path_utils import get_data_dir
        result = get_data_dir()
        expected = os.path.join("/home/testuser", ".local", "share", "TrayPrint")
        assert result == expected

    def test_get_data_dir_creates_directory(self, mocker):
        """get_data_dir creates the directory if it doesn't exist."""
        mocker.patch("sys.frozen", False, create=True)
        mock_makedirs = mocker.patch("os.makedirs")

        from path_utils import get_data_dir
        get_data_dir()
        mock_makedirs.assert_called_once()


# ===================================================================
# autostart
# ===================================================================

class TestAutostart:
    """Test autostart.enable_autostart and disable_autostart on each platform."""

    def test_enable_windows(self, mocker, temp_dir):
        """On Windows, enable_autostart writes to HKCU\...\Run."""
        # Mock logger's get_data_dir to prevent file handler creation in real dirs
        mocker.patch("logger.get_data_dir", return_value=temp_dir)
        from logger import get_logger
        _clean_logger = get_logger()
        import winreg
        mocker.patch("autostart.is_windows", return_value=True)
        mock_key = mocker.MagicMock(spec=winreg.HKEY_CURRENT_USER.__class__)
        mock_OpenKey = mocker.patch("winreg.OpenKey", return_value=mock_key)
        mock_SetValueEx = mocker.patch("winreg.SetValueEx")
        mocker.patch("winreg.CloseKey")

        from autostart import enable_autostart
        result = enable_autostart()

        assert result is True
        mock_OpenKey.assert_called_once()
        mock_SetValueEx.assert_called_once()
        # Clean up logger handlers so temp_dir can be removed
        for h in _clean_logger.handlers[:]:
            h.close()
            _clean_logger.removeHandler(h)

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
        # Use a real winreg.HKEY object type for the mock return
        import winreg
        mock_key = mocker.MagicMock(spec=winreg.HKEY_CURRENT_USER.__class__)
        mocker.patch("winreg.OpenKey", return_value=mock_key)
        mock_DeleteValue = mocker.patch("winreg.DeleteValue")
        mocker.patch("winreg.CloseKey")

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
        # The code uses os.unlink (not os.remove) and checks os.path.exists first
        mocker.patch("os.path.exists", return_value=True)
        mock_unlink = mocker.patch("os.unlink")

        from autostart import disable_autostart
        result = disable_autostart()
        assert result is True
        mock_unlink.assert_called_once()

    def test_disable_linux_no_file(self, mocker):
        """If the .desktop file doesn't exist, disable_autostart still returns True."""
        mocker.patch("autostart.is_windows", return_value=False)
        mocker.patch("autostart.is_macos", return_value=False)
        mocker.patch("os.path.expanduser", return_value="/tmp/.config/autostart")
        mocker.patch("os.path.exists", return_value=False)
        mock_unlink = mocker.patch("os.unlink")

        from autostart import disable_autostart
        result = disable_autostart()
        assert result is True  # Should be idempotent — no file == already disabled
        mock_unlink.assert_not_called()

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
