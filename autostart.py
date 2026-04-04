import sys
import os
from logger import get_logger

log = get_logger()

def is_windows():
    return sys.platform == 'win32'

def is_macos():
    return sys.platform == 'darwin'

def enable_autostart():
    """Register the application to start automatically on login."""
    app_path = os.path.abspath(sys.argv[0])

    if is_windows():
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_SET_VALUE
            )
            winreg.SetValueEx(key, "Trayprint", 0, winreg.REG_SZ, f'"{sys.executable}" "{app_path}"')
            winreg.CloseKey(key)
            log.info("Autostart enabled (Windows registry)")
            return True
        except Exception as e:
            log.error("Failed to enable autostart on Windows: %s", e)
            return False

    elif is_macos():
        plist_dir = os.path.expanduser("~/Library/LaunchAgents")
        plist_path = os.path.join(plist_dir, "com.trayprint.agent.plist")
        plist_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.trayprint.agent</string>
    <key>ProgramArguments</key>
    <array>
        <string>{sys.executable}</string>
        <string>{app_path}</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
</dict>
</plist>"""
        try:
            os.makedirs(plist_dir, exist_ok=True)
            with open(plist_path, 'w') as f:
                f.write(plist_content)
            log.info("Autostart enabled (macOS LaunchAgent)")
            return True
        except Exception as e:
            log.error("Failed to enable autostart on macOS: %s", e)
            return False

    else:
        # Linux — XDG autostart
        autostart_dir = os.path.expanduser("~/.config/autostart")
        desktop_path = os.path.join(autostart_dir, "trayprint.desktop")
        desktop_content = f"""[Desktop Entry]
Type=Application
Name=Trayprint
Comment=Local Print Service
Exec={sys.executable} {app_path}
Icon=printer
Terminal=false
X-GNOME-Autostart-enabled=true
"""
        try:
            os.makedirs(autostart_dir, exist_ok=True)
            with open(desktop_path, 'w') as f:
                f.write(desktop_content)
            log.info("Autostart enabled (Linux XDG)")
            return True
        except Exception as e:
            log.error("Failed to enable autostart on Linux: %s", e)
            return False


def disable_autostart():
    """Remove the application from auto-start."""
    if is_windows():
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_SET_VALUE
            )
            winreg.DeleteValue(key, "Trayprint")
            winreg.CloseKey(key)
            log.info("Autostart disabled (Windows)")
            return True
        except Exception as e:
            log.error("Failed to disable autostart on Windows: %s", e)
            return False

    elif is_macos():
        plist_path = os.path.expanduser("~/Library/LaunchAgents/com.trayprint.agent.plist")
        try:
            if os.path.exists(plist_path):
                os.unlink(plist_path)
            log.info("Autostart disabled (macOS)")
            return True
        except Exception as e:
            log.error("Failed to disable autostart on macOS: %s", e)
            return False

    else:
        desktop_path = os.path.expanduser("~/.config/autostart/trayprint.desktop")
        try:
            if os.path.exists(desktop_path):
                os.unlink(desktop_path)
            log.info("Autostart disabled (Linux)")
            return True
        except Exception as e:
            log.error("Failed to disable autostart on Linux: %s", e)
            return False


def is_autostart_enabled():
    """Check if autostart is currently configured."""
    if is_windows():
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_READ
            )
            winreg.QueryValueEx(key, "Trayprint")
            winreg.CloseKey(key)
            return True
        except:
            return False
    elif is_macos():
        return os.path.exists(os.path.expanduser("~/Library/LaunchAgents/com.trayprint.agent.plist"))
    else:
        return os.path.exists(os.path.expanduser("~/.config/autostart/trayprint.desktop"))
