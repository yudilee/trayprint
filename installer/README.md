# TrayPrint — Installer Builders

This directory contains everything needed to build installers for all supported
platforms:

| Platform | Format | Builder |
|----------|--------|---------|
| Windows  | MSI    | [`build_msi.py`](build_msi.py) / WiX Toolset |
| Linux    | AppImage | [`build_appimage.py`](build_appimage.py) |
| Linux    | .deb   | [`build_deb.sh`](build_deb.sh) |
| Linux    | .rpm   | [`build_rpm.sh`](build_rpm.sh) |
| macOS    | DMG    | [`build_dmg.py`](build_dmg.py) |

---

# Linux Installer

TrayPrint offers three Linux packaging formats, all built from a single source
codebase:

| Format | Distro | Dependencies | Package Size |
|--------|--------|-------------|--------------|
| **AppImage** | Any glibc Linux (preferred) | CUPS only (bundles Python + PySide6) | ~80-120 MB |
| **.deb** | Debian/Ubuntu/Mint | CUPS + system Python packages | ~500 KB (thin) |
| **.rpm** | Fedora/RHEL/CentOS | CUPS + system Python packages | ~500 KB (thin) |

> **Recommendation:** Use the AppImage for maximum portability. Use .deb/.rpm
> for tighter system integration (automatic dependency resolution, systemd
> service management).

---

## Quick Start (End Users)

### AppImage — Portable (Any Linux)

```bash
# Download and run (no installation needed)
chmod +x TrayPrint-x86_64.AppImage
./TrayPrint-x86_64.AppImage

# Or install system-wide
./installer/install.sh --appimage
```

### .deb — Debian/Ubuntu

```bash
# Build and install
./installer/build_deb.sh
sudo apt install ./dist/trayprint_*.deb
```

### .rpm — Fedora/RHEL

```bash
# Build and install
./installer/build_rpm.sh
sudo dnf install ./dist/trayprint-*.rpm
```

### Universal Installer (Auto-Detect)

```bash
# Detects your distro and installs the best format
./installer/install.sh
```

---

## Building

### Prerequisites

**All formats:**
- Python 3.10+
- PyInstaller (`pip install pyinstaller`)
- Pillow (`pip install Pillow`)

**AppImage:**
- `appimagetool` (downloaded automatically by the build script)

**Debian package:**
- `dpkg-deb` (from `dpkg-dev`)
- `fakeroot`

**RPM package:**
- `rpmbuild` (from `rpm-build`)

### Build Commands

```bash
# Build AppImage (portable, self-contained)
python installer/build_appimage.py

# Build .deb package (Debian/Ubuntu)
./installer/build_deb.sh

# Build .rpm package (Fedora/RHEL)
./installer/build_rpm.sh

# Build all formats at once
python installer/build_appimage.py && \
./installer/build_deb.sh && \
./installer/build_rpm.sh
```

### Build Options

All builders accept `--version X.Y.Z` and `--clean` flags. Example:

```bash
# Build AppImage with custom version
python installer/build_appimage.py --version 3.1.0 --clean

# Build .deb with custom version
./installer/build_deb.sh --version 3.1.0 --clean
```

---

## Installation

### Universal Installer (`installer/install.sh`)

The [`install.sh`](install.sh) script provides an interactive installation
experience that auto-detects the best format for your system:

```bash
# Interactive (auto-detect)
./installer/install.sh

# Force AppImage
./installer/install.sh --appimage

# Force .deb
./installer/install.sh --deb

# Force .rpm
./installer/install.sh --rpm

# User-only install (no sudo)
./installer/install.sh --appimage --user

# Non-interactive (automated scripts)
./installer/install.sh --appimage --no-prompt
```

### Manual Installation

#### AppImage

```bash
# Option 1: Run in-place (no installation)
chmod +x dist/TrayPrint-x86_64.AppImage
./dist/TrayPrint-x86_64.AppImage

# Option 2: System-wide install
sudo mkdir -p /opt/trayprint
sudo cp dist/TrayPrint-x86_64.AppImage /opt/trayprint/
sudo chmod +x /opt/trayprint/TrayPrint-x86_64.AppImage
sudo ln -s /opt/trayprint/TrayPrint-x86_64.AppImage /usr/local/bin/trayprint

# Desktop integration
sudo cp installer/trayprint.desktop /usr/share/applications/
sudo cp installer/trayprint.png /usr/share/icons/hicolor/256x256/apps/
sudo update-desktop-database

# systemd user service
sudo cp installer/trayprint.service /usr/lib/systemd/user/
systemctl --user enable trayprint.service
```

#### Debian Package

```bash
sudo dpkg -i dist/trayprint_3.0.0-1_amd64.deb
sudo apt-get install -f   # Install any missing dependencies
```

#### RPM Package

```bash
sudo dnf install dist/trayprint-3.0.0-1.x86_64.rpm
```

---

## What Each Installer Does

| Component | AppImage | .deb | .rpm |
|-----------|----------|------|------|
| Install Path | `/opt/trayprint/` or `~/.local/bin/` | `/usr/lib/trayprint/` | `/usr/lib/trayprint/` |
| Binary | `TrayPrint.AppImage` | `/usr/bin/trayprint` (launcher) | `/usr/bin/trayprint` (launcher) |
| GUI Launcher | App menu entry | App menu entry | App menu entry |
| systemd Service | `/usr/lib/systemd/user/` | `/usr/lib/systemd/user/` | `/usr/lib/systemd/user/` |
| Config | `/etc/trayprint/config.json` | `/etc/trayprint/config.json` | `/etc/trayprint/config.json` |
| Auto-start | XDG autostart (via app menu) | XDG autostart (via app menu) | XDG autostart (via app menu) |
| Dependency Resolution | CUPS only (manual) | Automatic via apt | Automatic via dnf |

---

## Running

### GUI Mode (Default)

```bash
# From terminal
trayprint

# From application menu
# → Look for "TrayPrint" in your app launcher
```

### Headless / Service Mode

```bash
# Run without GUI (for servers)
trayprint --silent

# Or via systemd (user service — starts on login)
systemctl --user start trayprint.service
systemctl --user enable trayprint.service
systemctl --user status trayprint.service
```

### Configuration

```bash
# Default config
/etc/trayprint/config.json

# Override via environment variable
TRAYPRINT_CONFIG_PATH=/path/to/custom/config.json trayprint

# Or via command line
trayprint --hub-url https://hub.example.com --agent-key your-key
```

### Viewing Logs

```bash
# Default log location
~/.local/share/trayprint/logs/trayprint.log

# Or use the tray menu → "View Logs"
```

---

## Uninstall

### AppImage

```bash
# If installed via install.sh --user
rm ~/.local/bin/TrayPrint.AppImage
rm ~/.local/share/applications/trayprint.desktop
rm ~/.local/share/icons/hicolor/256x256/apps/trayprint.png

# If installed system-wide
sudo rm /opt/trayprint/TrayPrint.AppImage
sudo rm /usr/local/bin/trayprint
sudo rm /usr/share/applications/trayprint.desktop
sudo rm /usr/share/icons/hicolor/256x256/apps/trayprint.png
sudo rm /usr/lib/systemd/user/trayprint.service
```

### Debian Package

```bash
sudo apt remove trayprint
# Or
sudo dpkg -r trayprint
# Purge config too:
sudo apt purge trayprint
```

### RPM Package

```bash
sudo dnf remove trayprint
# Or
sudo rpm -e trayprint
```

---

## System Requirements

- **Linux x86_64** with glibc (any modern distribution)
- **CUPS** installed and running (`systemctl status cups`)
- **D-Bus** session bus (for system tray — all modern DEs have this)
- **X11 or Wayland** compositor (for GUI mode — all modern DEs)

---

## Files Reference

| File | Purpose |
|------|---------|
| [`build_appimage.py`](build_appimage.py) | AppImage build script |
| [`AppRun`](AppRun) | AppImage entry-point |
| [`trayprint.desktop`](trayprint.desktop) | Desktop entry (all formats) |
| [`trayprint.png`](trayprint.png) | 256×256 application icon |
| [`trayprint.appdata.xml`](trayprint.appdata.xml) | AppStream metadata |
| [`trayprint.service`](trayprint.service) | systemd user service unit |
| [`build_deb.sh`](build_deb.sh) | Debian package build script |
| [`debian/`](debian/) | Debian package control files |
| [`build_rpm.sh`](build_rpm.sh) | RPM package build script |
| [`trayprint.spec`](trayprint.spec) | RPM spec file |
| [`install.sh`](install.sh) | Universal Linux installer |

---

# Windows Installer

This section covers the Windows MSI installer, which was the original content
of this README.

## Prerequisites

## Prerequisites

### Option A: cx_Freeze (Recommended)

1. Install cx_Freeze:
   ```cmd
   pip install cx_Freeze
   ```

2. Run the build script:
   ```cmd
   python installer\build_msi.py
   ```

### Option B: PyInstaller + WiX Toolset

1. **PyInstaller** — Install via pip:
   ```cmd
   pip install pyinstaller
   ```

2. **WiX Toolset v3.14+** — Download from https://wixtoolset.org
   - Ensure `candle.exe` and `light.exe` are in your `PATH`, or set the
     `WIX` environment variable to the WiX bin directory.

3. **Icon File** — Place a `trayprint.ico` file in the project root.

4. **Configuration** — Ensure `config.json` exists in the project root.

5. Run the build script:
   ```cmd
   python installer\build_msi.py --builder pyinstaller
   ```

## Build Script Usage

```
python installer\build_msi.py [options]

Options:
  --builder {cx_freeze,pyinstaller,auto}
                        Build backend (default: auto-detect)
  --version VERSION     Installer version (default: 3.0.0)
  --output-dir DIR      Output directory (default: installer/dist/)
  --clean               Clean build artifacts before building
```

### Examples

```cmd
:: Auto-detect and build
python installer\build_msi.py

:: Build with specific version
python installer\build_msi.py --version 3.1.0

:: Force PyInstaller + WiX
python installer\build_msi.py --builder pyinstaller

:: Clean and rebuild
python installer\build_msi.py --clean
```

## What the Installer Does

| Component            | Description                                                       |
|----------------------|-------------------------------------------------------------------|
| Install Directory    | `%ProgramFiles%\PrintHub\TrayPrint\`                              |
| `trayprint.exe`      | Main application executable                                       |
| `config.json`        | Configuration file (installed alongside executable)               |
| Start Menu Shortcut  | Added under "PrintHub" in the Start Menu                          |
| Desktop Shortcut     | Created for quick launch (optional)                               |
| Auto-start           | Registers `HKCU\...\Run` so TrayPrint starts on Windows login     |
| Uninstaller          | Registered with Windows Installer — uninstall via Settings        |
| Major Upgrade        | Automatically upgrades previous versions                          |

## Uninstall

The MSI registers itself with Windows Installer. Uninstall through:

- **Settings → Apps → Apps & features → TrayPrint**
- Or run the MSI again and choose **Remove**.

## Manual WiX Build

If you prefer to build manually with WiX:

```cmd
cd installer
candle.exe -dVersion=3.0.0 trayprint.wxs
light.exe -out TrayPrint-3.0.0.msi trayprint.wixobj
```

## Customization

Edit `trayprint.wxs` to change:

- **Version**: Pass via `-dVersion=X.Y.Z` to candle.exe, or use `--version`
- **UpgradeCode**: Change the GUID for a different product family
- **Install path**: Modify the `PrintHub\TrayPrint` directory structure
- **Auto-start flags**: Change the `--silent` flag in the auto-start registry value
