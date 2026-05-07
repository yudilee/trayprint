# TrayPrint — MSI Installer Builder

This directory contains everything needed to build a Windows MSI installer for
TrayPrint. The installer packages the TrayPrint tray application into a
standard Windows Installer package with Start Menu shortcuts, auto-start
registration, desktop shortcut, and uninstaller support.

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
