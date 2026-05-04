# TrayPrint — MSI Installer Builder

This directory contains WiX Toolset v3.x source files to build a Windows MSI
installer for TrayPrint.

## Prerequisites

1. **WiX Toolset v3.14+**  
   Download from: https://wixtoolset.org  
   Ensure `candle.exe` and `light.exe` are in your `PATH`, or set the
   `WIX_TOOLSET_PATH` environment variable to the WiX bin directory.

2. **PyInstaller Build**  
   Run `python build.py` from the project root **before** building the MSI.
   The installer expects `dist\trayprint.exe` to exist.

3. **Icon File**  
   Place a `trayprint.ico` file in the project root for the installer icon and
   shortcuts. This is optional but recommended.

## Build

### Using `build_msi.bat` (Windows)

Open a **Command Prompt** (cmd) and run:

```cmd
cd installer
build_msi.bat
```

The script will:

1. Locate WiX candle.exe and light.exe (via `PATH` or `WIX_TOOLSET_PATH`).
2. Verify that `dist\trayprint.exe` exists.
3. Compile `product.wxs` → `product.wixobj`.
4. Link `product.wixobj` → `TrayPrint.msi`.

The resulting MSI is written to the `installer\` directory.

### Manual Build

If the batch script does not suit your environment:

```cmd
candle.exe product.wxs
light.exe -out TrayPrint.msi product.wixobj
```

## What the Installer Does

| Component          | Description                                                    |
|--------------------|----------------------------------------------------------------|
| `trayprint.exe`    | Installed to `%ProgramFiles%\TrayPrint\`                       |
| `config.json`      | Installed alongside the executable                             |
| Start Menu Shortcut | Added under "TrayPrint" in the Start Menu                      |
| Desktop Shortcut   | Created for quick launch                                        |
| Auto-start         | Registers `HKCU\...\Run` so TrayPrint starts on Windows login  |
| Major Upgrade      | Installs over previous versions automatically (uninstalls old) |

## Uninstall

The MSI registers itself with Windows Installer. Uninstall through:

- **Settings → Apps → Apps & features → TrayPrint**
- Or run the MSI again and choose **Remove**.

## Customization

Edit `product.wxs` to change:

- **Version**: Update the `Version` attribute in `<Product>`.
- **UpgradeCode**: Change the `UpgradeCode` GUID for a different product family.
- **Start Menu name**: Modify the `Name` attribute in the Start Menu `<Directory>`.
- **Install path**: Change `APPLICATIONFOLDER` under `ProgramFiles64Folder`.
