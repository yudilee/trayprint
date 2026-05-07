#!/usr/bin/env python3
"""
TrayPrint — MSI Installer Build Script (Python)

This script builds a Windows MSI installer for TrayPrint using either:
  1. cx_Freeze (preferred) — generates an MSI via cx_Freeze's bdist_msi
  2. PyInstaller + WiX Toolset — builds an executable then packages with WiX

Usage:
    python build_msi.py                    # Auto-detect best available builder
    python build_msi.py --builder cx_freeze # Force cx_Freeze
    python build_msi.py --builder pyinstaller # Force PyInstaller + WiX
    python build_msi.py --version 3.0.0    # Set installer version
    python build_msi.py --output-dir ./dist # Custom output directory

Prerequisites:
    - For cx_Freeze: pip install cx_Freeze
    - For PyInstaller+WiX: pip install pyinstaller, WiX Toolset v3.14+
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import json

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
INSTALLER_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_VERSION = "3.0.0"
DEFAULT_OUTPUT_DIR = os.path.join(INSTALLER_DIR, "dist")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Build TrayPrint MSI installer"
    )
    parser.add_argument(
        "--builder", choices=["cx_freeze", "pyinstaller", "auto"],
        default="auto",
        help="Build backend to use (default: auto-detect)"
    )
    parser.add_argument(
        "--version", default=DEFAULT_VERSION,
        help=f"Installer version (default: {DEFAULT_VERSION})"
    )
    parser.add_argument(
        "--output-dir", default=DEFAULT_OUTPUT_DIR,
        help=f"Output directory (default: {DEFAULT_OUTPUT_DIR})"
    )
    parser.add_argument(
        "--clean", action="store_true",
        help="Clean build artifacts before building"
    )
    return parser.parse_args()


def check_cx_freeze():
    """Check if cx_Freeze is available."""
    try:
        import cx_Freeze  # noqa: F401
        return True
    except ImportError:
        return False


def check_pyinstaller():
    """Check if PyInstaller is available."""
    try:
        import PyInstaller  # noqa: F401
        return True
    except ImportError:
        return False


def check_wix():
    """Check if WiX Toolset (candle.exe / light.exe) is available."""
    if sys.platform != "win32":
        return False
    try:
        subprocess.run(
            ["candle.exe", "/?"],
            capture_output=True, timeout=10
        )
        return True
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    # Check common WiX paths
    wix_paths = [
        os.environ.get("WIX", ""),
        r"C:\Program Files (x86)\WiX Toolset v3.14\bin",
        r"C:\Program Files\WiX Toolset v3.14\bin",
    ]
    for p in wix_paths:
        if p and os.path.exists(os.path.join(p, "candle.exe")):
            return True
    return False


def build_with_cx_freeze(version, output_dir):
    """Build MSI using cx_Freeze's bdist_msi command."""
    print("=" * 70)
    print("  Building TrayPrint MSI with cx_Freeze")
    print("=" * 70)

    # Create a temporary setup.py for cx_Freeze
    setup_path = os.path.join(tempfile.mkdtemp(), "setup.py")
    with open(setup_path, "w") as f:
        f.write(f'''
import sys
from cx_Freeze import setup, Executable

# Dependencies
build_exe_options = {{
    "packages": [
        "os", "sys", "json", "threading", "time", "uuid",
        "subprocess", "webbrowser", "datetime", "platform",
        "tempfile", "collections", "re", "socket", "base64",
        "queue", "logging", "sqlite3",
    ],
    "excludes": ["tkinter", "test", "unittest"],
    "include_files": [
        ("{os.path.join(PROJECT_ROOT, 'templates').replace(chr(92), '/')}", "templates"),
        ("{os.path.join(PROJECT_ROOT, 'config.json').replace(chr(92), '/')}", "config.json"),
    ],
}}

# GUI executable (no console window)
gui_exe = Executable(
    script="{os.path.join(PROJECT_ROOT, 'app.py').replace(chr(92), '/')}",
    base="Win32GUI" if sys.platform == "win32" else None,
    target_name="trayprint.exe",
    icon="{os.path.join(PROJECT_ROOT, 'trayprint.ico').replace(chr(92), '/')}",
)

setup(
    name="TrayPrint",
    version="{version}",
    description="TrayPrint - Local Print Agent for Print Hub",
    author="Print Hub",
    options={{"build_exe": build_exe_options}},
    executables=[gui_exe],
)
''')

    # Run cx_Freeze bdist_msi
    result = subprocess.run(
        [sys.executable, setup_path, "bdist_msi"],
        cwd=PROJECT_ROOT,
        capture_output=False,
    )
    if result.returncode != 0:
        print("[ERROR] cx_Freeze build failed")
        sys.exit(1)

    # Find the generated MSI
    dist_dir = os.path.join(PROJECT_ROOT, "dist")
    if not os.path.exists(dist_dir):
        print("[ERROR] dist directory not found after build")
        sys.exit(1)

    msis = [f for f in os.listdir(dist_dir) if f.endswith(".msi")]
    if not msis:
        print("[ERROR] No MSI found in dist/ after build")
        sys.exit(1)

    # Copy to output directory
    os.makedirs(output_dir, exist_ok=True)
    src = os.path.join(dist_dir, msis[0])
    dst = os.path.join(output_dir, f"TrayPrint-{version}.msi")
    shutil.copy2(src, dst)
    print(f"\n[SUCCESS] MSI installer created: {dst}")
    return dst


def build_with_pyinstaller_wix(version, output_dir):
    """Build executable with PyInstaller, then package with WiX."""
    print("=" * 70)
    print("  Building TrayPrint MSI with PyInstaller + WiX")
    print("=" * 70)

    # Step 1: Build executable with PyInstaller
    print("\n[1/2] Building executable with PyInstaller...")
    spec_path = os.path.join(PROJECT_ROOT, "trayprint.spec")
    if not os.path.exists(spec_path):
        # Create spec file if it doesn't exist
        spec_path = os.path.join(PROJECT_ROOT, "build.spec")
        with open(spec_path, "w") as f:
            f.write(f'''
# -*- mode: python ; coding: utf-8 -*-
a = Analysis(
    ['{os.path.join(PROJECT_ROOT, "app.py").replace(chr(92), "/")}'],
    pathex=['{PROJECT_ROOT.replace(chr(92), "/")}'],
    binaries=[],
    datas=[
        ('{os.path.join(PROJECT_ROOT, "templates").replace(chr(92), "/")}', 'templates'),
        ('{os.path.join(PROJECT_ROOT, "config.json").replace(chr(92), "/")}', '.'),
    ],
    hiddenimports=[
        'sqlite3', 'queue', 'base64', 'uuid', 'ctypes',
        'logging.handlers', 'http.server', 'email',
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=['tkinter', 'test', 'unittest'],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='trayprint',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='{os.path.join(PROJECT_ROOT, "trayprint.ico").replace(chr(92), "/")}',
)
''')

    result = subprocess.run(
        [sys.executable, "-m", "PyInstaller", spec_path, "--clean", "--noconfirm"],
        cwd=PROJECT_ROOT,
        capture_output=False,
    )
    if result.returncode != 0:
        print("[ERROR] PyInstaller build failed")
        sys.exit(1)

    exe_path = os.path.join(PROJECT_ROOT, "dist", "trayprint.exe")
    if not os.path.exists(exe_path):
        print(f"[ERROR] Expected executable not found: {exe_path}")
        sys.exit(1)

    # Step 2: Build MSI with WiX
    print("\n[2/2] Building MSI with WiX Toolset...")
    wxs_path = os.path.join(INSTALLER_DIR, "trayprint.wxs")
    wixobj_path = os.path.join(INSTALLER_DIR, "trayprint.wixobj")
    msi_path = os.path.join(output_dir, f"TrayPrint-{version}.msi")

    # Locate WiX tools
    candle = "candle.exe"
    light = "light.exe"
    wix_path = os.environ.get("WIX", "")
    if wix_path:
        candle = os.path.join(wix_path, "candle.exe")
        light = os.path.join(wix_path, "light.exe")

    # Compile .wxs -> .wixobj
    result = subprocess.run(
        [candle, "-dVersion=%s" % version, wxs_path, "-out", wixobj_path],
        capture_output=False,
    )
    if result.returncode != 0:
        print("[ERROR] WiX compilation (candle.exe) failed")
        sys.exit(1)

    # Link .wixobj -> .msi
    os.makedirs(output_dir, exist_ok=True)
    result = subprocess.run(
        [light, "-out", msi_path, wixobj_path],
        capture_output=False,
    )
    if result.returncode != 0:
        print("[ERROR] WiX linking (light.exe) failed")
        sys.exit(1)

    # Cleanup intermediate files
    if os.path.exists(wixobj_path):
        os.remove(wixobj_path)

    print(f"\n[SUCCESS] MSI installer created: {msi_path}")
    return msi_path


def main():
    args = parse_args()

    if args.clean:
        # Clean build artifacts
        for d in ["build", "dist", "__pycache__"]:
            path = os.path.join(PROJECT_ROOT, d)
            if os.path.exists(path):
                shutil.rmtree(path)
                print(f"Cleaned: {path}")
        for f in ["trayprint.spec", "build.spec"]:
            path = os.path.join(PROJECT_ROOT, f)
            if os.path.exists(path):
                os.remove(path)
                print(f"Cleaned: {path}")

    os.makedirs(args.output_dir, exist_ok=True)

    # Detect available builders
    has_cx = check_cx_freeze()
    has_py = check_pyinstaller()
    has_wix = check_wix()

    print(f"Builder detection:")
    print(f"  cx_Freeze:     {'✓' if has_cx else '✗'}")
    print(f"  PyInstaller:   {'✓' if has_py else '✗'}")
    print(f"  WiX Toolset:   {'✓' if has_wix else '✗'}")

    builder = args.builder
    if builder == "auto":
        if has_cx:
            builder = "cx_freeze"
        elif has_py and has_wix:
            builder = "pyinstaller"
        else:
            print("[ERROR] No suitable build backend found.")
            print("  Install cx_Freeze: pip install cx_Freeze")
            print("  Or install PyInstaller + WiX Toolset v3.14+")
            sys.exit(1)

    print(f"\nUsing builder: {builder}")

    if builder == "cx_freeze":
        build_with_cx_freeze(args.version, args.output_dir)
    else:
        build_with_pyinstaller_wix(args.version, args.output_dir)


if __name__ == "__main__":
    main()
