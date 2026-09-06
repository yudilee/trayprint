#!/usr/bin/env python3
"""
TrayPrint — AppImage Builder

Builds a distro-agnostic AppImage from the PyInstaller-built TrayPrint binary.

Usage:
    python installer/build_appimage.py [--version X.Y.Z] [--clean]

Prerequisites:
    - PyInstaller build output at dist/trayprint (run build.py first)
    - appimagetool (downloaded automatically if not found)
    - installer/trayprint.desktop
    - installer/trayprint.png
    - installer/trayprint.appdata.xml
    - installer/AppRun

Output:
    dist/TrayPrint-x86_64.AppImage
"""

import os
import sys
import shutil
import stat
import subprocess
import argparse
import tempfile
import urllib.request
import json
import hashlib

# ── Constants ──

APP_NAME = "TrayPrint"
APP_VERSION = "3.0.0"
ARCH = "x86_64"
APPDIR_NAME = f"{APP_NAME}.AppDir"
OUTPUT_NAME = f"{APP_NAME}-{ARCH}.AppImage"

APPIMAGETOOL_RELEASE = "continuous"
APPIMAGETOOL_URL = (
    f"https://github.com/AppImage/AppImageKit/releases/download/"
    f"{APPIMAGETOOL_RELEASE}/appimagetool-{ARCH}.AppImage"
)
APPIMAGETOOL_MAGIC = b'\x41\x49\x02'  # AppImage magic bytes

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
INSTALLER_DIR = os.path.join(PROJECT_ROOT, "installer")
DIST_DIR = os.path.join(PROJECT_ROOT, "dist")


# ── Helper Functions ──


def log(msg, level="INFO"):
    print(f"[{level}] {msg}")


def _find_appimagetool():
    """Locate appimagetool in PATH or project directory."""
    # Check PATH first
    tool = shutil.which("appimagetool")
    if tool:
        return tool

    # Check project directory
    local = os.path.join(PROJECT_ROOT, "appimagetool")
    if os.path.isfile(local) and os.access(local, os.X_OK):
        return local

    local = os.path.join(PROJECT_ROOT, "dist", "appimagetool")
    if os.path.isfile(local) and os.access(local, os.X_OK):
        return local

    return None


def _download_appimagetool(dest_dir):
    """Download appimagetool from GitHub releases."""
    dest = os.path.join(dest_dir, "appimagetool")
    log(f"Downloading appimagetool from {APPIMAGETOOL_URL} ...")

    try:
        urllib.request.urlretrieve(APPIMAGETOOL_URL, dest)
        os.chmod(dest, os.stat(dest).st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)
        log(f"Downloaded appimagetool to {dest}")
        return dest
    except Exception as e:
        log(f"Failed to download appimagetool: {e}", "ERROR")
        return None


def _ensure_appimagetool():
    """Ensure appimagetool is available, downloading if necessary."""
    tool = _find_appimagetool()
    if tool:
        log(f"Found appimagetool at {tool}")
        return tool

    log("appimagetool not found. Downloading...")
    dest_dir = os.path.join(PROJECT_ROOT, "dist")
    os.makedirs(dest_dir, exist_ok=True)
    return _download_appimagetool(dest_dir)


def _check_pyinstaller_output():
    """Check that the PyInstaller binary exists; if not, run build.py."""
    exe_path = os.path.join(DIST_DIR, "trayprint")
    if os.path.isfile(exe_path) and os.access(exe_path, os.X_OK):
        log(f"Found PyInstaller output at {exe_path}")
        return exe_path

    log("PyInstaller output not found. Running build.py first...")
    build_script = os.path.join(PROJECT_ROOT, "build.py")
    if not os.path.isfile(build_script):
        log(f"build.py not found at {build_script}", "ERROR")
        return None

    try:
        subprocess.run([sys.executable, build_script], cwd=PROJECT_ROOT, check=True)
    except subprocess.CalledProcessError as e:
        log(f"build.py failed with exit code {e.returncode}", "ERROR")
        return None

    if os.path.isfile(exe_path) and os.access(exe_path, os.X_OK):
        log(f"PyInstaller build successful: {exe_path}")
        return exe_path

    log("PyInstaller output still not found after build.py", "ERROR")
    return None


def _check_assets():
    """Verify that all required assets exist."""
    required = [
        ("Desktop file", os.path.join(INSTALLER_DIR, "trayprint.desktop")),
        ("PNG icon", os.path.join(INSTALLER_DIR, "trayprint.png")),
        ("AppStream metadata", os.path.join(INSTALLER_DIR, "trayprint.appdata.xml")),
        ("AppRun script", os.path.join(INSTALLER_DIR, "AppRun")),
    ]

    missing = []
    for name, path in required:
        if not os.path.isfile(path):
            missing.append(f"  {name}: {path}")

    if missing:
        log("Missing required assets:", "ERROR")
        for m in missing:
            print(m)
        return False
    return True


def _create_appdir(appdir_root, pyinstaller_exe, version):
    """Construct the AppDir directory structure."""
    log(f"Creating AppDir at {appdir_root} ...")

    # Create directory structure
    usr_bin = os.path.join(appdir_root, "usr", "bin")
    usr_share_apps = os.path.join(appdir_root, "usr", "share", "applications")
    usr_share_icons = os.path.join(appdir_root, "usr", "share", "icons", "hicolor", "256x256", "apps")
    usr_share_meta = os.path.join(appdir_root, "usr", "share", "metainfo")

    for d in [usr_bin, usr_share_apps, usr_share_icons, usr_share_meta]:
        os.makedirs(d, exist_ok=True)

    # Copy PyInstaller binary
    dest_exe = os.path.join(usr_bin, "trayprint")
    shutil.copy2(pyinstaller_exe, dest_exe)
    os.chmod(dest_exe, os.stat(dest_exe).st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)
    log(f"Copied executable to {dest_exe}")

    # Copy desktop file
    shutil.copy2(
        os.path.join(INSTALLER_DIR, "trayprint.desktop"),
        os.path.join(usr_share_apps, "trayprint.desktop"),
    )
    # Also copy to AppDir root (required by AppImage spec)
    shutil.copy2(
        os.path.join(INSTALLER_DIR, "trayprint.desktop"),
        os.path.join(appdir_root, "trayprint.desktop"),
    )
    log("Copied desktop file")

    # Copy icon
    shutil.copy2(
        os.path.join(INSTALLER_DIR, "trayprint.png"),
        os.path.join(usr_share_icons, "trayprint.png"),
    )
    # Also copy to AppDir root (required by AppImage spec)
    shutil.copy2(
        os.path.join(INSTALLER_DIR, "trayprint.png"),
        os.path.join(appdir_root, "trayprint.png"),
    )
    log("Copied icon")

    # Copy AppStream metadata
    shutil.copy2(
        os.path.join(INSTALLER_DIR, "trayprint.appdata.xml"),
        os.path.join(usr_share_meta, "trayprint.appdata.xml"),
    )
    log("Copied AppStream metadata")

    # Copy AppRun entry point
    apprun_src = os.path.join(INSTALLER_DIR, "AppRun")
    apprun_dest = os.path.join(appdir_root, "AppRun")
    shutil.copy2(apprun_src, apprun_dest)
    os.chmod(apprun_dest, os.stat(apprun_dest).st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)
    log("Copied AppRun entry point")

    # Write .DirIcon (symlink or copy of trayprint.png)
    diricon = os.path.join(appdir_root, ".DirIcon")
    if not os.path.exists(diricon):
        shutil.copy2(os.path.join(INSTALLER_DIR, "trayprint.png"), diricon)

    log("AppDir structure created successfully")
    return True


def _run_appimagetool(appimagetool_path, appdir_root, output_path):
    """Run appimagetool to produce the final AppImage."""
    log("Running appimagetool to produce AppImage...")

    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    cmd = [
        appimagetool_path,
        "--verbose",
        "--no-appstream",  # Skip built-in AppStream check (we provide our own)
        appdir_root,
        output_path,
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if result.returncode != 0:
            log(f"appimagetool failed (exit code {result.returncode})", "ERROR")
            if result.stderr:
                print(result.stderr[-2000:])
            return False

        log("appimagetool completed successfully")
        return True

    except subprocess.TimeoutExpired:
        log("appimagetool timed out after 300 seconds", "ERROR")
        return False
    except FileNotFoundError:
        log(f"appimagetool not found at {appimagetool_path}", "ERROR")
        return False
    except Exception as e:
        log(f"appimagetool error: {e}", "ERROR")
        return False


def _sign_appimage(appimage_path, gpg_key=None):
    """Optionally sign the AppImage with GPG."""
    if not gpg_key:
        # Check if a default GPG key exists
        try:
            result = subprocess.run(
                ["gpg", "--list-secret-keys", "--keyid-format=long"],
                capture_output=True, text=True, timeout=10,
            )
            if "sec" not in result.stdout:
                log("No GPG key found — skipping AppImage signing")
                return True
        except Exception:
            log("GPG not available — skipping AppImage signing")
            return True

    cmd = ["gpg", "--detach-sign", "--armor"]
    if gpg_key:
        cmd.extend(["--local-user", gpg_key])
    cmd.append(appimage_path)

    try:
        subprocess.run(cmd, check=True, timeout=30)
        log(f"AppImage signed with GPG")
        return True
    except Exception as e:
        log(f"GPG signing failed (non-fatal): {e}")
        return True  # Non-fatal


def build(version=None, clean=False):
    """Main build function."""
    log("=" * 50)
    log(f"  TrayPrint AppImage Builder v{APP_VERSION}")
    log("=" * 50)

    version = version or APP_VERSION
    appdir_root = os.path.join(DIST_DIR, APPDIR_NAME)
    output_path = os.path.join(DIST_DIR, f"{APP_NAME}-{version}-{ARCH}.AppImage")

    # Step 1: Verify assets
    if not _check_assets():
        sys.exit(1)

    # Step 2: Ensure PyInstaller output exists
    pyinstaller_exe = _check_pyinstaller_output()
    if not pyinstaller_exe:
        sys.exit(1)

    # Step 3: Clean previous AppDir if requested
    if clean and os.path.isdir(appdir_root):
        log(f"Cleaning previous AppDir: {appdir_root}")
        shutil.rmtree(appdir_root)

    # Step 4: Create AppDir
    if not os.path.isdir(appdir_root):
        _create_appdir(appdir_root, pyinstaller_exe, version)
    else:
        log(f"AppDir already exists at {appdir_root} (use --clean to rebuild)")

    # Step 5: Ensure appimagetool
    appimagetool_path = _ensure_appimagetool()
    if not appimagetool_path:
        sys.exit(1)

    # Step 6: Build AppImage
    if not _run_appimagetool(appimagetool_path, appdir_root, output_path):
        sys.exit(1)

    # Step 7: Verify output
    if not os.path.isfile(output_path):
        log(f"Output file not found at {output_path}", "ERROR")
        sys.exit(1)

    size_mb = os.path.getsize(output_path) / (1024 * 1024)
    log(f"AppImage created successfully!")
    log(f"  Output: {output_path}")
    log(f"  Size:   {size_mb:.1f} MB")
    log(f"  Run:    chmod +x {output_path} && ./{output_path}")

    # Step 8: Sign (optional)
    _sign_appimage(output_path)

    return 0


def main():
    parser = argparse.ArgumentParser(
        description="Build TrayPrint AppImage",
    )
    parser.add_argument(
        "--version", default=None,
        help=f"App version (default: {APP_VERSION})",
    )
    parser.add_argument(
        "--clean", action="store_true",
        help="Clean build artifacts before building",
    )
    args = parser.parse_args()

    sys.exit(build(version=args.version, clean=args.clean))


if __name__ == "__main__":
    main()
