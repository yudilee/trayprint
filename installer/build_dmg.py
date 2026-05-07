#!/usr/bin/env python3
"""
macOS DMG Packaging Script for TrayPrint.

Builds a macOS .dmg disk image containing the TrayPrint application bundle.

Usage:
    python build_dmg.py [--output TrayPrint.dmg] [--app-dir ../dist/TrayPrint.app]

Requires:
    - dmgbuild (pip install dmgbuild) — optional, falls back to hdiutil
    - A valid .app bundle (e.g., from PyInstaller or py2app)
"""

import argparse
import os
import plistlib
import shutil
import subprocess
import sys
import tempfile

# ── Constants ────────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, ".."))

DEFAULT_APP_NAME = "TrayPrint.app"
DEFAULT_DMG_NAME = "TrayPrint.dmg"
DEFAULT_ICON_NAME = "trayprint.icns"

# Minimal Info.plist template (used if the .app lacks one)
INFO_PLIST_TEMPLATE = {
    "CFBundleName": "TrayPrint",
    "CFBundleDisplayName": "TrayPrint",
    "CFBundleIdentifier": "com.printhub.trayprint",
    "CFBundleVersion": "1.0.0",
    "CFBundleShortVersionString": "1.0.0",
    "CFBundleExecutable": "TrayPrint",
    "CFBundlePackageType": "APPL",
    "CFBundleInfoDictionaryVersion": "6.0",
    "NSHighResolutionCapable": True,
    "LSUIElement": True,  # Menu bar app, no dock icon
}


# ── Helpers ──────────────────────────────────────────────────────────────────


def find_app_bundle(search_dirs=None):
    """Locate the .app bundle in common locations."""
    if search_dirs is None:
        search_dirs = [
            os.path.join(PROJECT_DIR, "dist"),
            os.path.join(PROJECT_DIR, "build"),
            PROJECT_DIR,
        ]

    for d in search_dirs:
        if not os.path.isdir(d):
            continue
        for entry in os.listdir(d):
            if entry.endswith(".app"):
                return os.path.join(d, entry)
    return None


def ensure_info_plist(app_path):
    """Ensure the .app bundle has a valid Info.plist; create one if missing."""
    plist_path = os.path.join(app_path, "Contents", "Info.plist")
    if os.path.exists(plist_path):
        return  # Already has one

    print("[INFO] No Info.plist found in app bundle. Creating a minimal one.")
    os.makedirs(os.path.dirname(plist_path), exist_ok=True)
    with open(plist_path, "wb") as f:
        plistlib.dump(INFO_PLIST_TEMPLATE, f)


def create_placeholder_icns(output_path):
    """Create a minimal placeholder .icns file.

    This generates a tiny 16x16 PNG, wraps it in a basic .icns container.
    For a production build, replace this with a proper icon.
    """
    if os.path.exists(output_path):
        print(f"[INFO] Icon already exists at {output_path}, skipping creation.")
        return

    print("[INFO] Creating placeholder trayprint.icns (16x16).")
    try:
        from PIL import Image

        # Create a simple 16x16 icon with a printer-like symbol
        img = Image.new("RGBA", (16, 16), (0, 0, 0, 0))
        pixels = img.load()
        # Draw a simple printer shape
        for x in range(2, 14):
            for y in range(4, 10):
                pixels[x, y] = (52, 120, 246, 255)  # Blue
        for x in range(4, 12):
            for y in range(10, 14):
                pixels[x, y] = (52, 120, 246, 255)
        # Paper coming out
        for x in range(6, 10):
            for y in range(1, 4):
                pixels[x, y] = (255, 255, 255, 255)

        # Save as PNG then convert to icns
        png_path = output_path + ".png"
        img.save(png_path, "PNG")

        # Use iconutil if available (macOS)
        iconset_path = output_path + ".iconset"
        os.makedirs(iconset_path, exist_ok=True)
        # Copy the same image for all sizes (placeholder)
        for size in [16, 32, 64, 128, 256, 512]:
            resized = img.resize((size, size), Image.NEAREST)
            resized.save(os.path.join(iconset_path, f"icon_{size}x{size}.png"))
            if size in [16, 32, 128, 256]:
                resized.save(
                    os.path.join(iconset_path, f"icon_{size // 2}x{size // 2}@2x.png")
                )

        subprocess.run(
            ["iconutil", "-c", "icns", iconset_path, "-o", output_path],
            capture_output=True,
            timeout=30,
        )

        # Cleanup
        shutil.rmtree(iconset_path, ignore_errors=True)
        os.remove(png_path)

        if os.path.exists(output_path):
            print(f"[INFO] Created placeholder icon: {output_path}")
        else:
            print("[WARN] iconutil failed; creating minimal binary icns.")
            _create_minimal_icns_binary(output_path)

    except ImportError:
        print("[WARN] PIL/Pillow not available. Creating minimal binary icns.")
        _create_minimal_icns_binary(output_path)


def _create_minimal_icns_binary(output_path):
    """Create the absolute minimum valid .icns file (one 16x16 icon entry)."""
    # A minimal valid .icns with a single 16x16 icon entry
    # icns header: magic 'icns' + file size (4+4 bytes)
    # Icon entry: type 'icp4' (16x16 PNG) + size + data
    # Minimal 16x16 PNG (68 bytes)
    minimal_png = (
        b"\x89PNG\r\n\x1a\n"  # PNG signature
        b"\x00\x00\x00\rIHDR\x00\x00\x00\x10\x00\x00\x00\x10"
        b"\x08\x02\x00\x00\x00\x90\x91\xb6\x0e"
        b"\x00\x00\x00\x01sRGB\x00\xae\xce\x1c\xe9"
        b"\x00\x00\x00\x04gAMA\x00\x00\xb1\x8f\x0b\xfca\x05"
        b"\x00\x00\x00\x15IDATx\x9cc\xfc\xff\xff?\x03\x10\x00"
        b"\x00\xff\xff\x03\x00\x01\x20\x01\x8d"
        b"\x00\x00\x00\x00IEND\xaeB`\x82"
    )
    icon_entry = b"icp4" + (len(minimal_png) + 8).to_bytes(4, "big") + minimal_png
    icns_data = b"icns" + (len(icon_entry) + 8).to_bytes(4, "big") + icon_entry

    with open(output_path, "wb") as f:
        f.write(icns_data)
    print(f"[INFO] Created minimal binary icns: {output_path}")


# ── Build Methods ────────────────────────────────────────────────────────────


def build_with_dmgbuild(app_path, output_path, icon_path, volume_name):
    """Build DMG using the dmgbuild Python library."""
    try:
        import dmgbuild
    except ImportError:
        print("[INFO] dmgbuild not available, falling back to hdiutil.")
        return False

    settings = {
        "volume_name": volume_name,
        "format": "UDBZ",
        "size": None,  # auto
        "files": [app_path],
        "symlinks": {"Applications": "/Applications"},
        "icon_locations": {os.path.basename(app_path): (140, 120), "Applications": (420, 120)},
        "background": None,
        "show_status_bar": False,
        "show_tab_view": False,
        "show_toolbar": False,
        "show_pathbar": False,
        "show_sidebar": False,
        "sidebar_width": 180,
        "window_rect": ((100, 100), (640, 320)),
        "default_view": "icon-view",
        "icon_size": 80,
        "text_size": 12,
        "arranged_by": None,
        "grid_offset": (0, 0),
        "grid_spacing": 100,
        "scroll_position": (0, 0),
        "label_pos": "bottom",
    }

    if icon_path and os.path.exists(icon_path):
        settings["icon"] = icon_path

    print(f"[INFO] Building DMG with dmgbuild: {output_path}")
    dmgbuild.build_dmg(output_path, volume_name, settings)
    return True


def build_with_hdiutil(app_path, output_path, icon_path, volume_name):
    """Build DMG using the hdiutil CLI tool (native macOS)."""
    print("[INFO] Building DMG with hdiutil.")

    with tempfile.TemporaryDirectory(prefix="trayprint_dmg_") as tmpdir:
        # Create a temporary directory with the app and Applications symlink
        dmg_contents = os.path.join(tmpdir, "contents")
        os.makedirs(dmg_contents)

        # Copy app bundle
        dest_app = os.path.join(dmg_contents, os.path.basename(app_path))
        print(f"[INFO] Copying app bundle to staging directory...")
        if os.path.isdir(app_path):
            shutil.copytree(app_path, dest_app, symlinks=True)
        else:
            shutil.copy2(app_path, dest_app)

        # Create Applications symlink
        os.symlink("/Applications", os.path.join(dmg_contents, "Applications"))

        # Copy icon as volume icon if provided
        if icon_path and os.path.exists(icon_path):
            shutil.copy2(icon_path, os.path.join(dmg_contents, ".VolumeIcon.icns"))

        # Create the DMG
        temp_dmg = os.path.join(tmpdir, "temp.dmg")
        print(f"[INFO] Creating DMG: {output_path}")

        # Create a read/write DMG first
        subprocess.run(
            [
                "hdiutil",
                "create",
                "-volname", volume_name,
                "-srcfolder", dmg_contents,
                "-ov",
                "-format", "UDRW",
                temp_dmg,
            ],
            check=True,
            capture_output=True,
            timeout=120,
        )

        # Attach and set icon if needed
        if icon_path and os.path.exists(icon_path):
            attach_result = subprocess.run(
                ["hdiutil", "attach", temp_dmg, "-mountpoint", os.path.join(tmpdir, "mount")],
                check=True,
                capture_output=True,
                timeout=30,
            )
            # Set custom icon
            mount_point = os.path.join(tmpdir, "mount")
            icon_dest = os.path.join(mount_point, ".VolumeIcon.icns")
            if os.path.exists(icon_path) and not os.path.exists(icon_dest):
                shutil.copy2(icon_path, icon_dest)
                # Set the custom icon attribute
                subprocess.run(
                    ["SetFile", "-a", "C", mount_point],
                    capture_output=True,
                    timeout=10,
                )
            subprocess.run(
                ["hdiutil", "detach", mount_point],
                check=True,
                capture_output=True,
                timeout=30,
            )

        # Convert to compressed DMG
        subprocess.run(
            [
                "hdiutil",
                "convert",
                temp_dmg,
                "-format", "UDZO",
                "-o", output_path,
                "-ov",
            ],
            check=True,
            capture_output=True,
            timeout=120,
        )

    print(f"[SUCCESS] DMG created: {output_path}")
    return True


# ── Main ─────────────────────────────────────────────────────────────────────


def main():
    parser = argparse.ArgumentParser(
        description="Build a macOS .dmg package for TrayPrint."
    )
    parser.add_argument(
        "--output",
        default=None,
        help=f"Output DMG path (default: ./{DEFAULT_DMG_NAME})",
    )
    parser.add_argument(
        "--app-dir",
        default=None,
        help=f"Path to TrayPrint.app bundle (default: auto-detect in dist/, build/, project root)",
    )
    parser.add_argument(
        "--icon",
        default=None,
        help=f"Path to .icns icon file (default: look for {DEFAULT_ICON_NAME} in script dir)",
    )
    parser.add_argument(
        "--volume-name",
        default="TrayPrint",
        help="Volume name for the DMG (default: TrayPrint)",
    )
    parser.add_argument(
        "--force-hdiutil",
        action="store_true",
        help="Force using hdiutil even if dmgbuild is available",
    )

    args = parser.parse_args()

    # ── Resolve app bundle ────────────────────────────────────
    app_path = args.app_dir
    if not app_path:
        app_path = find_app_bundle()
        if not app_path:
            print(
                "[ERROR] Could not find TrayPrint.app. Specify with --app-dir or "
                "place the .app bundle in dist/, build/, or the project root."
            )
            sys.exit(1)
        print(f"[INFO] Auto-detected app bundle: {app_path}")

    if not os.path.exists(app_path):
        print(f"[ERROR] App bundle not found: {app_path}")
        sys.exit(1)

    # ── Ensure Info.plist ─────────────────────────────────────
    try:
        ensure_info_plist(app_path)
    except Exception as e:
        print(f"[WARN] Could not ensure Info.plist: {e}")

    # ── Resolve icon ──────────────────────────────────────────
    icon_path = args.icon
    if not icon_path:
        candidate = os.path.join(SCRIPT_DIR, DEFAULT_ICON_NAME)
        if os.path.exists(candidate):
            icon_path = candidate
            print(f"[INFO] Using icon: {icon_path}")
        else:
            # Create placeholder icon
            icon_path = candidate
            try:
                create_placeholder_icns(icon_path)
            except Exception as e:
                print(f"[WARN] Could not create placeholder icon: {e}")
                icon_path = None

    # ── Resolve output path ───────────────────────────────────
    output_path = args.output
    if not output_path:
        output_path = os.path.join(SCRIPT_DIR, DEFAULT_DMG_NAME)
    output_path = os.path.abspath(output_path)

    # ── Build DMG ─────────────────────────────────────────────
    volume_name = args.volume_name

    success = False
    if not args.force_hdiutil:
        try:
            success = build_with_dmgbuild(app_path, output_path, icon_path, volume_name)
        except Exception as e:
            print(f"[WARN] dmgbuild failed: {e}")

    if not success:
        try:
            success = build_with_hdiutil(app_path, output_path, icon_path, volume_name)
        except Exception as e:
            print(f"[ERROR] hdiutil also failed: {e}")
            sys.exit(1)

    # ── Verify ────────────────────────────────────────────────
    if os.path.exists(output_path):
        size_mb = os.path.getsize(output_path) / (1024 * 1024)
        print(f"[SUCCESS] DMG package created: {output_path} ({size_mb:.1f} MB)")
    else:
        print(f"[ERROR] DMG was not created at {output_path}")
        sys.exit(1)


if __name__ == "__main__":
    main()
