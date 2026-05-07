#!/bin/bash
# ──────────────────────────────────────────────────────────────────────────────
# TrayPrint macOS DMG Builder — Shell Script Alternative
# ──────────────────────────────────────────────────────────────────────────────
# Builds a macOS .dmg disk image from a TrayPrint.app bundle using hdiutil.
#
# Usage:
#   ./build_dmg.sh [--app /path/to/TrayPrint.app] [--output TrayPrint.dmg]
#
# Requirements:
#   - macOS with hdiutil (included with macOS)
#   - A valid TrayPrint.app bundle (e.g., from PyInstaller or py2app)
# ──────────────────────────────────────────────────────────────────────────────

set -euo pipefail

# ── Configuration ────────────────────────────────────────────────────────────

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

APP_NAME="TrayPrint.app"
DMG_NAME="TrayPrint.dmg"
VOLUME_NAME="TrayPrint"
ICON_NAME="trayprint.icns"

# ── Argument Parsing ─────────────────────────────────────────────────────────

APP_PATH=""
OUTPUT_PATH=""
ICON_PATH=""

while [[ $# -gt 0 ]]; do
    case "$1" in
        --app)
            APP_PATH="$2"
            shift 2
            ;;
        --output)
            OUTPUT_PATH="$2"
            shift 2
            ;;
        --icon)
            ICON_PATH="$2"
            shift 2
            ;;
        --volume-name)
            VOLUME_NAME="$2"
            shift 2
            ;;
        --help)
            echo "Usage: $0 [--app /path/to/TrayPrint.app] [--output TrayPrint.dmg] [--icon icon.icns] [--volume-name TrayPrint]"
            exit 0
            ;;
        *)
            echo "Unknown option: $1"
            echo "Usage: $0 [--app /path/to/TrayPrint.app] [--output TrayPrint.dmg] [--icon icon.icns] [--volume-name TrayPrint]"
            exit 1
            ;;
    esac
done

# ── Auto-detect App Bundle ───────────────────────────────────────────────────

if [[ -z "$APP_PATH" ]]; then
    for dir in "$PROJECT_DIR/dist" "$PROJECT_DIR/build" "$PROJECT_DIR"; do
        if [[ -d "$dir" ]]; then
            found_app=$(find "$dir" -maxdepth 2 -name "*.app" -type d 2>/dev/null | head -1)
            if [[ -n "$found_app" ]]; then
                APP_PATH="$found_app"
                echo "[INFO] Auto-detected app bundle: $APP_PATH"
                break
            fi
        fi
    done
fi

if [[ -z "$APP_PATH" ]]; then
    echo "[ERROR] Could not find $APP_NAME. Specify with --app or place the .app bundle in dist/, build/, or the project root."
    exit 1
fi

if [[ ! -d "$APP_PATH" ]]; then
    echo "[ERROR] App bundle not found: $APP_PATH"
    exit 1
fi

# ── Ensure Info.plist ────────────────────────────────────────────────────────

PLIST_PATH="$APP_PATH/Contents/Info.plist"
if [[ ! -f "$PLIST_PATH" ]]; then
    echo "[INFO] No Info.plist found. Creating a minimal one."
    mkdir -p "$APP_PATH/Contents"
    cat > "$PLIST_PATH" <<-EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key>
    <string>TrayPrint</string>
    <key>CFBundleDisplayName</key>
    <string>TrayPrint</string>
    <key>CFBundleIdentifier</key>
    <string>com.printhub.trayprint</string>
    <key>CFBundleVersion</key>
    <string>1.0.0</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0.0</string>
    <key>CFBundleExecutable</key>
    <string>TrayPrint</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleInfoDictionaryVersion</key>
    <string>6.0</string>
    <key>NSHighResolutionCapable</key>
    <true/>
    <key>LSUIElement</key>
    <true/>
</dict>
</plist>
EOF
    echo "[INFO] Created minimal Info.plist at $PLIST_PATH"
fi

# ── Resolve Icon ─────────────────────────────────────────────────────────────

if [[ -z "$ICON_PATH" ]]; then
    if [[ -f "$SCRIPT_DIR/$ICON_NAME" ]]; then
        ICON_PATH="$SCRIPT_DIR/$ICON_NAME"
        echo "[INFO] Using icon: $ICON_PATH"
    else
        echo "[WARN] No icon found at $SCRIPT_DIR/$ICON_NAME. Creating a placeholder..."
        # Create a minimal 16x16 PNG as placeholder
        ICON_PATH="$SCRIPT_DIR/$ICON_NAME"
        python3 -c "
import struct, zlib

def create_minimal_icns(path):
    # Minimal 16x16 PNG
    png_data = b'\x89PNG\r\n\x1a\n' + \
        struct.pack('>I', 13) + b'IHDR' + struct.pack('>IIBBBBB', 16, 16, 8, 2, 0, 0, 0) + \
        struct.pack('>I', zlib.crc32(b'IHDR' + struct.pack('>IIBBBBB', 16, 16, 8, 2, 0, 0, 0)) & 0xffffffff) + \
        struct.pack('>I', 1) + b'sRGB' + struct.pack('>I', 0xaece1ce9) + \
        struct.pack('>I', 4) + b'gAMA' + struct.pack('>I', 0x00b18f0b) + struct.pack('>I', zlib.crc32(b'gAMA' + struct.pack('>I', 0x00b18f0b)) & 0xffffffff) + \
        struct.pack('>I', 21) + b'IDAT' + zlib.compress(bytes([0x00] * 256)) + \
        struct.pack('>I', zlib.crc32(b'IDAT' + zlib.compress(bytes([0x00] * 256))) & 0xffffffff) + \
        struct.pack('>I', 0) + b'IEND' + struct.pack('>I', zlib.crc32(b'IEND') & 0xffffffff)
    
    icon_entry = b'icp4' + struct.pack('>I', len(png_data) + 8) + png_data
    icns_data = b'icns' + struct.pack('>I', len(icon_entry) + 8) + icon_entry
    
    with open(path, 'wb') as f:
        f.write(icns_data)
    print(f'[INFO] Created placeholder icon: {path}')

create_minimal_icns('$ICON_PATH')
" 2>/dev/null || echo "[WARN] Could not create placeholder icon; continuing without one."
    fi
fi

# ── Resolve Output Path ──────────────────────────────────────────────────────

if [[ -z "$OUTPUT_PATH" ]]; then
    OUTPUT_PATH="$SCRIPT_DIR/$DMG_NAME"
fi
OUTPUT_PATH="$(cd "$(dirname "$OUTPUT_PATH")" && pwd)/$(basename "$OUTPUT_PATH")"

# ── Build DMG ────────────────────────────────────────────────────────────────

echo "[INFO] Building DMG: $OUTPUT_PATH"
echo "[INFO] App bundle: $APP_PATH"
echo "[INFO] Volume name: $VOLUME_NAME"

# Create a temporary directory for DMG contents
TMPDIR=$(mktemp -d -t trayprint_dmg)
trap 'rm -rf "$TMPDIR"' EXIT

DMG_CONTENTS="$TMPDIR/contents"
mkdir -p "$DMG_CONTENTS"

# Copy app bundle
echo "[INFO] Copying app bundle..."
cp -R "$APP_PATH" "$DMG_CONTENTS/"

# Create Applications symlink
ln -s /Applications "$DMG_CONTENTS/Applications"

# Copy icon as volume icon if provided
if [[ -n "$ICON_PATH" && -f "$ICON_PATH" ]]; then
    cp "$ICON_PATH" "$DMG_CONTENTS/.VolumeIcon.icns"
fi

# Create read/write DMG
TEMP_DMG="$TMPDIR/temp.dmg"
echo "[INFO] Creating temporary DMG..."
hdiutil create \
    -volname "$VOLUME_NAME" \
    -srcfolder "$DMG_CONTENTS" \
    -ov \
    -format UDRW \
    "$TEMP_DMG"

# Attach and set volume icon if available
if [[ -n "$ICON_PATH" && -f "$ICON_PATH" ]]; then
    MOUNT_POINT="$TMPDIR/mount"
    echo "[INFO] Attaching DMG to set volume icon..."
    hdiutil attach "$TEMP_DMG" -mountpoint "$MOUNT_POINT"
    
    # Copy icon and set custom icon attribute
    if [[ ! -f "$MOUNT_POINT/.VolumeIcon.icns" ]]; then
        cp "$ICON_PATH" "$MOUNT_POINT/.VolumeIcon.icns"
    fi
    
    # Set the custom icon attribute (requires SetFile or Python)
    if command -v SetFile &>/dev/null; then
        SetFile -a C "$MOUNT_POINT"
    else
        # Fallback: use Python to set the custom icon bit
        python3 -c "
import struct
with open('$MOUNT_POINT/.VolumeIcon.icns', 'ab') as f:
    pass
# Set the custom icon bit using os.stat
import os
# The custom icon bit is stored in the Finder info
# We use the 'C' attribute via xattr
os.system('xattr -w com.apple.FinderInfo \$(printf \"\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x04\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\\x00\") \"$MOUNT_POINT\" 2>/dev/null')" 2>/dev/null || true
    fi
    
    hdiutil detach "$MOUNT_POINT"
fi

# Convert to compressed DMG
echo "[INFO] Compressing DMG..."
hdiutil convert "$TEMP_DMG" \
    -format UDZO \
    -o "$OUTPUT_PATH" \
    -ov

# ── Verify ───────────────────────────────────────────────────────────────────

if [[ -f "$OUTPUT_PATH" ]]; then
    SIZE_MB=$(du -h "$OUTPUT_PATH" | cut -f1)
    echo "[SUCCESS] DMG package created: $OUTPUT_PATH ($SIZE_MB)"
else
    echo "[ERROR] DMG was not created at $OUTPUT_PATH"
    exit 1
fi
