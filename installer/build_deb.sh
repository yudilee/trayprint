#!/bin/bash
#
# TrayPrint — Debian Package Builder
#
# Builds a .deb package for Debian/Ubuntu from source (using system Python deps).
# For a self-contained binary, use the AppImage instead.
#
# Usage:
#   ./installer/build_deb.sh [--version X.Y.Z] [--clean]
#
# Prerequisites:
#   - dpkg-deb
#   - All Python dependencies listed in requirements.txt
#
# Output:
#   dist/trayprint_X.Y.Z-1_amd64.deb
#

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
DIST_DIR="$PROJECT_ROOT/dist"
BUILD_DIR="$DIST_DIR/deb-build"

# Default version — read from config.json if available
DEFAULT_VERSION="3.0.0"
if [ -f "$PROJECT_ROOT/config.json" ]; then
    CONFIG_VER=$(python3 -c "import json; print(json.load(open('$PROJECT_ROOT/config.json')).get('version', '$DEFAULT_VERSION'))" 2>/dev/null || echo "$DEFAULT_VERSION")
else
    CONFIG_VER="$DEFAULT_VERSION"
fi

VERSION="${CONFIG_VER}"
CLEAN=false

# Parse arguments
while [[ $# -gt 0 ]]; do
    case "$1" in
        --version)
            VERSION="$2"
            shift 2
            ;;
        --clean)
            CLEAN=true
            shift
            ;;
        *)
            echo "Unknown option: $1"
            echo "Usage: $0 [--version X.Y.Z] [--clean]"
            exit 1
            ;;
    esac
done

PACKAGE_NAME="trayprint"
PACKAGE_VERSION="${VERSION}-1"
ARCH="amd64"
DEB_NAME="${PACKAGE_NAME}_${PACKAGE_VERSION}_${ARCH}.deb"
DEB_OUTPUT="$DIST_DIR/$DEB_NAME"

echo "================================================"
echo "  TrayPrint Debian Package Builder"
echo "  Version: $VERSION"
echo "================================================"
echo ""

# ── Step 1: Check prerequisites ──
echo "[1/5] Checking prerequisites..."

if ! command -v dpkg-deb &>/dev/null; then
    echo "ERROR: dpkg-deb not found. Is dpkg-dev installed?"
    echo "  sudo apt install dpkg-dev"
    exit 1
fi

echo "  ✓ dpkg-deb found"

# ── Step 2: Clean previous build ──
echo "[2/5] Preparing build directory..."

if [ "$CLEAN" = true ] && [ -d "$BUILD_DIR" ]; then
    echo "  Cleaning previous build..."
    rm -rf "$BUILD_DIR"
fi

if [ -d "$BUILD_DIR" ]; then
    echo "  Build directory already exists at $BUILD_DIR (use --clean to rebuild)"
else
    mkdir -p "$BUILD_DIR"
    echo "  Created build directory: $BUILD_DIR"
fi

# ── Step 3: Populate package directory structure ──
echo "[3/5] Populating package directory..."

# Create directory structure
PKG_DIR="$BUILD_DIR/package"
DEBIAN_DIR="$PKG_DIR/DEBIAN"
USR_BIN="$PKG_DIR/usr/bin"
USR_LIB="$PKG_DIR/usr/lib/trayprint"
USR_SHARE_APPS="$PKG_DIR/usr/share/applications"
USR_SHARE_ICONS="$PKG_DIR/usr/share/icons/hicolor/256x256/apps"
USR_SHARE_META="$PKG_DIR/usr/share/metainfo"
USR_SHARE_DOC="$PKG_DIR/usr/share/doc/trayprint"
USR_UNIT="$PKG_DIR/usr/lib/systemd/user"
ETC_DIR="$PKG_DIR/etc/trayprint"

mkdir -p "$DEBIAN_DIR" \
    "$USR_BIN" \
    "$USR_LIB" \
    "$USR_SHARE_APPS" \
    "$USR_SHARE_ICONS" \
    "$USR_SHARE_META" \
    "$USR_SHARE_DOC" \
    "$USR_UNIT" \
    "$ETC_DIR"

# Copy control file
cp "$SCRIPT_DIR/debian/control" "$DEBIAN_DIR/control"
# Update version in control file
sed -i "s/^Version:.*/Version: $PACKAGE_VERSION/" "$DEBIAN_DIR/control"

# Copy maintainer scripts
cp "$SCRIPT_DIR/debian/postinst" "$DEBIAN_DIR/postinst"
cp "$SCRIPT_DIR/debian/prerm" "$DEBIAN_DIR/prerm"
cp "$SCRIPT_DIR/debian/postrm" "$DEBIAN_DIR/postrm"
chmod 755 "$DEBIAN_DIR/postinst" "$DEBIAN_DIR/prerm" "$DEBIAN_DIR/postrm"

# Create launcher script (/usr/bin/trayprint)
cat > "$USR_BIN/trayprint" << 'LAUNCHER'
#!/bin/bash
# TrayPrint launcher — resolves the application directory and runs the Python app
SCRIPT_DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
INSTALL_DIR="/usr/lib/trayprint"
CONFIG_DIR="${TRAYPRINT_CONFIG_PATH:-/etc/trayprint}"

# If a custom config path is provided via env var, pass it
if [ -n "${TRAYPRINT_CONFIG_PATH:-}" ]; then
    exec python3 "$INSTALL_DIR/app.py" --config-path "$TRAYPRINT_CONFIG_PATH" "$@"
else
    exec python3 "$INSTALL_DIR/app.py" "$@"
fi
LAUNCHER
chmod 755 "$USR_BIN/trayprint"

# Copy application source files (excluding tests, installer, etc.)
rsync -a --exclude='__pycache__' \
    --exclude='*.pyc' \
    --exclude='.git' \
    --exclude='tests' \
    --exclude='installer' \
    --exclude='dist' \
    --exclude='venv' \
    --exclude='*.db' \
    --include='*.py' \
    --include='*.html' \
    --include='*.json' \
    --include='*.ico' \
    "$PROJECT_ROOT/" "$USR_LIB/"

# Remove build system files from the package
rm -f "$USR_LIB/build.py"

# Copy desktop file
cp "$SCRIPT_DIR/trayprint.desktop" "$USR_SHARE_APPS/trayprint.desktop"

# Copy icon
cp "$SCRIPT_DIR/trayprint.png" "$USR_SHARE_ICONS/trayprint.png"

# Copy AppStream metadata
cp "$SCRIPT_DIR/trayprint.appdata.xml" "$USR_SHARE_META/trayprint.appdata.xml"

# Copy systemd user service
cp "$SCRIPT_DIR/trayprint.service" "$USR_UNIT/trayprint.service"
# Fix executable path for Debian install
sed -i "s|/usr/bin/trayprint|/usr/bin/trayprint|" "$USR_UNIT/trayprint.service"

# Copy default config
cp "$PROJECT_ROOT/config.json" "$ETC_DIR/config.json"

# Copy README and license
cp "$PROJECT_ROOT/README.md" "$USR_SHARE_DOC/README.md"
cp "$PROJECT_ROOT/requirements.txt" "$USR_SHARE_DOC/requirements.txt"

# Copy templates
if [ -d "$USR_LIB/templates" ]; then
    cp -r "$PROJECT_ROOT/templates" "$USR_LIB/"
fi

echo "  ✓ Package directory populated"

# ── Step 4: Build the .deb ──
echo "[4/5] Building .deb package..."

# Compute installed size (in KB)
INSTALLED_SIZE=$(du -sk "$PKG_DIR" | cut -f1)
sed -i "/^Installed-Size:/d" "$DEBIAN_DIR/control"
echo "Installed-Size: $INSTALLED_SIZE" >> "$DEBIAN_DIR/control"

# Fix permissions
find "$PKG_DIR" -type d -exec chmod 755 {} \;
find "$PKG_DIR" -type f -exec chmod 644 {} \;
chmod 755 "$DEBIAN_DIR/postinst" "$DEBIAN_DIR/prerm" "$DEBIAN_DIR/postrm"
chmod 755 "$USR_BIN/trayprint"

# Build the package
fakeroot dpkg-deb --build "$PKG_DIR" "$DEB_OUTPUT" 2>&1

if [ ! -f "$DEB_OUTPUT" ]; then
    echo "ERROR: Failed to build .deb package" >&2
    exit 1
fi

echo "  ✓ .deb package built"

# ── Step 5: Verify ──
echo "[5/5] Verifying package..."

SIZE_MB=$(du -h "$DEB_OUTPUT" | cut -f1)
echo "  Package: $DEB_OUTPUT"
echo "  Size:    $SIZE_MB"

# Show package info
dpkg-deb --info "$DEB_OUTPUT" 2>/dev/null || true

echo ""
echo "================================================"
echo "  BUILD SUCCESS!"
echo "  Output: $DEB_OUTPUT"
echo "  Size:   $SIZE_MB"
echo "================================================"
echo ""
echo "Install with:"
echo "  sudo apt install ./$DEB_NAME"
echo ""
echo "Or on Debian-based systems:"
echo "  sudo dpkg -i $DEB_OUTPUT"
echo "  sudo apt-get install -f    # Install missing dependencies"
echo ""
