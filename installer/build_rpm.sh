#!/bin/bash
#
# TrayPrint — RPM Package Builder
#
# Builds an .rpm package for Fedora/RHEL/CentOS from source
# (using system Python dependencies via dnf).
#
# Usage:
#   ./installer/build_rpm.sh [--version X.Y.Z] [--clean]
#
# Prerequisites:
#   - rpmbuild (install: dnf install rpm-build)
#   - All Python dependencies listed in requirements.txt
#
# Output:
#   dist/trayprint-X.Y.Z-1.x86_64.rpm
#

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
DIST_DIR="$PROJECT_ROOT/dist"

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

RELEASE="1"
PACKAGE_NAME="trayprint"
RPM_NAME="${PACKAGE_NAME}-${VERSION}-${RELEASE}.x86_64.rpm"
RPM_OUTPUT="$DIST_DIR/$RPM_NAME"

echo "================================================"
echo "  TrayPrint RPM Package Builder"
echo "  Version: $VERSION-$RELEASE"
echo "================================================"
echo ""

# ── Step 1: Check prerequisites ──
echo "[1/5] Checking prerequisites..."

if ! command -v rpmbuild &>/dev/null; then
    echo "ERROR: rpmbuild not found. Install with:"
    echo "  dnf install rpm-build"
    exit 1
fi

echo "  ✓ rpmbuild found"

# ── Step 2: Create RPM build environment ──
echo "[2/5] Setting up RPM build environment..."

RPM_BUILD_ROOT="${HOME}/rpmbuild"
RPM_SOURCES="${RPM_BUILD_ROOT}/SOURCES"
RPM_SPECS="${RPM_BUILD_ROOT}/SPECS"
RPM_RPMS="${RPM_BUILD_ROOT}/RPMS/x86_64"

for dir in SOURCES SPECS BUILD RPMS SRPMS; do
    mkdir -p "${RPM_BUILD_ROOT}/${dir}"
done

# ── Step 3: Create source tarball ──
echo "[3/5] Creating source tarball..."

TARBALL_NAME="${PACKAGE_NAME}-${VERSION}"
TARBALL_DIR="/tmp/${TARBALL_NAME}"
TARBALL_PATH="${RPM_SOURCES}/${TARBALL_NAME}.tar.gz"

# Clean previous
if [ "$CLEAN" = true ]; then
    rm -rf "$TARBALL_DIR" "$TARBALL_PATH"
fi

if [ -f "$TARBALL_PATH" ]; then
    echo "  Source tarball already exists at $TARBALL_PATH"
else
    # Copy project files (excluding build artifacts)
    mkdir -p "$TARBALL_DIR"

    # Use rsync or cp to copy files
    if command -v rsync &>/dev/null; then
        rsync -a \
            --exclude='__pycache__' \
            --exclude='*.pyc' \
            --exclude='.git' \
            --exclude='tests' \
            --exclude='dist' \
            --exclude='venv' \
            --exclude='*.db' \
            --exclude='plans' \
            "$PROJECT_ROOT/" "$TARBALL_DIR/"
    else
        # Fallback: copy individual files/directories
        mkdir -p "$TARBALL_DIR/installer" "$TARBALL_DIR/templates"

        for item in app.py server.py printer.py capabilities.py theme.py \
                    ui_settings.py autostart.py logger.py log_utils.py \
                    path_utils.py updater.py websocket_client.py \
                    notification_history.py queue_dialog.py \
                    diagnostics_dialog.py platform_darwin.py service.py \
                    build.py config.json trayprint.ico requirements.txt \
                    README.md; do
            [ -f "$PROJECT_ROOT/$item" ] && cp "$PROJECT_ROOT/$item" "$TARBALL_DIR/"
        done

        cp -r "$PROJECT_ROOT/templates"/* "$TARBALL_DIR/templates/" 2>/dev/null || true
        cp -r "$PROJECT_ROOT/installer/trayprint.desktop" "$TARBALL_DIR/installer/"
        cp -r "$PROJECT_ROOT/installer/trayprint.png" "$TARBALL_DIR/installer/"
        cp -r "$PROJECT_ROOT/installer/trayprint.appdata.xml" "$TARBALL_DIR/installer/"
        cp -r "$PROJECT_ROOT/installer/trayprint.service" "$TARBALL_DIR/installer/"
        cp -r "$PROJECT_ROOT/installer/trayprint.spec" "$TARBALL_DIR/installer/"

        # Copy license if exists
        [ -f "$PROJECT_ROOT/LICENSE" ] && cp "$PROJECT_ROOT/LICENSE" "$TARBALL_DIR/"
    fi

    # Create tarball
    cd /tmp
    tar czf "$TARBALL_PATH" "$TARBALL_NAME" 2>/dev/null
    rm -rf "$TARBALL_DIR"

    echo "  Created source tarball: $TARBALL_PATH"
fi

# ── Step 4: Prepare spec and build ──
echo "[4/5] Building RPM package..."

# Copy spec file
SPEC_SRC="$SCRIPT_DIR/trayprint.spec"
SPEC_DEST="${RPM_SPECS}/trayprint.spec"

# Update version in spec
sed "s/%global _version .*/%global _version ${VERSION}/" "$SPEC_SRC" > "$SPEC_DEST"
sed -i "s/%global _release .*/%global _release ${RELEASE}/" "$SPEC_DEST"

# Build the RPM
rpmbuild -ba "$SPEC_DEST" 2>&1

if [ $? -ne 0 ]; then
    echo "ERROR: rpmbuild failed" >&2
    exit 1
fi

echo "  ✓ RPM package built"

# ── Step 5: Copy output and verify ──
echo "[5/5] Copying output to dist/..."

mkdir -p "$DIST_DIR"
# Find the built RPM in the standard location
BUILT_RPM="${RPM_RPMS}/${RPM_NAME}"

if [ -f "$BUILT_RPM" ]; then
    cp "$BUILT_RPM" "$RPM_OUTPUT"
    echo "  Copied to: $RPM_OUTPUT"
else
    # Try alternate locations
    ALT_RPM=$(find "${RPM_BUILD_ROOT}/RPMS" -name "${PACKAGE_NAME}-${VERSION}-*.rpm" 2>/dev/null | head -1)
    if [ -n "$ALT_RPM" ]; then
        cp "$ALT_RPM" "$RPM_OUTPUT"
        echo "  Copied from $ALT_RPM"
    else
        echo "WARNING: Could not find built RPM. Check ${RPM_BUILD_ROOT}/RPMS/" >&2
        exit 1
    fi
fi

SIZE_MB=$(du -h "$RPM_OUTPUT" | cut -f1)

echo ""
echo "================================================"
echo "  BUILD SUCCESS!"
echo "  Output: $RPM_OUTPUT"
echo "  Size:   $SIZE_MB"
echo "================================================"
echo ""
echo "Install with:"
echo "  sudo dnf install $RPM_OUTPUT"
echo ""
echo "Or:"
echo "  sudo rpm -ivh $RPM_OUTPUT"
echo ""
