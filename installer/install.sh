#!/bin/bash
#
# TrayPrint — Universal Linux Installer
#
# Installs TrayPrint on any Linux system. Auto-detects the best format
# (.deb, .rpm, or AppImage) based on the available files in the installer/
# or dist/ directory.
#
# Usage:
#   ./installer/install.sh                  # Interactive (auto-detect format)
#   ./installer/install.sh --appimage       # Force AppImage install
#   ./installer/install.sh --deb            # Force .deb install
#   ./installer/install.sh --rpm            # Force .rpm install
#   ./installer/install.sh --user           # Install for current user only
#   ./installer/install.sh --prefix /opt    # Custom install prefix
#   ./installer/install.sh --help           # Show help
#
# Silent/automated:
#   ./installer/install.sh --appimage --prefix /opt/trayprint --no-prompt
#

set -euo pipefail

# ── Colors ──
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# ── Configuration ──
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
DIST_DIR="$PROJECT_ROOT/dist"

APP_NAME="TrayPrint"
APP_BINARY="trayprint"

# Auto-detect version
VERSION="3.0.0"
if [ -f "$PROJECT_ROOT/config.json" ]; then
    VERSION=$(python3 -c "import json; print(json.load(open('$PROJECT_ROOT/config.json')).get('version', '$VERSION'))" 2>/dev/null || echo "$VERSION")
fi

# Install mode
MODE="auto"       # auto, appimage, deb, rpm
PREFIX=""         # empty = use defaults (system or user)
NO_PROMPT=false
USER_INSTALL=false

# ── Help ──

show_help() {
    cat << EOF
TrayPrint v${VERSION} — Universal Linux Installer

Installs TrayPrint, the local print agent for Print Hub, on any Linux system.

USAGE:
    $0 [OPTIONS]

OPTIONS:
    --appimage          Force AppImage installation (portable)
    --deb               Force .deb package installation (Debian/Ubuntu)
    --rpm               Force .rpm package installation (Fedora/RHEL)
    --user              Install for current user only (default: system-wide)
    --prefix DIR        Install to a custom prefix directory
    --no-prompt         Non-interactive mode (answer yes to all prompts)
    --help              Show this help message

EXAMPLES:
    # Interactive install (auto-detects best format)
    $0

    # Install AppImage system-wide
    $0 --appimage

    # Install to user home directory
    $0 --appimage --user

    # Install .deb on Ubuntu without prompts
    $0 --deb --no-prompt

    # Custom prefix
    $0 --prefix /opt/trayprint
EOF
    exit 0
}

# ── Parse Arguments ──

while [[ $# -gt 0 ]]; do
    case "$1" in
        --appimage) MODE="appimage"; shift ;;
        --deb)      MODE="deb"; shift ;;
        --rpm)      MODE="rpm"; shift ;;
        --user)     USER_INSTALL=true; shift ;;
        --prefix)   PREFIX="$2"; shift 2 ;;
        --no-prompt) NO_PROMPT=true; shift ;;
        --help)     show_help ;;
        *)          echo "Unknown option: $1"; show_help ;;
    esac
done

# ── Utility Functions ──

log_info()  { echo -e "${BLUE}[INFO]${NC} $*"; }
log_ok()    { echo -e "${GREEN}[OK]${NC} $*"; }
log_warn()  { echo -e "${YELLOW}[WARN]${NC} $*"; }
log_error() { echo -e "${RED}[ERROR]${NC} $*"; }

prompt_yes_no() {
    local msg="$1"
    local default="${2:-Y}"
    if [ "$NO_PROMPT" = true ]; then
        return 0
    fi
    if [ "$default" = "Y" ]; then
        read -r -p "$msg [Y/n] " response
        case "$response" in
            [nN]|[nN][oO]) return 1 ;;
            *) return 0 ;;
        esac
    else
        read -r -p "$msg [y/N] " response
        case "$response" in
            [yY]|[yY][eE][sS]) return 0 ;;
            *) return 1 ;;
        esac
    fi
}

check_command() {
    if ! command -v "$1" &>/dev/null; then
        log_warn "$1 not found. Install with: $2"
        return 1
    fi
    return 0
}

# ── Detect System ──

detect_distro() {
    if [ -f /etc/os-release ]; then
        . /etc/os-release
        echo "${ID:-linux}"
    elif [ -f /etc/debian_version ]; then
        echo "debian"
    elif [ -f /etc/redhat-release ]; then
        echo "rhel"
    elif command -v lsb_release &>/dev/null; then
        lsb_release -si | tr '[:upper:]' '[:lower:]'
    else
        echo "linux"
    fi
}

detect_package_manager() {
    if command -v apt-get &>/dev/null; then
        echo "apt"
    elif command -v dnf &>/dev/null; then
        echo "dnf"
    elif command -v yum &>/dev/null; then
        echo "yum"
    elif command -v zypper &>/dev/null; then
        echo "zypper"
    elif command -v pacman &>/dev/null; then
        echo "pacman"
    else
        echo "unknown"
    fi
}

# ── Check Dependencies ──

check_cups() {
    if command -v cupsctl &>/dev/null || command -v lpinfo &>/dev/null; then
        log_ok "CUPS is installed"
        return 0
    fi
    if [ -f /usr/sbin/cupsd ] || [ -f /usr/bin/cupsd ]; then
        log_ok "CUPS is installed"
        return 0
    fi
    log_warn "CUPS does not appear to be installed."
    log_warn "TrayPrint requires CUPS for printer discovery and printing."
    log_warn "Install with your package manager:"
    log_warn "  Debian/Ubuntu: sudo apt install cups"
    log_warn "  Fedora:        sudo dnf install cups"
    log_warn "  Arch:          sudo pacman -S cups"
    echo ""

    if prompt_yes_no "Continue anyway?" "N"; then
        return 0
    fi
    return 1
}

# ── AppImage Install ──

find_appimage() {
    # Look for AppImage in common locations
    local search_paths=(
        "$DIST_DIR/TrayPrint-${VERSION}-x86_64.AppImage"
        "$DIST_DIR/TrayPrint-x86_64.AppImage"
        "$DIST_DIR/TrayPrint-*.AppImage"
        "$PROJECT_ROOT/TrayPrint-*.AppImage"
    )
    for pattern in "${search_paths[@]}"; do
        for f in $pattern; do
            if [ -f "$f" ]; then
                echo "$f"
                return 0
            fi
        done
    done
    return 1
}

install_appimage() {
    local appimage_path
    appimage_path=$(find_appimage) || true

    if [ -z "$appimage_path" ] || [ ! -f "$appimage_path" ]; then
        log_info "AppImage not found. Building from source..."
        if [ -f "$SCRIPT_DIR/build_appimage.py" ]; then
            python3 "$SCRIPT_DIR/build_appimage.py" --version "$VERSION"
            appimage_path=$(find_appimage) || true
        else
            log_error "build_appimage.py not found. Cannot build AppImage."
            return 1
        fi
    fi

    if [ -z "$appimage_path" ] || [ ! -f "$appimage_path" ]; then
        log_error "Could not find or build AppImage."
        return 1
    fi

    log_info "Installing AppImage..."

    if [ "$USER_INSTALL" = true ]; then
        # User install
        local install_dir="${PREFIX:-$HOME/.local/bin}"
        local apps_dir="${HOME}/.local/share/applications"
        local icons_dir="${HOME}/.local/share/icons/hicolor/256x256/apps"

        mkdir -p "$install_dir" "$apps_dir" "$icons_dir"

        # Install AppImage
        local dest="$install_dir/${APP_NAME}.AppImage"
        cp "$appimage_path" "$dest"
        chmod +x "$dest"

        # Install .desktop file (pointing to AppImage)
        sed "s|Exec=trayprint|Exec=$dest|" "$SCRIPT_DIR/trayprint.desktop" > "$apps_dir/trayprint.desktop"

        # Install icon
        cp "$SCRIPT_DIR/trayprint.png" "$icons_dir/trayprint.png"

        # Update desktop database
        update-desktop-database "$apps_dir" 2>/dev/null || true

        log_ok "TrayPrint AppImage installed to $dest"
        log_info "Run with: $dest"

    else
        # System install
        local install_dir="${PREFIX:-/opt/trayprint}"

        if [ "$(id -u)" -ne 0 ]; then
            log_warn "System-wide install requires root privileges."
            if ! prompt_yes_no "Use sudo?" "Y"; then
                log_info "Falling back to user install."
                USER_INSTALL=true
                install_appimage
                return $?
            fi
        fi

        sudo mkdir -p "$install_dir"
        local dest="$install_dir/TrayPrint.AppImage"
        sudo cp "$appimage_path" "$dest"
        sudo chmod +x "$dest"

        # Create symlink in /usr/local/bin
        sudo ln -sf "$dest" "/usr/local/bin/trayprint"

        # Install .desktop file (system-wide)
        sudo mkdir -p /usr/share/applications /usr/share/icons/hicolor/256x256/apps
        sudo sed "s|Exec=trayprint|Exec=$dest|" "$SCRIPT_DIR/trayprint.desktop" > /tmp/trayprint.desktop
        sudo mv /tmp/trayprint.desktop /usr/share/applications/trayprint.desktop
        sudo cp "$SCRIPT_DIR/trayprint.png" /usr/share/icons/hicolor/256x256/apps/trayprint.png

        # Install systemd user service
        sudo mkdir -p /usr/lib/systemd/user
        sudo cp "$SCRIPT_DIR/trayprint.service" /usr/lib/systemd/user/trayprint.service

        # Update desktop database
        update-desktop-database 2>/dev/null || true
        gtk-update-icon-cache -f /usr/share/icons/hicolor 2>/dev/null || true

        # Enable systemd user service globally
        systemctl --global enable trayprint.service 2>/dev/null || true

        log_ok "TrayPrint AppImage installed to $dest"
        log_ok "Symlink: /usr/local/bin/trayprint → $dest"
    fi

    return 0
}

# ── .deb Install ──

install_deb() {
    local deb_path="$DIST_DIR/trayprint_${VERSION}-1_amd64.deb"

    if [ ! -f "$deb_path" ]; then
        log_info ".deb not found. Building from source..."
        if [ -f "$SCRIPT_DIR/build_deb.sh" ]; then
            bash "$SCRIPT_DIR/build_deb.sh" --version "$VERSION"
        else
            log_error "build_deb.sh not found."
            return 1
        fi
    fi

    if [ ! -f "$deb_path" ]; then
        log_error "Could not find or build .deb package."
        return 1
    fi

    log_info "Installing .deb package..."
    log_info "Package: $deb_path"

    if [ "$(id -u)" -ne 0 ]; then
        log_warn "This operation requires root privileges."
        echo "Running: sudo dpkg -i $deb_path"
    fi

    sudo dpkg -i "$deb_path" || {
        log_warn "dpkg install had errors. Attempting to fix dependencies..."
        sudo apt-get install -f -y
    }

    log_ok "TrayPrint .deb package installed."
    return 0
}

# ── .rpm Install ──

install_rpm() {
    local rpm_path="$DIST_DIR/trayprint-${VERSION}-1.x86_64.rpm"

    if [ ! -f "$rpm_path" ]; then
        log_info ".rpm not found. Building from source..."
        if [ -f "$SCRIPT_DIR/build_rpm.sh" ]; then
            bash "$SCRIPT_DIR/build_rpm.sh" --version "$VERSION"
        else
            log_error "build_rpm.sh not found."
            return 1
        fi
    fi

    if [ ! -f "$rpm_path" ]; then
        log_error "Could not find or build .rpm package."
        return 1
    fi

    log_info "Installing .rpm package..."
    log_info "Package: $rpm_path"

    if [ "$(id -u)" -ne 0 ]; then
        log_warn "This operation requires root privileges."
    fi

    detect_package_manager
    local pm=$(detect_package_manager)

    case "$pm" in
        dnf)
            sudo dnf install -y "$rpm_path"
            ;;
        yum)
            sudo yum install -y "$rpm_path"
            ;;
        zypper)
            sudo zypper install -y "$rpm_path"
            ;;
        *)
            sudo rpm -ivh "$rpm_path"
            ;;
    esac

    log_ok "TrayPrint .rpm package installed."
    return 0
}

# ── Auto-detect Best Format ──

auto_detect() {
    local distro
    local pm
    distro=$(detect_distro)
    pm=$(detect_package_manager)

    log_info "Detected distribution: $distro"
    log_info "Package manager: $pm"

    case "$distro" in
        ubuntu|debian|linuxmint|pop|kali|elementary|zorin)
            log_info "Recommended format: .deb"
            if prompt_yes_no "Install via .deb package?" "Y"; then
                install_deb && return 0
                log_warn ".deb install failed, falling back to AppImage..."
            fi
            ;;
        fedora|rhel|centos|rocky|alma)
            log_info "Recommended format: .rpm"
            if prompt_yes_no "Install via .rpm package?" "Y"; then
                install_rpm && return 0
                log_warn ".rpm install failed, falling back to AppImage..."
            fi
            ;;
        *)
            log_info "Recommended format: AppImage (portable)"
            ;;
    esac

    # Fallback to AppImage
    log_info "Attempting AppImage install..."
    install_appimage && return 0

    log_error "All installation methods failed."
    return 1
}

# ── Post-installation Summary ──

print_summary() {
    echo ""
    echo "================================================"
    echo -e "${GREEN}  TrayPrint v${VERSION} installed successfully!${NC}"
    echo "================================================"
    echo ""
    echo "  GUI mode:"
    echo "    Run 'trayprint' from the terminal or application menu"
    echo ""
    echo "  Headless service mode:"
    echo "    systemctl --user start trayprint.service"
    echo "    systemctl --user enable trayprint.service"
    echo ""
    echo "  Command-line options:"
    echo "    trayprint --help            Show all options"
    echo "    trayprint --silent          Run without GUI"
    echo "    trayprint --version         Show version"
    echo "    trayprint --hub-url URL     Configure hub URL"
    echo "    trayprint --agent-key KEY   Configure agent key"
    echo ""
    echo "  Configuration: /etc/trayprint/config.json"
    echo "  Logs:          ~/.local/share/trayprint/logs/"
    echo ""
    echo "  Prerequisites: CUPS must be installed and running"
    echo "  Check:         systemctl status cups"
    echo ""
}

# ── Main ──

main() {
    echo ""
    echo "================================================"
    echo "  TrayPrint v${VERSION} — Linux Installer"
    echo "================================================"
    echo ""

    # Check CUPS
    check_cups || {
        log_error "CUPS is required but not found. Aborting."
        exit 1
    }

    # Install
    case "$MODE" in
        appimage)
            install_appimage
            ;;
        deb)
            install_deb
            ;;
        rpm)
            install_rpm
            ;;
        auto)
            auto_detect
            ;;
        *)
            log_error "Unknown install mode: $MODE"
            exit 1
            ;;
    esac

    if [ $? -eq 0 ]; then
        print_summary
    else
        log_error "Installation failed."
        exit 1
    fi
}

main
