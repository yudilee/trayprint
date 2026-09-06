# TrayPrint — RPM Spec File
#
# Build with:
#   rpmbuild -ba installer/trayprint.spec
#
# Or use the build_rpm.sh wrapper which handles PyInstaller and paths.

%global _appname trayprint

# Pure-python app: no compiled artifacts -> disable debuginfo/debugsource subpackages
%global debug_package %{nil}

%global _version 3.0.0
%global _release 1

Name:          %{_appname}
Version:       %{_version}
Release:       %{_release}%{?dist}
Summary:       Local Print Agent for Print Hub

License:       MIT
URL:           https://print-hub.example.com
Source0:       %{name}-%{version}.tar.gz

BuildRequires: python3-devel
BuildRequires: python3-pip
BuildRequires: desktop-file-utils

Requires:      cups
Requires:      python3-pyside6 >= 6.5
Requires:      python3-flask >= 3.0
Requires:      python3-requests
Requires:      python3-pillow
Requires:      python3-websockets >= 14.0
Requires:      python3-packaging

%description
TrayPrint is a system tray application that connects printers to your
enterprise print management system. It runs in the background, automatically
discovers printers via CUPS, and processes print jobs from Print Hub.

Features:
- System tray GUI with dark/light theme support
- Automatic printer discovery via CUPS
- Real-time job queue with WebSocket updates
- Per-printer configuration with capability discovery
- Headless server mode for service deployment
- Auto-update with SHA-256 verification

%prep
%setup -q -n %{_appname}-%{version}

%build
# No compilation needed — pure Python application
echo "TrayPrint %{version}-%{release}" > .build-version

%install
# Remove build artifacts
rm -rf %{buildroot}

# Application directory
install -d %{buildroot}%{_libdir}/%{_appname}

# Copy Python source files (dynamic — nama file bisa berubah antar versi)
for src in *.py config.json requirements.txt trayprint.ico; do
    [ -e "$src" ] && cp -pr "$src" %{buildroot}%{_libdir}/%{_appname}/
done

# Copy templates directory
cp -pr templates %{buildroot}%{_libdir}/%{_appname}/templates

# Remove __pycache__ and *.pyc
find %{buildroot}%{_libdir}/%{_appname} -name '__pycache__' -type d -exec rm -rf {} + 2>/dev/null || true
find %{buildroot}%{_libdir}/%{_appname} -name '*.pyc' -delete

# Launcher script
install -d %{buildroot}%{_bindir}
cat > %{buildroot}%{_bindir}/%{_appname} << 'LAUNCHER'
#!/bin/bash
SCRIPT_DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
INSTALL_DIR="/usr/lib/trayprint"
CONFIG_DIR="${TRAYPRINT_CONFIG_PATH:-/etc/trayprint}"

if [ -n "${TRAYPRINT_CONFIG_PATH:-}" ]; then
    exec python3 "$INSTALL_DIR/app.py" --config-path "$TRAYPRINT_CONFIG_PATH" "$@"
else
    exec python3 "$INSTALL_DIR/app.py" "$@"
fi
LAUNCHER
chmod 755 %{buildroot}%{_bindir}/%{_appname}

# Desktop entry
install -d %{buildroot}%{_datadir}/applications
install -m 644 installer/trayprint.desktop %{buildroot}%{_datadir}/applications/%{_appname}.desktop

# Icon
install -d %{buildroot}%{_datadir}/icons/hicolor/256x256/apps
install -m 644 installer/trayprint.png %{buildroot}%{_datadir}/icons/hicolor/256x256/apps/%{_appname}.png

# AppStream metadata
install -d %{buildroot}%{_datadir}/metainfo
install -m 644 installer/trayprint.appdata.xml %{buildroot}%{_datadir}/metainfo/%{_appname}.appdata.xml

# systemd user service
install -d %{buildroot}/usr/lib/systemd/user
install -m 644 installer/trayprint.service %{buildroot}/usr/lib/systemd/user/%{_appname}.service

# Default configuration
install -d %{buildroot}%{_sysconfdir}/%{_appname}
install -m 644 config.json %{buildroot}%{_sysconfdir}/%{_appname}/config.json

# Documentation
install -d %{buildroot}%{_docdir}/%{_appname}
install -m 644 README.md %{buildroot}%{_docdir}/%{_appname}/README.md
install -m 644 requirements.txt %{buildroot}%{_docdir}/%{_appname}/requirements.txt

# License (opsional — repo boleh belum punya file LICENSE)
install -d %{buildroot}%{_licensedir}/%{_appname}
[ -f LICENSE ] && install -m 644 LICENSE %{buildroot}%{_licensedir}/%{_appname}/LICENSE || true

%check
# Validate desktop file
desktop-file-validate %{buildroot}%{_datadir}/applications/%{_appname}.desktop

%post
# Update desktop database
%{_bindir}/update-desktop-database &>/dev/null || :

# Update icon cache
%{_bindir}/gtk-update-icon-cache %{_datadir}/icons/hicolor &>/dev/null || :

# Enable systemd user service globally
%{_bindir}/systemctl --global enable %{_appname}.service 2>/dev/null || :

%preun
if [ "$1" = "0" ]; then
    # Stop service for all logged-in users before removal
    for session_dir in /run/user/*; do
        uid="${session_dir##*/}"
        [ -n "$uid" ] && [ "$uid" != "0" ] || continue
        %{_bindir}/systemctl --machine="$uid@" --user stop %{_appname}.service 2>/dev/null || true
    done
    %{_bindir}/systemctl --global disable %{_appname}.service 2>/dev/null || true
fi

%postun
if [ "$1" = "0" ]; then
    # Clean up on removal
    %{_bindir}/update-desktop-database &>/dev/null || :
    %{_bindir}/gtk-update-icon-cache %{_datadir}/icons/hicolor &>/dev/null || :
fi

%files
%defattr(-,root,root,-)
%{_bindir}/%{_appname}
%{_libdir}/%{_appname}/
%{_datadir}/applications/%{_appname}.desktop
%{_datadir}/icons/hicolor/256x256/apps/%{_appname}.png
%{_datadir}/metainfo/%{_appname}.appdata.xml
/usr/lib/systemd/user/%{_appname}.service
%config(noreplace) %{_sysconfdir}/%{_appname}/config.json
%doc %{_docdir}/%{_appname}/

%changelog
* Thu May 29 2026 Print Hub Team <support@print-hub.example.com> - 3.0.0-1
- Initial RPM package for TrayPrint
- AppImage + DEB + RPM multi-format Linux support
