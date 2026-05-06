# TrayPrint 🖨️

> Cross-platform print agent for Print Hub — connects printers to your enterprise
> print management system. Runs as a system tray application on Windows, Linux, and macOS.

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python)](https://python.org)
[![PySide6](https://img.shields.io/badge/PySide6-6.5-41CD52?logo=qt)](https://www.qt.io/qt-for-python)
[![Flask](https://img.shields.io/badge/Flask-3.0-000000?logo=flask)](https://flask.palletsprojects.com)
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

---

## ✨ Features

### Core Functionality

- **System Tray App** — Runs in the background with easy access from the system tray/notification area
- **Automatic Printer Discovery** — Detects all installed printers on the system via platform-specific APIs
- **Real-Time Job Queue** — Fetches and processes jobs from Print Hub automatically with persistent SQLite history
- **Multi-Platform** — Windows (win32print/DEVMODE), Linux (CUPS), macOS (CUPS via `lp`)

### Printing Capabilities

- **Full Printer Control** — Tray source, color mode, print quality, scaling percentage, media type, collation, reverse order, copies, duplex
- **Finishing Options** — Stapling, hole punch, booklet, folding, binding (CUPS printers)
- **Multiple Renderers** — CUPS `lp` (Linux/macOS), GDI `StretchDIBits` (Windows), SumatraPDF fallback (Windows for unsupported formats)
- **High Resolution** — 720 DPI rendering with PyMuPDF for quality PDF output on Windows
- **Per-Job Options** — Each job carries its own print options, merged with per-printer saved defaults

### Printer Management

- **Per-Printer Configuration** — Configure defaults for each printer independently, saved in [`config.json`](config.json) as `printer_configs`
- **Capability Discovery** — Automatically detects supported trays, resolutions, media sizes, color modes, and duplex via [`capabilities.py`](capabilities.py) using win32print.DeviceCapabilities (Windows) or `lpoptions -l` (CUPS)
- **Refresh Capabilities Button** — Per-printer "Refresh Capabilities" button in the settings dialog that re-queries the print subsystem for the latest supported options and updates dropdowns in real time
- **Thread-Safe UI Update** — Capability refresh runs in a background thread and delivers results to the UI via a polling timer (avoids PySide6 `QTimer.singleShot` cross-thread issues)
- **Configuration UI** — Visual settings dialog ([`ui_settings.py`](ui_settings.py)) with per-printer tabbed interface
- **Batch Apply** — Apply settings to all printers at once via "Apply to All" button
- **Config Merge** — Runtime merge of saved configs with request-level options via [`merge_printer_config()`](server.py:36)

### Monitoring & Self-Healing

- **Spooler Watchdog** — Monitors print spooler health; waits for spooler to pick up jobs and cleans up temp files after completion
- **Diagnostics Dialog** — Built-in system diagnostics ([`diagnostics_dialog.py`](diagnostics_dialog.py)) showing app version, Python version, platform, config status, hub connection, printer count, queue length, uptime, and watchdog status — with one-click "Copy All" for support
- **Connection Status** — Live hub connection indicator ("Connected" / "Disconnected / "No Hub" / "Unknown") in tray tooltip and menu
- **Print Queue Viewer** — Real-time queue dialog ([`queue_dialog.py`](queue_dialog.py)) with per-job status display and cancel capability
- **Notification History** — Persistent notification log ([`notification_history.py`](notification_history.py)) with configurable max length

### Hub Integration

- **WebSocket Client** — Real-time queue updates via Laravel Reverb (Pusher protocol) in [`websocket_client.py`](websocket_client.py)
- **Status Reporting** — Periodic health and printer status reports to the hub via configurable sync interval (default: 60s)
- **Job Acknowledgment** — Proper job lifecycle: `picked` → `processing` → `completed`/`failed` with status reporting
- **Approval Awareness** — Skips pending approval jobs; only picks jobs with `approved` status
- **Graceful Fallback** — Falls back to polling if WebSocket connection fails

### User Experience

- **Dark/Light Theme** — Built-in theme switcher ([`theme.py`](theme.py)) with WCAG AA compliant contrast ratios and comprehensive Qt stylesheet
- **Auto-Start Option** — Register to start with the operating system via [`autostart.py`](autostart.py)
- **Auto-Update** — Checks for new versions from Print Hub, downloads with SHA-256 verification, and installs updates with rollback on failure ([`updater.py`](updater.py))
- **Minimize to Tray** — Closing the window minimizes to tray instead of quitting; configurable behavior
- **Rich Tooltips** — System tray tooltip shows app version, hub status, printer count, queue depth, and live printing indicator
- **Printing Indicator** — Green dot overlay on tray icon when actively printing, updated every 2 seconds
- **Recent Jobs Menu** — Quick access to last 10 recent jobs from the tray context menu
- **View Logs** — One-click access to log file via system default viewer

---

## 📋 Requirements

| OS | Requirements |
|----|-------------|
| **Windows** | Windows 10/11 64-bit, Python 3.10+ |
| **Linux** | CUPS installed, Python 3.10+, PySide6 (system package recommended) |
| **macOS** | macOS 12+ (Monterey+), Python 3.10+, PySide6 (via Homebrew), CUPS (built-in) |

---

## 🚀 Quick Start

### Development Mode

```bash
# Clone the repository
git clone https://github.com/your-org/trayprint.git
cd trayprint

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Production dependencies include:
#   PySide6>=6.5.0    — Qt GUI framework
#   Flask>=3.0.0      — REST API server
#   requests>=2.31.0  — HTTP client for hub communication
#   PyMuPDF>=1.23.0   — PDF rendering (Windows only)
#   pywin32>=306      — Windows print API (Windows only)
#   Pillow>=10.2.0    — Image processing
#   pystray>=0.19.5   — Cross-platform tray (functionality reference)
#
# Optional:
#   websockets>=12.0  — WebSocket client for real-time updates
#   packaging>=23.0   — Semantic version comparison for updater

# Copy and edit config
cp config.json config.json  # Edit hub_url and agent_key

# Run the app
python app.py
```

### Production (Packaged Build)

```bash
# Build with PyInstaller
pip install pyinstaller
python build.py

# The built executable will be in dist/TrayPrint/
```

### MSI Installer (Windows)

```bash
# Prerequisites: WiX Toolset v3.x
cd installer
build_msi.bat
# Output: installer/TrayPrint.msi
```

---

## ⚙️ Configuration

Edit [`config.json`](config.json):

```json
{
  "port": 49211,
  "hub_url": "http://localhost:8000",
  "agent_key": "your-print-hub-agent-key",
  "autostart": false,
  "sync_interval_seconds": 60,
  "max_job_history": 50,
  "version": "3.0.0",
  "printer_configs": {
    "Office Printer": {
      "tray_source": "Tray 1",
      "color_mode": "color",
      "print_quality": "600dpi",
      "scaling_percentage": 100,
      "media_type": "plain",
      "collate": true,
      "copies": 1,
      "duplex": "long-edge"
    }
  }
}
```

| Key | Default | Description |
|-----|---------|-------------|
| `port` | `49211` | Local Flask API server port |
| `hub_url` | — | Print Hub server URL |
| `agent_key` | — | API key for authenticating with Print Hub |
| `autostart` | `false` | Register agent to start with OS |
| `sync_interval_seconds` | `60` | Interval between hub sync cycles |
| `max_job_history` | `50` | Max recent jobs stored in [`jobs.db`](jobs.db) |
| `printer_configs` | `{}` | Per-printer default settings |

---

## 🖥️ Usage

### System Tray

After launching, TrayPrint appears in the system tray (notification area). Right-click for the context menu:

| Menu Item | Description |
|-----------|-------------|
| **TrayPrint vX.X.X** | App version header (non-interactive) |
| **Hub: Connected/Disconnected** | Current hub connection status (non-interactive) |
| **Printers: N found** | Printer discovery count (non-interactive) |
| **View Queue...** | Opens the print queue viewer dialog |
| **Diagnostics...** | Opens system diagnostics dialog |
| **Show Notifications (N)...** | Opens notification history |
| **Check for Updates...** | Manually check for new version from Print Hub |
| **Settings** | Opens configuration dialog |
| **Recent Jobs** | Submenu with last 10 jobs (success/failed/pending) |
| **View Logs** | Opens log file in system default viewer |
| **Auto-start on Login** | Toggle auto-start registration |
| **Restart App** | Restarts the application |
| **Exit** | Quits TrayPrint entirely |

### Settings Dialog

The settings dialog ([`ui_settings.py`](ui_settings.py)) has tabbed sections:

1. **Connection** — Hub URL, API key, test connection button
2. **Appearance** — Dark/Light theme toggle
3. **General** — Auto-start on login, auto-update preferences
4. **Printer Configuration** — Per-printer settings for each discovered printer with "Apply to All" button; each printer has a **"Refresh Capabilities"** button that re-queries the print subsystem for supported trays, media sizes, color modes, resolutions, and duplex modes, then updates the dropdown widgets accordingly
5. **Software Updates** — Current version display, check for updates, download & install

### Diagnostics

The diagnostics dialog ([`diagnostics_dialog.py`](diagnostics_dialog.py)) shows:

- App version, Python version, platform
- Configuration status (hub_url and agent_key configured or missing)
- Hub connection status
- Printer count and full printer list
- Queue length and agent uptime
- Watchdog status

---

## 🔌 API Endpoints

TrayPrint runs a local Flask HTTP server ([`server.py`](server.py)) for hub communication. Default port: `49211`.

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/status` | GET | Agent and queue status summary |
| `/api/printers` | GET | List discovered printers with details |
| `/api/profiles` | GET | Get available print profiles |
| `/api/queue` | GET | Get current job queue |
| `/api/queue/status` | GET | Queue status (queued, processing, completed, failed counts) |
| `/api/capabilities?printer=X` | GET | Capabilities for a specific printer |
| `/api/capabilities/all` | GET | Capabilities for all printers |
| `/api/diagnostics` | GET | Full diagnostics data (JSON) |
| `/api/watchdog/status` | GET | Watchdog health status |
| `/api/jobs/pick` | POST | Pick the next available approved job |
| `/api/job/<id>/start` | POST | Mark a job as processing |
| `/api/job/<id>/complete` | POST | Mark a job as completed successfully |
| `/api/job/<id>/fail` | POST | Mark a job as failed |
| `/api/job/<id>/cancel` | POST | Cancel a queued job |

### Job Lifecycle

```
pending (hub)
   │
   ▼
picked ──→ processing ──→ completed
   │                        │
   │                        ▼
   └──→ failed          success
```

---

## 🏗 Architecture

```
┌─────────────────────────────────────────────┐
│              TrayPrint Agent                 │
│                                             │
│  ┌──────────┐    ┌──────────────────────┐   │
│  │ app.py   │────│   Qt System Tray     │   │
│  │ (Main)   │    │   Application        │   │
│  └────┬─────┘    └──────────────────────┘   │
│       │                                      │
│  ┌────▼─────┐    ┌──────────────────────┐   │
│  │server.py │────│   Flask REST API     │   │
│  │ (Agent)  │    │   :49211             │   │
│  └────┬─────┘    └──────────────────────┘   │
│       │                                      │
│  ┌────▼─────┐    ┌──────────────────────┐   │
│  │printer.py│────│   Print Backend      │   │
│  │          │    │   CUPS/DEVMODE/GDI   │   │
│  └──────────┘    └──────────────────────┘   │
│                                             │
│  ┌──────────────┐  ┌──────────────────────┐ │
│  │ capabilities │  │   Printer Capability  │ │
│  │ .py          │  │   Discovery          │ │
│  ├──────────────┤  ├──────────────────────┤ │
│  │ updater.py   │  │   Auto-Update        │ │
│  │              │  │   (SHA-256 verify)   │ │
│  ├──────────────┤  ├──────────────────────┤ │
│  │ theme.py     │  │   Dark/Light Theme   │ │
│  │              │  │   (WCAG AA)          │ │
│  ├──────────────┤  ├──────────────────────┤ │
│  │ autostart.py │  │   OS Auto-Start     │ │
│  └──────────────┘  └──────────────────────┘ │
│                                             │
│  ┌──────────────┐  ┌──────────────────────┐ │
│  │ ui_settings  │  │   Settings Dialog    │ │
│  │ .py          │  │   (Tabbed UI)        │ │
│  ├──────────────┤  ├──────────────────────┤ │
│  │ queue_dialog │  │   Queue Viewer       │ │
│  │ .py          │  │                      │ │
│  ├──────────────┤  ├──────────────────────┤ │
│  │ diagnostics_ │  │   Diagnostics        │ │
│  │ dialog.py    │  │   Dialog             │ │
│  ├──────────────┤  ├──────────────────────┤ │
│  │ notification │  │   Notification       │ │
│  │ _history.py  │  │   History            │ │
│  ├──────────────┤  ├──────────────────────┤ │
│  │ websocket_   │  │   Reverb WebSocket   │ │
│  │ client.py    │  │   Client (Pusher)    │ │
│  ├──────────────┤  ├──────────────────────┤ │
│  │ platform_    │  │   macOS-Specific     │ │
│  │ darwin.py    │  │   Helpers            │ │
│  └──────────────┘  └──────────────────────┘ │
└─────────────────────────────────────────────┘
         │                    ▲
         ▼                    │
  ┌──────────────────────────────────┐
  │         Print Hub (Server)       │
  │  Job Queue, Profiles, Webhooks   │
  └──────────────────────────────────┘
```

### Key Modules

| Module | Responsibility |
|--------|---------------|
| [`app.py`](app.py) | Main entry point; initializes QApplication, tray icon, menu, dialogs, WebSocket client, update checker |
| [`server.py`](server.py) | Flask HTTP server; job queue (SQLite), hub sync loop, printer config merge, REST API endpoints |
| [`printer.py`](printer.py) | Print backend: printer discovery (win32print/CUPS), job printing (DEVMODE/GDI/lp), temp file management |
| [`capabilities.py`](capabilities.py) | Printer capability discovery (trays, resolutions, media sizes, color modes, duplex); uses `win32print.DeviceCapabilities` on Windows and `lpoptions -l` on CUPS (Linux/macOS) |
| [`ui_settings.py`](ui_settings.py) | Settings dialog with tabs for connection, appearance, general, printer configuration, software updates; includes per-printer "Refresh Capabilities" button with thread-safe UI update via polling timer |
| [`queue_dialog.py`](queue_dialog.py) | Print queue viewer with real-time status, job cancel button |
| [`diagnostics_dialog.py`](diagnostics_dialog.py) | System health overview with copy-to-clipboard |
| [`notification_history.py`](notification_history.py) | Persistent notification store and viewer dialog |
| [`websocket_client.py`](websocket_client.py) | Optional Reverb WebSocket client (Pusher protocol) with auto-reconnect |
| [`updater.py`](updater.py) | Automatic update checker with download, SHA-256 verification, and atomic install |
| [`theme.py`](theme.py) | Dark/light theme system with WCAG AA compliant Qt stylesheet generation |
| [`autostart.py`](autostart.py) | OS auto-start registration and management |
| [`platform_darwin.py`](platform_darwin.py) | macOS-specific printer discovery and printing via CUPS `lp` |
| [`build.py`](build.py) | PyInstaller build script for packaged executable |
| [`logger.py`](logger.py) | Logging configuration |

---

## 🧪 Testing

```bash
# Install test dependencies
pip install -r tests/requirements-test.txt

# Run all tests
pytest tests/ -v

# With coverage report
pytest tests/ --cov=. --cov-report=term-missing
```

### Test Coverage

| Test Suite | File | Tests | Scope |
|------------|------|-------|-------|
| Server | [`tests/test_server.py`](tests/test_server.py) | 19 | Flask endpoints, job lifecycle, queue management |
| Printer | [`tests/test_printer.py`](tests/test_printer.py) | 22 | Print option parsing, DEVMODE building, option validation |
| Capabilities | [`tests/test_capabilities.py`](tests/test_capabilities.py) | 13 | CUPS `lpoptions` parsing, capability discovery |
| Utilities | [`tests/test_utils.py`](tests/test_utils.py) | 11 | Config validation, path utilities, helpers |

---

## 🛠️ Development

### Project Structure

```
trayprint/
├── app.py                   # Main entry point / Qt system tray app
├── server.py                # Flask REST API server (1233 lines)
├── printer.py               # Print backend (CUPS/DEVMODE/GDI) (1352 lines)
├── capabilities.py          # Printer capability discovery (270 lines)
├── ui_settings.py           # Settings dialog (tabbed interface)
├── queue_dialog.py          # Queue viewer dialog
├── notification_history.py  # Notification history store and dialog
├── diagnostics_dialog.py    # Diagnostics dialog (190 lines)
├── websocket_client.py      # Reverb WebSocket client (Pusher protocol)
├── updater.py               # Auto-update checker and installer (283 lines)
├── theme.py                 # Dark/light theme with Qt stylesheets (674 lines)
├── autostart.py             # OS auto-start registration
├── platform_darwin.py       # macOS-specific helpers (175 lines)
├── build.py                 # PyInstaller build script
├── logger.py                # Logging configuration
├── path_utils.py            # File path resolution
├── config.json              # Agent configuration
├── jobs.db                  # SQLite job history database (auto-created)
│
├── installer/               # MSI installer for Windows
│   ├── product.wxs          # WiX product definition
│   ├── build_msi.bat        # MSI build script
│   └── README.md            # Installer documentation
│
├── templates/               # Flask HTML templates
│   └── settings.html        # Web-based settings (legacy/fallback)
│
└── tests/                   # Pytest test suite
    ├── conftest.py          # Test fixtures and configuration
    ├── requirements-test.txt # Test dependencies
    ├── test_server.py       # Server endpoint tests
    ├── test_printer.py      # Printer option tests
    ├── test_capabilities.py # Capability parsing tests
    └── test_utils.py        # Utility function tests
```

---

## 🔧 Troubleshooting

### Windows: PyMuPDF Import Error in Packaged Build

If using PyInstaller, ensure hidden imports are included. The [`build.py`](build.py) script handles this automatically, but if building manually:

```bash
pyinstaller --hidden-import=fitz --hidden-import=win32print app.py
```

### Windows: Printer Shows 0 Count

Ensure hub URL is configured and agent is registered. The cache populates after the first successful status report to the hub.

### Windows: SumatraPDF Not Found

Download [`SumatraPDF.exe`](https://www.sumatrapdfreader.org) and place it in the same directory as the executable or source files. Used as fallback renderer for non-PDF formats.

### WebSocket Not Connecting

The app falls back to polling if WebSocket connection fails. Check that:
- Laravel Reverb is running on the hub
- `reverb_host`, `reverb_port`, and `reverb_app_key` are configured
- The `websockets` Python package is installed (`pip install websockets`)

### macOS: PySide6 Installation

```bash
brew install pyside@6
```

### Linux: System Package for PySide6

```bash
# Ubuntu/Debian
sudo apt install python3-pyside6

# Fedora
sudo dnf install python3-pyside6
```

---

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## 📄 License

MIT License — see [LICENSE](LICENSE) for details.
