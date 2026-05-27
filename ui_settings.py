import os
import json
import threading
import requests
from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QLabel, QLineEdit, QPushButton, QSpinBox, QCheckBox, QComboBox,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView,
    QFormLayout, QGroupBox, QFrame, QListWidget, QSplitter,
    QScrollArea, QProgressBar, QTextEdit, QGridLayout
)
from PySide6.QtCore import Qt, QTimer, QSize
from PySide6.QtGui import QFont, QIcon, QColor, QPalette

from path_utils import get_root_dir, get_data_dir
from logger import get_logger
from theme import get_stylesheet, get_palette
import server
import autostart
import printer
import capabilities

log = get_logger()


# ─────────────────────────────────────────────
#  Printer Control Defaults
# ─────────────────────────────────────────────

PRINTER_CONTROL_FIELDS = {
    'tray_source': {
        'type': 'combobox',
        'label': 'Tray Source',
        'default': 'AutoSelect',
        'options': ["AutoSelect", "Tray1", "Tray2", "Tray3", "ManualFeed", "Envelope", "PaperCassette", "SmallFormat", "LargeCapacity"],
    },
    'color_mode': {
        'type': 'combobox',
        'label': 'Color Mode',
        'default': 'color',
        'options': ["color", "monochrome"],
    },
    'print_quality': {
        'type': 'combobox',
        'label': 'Print Quality',
        'default': 'normal',
        'options': ["draft", "low", "normal", "high"],
    },
    'scaling_percentage': {
        'type': 'spinbox',
        'label': 'Scaling',
        'default': 100,
        'range': (1, 200),
        'suffix': '%',
    },
    'media_type': {
        'type': 'combobox',
        'label': 'Media Type',
        'default': 'plain',
        'options': ["plain", "glossy", "transparency", "envelope", "labels", "recycled", "cardstock"],
    },
    'collate': {
        'type': 'checkbox',
        'label': 'Collate',
        'default': False,
    },
    'reverse_order': {
        'type': 'checkbox',
        'label': 'Reverse Order',
        'default': False,
    },
}

class SettingsWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("TrayPrint Settings")
        self.setMinimumSize(900, 650)
        self.resize(900, 650)
        
        # Callback for applying settings without restart (set by app.py)
        self.apply_callback = None
        # Optional update checker reference (set by app.py)
        self.update_checker = None
        
        self.config_path = os.path.join(get_data_dir(), 'config.json')
        self.config_data = {
            "port": 49211,
            "sync_interval_seconds": 60,
            "max_retries": 3,
            "retry_delay_seconds": 60,
            "hub_url": "",
            "agent_key": "",
            "printer_configs": {},
            "theme": "dark",
        }
        self.load_config()
        
        self._theme_id = self.config_data.get('theme', 'dark')
        self._theme = get_palette(self._theme_id)
        
        self._last_sync_time = None
        
        self.setup_ui()
        self.populate_data()
        
        # Auto-refresh jobs and status while settings is open
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.refresh_status)
        self.refresh_timer.start(5000)

        # Thread-safe capability refresh via polled result variable
        self._pending_capability_result = None
        self._capability_poll_timer = QTimer(self)
        self._capability_poll_timer.timeout.connect(self._process_pending_capability)
        self._capability_poll_timer.start(200)  # Check for background-thread results every 200ms

    def load_config(self):
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r') as f:
                    self.config_data.update(json.load(f))
        except Exception:
            pass

    def save_config(self):
        try:
            with open(self.config_path, 'w') as f:
                json.dump(self.config_data, f, indent=2)
            return True
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save config: {e}")
            return False

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(16, 16, 16, 16)
        
        # Tabs
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Setup individual tabs
        self.setup_connection_tab()
        self.setup_general_tab()
        self.setup_printers_tab()
        self.setup_updates_tab()
        self.setup_jobs_tab()
        
        # Bottom button row
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.addStretch()
        
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setObjectName("cancel_btn")
        self.btn_cancel.clicked.connect(self.reject)
        
        self.btn_save = QPushButton("Save && Restart")
        self.btn_save.setObjectName("save_btn")
        self.btn_save.setDefault(True)
        self.btn_save.clicked.connect(self.on_save_clicked)
        
        btn_layout.addWidget(self.btn_cancel)
        btn_layout.addWidget(self.btn_save)
        
        main_layout.addLayout(btn_layout)
        
        # ── Footer / Status Bar ──
        footer_frame = QFrame()
        footer_frame.setFrameShape(QFrame.StyledPanel)
        footer_frame.setObjectName("footer_frame")
        footer_layout = QHBoxLayout(footer_frame)
        footer_layout.setContentsMargins(12, 6, 12, 6)
        footer_layout.setSpacing(16)

        self.lbl_version = QLabel(f"v{server.APP_VERSION}")
        self.lbl_version.setObjectName("lbl_version")
        
        self.lbl_footer_status = QLabel("● Disconnected")
        self.lbl_footer_status.setObjectName("lbl_footer_status")
        
        self.lbl_last_sync = QLabel("Last sync: --")
        self.lbl_last_sync.setObjectName("lbl_last_sync")

        self.btn_theme = QPushButton()
        self.btn_theme.setObjectName("theme_btn")
        self.btn_theme.setToolTip("Toggle between Dark and Light theme")
        self.btn_theme.clicked.connect(self.toggle_theme)

        footer_layout.addWidget(self.lbl_version)
        footer_layout.addWidget(self.btn_theme)
        footer_layout.addStretch()
        footer_layout.addWidget(self.lbl_footer_status)
        footer_layout.addStretch()
        footer_layout.addWidget(self.lbl_last_sync)

        main_layout.addWidget(footer_frame)
        
        self.apply_styles()

    def _style_section_group(self, title):
        """Create a styled QGroupBox using current theme colors."""
        t = self._theme
        group = QGroupBox(title)
        group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: 700; font-size: 14px; color: {t['group_title']};
                border: 1px solid {t['group_border']}; border-radius: 8px;
                margin-top: 16px; padding: 18px 12px 12px 12px;
                background: {t['group_bg']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin; subcontrol-position: top left;
                padding: 4px 12px; background: {t['group_bg']};
                color: {t['group_title']};
                border: 1px solid {t['group_border']};
                border-radius: 4px;
            }}
        """)
        return group

    def _style_field_label(self, text):
        """Create a styled form label."""
        label = QLabel(text)
        label.setStyleSheet(
            "font-weight: 600; font-size: 13px; color: %s;" % self._theme['text_secondary']
        )
        return label

    # ── Connection Tab ──

    def setup_connection_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        layout.setContentsMargins(4, 4, 4, 4)

        # ── Server Configuration ──
        server_group = self._style_section_group("Server Configuration")
        server_form = QFormLayout(server_group)
        server_form.setSpacing(10)
        server_form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self.input_hub_url = QLineEdit()
        self.input_hub_url.setPlaceholderText("https://your-print-hub.com")
        self.input_hub_url.setMinimumWidth(300)
        
        self.input_agent_key = QLineEdit()
        self.input_agent_key.setEchoMode(QLineEdit.PasswordEchoOnEdit)
        self.input_agent_key.setPlaceholderText("Enter Agent Key")
        self.input_agent_key.setMinimumWidth(300)

        server_form.addRow(self._style_field_label("Hub URL:"), self.input_hub_url)
        server_form.addRow(self._style_field_label("Agent Key:"), self.input_agent_key)

        layout.addWidget(server_group)

        # ── Connection Status ──
        status_group = self._style_section_group("Connection Status")
        status_layout = QVBoxLayout(status_group)
        status_layout.setSpacing(10)

        # Hub URL display
        url_layout = QHBoxLayout()
        url_label = QLabel("Target:")
        url_label.setStyleSheet(
            "font-weight: 600; font-size: 12px; color: %s;" % self._theme['text_secondary']
        )
        self.lbl_hub_url_display = QLabel("Not configured")
        self.lbl_hub_url_display.setStyleSheet(
            "font-size: 12px; color: %s;" % self._theme['text_muted']
        )
        url_layout.addWidget(url_label)
        url_layout.addWidget(self.lbl_hub_url_display)
        url_layout.addStretch()
        status_layout.addLayout(url_layout)

        # Status indicator row
        indicator_layout = QHBoxLayout()
        self.lbl_status_dot = QLabel("●")
        self.lbl_status_dot.setStyleSheet("font-size: 26px; color: %s;" % self._theme['text_muted'])
        self.lbl_status_text = QLabel("Checking...")
        self.lbl_status_text.setStyleSheet(
            "font-size: 14px; color: %s; font-weight: 500;" % self._theme['text_secondary']
        )
        indicator_layout.addWidget(self.lbl_status_dot)
        indicator_layout.addSpacing(6)
        indicator_layout.addWidget(self.lbl_status_text)
        indicator_layout.addStretch()

        self.lbl_conn_status = QLabel("")
        self.lbl_conn_status.setStyleSheet(
            "color: %s; font-size: 12px;" % self._theme['text_muted']
        )

        status_layout.addLayout(indicator_layout)
        status_layout.addWidget(self.lbl_conn_status)

        # Test connection button
        test_btn_layout = QHBoxLayout()
        test_btn_layout.addStretch()
        self.btn_test_conn = QPushButton("Test Connection")
        self.btn_test_conn.setObjectName("test_btn")
        self.btn_test_conn.clicked.connect(self.on_test_connection)
        test_btn_layout.addWidget(self.btn_test_conn)
        status_layout.addLayout(test_btn_layout)

        layout.addWidget(status_group)
        layout.addStretch()
        self.tabs.addTab(tab, "Connection")

    # ── General Tab ──

    def setup_general_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        layout.setContentsMargins(4, 4, 4, 4)

        # ── Theme Selection ──
        theme_group = self._style_section_group("Appearance")
        theme_layout = QFormLayout(theme_group)
        theme_layout.setSpacing(10)
        theme_layout.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self.cmb_theme = QComboBox()
        self.cmb_theme.addItem("☀️  Light", "light")
        self.cmb_theme.addItem("🌙  Dark", "dark")
        idx = self.cmb_theme.findData(self._theme_id)
        if idx >= 0:
            self.cmb_theme.setCurrentIndex(idx)
        self.cmb_theme.currentIndexChanged.connect(self._on_theme_combo_changed)
        self.cmb_theme.setMinimumWidth(200)
        self.cmb_theme.setToolTip("Select UI color scheme (changes apply immediately)")

        theme_layout.addRow(self._style_field_label("Theme:"), self.cmb_theme)
        layout.addWidget(theme_group)

        # ── Application Behavior ──
        behavior_group = self._style_section_group("Application Behavior")
        behavior_layout = QVBoxLayout(behavior_group)
        behavior_layout.setSpacing(8)

        self.chk_autostart = QCheckBox("Start automatically on login")
        self.chk_autostart.setChecked(autostart.is_autostart_enabled())
        self.chk_autostart.toggled.connect(self.on_autostart_toggled)
        behavior_layout.addWidget(self.chk_autostart)

        # "Start TrayPrint when Windows starts" checkbox (persisted in config.json)
        self.chk_windows_startup = QCheckBox("Start TrayPrint when Windows starts")
        startup_enabled = self.config_data.get('windows_startup', False)
        self.chk_windows_startup.setChecked(startup_enabled)
        self.chk_windows_startup.toggled.connect(self._on_windows_startup_toggled)
        behavior_layout.addWidget(self.chk_windows_startup)

        layout.addWidget(behavior_group)

        # ── Local Service Settings ──
        service_group = self._style_section_group("Local Service Settings")
        service_form = QFormLayout(service_group)
        service_form.setSpacing(10)
        service_form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self.spin_port = QSpinBox()
        self.spin_port.setRange(1024, 65535)
        self.spin_port.setMinimumWidth(200)

        self.spin_interval = QSpinBox()
        self.spin_interval.setRange(5, 3600)
        self.spin_interval.setSuffix(" sec")
        self.spin_interval.setMinimumWidth(200)

        self.spin_retries = QSpinBox()
        self.spin_retries.setRange(0, 10)
        self.spin_retries.setMinimumWidth(200)

        self.spin_retry_delay = QSpinBox()
        self.spin_retry_delay.setRange(5, 3600)
        self.spin_retry_delay.setSuffix(" sec")
        self.spin_retry_delay.setMinimumWidth(200)

        service_form.addRow(self._style_field_label("Local API Port:"), self.spin_port)
        service_form.addRow(self._style_field_label("Sync Interval:"), self.spin_interval)
        service_form.addRow(self._style_field_label("Max Retries:"), self.spin_retries)
        service_form.addRow(self._style_field_label("Retry Delay:"), self.spin_retry_delay)

        layout.addWidget(service_group)
        layout.addStretch()
        self.tabs.addTab(tab, "General")

    # ── Updates Tab ──

    def setup_updates_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        layout.setContentsMargins(4, 4, 4, 4)

        # ── Software Updates ──
        updates_group = self._style_section_group("Software Updates")
        updates_layout = QVBoxLayout(updates_group)
        updates_layout.setSpacing(10)

        # Current version
        version_row = QHBoxLayout()
        version_row.addWidget(QLabel("Current Version:"))
        self.lbl_current_version = QLabel(f"v{server.APP_VERSION}")
        self.lbl_current_version.setStyleSheet(
            "font-weight: 700; font-size: 14px; color: %s;" % self._theme['accent']
        )
        version_row.addWidget(self.lbl_current_version)
        version_row.addStretch()
        updates_layout.addLayout(version_row)

        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("color: %s;" % self._theme['border'])
        updates_layout.addWidget(line)

        # Update status / info area
        self.lbl_update_status = QLabel("Update checker not initialized.")
        self.lbl_update_status.setWordWrap(True)
        self.lbl_update_status.setStyleSheet(
            "font-size: 13px; padding: 8px 0; color: %s;" % self._theme['text_secondary']
        )
        updates_layout.addWidget(self.lbl_update_status)

        # Release notes
        self.lbl_release_notes = QLabel("")
        self.lbl_release_notes.setWordWrap(True)
        self.lbl_release_notes.setStyleSheet(
            "font-size: 12px; color: %s; padding: 4px 12px; "
            "background: %s; border-radius: 6px; border: 1px solid %s;"
            % (self._theme['text_muted'], self._theme['bg_card'], self._theme['border'])
        )
        self.lbl_release_notes.setVisible(False)
        updates_layout.addWidget(self.lbl_release_notes)

        # Progress bar
        self.progress_update = QProgressBar()
        self.progress_update.setVisible(False)
        self.progress_update.setMinimum(0)
        self.progress_update.setMaximum(100)
        self.progress_update.setValue(0)
        updates_layout.addWidget(self.progress_update)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        self.btn_check_updates = QPushButton("Check for Updates")
        self.btn_check_updates.setObjectName("test_btn")
        self.btn_check_updates.clicked.connect(self.on_check_updates)
        btn_row.addWidget(self.btn_check_updates)

        self.btn_download_update = QPushButton("Download & Install")
        self.btn_download_update.setObjectName("save_btn")
        self.btn_download_update.setVisible(False)
        self.btn_download_update.clicked.connect(self.on_download_update)
        btn_row.addWidget(self.btn_download_update)

        btn_row.addStretch()
        updates_layout.addLayout(btn_row)

        updates_layout.addStretch()
        layout.addWidget(updates_group)
        layout.addStretch()
        self.tabs.addTab(tab, "Updates")

    def _refresh_updates_tab(self):
        """Refresh the updates tab UI based on update checker state."""
        if not self.update_checker:
            self.lbl_update_status.setText(
                "Update checker not configured.\nConfigure Hub URL and Agent Key in the Connection tab."
            )
            self.btn_check_updates.setEnabled(False)
            self.btn_download_update.setVisible(False)
            return

        self.btn_check_updates.setEnabled(True)
        checker = self.update_checker

        if checker.update_ready:
            self.lbl_update_status.setText(
                f"\u2b06 Update v{checker.latest_version} is available!"
            )
            self.lbl_update_status.setStyleSheet(
                "font-size: 13px; padding: 8px 0; color: %s; font-weight: 600;" % self._theme['success']
            )
            if checker.release_notes:
                self.lbl_release_notes.setText(f"Release notes:\n{checker.release_notes}")
                self.lbl_release_notes.setVisible(True)
            else:
                self.lbl_release_notes.setVisible(False)
            self.btn_download_update.setVisible(True)
        else:
            self.lbl_update_status.setText("You are running the latest version.")
            self.lbl_update_status.setStyleSheet(
                "font-size: 13px; padding: 8px 0; color: %s;" % self._theme['text_muted']
            )
            self.lbl_release_notes.setVisible(False)
            self.btn_download_update.setVisible(False)

    # ── Update check / download handlers ──

    def on_check_updates(self):
        """Trigger an immediate update check."""
        if not self.update_checker:
            return

        self.btn_check_updates.setEnabled(False)
        self.btn_check_updates.setText("Checking...")
        self.lbl_update_status.setText("Checking for updates...")
        QApplication.processEvents()

        def _do_check():
            try:
                self.update_checker.check_for_updates()
                # Schedule UI update on the main thread
                QTimer.singleShot(0, self._refresh_updates_tab)
                QTimer.singleShot(0, lambda: self.btn_check_updates.setText("Check for Updates"))
                QTimer.singleShot(0, lambda: self.btn_check_updates.setEnabled(True))
            except Exception as e:
                QTimer.singleShot(0, lambda: self.lbl_update_status.setText(f"Check failed: {e}"))
                QTimer.singleShot(0, lambda: self.btn_check_updates.setText("Check for Updates"))
                QTimer.singleShot(0, lambda: self.btn_check_updates.setEnabled(True))

        threading.Thread(target=_do_check, daemon=True).start()

    def on_download_update(self):
        """Download and install the update."""
        if not self.update_checker or not self.update_checker.update_ready:
            return

        self.btn_download_update.setEnabled(False)
        self.btn_download_update.setText("Downloading...")
        self.progress_update.setVisible(True)
        self.progress_update.setValue(0)
        QApplication.processEvents()

        def _progress(pct):
            QTimer.singleShot(0, lambda: self.progress_update.setValue(int(pct * 100)))

        def _do_download():
            try:
                path = self.update_checker.download_update(progress_callback=_progress)
                if path:
                    QTimer.singleShot(0, lambda: self.lbl_update_status.setText(
                        "Download complete! Installing..."))
                    QTimer.singleShot(0, lambda: self.progress_update.setValue(100))
                    # Install
                    success = self.update_checker.install_update(path)
                    if success:
                        QTimer.singleShot(0, lambda: self.lbl_update_status.setText(
                            "Update installed successfully. The application will restart."))
                    else:
                        QTimer.singleShot(0, lambda: self.lbl_update_status.setText(
                            "Installation failed. See log for details."))
                else:
                    QTimer.singleShot(0, lambda: self.lbl_update_status.setText(
                        "Download failed. See log for details."))
            except Exception as e:
                QTimer.singleShot(0, lambda: self.lbl_update_status.setText(
                    f"Update failed: {e}"))
            finally:
                QTimer.singleShot(0, lambda: self.btn_download_update.setEnabled(True))
                QTimer.singleShot(0, lambda: self.btn_download_update.setText("Download & Install"))
                QTimer.singleShot(0, lambda: self.progress_update.setVisible(False))

        threading.Thread(target=_do_download, daemon=True).start()

    # ─────────────────────────────────────────────
    #  Printer Config Tab (split-panel layout)
    # ─────────────────────────────────────────────

    def setup_printers_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        layout.setContentsMargins(4, 4, 4, 4)

        self.lbl_printer_count = QLabel("Found 0 printers on this system.")
        self.lbl_printer_count.setStyleSheet(
            "font-weight: 600; font-size: 13px; color: %s; margin-bottom: 4px;" % self._theme['text_secondary']
        )
        layout.addWidget(self.lbl_printer_count)

        # Splitter: left = printer list, right = config panel
        splitter = QSplitter(Qt.Horizontal)

        # ── Left panel: Printer List ──
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(6, 6, 4, 6)

        lbl_printers_title = QLabel("Available Printers")
        lbl_printers_title.setStyleSheet(
            "font-weight: 700; color: %s; padding: 6px 0; font-size: 13px;" % self._theme['text_primary']
        )
        left_layout.addWidget(lbl_printers_title)

        self.list_printers = QListWidget()
        self.list_printers.setAlternatingRowColors(True)
        self.list_printers.currentRowChanged.connect(self.on_printer_selected)
        left_layout.addWidget(self.list_printers)

        splitter.addWidget(left_panel)

        # ── Right panel: Config Panel ──
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 6, 6, 6)
        right_layout.setSpacing(8)

        self.lbl_selected_printer = QLabel("No printer selected")
        self.lbl_selected_printer.setStyleSheet(
            "font-weight: 700; color: %s; font-size: 15px; padding: 6px 0;" % self._theme['text_primary']
        )
        right_layout.addWidget(self.lbl_selected_printer)

        # Build control widgets based on PRINTER_CONTROL_FIELDS definition
        self.printer_control_widgets = {}
        for field_name, field_def in PRINTER_CONTROL_FIELDS.items():
            w = None
            if field_def['type'] == 'combobox':
                w = QComboBox()
                w.addItems(field_def['options'])
                w.setCurrentText(field_def['default'])
                w.setMinimumWidth(200)
            elif field_def['type'] == 'spinbox':
                w = QSpinBox()
                w.setRange(*field_def['range'])
                w.setValue(field_def['default'])
                if 'suffix' in field_def:
                    w.setSuffix(field_def['suffix'])
                w.setMinimumWidth(200)
            elif field_def['type'] == 'checkbox':
                w = QCheckBox()
                w.setChecked(field_def['default'])
            if w is not None:
                self.printer_control_widgets[field_name] = w

        # ── Scroll area for config fields ──
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        config_container = QWidget()
        config_container_layout = QVBoxLayout(config_container)
        config_container_layout.setSpacing(6)
        config_container_layout.setContentsMargins(0, 0, 0, 0)

        # Group: Color & Quality
        color_group = self._style_section_group("Color & Quality")
        color_form = QFormLayout(color_group)
        color_form.setSpacing(8)
        color_form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)

        for field_name in ['color_mode', 'print_quality', 'media_type']:
            w = self.printer_control_widgets.get(field_name)
            if w:
                color_form.addRow(
                    self._style_field_label(f"{PRINTER_CONTROL_FIELDS[field_name]['label']}:"),
                    w
                )
        config_container_layout.addWidget(color_group)

        # Group: Layout & Output
        layout_group = self._style_section_group("Layout & Output")
        layout_form = QFormLayout(layout_group)
        layout_form.setSpacing(8)
        layout_form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)

        for field_name in ['tray_source', 'collate', 'reverse_order', 'scaling_percentage']:
            w = self.printer_control_widgets.get(field_name)
            if w:
                layout_form.addRow(
                    self._style_field_label(f"{PRINTER_CONTROL_FIELDS[field_name]['label']}:"),
                    w
                )
        config_container_layout.addWidget(layout_group)

        config_container_layout.addStretch()
        scroll.setWidget(config_container)
        right_layout.addWidget(scroll, stretch=1)

        # Buttons row
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        self.btn_refresh_caps = QPushButton("Refresh Capabilities")
        self.btn_refresh_caps.setObjectName("printer_config_refresh_caps")
        self.btn_refresh_caps.clicked.connect(self.on_refresh_capabilities)
        self.btn_refresh_caps.setEnabled(False)
        self.btn_refresh_caps.setToolTip("Query printer for supported trays, resolutions, and media sizes")

        self.btn_reset_printer_config = QPushButton("Reset to Defaults")
        self.btn_reset_printer_config.setObjectName("printer_config_reset")
        self.btn_reset_printer_config.clicked.connect(self.on_reset_printer_config)
        self.btn_reset_printer_config.setEnabled(False)

        self.btn_save_printer_config = QPushButton("Save Config")
        self.btn_save_printer_config.setObjectName("printer_config_save")
        self.btn_save_printer_config.clicked.connect(self.on_save_printer_config)
        self.btn_save_printer_config.setEnabled(False)

        self.btn_view_caps = QPushButton("View Capabilities...")
        self.btn_view_caps.setObjectName("printer_config_view_caps")
        self.btn_view_caps.clicked.connect(self.on_view_capabilities)
        self.btn_view_caps.setEnabled(False)
        self.btn_view_caps.setToolTip("Open detailed capabilities dialog for the selected printer")

        btn_row.addWidget(self.btn_refresh_caps)
        btn_row.addWidget(self.btn_reset_printer_config)
        btn_row.addWidget(self.btn_view_caps)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_save_printer_config)
        right_layout.addLayout(btn_row)

        right_layout.addStretch()
        splitter.addWidget(right_panel)

        # Set initial splitter proportions (35% left, 65% right)
        splitter.setSizes([280, 520])
        layout.addWidget(splitter)
        self.tabs.addTab(tab, "Printers")

    # ─────────────────────────────────────────────
    #  Printer Config: load from config_data
    # ─────────────────────────────────────────────

    def get_printer_config(self, printer_name):
        """Return the saved config dict for a given printer, or None."""
        configs = self.config_data.get('printer_configs', {})
        return configs.get(printer_name)

    def set_printer_config(self, printer_name, config_dict):
        """Store a config dict for a given printer."""
        if 'printer_configs' not in self.config_data:
            self.config_data['printer_configs'] = {}
        self.config_data['printer_configs'][printer_name] = config_dict

    # ─────────────────────────────────────────────
    #  Printer Config: populate UI from saved data
    # ─────────────────────────────────────────────

    def populate_printer_config_ui(self, printer_name):
        """Load saved config for the given printer into the control widgets."""
        config = self.get_printer_config(printer_name)
        for field_name, field_def in PRINTER_CONTROL_FIELDS.items():
            w = self.printer_control_widgets.get(field_name)
            if w is None:
                continue
            saved_value = config.get(field_name, field_def['default']) if config else field_def['default']
            if field_def['type'] == 'combobox':
                idx = w.findText(str(saved_value))
                if idx >= 0:
                    w.setCurrentIndex(idx)
                else:
                    w.setCurrentText(field_def['default'])
            elif field_def['type'] == 'spinbox':
                w.setValue(int(saved_value))
            elif field_def['type'] == 'checkbox':
                w.setChecked(bool(saved_value))

    # ─────────────────────────────────────────────
    #  Printer Config: read current UI values
    # ─────────────────────────────────────────────

    def read_printer_config_from_ui(self):
        """Read the current values from the control widgets into a dict."""
        config = {}
        for field_name, field_def in PRINTER_CONTROL_FIELDS.items():
            w = self.printer_control_widgets.get(field_name)
            if w is None:
                config[field_name] = field_def['default']
                continue
            if field_def['type'] == 'combobox':
                config[field_name] = w.currentText()
            elif field_def['type'] == 'spinbox':
                config[field_name] = w.value()
            elif field_def['type'] == 'checkbox':
                config[field_name] = w.isChecked()
        return config

    # ─────────────────────────────────────────────
    #  Printer Config: event handlers
    # ─────────────────────────────────────────────

    def on_printer_selected(self, row):
        """When a printer is clicked in the list, show its config panel."""
        if row < 0:
            self.lbl_selected_printer.setText("No printer selected")
            self.btn_save_printer_config.setEnabled(False)
            self.btn_reset_printer_config.setEnabled(False)
            self.btn_refresh_caps.setEnabled(False)
            self.btn_view_caps.setEnabled(False)
            return

        item = self.list_printers.item(row)
        if not item:
            return

        printer_name = item.text()
        self.lbl_selected_printer.setText(printer_name)
        self.btn_save_printer_config.setEnabled(True)
        self.btn_reset_printer_config.setEnabled(True)
        self.btn_refresh_caps.setEnabled(True)
        self.btn_view_caps.setEnabled(True)
        self.populate_printer_config_ui(printer_name)

    def on_refresh_capabilities(self):
        """Fetch capabilities for the selected printer and update dropdowns."""
        printer_name = self.lbl_selected_printer.text()
        if not printer_name or printer_name == "No printer selected":
            log.debug("on_refresh_capabilities: no printer selected, skipping")
            return

        log.info("on_refresh_capabilities: starting refresh for '%s'", printer_name)
        self.btn_refresh_caps.setEnabled(False)
        self.btn_refresh_caps.setText("Refreshing...")

        def _do_refresh():
            """Run capability discovery in a background thread."""
            try:
                import capabilities as caps_mod
                caps = caps_mod.discover_capabilities(printer_name)
                # Store result for main-thread poll timer to pick up
                # (Python GIL makes single-variable assignment atomic)
                self._pending_capability_result = (printer_name, caps)
            except Exception as e:
                log.error("Capability refresh EXCEPTION for '%s': %s",
                          printer_name, e, exc_info=True)
                # Store error result so polling handler can reset UI
                self._pending_capability_result = (printer_name, {'error': str(e)})

        import threading
        threading.Thread(target=_do_refresh, daemon=True).start()

    def _process_pending_capability(self):
        """Poll-timer callback (main thread): pick up background-thread results."""
        result = self._pending_capability_result
        if result is not None:
            self._pending_capability_result = None
            printer_name, caps = result
            self._apply_capabilities_to_dropdowns(printer_name, caps)

    def _apply_capabilities_to_dropdowns(self, printer_name, capabilities):
        """Update PRINTER_CONTROL_FIELDS dropdown options from discovered capabilities."""
        log.debug("_apply_capabilities_to_dropdowns ENTER: printer='%s', caps_keys=%s, caps_error=%s",
                  printer_name, list(capabilities.keys()) if capabilities else 'NONE',
                  capabilities.get('error', 'None') if capabilities else 'N/A')

        # Guard: capabilities must be a dict
        if not isinstance(capabilities, dict):
            log.error("_apply_capabilities_to_dropdowns: capabilities is not a dict, got %s", type(capabilities).__name__)
            self.btn_refresh_caps.setText("Refresh Capabilities")
            self.btn_refresh_caps.setEnabled(True)
            return

        try:
            if 'error' in capabilities and capabilities['error']:
                log.warning("_apply_capabilities_to_dropdowns: capabilities contain error: %s", capabilities['error'])
                self.btn_refresh_caps.setText("Refresh Capabilities")
                self.btn_refresh_caps.setEnabled(True)
                QMessageBox.warning(self, "Capability Error", capabilities['error'])
                return

            updated_fields = 0

            # ── Trays ──
            trays = capabilities.get('trays', [])
            log.debug("_apply_capabilities_to_dropdowns: trays=%s", trays)
            if trays and 'tray_source' in self.printer_control_widgets:
                w = self.printer_control_widgets['tray_source']
                current_val = w.currentText()
                w.clear()
                w.addItems(trays)
                idx = w.findText(current_val)
                if idx >= 0:
                    w.setCurrentIndex(idx)
                updated_fields += 1
                log.debug("_apply_capabilities_to_dropdowns: updated tray_source with %d items", len(trays))

            # ── Media Sizes ──
            media_sizes = capabilities.get('media_sizes', [])
            if media_sizes:
                log.info("Discovered %d media sizes for '%s'", len(media_sizes), printer_name)

            # ── Resolutions ──
            resolutions = capabilities.get('resolutions', [])
            log.debug("_apply_capabilities_to_dropdowns: resolutions=%s", resolutions)
            if resolutions and 'print_quality' in self.printer_control_widgets:
                self._cached_resolutions = resolutions
                updated_fields += 1
                log.debug("_apply_capabilities_to_dropdowns: cached %d resolutions", len(resolutions))

            # ── Color Modes ──
            color_modes = capabilities.get('color_modes', [])
            log.debug("_apply_capabilities_to_dropdowns: color_modes=%s", color_modes)
            if color_modes and 'color_mode' in self.printer_control_widgets:
                w = self.printer_control_widgets['color_mode']
                current_val = w.currentText()
                w.clear()
                w.addItems(color_modes)
                idx = w.findText(current_val)
                if idx >= 0:
                    w.setCurrentIndex(idx)
                updated_fields += 1
                log.debug("_apply_capabilities_to_dropdowns: updated color_mode with %d items", len(color_modes))

            # ── Duplex ──
            duplex = capabilities.get('duplex', [])
            if duplex:
                log.info("Discovered %d duplex modes for '%s'", len(duplex), printer_name)
                self._cached_duplex = duplex

            log.debug("_apply_capabilities_to_dropdowns: resetting button (updated_fields=%d)", updated_fields)
            self.btn_refresh_caps.setText("Refresh Capabilities")
            self.btn_refresh_caps.setEnabled(True)

            if updated_fields > 0:
                log.info("Updated capability dropdowns for '%s' (%d fields)", printer_name, updated_fields)
                QMessageBox.information(
                    self, "Capabilities Updated",
                    f"Refreshed {updated_fields} capability fields for '{printer_name}'."
                )
            else:
                log.info("No capability fields to update for '%s'", printer_name)
                QMessageBox.information(
                    self, "Capabilities",
                    f"No capability fields to update for '{printer_name}'."
                )
            log.debug("_apply_capabilities_to_dropdowns EXIT: success")

        except Exception as e:
            log.error("_apply_capabilities_to_dropdowns UNHANDLED EXCEPTION: %s", e, exc_info=True)
            # Reset button regardless so it doesn't stay stuck
            self.btn_refresh_caps.setText("Refresh Capabilities")
            self.btn_refresh_caps.setEnabled(True)

    def on_save_printer_config(self):
        """Save the current control values for the selected printer."""
        printer_name = self.lbl_selected_printer.text()
        if not printer_name or printer_name == "No printer selected":
            return

        config = self.read_printer_config_from_ui()
        self.set_printer_config(printer_name, config)
        QMessageBox.information(self, "Saved", f"Printer config saved for '{printer_name}'.")

    def on_reset_printer_config(self):
        """Reset the current printer's config to field defaults."""
        printer_name = self.lbl_selected_printer.text()
        if not printer_name or printer_name == "No printer selected":
            return

        reply = QMessageBox.question(
            self, "Reset Config",
            f"Reset configuration for '{printer_name}' to default values?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        # Remove saved config from config_data
        configs = self.config_data.get('printer_configs', {})
        if printer_name in configs:
            del configs[printer_name]
            if not configs:
                self.config_data.pop('printer_configs', None)
            else:
                self.config_data['printer_configs'] = configs

        # Reset UI to defaults
        for field_name, field_def in PRINTER_CONTROL_FIELDS.items():
            w = self.printer_control_widgets.get(field_name)
            if w is None:
                continue
            if field_def['type'] == 'combobox':
                w.setCurrentText(field_def['default'])
            elif field_def['type'] == 'spinbox':
                w.setValue(field_def['default'])
            elif field_def['type'] == 'checkbox':
                w.setChecked(field_def['default'])

        QMessageBox.information(self, "Reset", f"Configuration for '{printer_name}' reset to defaults.")

    def on_view_capabilities(self):
        """Open the capabilities drill-down dialog for the selected printer."""
        printer_name = self.lbl_selected_printer.text()
        if not printer_name or printer_name == "No printer selected":
            log.debug("on_view_capabilities: no printer selected, skipping")
            return

        dlg = PrinterCapabilitiesDialog(printer_name, self)
        dlg.exec()

    def setup_jobs_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        layout.setContentsMargins(4, 4, 4, 4)

        # Header
        jobs_header = QLabel("Recent Print Jobs")
        jobs_header.setStyleSheet(
            "font-weight: 700; font-size: 14px; color: %s; padding: 4px 0;" % self._theme['text_primary']
        )
        layout.addWidget(jobs_header)

        self.table_jobs = QTableWidget(0, 5)
        self.table_jobs.setHorizontalHeaderLabels(["Time", "Printer", "Type", "Status", "Preview/Error"])
        self.table_jobs.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table_jobs.horizontalHeader().setMinimumSectionSize(80)
        self.table_jobs.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_jobs.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_jobs.setAlternatingRowColors(True)
        self.table_jobs.verticalHeader().setVisible(False)
        
        layout.addWidget(self.table_jobs)
        self.tabs.addTab(tab, "Recent Jobs")

    def populate_data(self):
        self.input_hub_url.setText(self.config_data.get("hub_url", ""))
        self.input_agent_key.setText(self.config_data.get("agent_key", ""))
        self.spin_port.setValue(self.config_data.get("port", 49211))
        self.spin_interval.setValue(self.config_data.get("sync_interval_seconds", 60))
        self.spin_retries.setValue(self.config_data.get("max_retries", 3))
        self.spin_retry_delay.setValue(self.config_data.get("retry_delay_seconds", 60))
        
        # Update hub URL display
        hub_url = self.config_data.get("hub_url", "")
        self.lbl_hub_url_display.setText(hub_url if hub_url else "Not configured")
        
        self.refresh_status()

    def _update_status_indicators(self, hub_status):
        """Update connection tab indicator + footer bar based on hub_status."""
        # ── Connection tab indicator ──
        if "Connected" in hub_status:
            # Green dot + text
            self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #4caf88;")
            self.lbl_status_text.setText("Connected")
            self.lbl_status_text.setStyleSheet("font-size: 14px; color: #4caf88; font-weight: 600;")
            self.lbl_conn_status.setText(hub_status)
            self.lbl_conn_status.setStyleSheet("color: #4caf88; font-weight: 500; font-size: 12px;")
            # Footer
            self.lbl_footer_status.setText("● Connected")
            self.lbl_footer_status.setStyleSheet("color: #4caf88; font-weight: 600; font-size: 12px;")
            self._last_sync_time = None  # will be updated by next sync
        elif "Offline" in hub_status or "Disconnected" in hub_status:
            # Red dot + text
            self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #e66565;")
            self.lbl_status_text.setText("Disconnected")
            self.lbl_status_text.setStyleSheet("font-size: 14px; color: #e66565; font-weight: 600;")
            self.lbl_conn_status.setText(hub_status)
            self.lbl_conn_status.setStyleSheet("color: #e66565; font-weight: 500; font-size: 12px;")
            # Footer
            self.lbl_footer_status.setText("● Disconnected")
            self.lbl_footer_status.setStyleSheet("color: #e66565; font-weight: 600; font-size: 12px;")
        else:
            # Amber dot + text
            self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #f0b34b;")
            self.lbl_status_text.setText("Checking...")
            self.lbl_status_text.setStyleSheet("font-size: 14px; color: #f0b34b; font-weight: 600;")
            self.lbl_conn_status.setText(hub_status)
            self.lbl_conn_status.setStyleSheet("color: #f0b34b; font-weight: 500; font-size: 12px;")
            # Footer
            self.lbl_footer_status.setText("● Checking...")
            self.lbl_footer_status.setStyleSheet("color: #f0b34b; font-weight: 600; font-size: 12px;")

    def refresh_status(self):
        # Hub Status
        hub_status = server.get_hub_status()
        self._update_status_indicators(hub_status)
        
        # Update hub url display in status group
        hub_url = self.input_hub_url.text().strip()
        self.lbl_hub_url_display.setText(hub_url if hub_url else "Not configured")
            
        # Printers — populate QListWidget
        try:
            printers_list = printer.get_printers()
            self.lbl_printer_count.setText(f"Found {len(printers_list)} printers on this system.")
            
            # Preserve selection across refresh
            current_item = self.list_printers.currentItem()
            current_name = current_item.text() if current_item else None
            
            self.list_printers.blockSignals(True)
            self.list_printers.clear()
            for p in printers_list:
                self.list_printers.addItem(p['name'])
            self.list_printers.blockSignals(False)
            
            # Restore selection if printer still exists
            if current_name:
                items = self.list_printers.findItems(current_name, Qt.MatchExactly)
                if items:
                    self.list_printers.setCurrentItem(items[0])
                elif self.list_printers.count() > 0:
                    self.list_printers.setCurrentRow(0)
            elif self.list_printers.count() > 0:
                self.list_printers.setCurrentRow(0)
        except Exception:
            pass

        # Jobs
        try:
            jobs = server._job_queue.list_recent(20)
            self.table_jobs.setRowCount(0)
            for j in reversed(jobs):
                row = self.table_jobs.rowCount()
                self.table_jobs.insertRow(row)
                
                time_str = j['created_at'][11:19] if j['created_at'] else ""
                
                self.table_jobs.setItem(row, 0, QTableWidgetItem(time_str))
                self.table_jobs.setItem(row, 1, QTableWidgetItem(j['printer']))
                self.table_jobs.setItem(row, 2, QTableWidgetItem(j['type']))
                
                status_item = QTableWidgetItem(j['status'])
                if j['status'] == 'success':
                    status_item.setForeground(QColor("#4caf88"))
                elif j['status'] == 'failed':
                    status_item.setForeground(QColor("#e66565"))
                else:
                    status_item.setForeground(QColor("#f0b34b"))
                self.table_jobs.setItem(row, 3, status_item)
                
                info = j['error'] if j['status'] == 'failed' else j.get('data_preview', '')
                self.table_jobs.setItem(row, 4, QTableWidgetItem(info or ''))
        except Exception:
            pass

        # Update last sync time in footer
        try:
            if "Connected" in hub_status:
                from datetime import datetime
                self._last_sync_time = datetime.now()
                self.lbl_last_sync.setText(f"Last sync: {self._last_sync_time.strftime('%H:%M:%S')}")
        except Exception:
            pass

    def on_autostart_toggled(self, checked):
        if checked:
            autostart.enable_autostart()
        else:
            autostart.disable_autostart()

    def _on_windows_startup_toggled(self, checked):
        """Persist the 'Start TrayPrint when Windows starts' preference to config.json."""
        self.config_data['windows_startup'] = checked
        self.save_config()
        if checked:
            # Also enable the standard autostart mechanism
            autostart.enable_autostart()
            self.chk_autostart.setChecked(True)
        else:
            # Optionally disable standard autostart if user unchecks
            autostart.disable_autostart()
            self.chk_autostart.setChecked(False)

    def on_test_connection(self):
        hub_url = self.input_hub_url.text().strip()
        agent_key = self.input_agent_key.text().strip()
        
        if not hub_url or not agent_key:
            QMessageBox.warning(self, "Validation Error", "Hub URL and Agent Key are required.")
            return
            
        self.btn_test_conn.setEnabled(False)
        self.btn_test_conn.setText("Testing...")
        QApplication.processEvents()
        
        # Show testing state
        self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #f0b34b;")
        self.lbl_status_text.setText("Testing...")
        self.lbl_status_text.setStyleSheet("font-size: 14px; color: #f0b34b; font-weight: 600;")
        self.lbl_conn_status.setText("Connecting...")
        self.lbl_conn_status.setStyleSheet("color: #f0b34b; font-weight: 500; font-size: 12px;")
        
        try:
            headers = {'Authorization': f'Bearer {agent_key}'}
            resp = requests.get(f'{hub_url}/api/print-hub/profiles', headers=headers, timeout=5)
            
            if resp.status_code == 200:
                json_data = resp.json()
                # Handle both wrapped and unwrapped response
                data = json_data.get('data', json_data)
                profiles = data.get('profiles', {})
                count = len(profiles)
                
                # Update server status so the main UI knows we are connected
                server._hub_last_status = "Connected"
                
                # Ensure the hub sync loop is running with these credentials.
                # If the sync loop was started at app boot with empty/wrong values,
                # this (re)starts it with the correct hub_url and agent_key.
                try:
                    from server import start_hub_sync, _hub_sync_running
                    if not _hub_sync_running:
                        interval = self.spin_interval.value() if hasattr(self, 'spin_interval') else 60
                        max_retries = self.spin_retries.value() if hasattr(self, 'spin_retries') else 3
                        retry_delay = self.spin_retry_delay.value() if hasattr(self, 'spin_retry_delay') else 60
                        start_hub_sync(hub_url, agent_key, interval, max_retries, retry_delay)
                except Exception:
                    pass
                    
                QMessageBox.information(
                    self, "Success",
                    f"Connected successfully!\n\n"
                    f"Hub found {count} printer queue(s).\n"
                    f"Your agent status has been updated in the dashboard.\n\n"
                    "NOTE: Remember to click 'Save & Restart' to finalize these settings."
                )
                self.refresh_status()
            elif resp.status_code == 401:
                QMessageBox.warning(self, "Auth Failed", "Invalid Agent Key.\nPlease verify your key in the Print Hub dashboard.")
                # Update status indicators for auth failure
                self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #e66565;")
                self.lbl_status_text.setText("Authentication Failed")
                self.lbl_status_text.setStyleSheet("font-size: 14px; color: #e66565; font-weight: 600;")
                self.lbl_conn_status.setText("Invalid Agent Key (HTTP 401)")
                self.lbl_conn_status.setStyleSheet("color: #e66565; font-weight: 500; font-size: 12px;")
            else:
                QMessageBox.warning(self, "Connection Error", f"Hub returned an error (HTTP {resp.status_code})")
                self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #e66565;")
                self.lbl_status_text.setText("Error")
                self.lbl_status_text.setStyleSheet("font-size: 14px; color: #e66565; font-weight: 600;")
                self.lbl_conn_status.setText(f"HTTP Error {resp.status_code}")
                self.lbl_conn_status.setStyleSheet("color: #e66565; font-weight: 500; font-size: 12px;")
        except Exception as e:
            QMessageBox.critical(self, "Connection Error", f"Cannot reach the Print Hub server.\n\nError details: {e}")
            self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #e66565;")
            self.lbl_status_text.setText("Connection Failed")
            self.lbl_status_text.setStyleSheet("font-size: 14px; color: #e66565; font-weight: 600;")
            self.lbl_conn_status.setText("Cannot reach server")
            self.lbl_conn_status.setStyleSheet("color: #e66565; font-weight: 500; font-size: 12px;")
        finally:
            self.btn_test_conn.setEnabled(True)
            self.btn_test_conn.setText("Test Connection")

    def apply_settings(self):
        """Apply settings changes live without restarting the app.

        Uses a QTimer.singleShot to signal the server to reload config after save,
        allowing the UI to close cleanly before the config reload takes effect.
        """
        # Use QTimer to signal reload — gives UI time to close before the
        # config reload (which may restart hub sync or change intervals)
        if self.apply_callback:
            QTimer.singleShot(200, self.apply_callback)
        
        # Show brief status message
        QMessageBox.information(
            self, "Settings Applied",
            "Settings saved and applied successfully.\n\n"
            "Changes take effect immediately. No restart required."
        )
        self.accept()

    def on_save_clicked(self):
        self.config_data["hub_url"] = self.input_hub_url.text().strip()
        self.config_data["agent_key"] = self.input_agent_key.text().strip()
        self.config_data["port"] = self.spin_port.value()
        self.config_data["sync_interval_seconds"] = self.spin_interval.value()
        self.config_data["max_retries"] = self.spin_retries.value()
        self.config_data["retry_delay_seconds"] = self.spin_retry_delay.value()
        
        # Save current printer config if a printer is selected
        current_printer = self.lbl_selected_printer.text()
        if current_printer and current_printer != "No printer selected":
            config = self.read_printer_config_from_ui()
            self.set_printer_config(current_printer, config)
        
        if self.save_config():
            self.apply_settings()

    def apply_styles(self):
        """Apply the current theme stylesheet to the window."""
        stylesheet = get_stylesheet(self._theme_id)
        # Add dialog-specific overrides
        stylesheet += """
            QPushButton#save_btn {
                background: %s; color: %s; font-weight: 700;
                font-size: 14px; padding: 10px 32px; border-radius: 8px; border: none;
            }
            QPushButton#save_btn:hover { background: %s; }
            QPushButton#save_btn:pressed { background: %s; }
            QPushButton#save_btn:disabled { background: %s; }

            QPushButton#cancel_btn {
                background: %s; color: %s; font-weight: 500;
                font-size: 14px; padding: 10px 32px; border-radius: 8px;
                border: 1.5px solid %s;
            }
            QPushButton#cancel_btn:hover { background: %s; border-color: %s; }

            QPushButton#theme_btn {
                background: transparent; border: 1.5px solid %s;
                border-radius: 7px; color: %s;
                font-size: 14px; padding: 4px 12px; min-width: 80px;
            }
            QPushButton#theme_btn:hover { background: %s; border-color: %s; color: %s; }

            QPushButton#test_btn {
                background: %s; color: %s; font-weight: 600;
                padding: 10px 24px; border-radius: 7px; border: none; font-size: 13px;
            }
            QPushButton#test_btn:hover { background: %s; }
            QPushButton#test_btn:pressed { background: %s; }
            QPushButton#test_btn:disabled { background: %s; }

            QPushButton#printer_config_save {
                background: %s; color: white; font-weight: 600;
                padding: 10px 24px; border-radius: 7px; border: none; font-size: 13px;
            }
            QPushButton#printer_config_save:hover { background: %s; }
            QPushButton#printer_config_save:disabled { background: %s; color: rgba(255,255,255,0.5); }

            QPushButton#printer_config_reset {
                background: %s; color: %s; font-weight: 500;
                padding: 10px 20px; border-radius: 7px; border: 1.5px solid %s; font-size: 13px;
            }
            QPushButton#printer_config_reset:hover { background: %s; border-color: %s; }

            QLabel#lbl_version {
                font-weight: 600; color: %s; font-size: 12px;
                background: transparent; padding: 2px 8px;
                border: 1px solid %s; border-radius: 4px;
            }
            QLabel#lbl_footer_status, QLabel#lbl_last_sync {
                font-size: 12px; background: transparent;
            }
        """ % (
            # save_btn
            self._theme['accent'], self._theme['accent_text'],
            self._theme['accent_hover'], self._theme['accent_pressed'],
            self._theme['accent_disabled'],
            # cancel_btn
            self._theme['btn_bg'], self._theme['btn_text'],
            self._theme['btn_border'],
            self._theme['btn_hover'], self._theme['btn_hover_border'],
            # theme_btn
            self._theme['border'], self._theme['text_secondary'],
            self._theme['surface2'], self._theme['border_focus'], self._theme['accent'],
            # test_btn
            self._theme['accent'], self._theme['accent_text'],
            self._theme['accent_hover'], self._theme['accent_pressed'],
            self._theme['accent_disabled'],
            # printer_config_save (teal → success)
            self._theme['success'],
            self._theme['success_hover'],
            self._theme['success_disabled'],
            # printer_config_reset
            self._theme['btn_bg'], self._theme['btn_text'],
            self._theme['btn_border'],
            self._theme['btn_hover'], self._theme['btn_hover_border'],
            # lbl_version
            self._theme['text_muted'], self._theme['border'],
        )
        self.setStyleSheet(stylesheet)
        
        # Update theme toggle button text
        opposite_id = 'light' if self._theme_id == 'dark' else 'dark'
        opposite = get_palette(opposite_id)
        self.btn_theme.setText(f"{opposite['icon']}  {opposite['name']}")

    def _on_theme_combo_changed(self, index):
        """Handle theme combo box selection."""
        new_theme = self.cmb_theme.itemData(index)
        if new_theme and new_theme != self._theme_id:
            self._theme_id = new_theme
            self._theme = get_palette(self._theme_id)
            self.config_data['theme'] = self._theme_id
            self.save_config()
            self.apply_styles()
            # Re-apply group box per-widget styles (they use inline stylesheets)
            self.refresh_status()

    def toggle_theme(self):
        """Toggle between light and dark theme without restarting."""
        self._theme_id = 'dark' if self._theme_id == 'light' else 'light'
        self._theme = get_palette(self._theme_id)
        self.config_data['theme'] = self._theme_id
        self.save_config()
        
        # Sync combo box
        idx = self.cmb_theme.findData(self._theme_id)
        if idx >= 0:
            self.cmb_theme.blockSignals(True)
            self.cmb_theme.setCurrentIndex(idx)
            self.cmb_theme.blockSignals(False)
        
        self.apply_styles()
        # Re-apply group box per-widget styles (they use inline stylesheets)
        self.refresh_status()


# ─────────────────────────────────────────────
#  Printer Capabilities Drill-Down Dialog
# ─────────────────────────────────────────────

class PrinterCapabilitiesDialog(QDialog):
    """Detailed printer capabilities viewer with test print support.

    Shows per-printer details for:
      - Input trays (AutoSelect, Tray1, Tray2, ManualFeed, etc.)
      - Media sizes (A4, Letter, Legal, custom sizes)
      - Resolutions (600dpi, 1200dpi, etc.)
      - Color modes (Color, Grayscale, CMYK)
      - Duplex modes (None, TwoSidedLong, TwoSidedShort)
    """

    def __init__(self, printer_name: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Printer Capabilities — {printer_name}")
        self.setMinimumSize(600, 500)
        self.resize(600, 500)

        self.printer_name = printer_name
        self._capabilities: dict = {}
        self._load_in_progress = False

        self.setup_ui()
        self.load_capabilities()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── Header ──
        header_layout = QHBoxLayout()
        self.lbl_printer_title = QLabel(f"<h2>{self.printer_name}</h2>")
        header_layout.addWidget(self.lbl_printer_title)
        header_layout.addStretch()

        self.lbl_status = QLabel("Loading...")
        self.lbl_status.setStyleSheet("color: #f0b34b; font-weight: 600;")
        header_layout.addWidget(self.lbl_status)
        layout.addLayout(header_layout)

        # ── Capabilities Grid ──
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        self.caps_container = QWidget()
        self.caps_layout = QVBoxLayout(self.caps_container)
        self.caps_layout.setSpacing(10)
        self.caps_layout.setContentsMargins(0, 0, 0, 0)

        # Placeholder label - will be populated after discovery
        self.lbl_placeholder = QLabel("Querying printer capabilities...")
        self.lbl_placeholder.setStyleSheet("color: #888; font-size: 13px; padding: 20px;")
        self.lbl_placeholder.setAlignment(Qt.AlignCenter)
        self.caps_layout.addWidget(self.lbl_placeholder)
        self.caps_layout.addStretch()

        scroll.setWidget(self.caps_container)
        layout.addWidget(scroll, stretch=1)

        # ── Bottom buttons ──
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self.btn_refresh = QPushButton("Refresh")
        self.btn_refresh.clicked.connect(self.load_capabilities)
        btn_layout.addWidget(self.btn_refresh)

        self.btn_test_print = QPushButton("Test Print...")
        self.btn_test_print.setObjectName("save_btn")
        self.btn_test_print.clicked.connect(self.on_test_print)
        self.btn_test_print.setEnabled(False)
        btn_layout.addWidget(self.btn_test_print)

        btn_layout.addStretch()

        self.btn_close = QPushButton("Close")
        self.btn_close.setObjectName("cancel_btn")
        self.btn_close.clicked.connect(self.accept)
        btn_layout.addWidget(self.btn_close)

        layout.addLayout(btn_layout)

    def load_capabilities(self):
        """Fetch capabilities in a background thread."""
        if self._load_in_progress:
            return
        self._load_in_progress = True
        self.lbl_status.setText("Loading...")
        self.lbl_status.setStyleSheet("color: #f0b34b; font-weight: 600;")
        self.btn_refresh.setEnabled(False)
        self.btn_test_print.setEnabled(False)

        def _do_load():
            try:
                caps = capabilities.discover_capabilities(self.printer_name)
                # Schedule UI update on main thread
                QTimer.singleShot(0, lambda: self._display_capabilities(caps))
            except Exception as e:
                QTimer.singleShot(0, lambda: self._display_error(str(e)))

        threading.Thread(target=_do_load, daemon=True).start()

    def _display_error(self, error_msg: str):
        """Show an error message in the capabilities area."""
        self._load_in_progress = False
        self.lbl_status.setText("Error")
        self.lbl_status.setStyleSheet("color: #e66565; font-weight: 600;")
        self.btn_refresh.setEnabled(True)

        # Clear layout
        self._clear_caps_layout()
        error_label = QLabel(f"Failed to discover capabilities:\n{error_msg}")
        error_label.setStyleSheet("color: #e66565; font-size: 13px; padding: 20px;")
        error_label.setAlignment(Qt.AlignCenter)
        error_label.setWordWrap(True)
        self.caps_layout.addWidget(error_label)
        self.caps_layout.addStretch()

    def _clear_caps_layout(self):
        """Remove all widgets from the capabilities layout."""
        while self.caps_layout.count():
            item = self.caps_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def _display_capabilities(self, caps: dict):
        """Populate the UI with discovered capabilities."""
        self._load_in_progress = False
        self._capabilities = caps
        self.lbl_status.setText("Ready")
        self.lbl_status.setStyleSheet("color: #4caf88; font-weight: 600;")
        self.btn_refresh.setEnabled(True)
        self.btn_test_print.setEnabled(True)

        # Clear existing widgets
        self._clear_caps_layout()

        # Check for error
        if caps.get('error'):
            error_label = QLabel(f"Capability error: {caps['error']}")
            error_label.setStyleSheet("color: #e66565; font-size: 13px; padding: 20px;")
            error_label.setAlignment(Qt.AlignCenter)
            error_label.setWordWrap(True)
            self.caps_layout.addWidget(error_label)
            self.caps_layout.addStretch()
            return

        # ── Trays ──
        trays = caps.get('trays', [])
        self._add_section("Input Trays", trays, "No trays discovered")

        # ── Media Sizes ──
        media_sizes = caps.get('media_sizes', [])
        self._add_section("Supported Media Sizes", media_sizes, "No media sizes discovered", is_media=True)

        # ── Resolutions ──
        resolutions = caps.get('resolutions', [])
        self._add_section("Resolutions (DPI)", resolutions, "No resolutions discovered")

        # ── Color Modes ──
        color_modes = caps.get('color_modes', [])
        self._add_section("Color Modes", color_modes, "No color modes discovered")

        # ── Duplex ──
        duplex = caps.get('duplex', [])
        self._add_section("Duplex Modes", duplex, "No duplex modes discovered")

        self.caps_layout.addStretch()

    def _add_section(self, title: str, items: list, empty_msg: str, is_media: bool = False):
        """Add a named section with capability items to the layout."""
        group = QGroupBox(title)
        group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: 700; font-size: 13px;
                border: 1px solid #444; border-radius: 6px;
                margin-top: 12px; padding: 14px 10px 8px 10px;
                background: rgba(255,255,255,0.03);
            }}
            QGroupBox::title {{
                subcontrol-origin: margin; subcontrol-position: top left;
                padding: 2px 8px;
            }}
        """)

        if not items:
            group_layout = QVBoxLayout(group)
            empty_label = QLabel(empty_msg)
            empty_label.setStyleSheet("color: #888; font-size: 12px; font-style: italic;")
            group_layout.addWidget(empty_label)
        elif is_media and len(items) > 15:
            # For large lists (many media sizes), use a scrollable text area
            group_layout = QVBoxLayout(group)
            text_area = QTextEdit()
            text_area.setReadOnly(True)
            text_area.setMaximumHeight(120)
            text_area.setPlainText("\n".join(items))
            text_area.setStyleSheet("font-size: 12px;")
            group_layout.addWidget(text_area)
        else:
            group_layout = QVBoxLayout(group)
            # Show as a flow of tags
            tags_widget = QWidget()
            tags_layout = QHBoxLayout(tags_widget)
            tags_layout.setSpacing(6)
            tags_layout.setContentsMargins(0, 0, 0, 0)
            tags_layout.setAlignment(Qt.AlignLeft)

            # Display in a grid-like fashion, up to 4 columns
            cols = min(4, len(items))
            grid = QGridLayout()
            grid.setSpacing(6)
            for idx, item in enumerate(items):
                label = QLabel(f"  {item}  ")
                label.setStyleSheet("""
                    background: #2a2a2a; color: #ccc; font-size: 12px;
                    padding: 4px 10px; border-radius: 4px;
                    border: 1px solid #444;
                """)
                grid.addWidget(label, idx // cols, idx % cols, Qt.AlignLeft)

            group_layout.addLayout(grid)

        self.caps_layout.addWidget(group)

    def on_test_print(self):
        """Send a test print job to the selected printer."""
        if not self._capabilities or self._capabilities.get('error'):
            QMessageBox.warning(self, "Cannot Test Print",
                                "Capabilities not available. Refresh capabilities first.")
            return

        reply = QMessageBox.question(
            self, "Test Print",
            f"Send a test page to '{self.printer_name}'?\n\n"
            "This will print a simple test page to verify printer connectivity.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.btn_test_print.setEnabled(False)
        self.btn_test_print.setText("Printing...")
        QApplication.processEvents()

        def _do_test():
            try:
                # Generate a simple test page (PostScript or raw text)
                test_data = (
                    f"TrayPrint Test Page\n"
                    f"====================\n\n"
                    f"Printer: {self.printer_name}\n"
                    f"Date: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
                    f"Capabilities:\n"
                    f"  Trays: {', '.join(self._capabilities.get('trays', ['N/A']))}\n"
                    f"  Media: {', '.join(self._capabilities.get('media_sizes', ['N/A'])[:5])}\n"
                    f"  Resolutions: {', '.join(self._capabilities.get('resolutions', ['N/A']))}\n"
                    f"  Color Modes: {', '.join(self._capabilities.get('color_modes', ['N/A']))}\n"
                    f"  Duplex: {', '.join(self._capabilities.get('duplex', ['N/A']))}\n\n"
                    f"If you can read this, the printer is working correctly.\n"
                )

                # Use the server's raw print function directly
                success, error = server.print_raw(self.printer_name, test_data)
                if success:
                    QTimer.singleShot(0, lambda: QMessageBox.information(
                        self, "Test Print", f"Test page sent to '{self.printer_name}' successfully."))
                else:
                    QTimer.singleShot(0, lambda: QMessageBox.warning(
                        self, "Test Print Failed", f"Failed to print test page:\n{error}"))
            except Exception as e:
                QTimer.singleShot(0, lambda: QMessageBox.critical(
                    self, "Test Print Error", str(e)))
            finally:
                QTimer.singleShot(0, lambda: self.btn_test_print.setEnabled(True))
                QTimer.singleShot(0, lambda: self.btn_test_print.setText("Test Print..."))

        threading.Thread(target=_do_test, daemon=True).start()


def show_settings():
    window = SettingsWindow()
    window.exec()
