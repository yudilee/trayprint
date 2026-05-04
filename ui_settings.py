import os
import json
import requests
from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QLabel, QLineEdit, QPushButton, QSpinBox, QCheckBox, QComboBox,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView,
    QFormLayout, QGroupBox, QFrame, QListWidget, QSplitter,
    QScrollArea
)
from PySide6.QtCore import Qt, QTimer, QSize
from PySide6.QtGui import QFont, QIcon, QColor, QPalette

from path_utils import get_root_dir
import server
import autostart
import printer

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
        self.setMinimumSize(800, 600)
        self.resize(900, 650)
        
        self.config_path = os.path.join(get_root_dir(), 'config.json')
        self.config_data = {
            "port": 49211,
            "sync_interval_seconds": 60,
            "max_retries": 3,
            "retry_delay_seconds": 60,
            "hub_url": "",
            "agent_key": "",
            "printer_configs": {}
        }
        self.load_config()
        
        self._last_sync_time = None
        
        self.setup_ui()
        self.populate_data()
        
        # Auto-refresh jobs and status while settings is open
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.refresh_status)
        self.refresh_timer.start(5000)

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

        footer_layout.addWidget(self.lbl_version)
        footer_layout.addStretch()
        footer_layout.addWidget(self.lbl_footer_status)
        footer_layout.addStretch()
        footer_layout.addWidget(self.lbl_last_sync)

        main_layout.addWidget(footer_frame)
        
        self.apply_styles()

    def _style_section_group(self, title):
        """Create a styled QGroupBox with consistent section header styling."""
        group = QGroupBox(title)
        group.setStyleSheet("""
            QGroupBox {
                font-weight: 700;
                font-size: 14px;
                color: #111827;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                margin-top: 16px;
                padding: 20px 12px 12px 12px;
                background: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 4px 12px;
                background: white;
            }
        """)
        return group

    def _style_field_label(self, text):
        """Create a styled form label."""
        label = QLabel(text)
        label.setStyleSheet("font-weight: 600; font-size: 13px; color: #374151;")
        return label

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
        url_label.setStyleSheet("font-weight: 600; font-size: 12px; color: #6b7280;")
        self.lbl_hub_url_display = QLabel("Not configured")
        self.lbl_hub_url_display.setStyleSheet("font-size: 12px; color: #9ca3af;")
        url_layout.addWidget(url_label)
        url_layout.addWidget(self.lbl_hub_url_display)
        url_layout.addStretch()
        status_layout.addLayout(url_layout)

        # Status indicator row
        indicator_layout = QHBoxLayout()
        self.lbl_status_dot = QLabel("●")
        self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #9ca3af;")
        self.lbl_status_text = QLabel("Checking...")
        self.lbl_status_text.setStyleSheet("font-size: 14px; color: #6b7280; font-weight: 500;")
        indicator_layout.addWidget(self.lbl_status_dot)
        indicator_layout.addSpacing(6)
        indicator_layout.addWidget(self.lbl_status_text)
        indicator_layout.addStretch()

        self.lbl_conn_status = QLabel("")
        self.lbl_conn_status.setStyleSheet("color: #9ca3af; font-size: 12px;")

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

    def setup_general_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        layout.setContentsMargins(4, 4, 4, 4)

        # ── Application Behavior ──
        behavior_group = self._style_section_group("Application Behavior")
        behavior_layout = QVBoxLayout(behavior_group)
        behavior_layout.setSpacing(8)

        self.chk_autostart = QCheckBox("Start automatically on login")
        self.chk_autostart.setChecked(autostart.is_autostart_enabled())
        self.chk_autostart.toggled.connect(self.on_autostart_toggled)
        behavior_layout.addWidget(self.chk_autostart)

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

    # ─────────────────────────────────────────────
    #  Printer Config Tab (split-panel layout)
    # ─────────────────────────────────────────────

    def setup_printers_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        layout.setContentsMargins(4, 4, 4, 4)

        self.lbl_printer_count = QLabel("Found 0 printers on this system.")
        self.lbl_printer_count.setStyleSheet("font-weight: 600; font-size: 13px; color: #374151; margin-bottom: 4px;")
        layout.addWidget(self.lbl_printer_count)

        # Splitter: left = printer list, right = config panel
        splitter = QSplitter(Qt.Horizontal)

        # ── Left panel: Printer List ──
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(6, 6, 4, 6)

        lbl_printers_title = QLabel("Available Printers")
        lbl_printers_title.setStyleSheet("font-weight: 700; color: #111827; padding: 6px 0; font-size: 13px;")
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
        self.lbl_selected_printer.setStyleSheet("font-weight: 700; color: #111827; font-size: 15px; padding: 6px 0;")
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

        self.btn_reset_printer_config = QPushButton("Reset to Defaults")
        self.btn_reset_printer_config.setObjectName("printer_config_reset")
        self.btn_reset_printer_config.clicked.connect(self.on_reset_printer_config)
        self.btn_reset_printer_config.setEnabled(False)

        self.btn_save_printer_config = QPushButton("Save Config")
        self.btn_save_printer_config.setObjectName("printer_config_save")
        self.btn_save_printer_config.clicked.connect(self.on_save_printer_config)
        self.btn_save_printer_config.setEnabled(False)

        btn_row.addWidget(self.btn_reset_printer_config)
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
            return

        item = self.list_printers.item(row)
        if not item:
            return

        printer_name = item.text()
        self.lbl_selected_printer.setText(printer_name)
        self.btn_save_printer_config.setEnabled(True)
        self.btn_reset_printer_config.setEnabled(True)
        self.populate_printer_config_ui(printer_name)

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

    def setup_jobs_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        layout.setContentsMargins(4, 4, 4, 4)

        # Header
        jobs_header = QLabel("Recent Print Jobs")
        jobs_header.setStyleSheet("font-weight: 700; font-size: 14px; color: #111827; padding: 4px 0;")
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
            self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #22c55e;")
            self.lbl_status_text.setText("Connected")
            self.lbl_status_text.setStyleSheet("font-size: 14px; color: #16a34a; font-weight: 600;")
            self.lbl_conn_status.setText(hub_status)
            self.lbl_conn_status.setStyleSheet("color: #22c55e; font-weight: 500; font-size: 12px;")
            # Footer
            self.lbl_footer_status.setText("● Connected")
            self.lbl_footer_status.setStyleSheet("color: #22c55e; font-weight: 600; font-size: 12px;")
            self._last_sync_time = None  # will be updated by next sync
        elif "Offline" in hub_status or "Disconnected" in hub_status:
            # Red dot + text
            self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #ef4444;")
            self.lbl_status_text.setText("Disconnected")
            self.lbl_status_text.setStyleSheet("font-size: 14px; color: #dc2626; font-weight: 600;")
            self.lbl_conn_status.setText(hub_status)
            self.lbl_conn_status.setStyleSheet("color: #ef4444; font-weight: 500; font-size: 12px;")
            # Footer
            self.lbl_footer_status.setText("● Disconnected")
            self.lbl_footer_status.setStyleSheet("color: #ef4444; font-weight: 600; font-size: 12px;")
        else:
            # Amber dot + text
            self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #f59e0b;")
            self.lbl_status_text.setText("Checking...")
            self.lbl_status_text.setStyleSheet("font-size: 14px; color: #d97706; font-weight: 600;")
            self.lbl_conn_status.setText(hub_status)
            self.lbl_conn_status.setStyleSheet("color: #f59e0b; font-weight: 500; font-size: 12px;")
            # Footer
            self.lbl_footer_status.setText("● Checking...")
            self.lbl_footer_status.setStyleSheet("color: #f59e0b; font-weight: 600; font-size: 12px;")

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
                    status_item.setForeground(QColor("#16a34a"))
                elif j['status'] == 'failed':
                    status_item.setForeground(QColor("#dc2626"))
                else:
                    status_item.setForeground(QColor("#d97706"))
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
        self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #f59e0b;")
        self.lbl_status_text.setText("Testing...")
        self.lbl_status_text.setStyleSheet("font-size: 14px; color: #d97706; font-weight: 600;")
        self.lbl_conn_status.setText("Connecting...")
        self.lbl_conn_status.setStyleSheet("color: #f59e0b; font-weight: 500; font-size: 12px;")
        
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
                self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #ef4444;")
                self.lbl_status_text.setText("Authentication Failed")
                self.lbl_status_text.setStyleSheet("font-size: 14px; color: #dc2626; font-weight: 600;")
                self.lbl_conn_status.setText("Invalid Agent Key (HTTP 401)")
                self.lbl_conn_status.setStyleSheet("color: #ef4444; font-weight: 500; font-size: 12px;")
            else:
                QMessageBox.warning(self, "Connection Error", f"Hub returned an error (HTTP {resp.status_code})")
                self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #ef4444;")
                self.lbl_status_text.setText("Error")
                self.lbl_status_text.setStyleSheet("font-size: 14px; color: #dc2626; font-weight: 600;")
                self.lbl_conn_status.setText(f"HTTP Error {resp.status_code}")
                self.lbl_conn_status.setStyleSheet("color: #ef4444; font-weight: 500; font-size: 12px;")
        except Exception as e:
            QMessageBox.critical(self, "Connection Error", f"Cannot reach the Print Hub server.\n\nError details: {e}")
            self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #ef4444;")
            self.lbl_status_text.setText("Connection Failed")
            self.lbl_status_text.setStyleSheet("font-size: 14px; color: #dc2626; font-weight: 600;")
            self.lbl_conn_status.setText("Cannot reach server")
            self.lbl_conn_status.setStyleSheet("color: #ef4444; font-weight: 500; font-size: 12px;")
        finally:
            self.btn_test_conn.setEnabled(True)
            self.btn_test_conn.setText("Test Connection")

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
            reply = QMessageBox.question(
                self, 'Restart Required',
                "Settings saved successfully. TrayPrint needs to restart to apply changes.\n\nRestart now?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            
            if reply == QMessageBox.Yes:
                import sys
                from PySide6.QtCore import QCoreApplication
                QCoreApplication.quit()
                os.execl(sys.executable, sys.executable, *sys.argv)
            else:
                self.accept()

    def apply_styles(self):
        # Modern Tailwind-inspired stylesheet with cohesive design
        self.setStyleSheet("""
            QDialog {
                background: #f9fafb;
                color: #111827;
            }
            
            /* Universal Widget Background fix for Linux/Dark themes */
            QWidget {
                background: transparent;
                color: #111827;
            }
            
            /* ── Tab Widget ── */
            QTabWidget::pane {
                border: 1px solid #d1d5db;
                border-radius: 8px;
                padding: 8px;
                background: white;
                top: -1px;
            }
            QTabWidget::tab-bar {
                left: 8px;
            }
            QTabBar::tab {
                padding: 10px 22px;
                margin-right: 4px;
                border: 1px solid transparent;
                border-bottom: none;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                background: #f3f4f6;
                color: #6b7280;
                font-weight: 500;
                font-size: 13px;
            }
            QTabBar::tab:hover:!selected {
                background: #e5e7eb;
                color: #374151;
            }
            QTabBar::tab:selected {
                background: white;
                color: #2563eb;
                font-weight: 600;
                border-color: #d1d5db;
                border-bottom: 2px solid #2563eb;
            }
            
            /* ── Group Box (base, overridden by _style_section_group per-widget) ── */
            QGroupBox {
                font-weight: 600;
                font-size: 13px;
                color: #374151;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                margin-top: 12px;
                padding: 12px;
                background: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 6px;
                color: #6b7280;
            }
            
            /* ── Labels ── */
            QLabel {
                color: #374151;
                background: transparent;
            }
            
            /* ── Input Fields ── */
            QLineEdit, QSpinBox, QComboBox {
                padding: 8px 12px;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background: white;
                color: #111827;
                font-size: 13px;
                min-height: 20px;
                selection-background-color: #2563eb;
                selection-color: white;
            }
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus {
                border-color: #2563eb;
            }
            QLineEdit:hover, QSpinBox:hover, QComboBox:hover {
                border-color: #9ca3af;
            }
            QLineEdit:disabled, QSpinBox:disabled, QComboBox:disabled {
                background: #f9fafb;
                color: #9ca3af;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 24px;
                border-left: 1px solid #d1d5db;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
            }
            QComboBox::down-arrow {
                width: 10px;
                height: 10px;
            }
            
            /* ── SpinBox arrows ── */
            QSpinBox::up-button, QSpinBox::down-button {
                subcontrol-origin: border;
                border-left: 1px solid #d1d5db;
                width: 20px;
            }
            QSpinBox::up-button {
                subcontrol-position: top right;
                border-bottom: 1px solid #d1d5db;
                border-top-right-radius: 6px;
            }
            QSpinBox::down-button {
                subcontrol-position: bottom right;
                border-bottom-right-radius: 6px;
            }
            
            /* ── Push Buttons ── */
            QPushButton {
                padding: 8px 20px;
                background: white;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                color: #374151;
                font-weight: 500;
                font-size: 13px;
                min-height: 20px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
                color: #111827;
            }
            QPushButton:pressed {
                background: #f3f4f6;
            }
            QPushButton:disabled {
                background: #f9fafb;
                color: #d1d5db;
                border-color: #e5e7eb;
            }
            
            /* Save Button */
            QPushButton#save_btn {
                background: #2563eb;
                color: white;
                font-weight: 600;
                font-size: 14px;
                padding: 10px 32px;
                border-radius: 8px;
                border: none;
            }
            QPushButton#save_btn:hover {
                background: #1d4ed8;
            }
            QPushButton#save_btn:pressed {
                background: #1e40af;
            }
            
            /* Cancel Button */
            QPushButton#cancel_btn {
                background: white;
                color: #374151;
                font-weight: 500;
                font-size: 14px;
                padding: 10px 32px;
                border-radius: 8px;
                border: 1px solid #d1d5db;
            }
            QPushButton#cancel_btn:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
            
            /* Test Connection Button */
            QPushButton#test_btn {
                background: #2563eb;
                color: white;
                font-weight: 600;
                padding: 10px 24px;
                border-radius: 6px;
                border: none;
                font-size: 13px;
            }
            QPushButton#test_btn:hover {
                background: #1d4ed8;
            }
            QPushButton#test_btn:pressed {
                background: #1e40af;
            }
            QPushButton#test_btn:disabled {
                background: #93c5fd;
                color: #dbeafe;
            }
            
            /* Printer Save Config Button (teal accent) */
            QPushButton#printer_config_save {
                background: #0d9488;
                color: white;
                font-weight: 600;
                padding: 10px 24px;
                border-radius: 6px;
                border: none;
                font-size: 13px;
            }
            QPushButton#printer_config_save:hover {
                background: #0f766e;
            }
            QPushButton#printer_config_save:pressed {
                background: #115e59;
            }
            QPushButton#printer_config_save:disabled {
                background: #99f6e4;
                color: #ccfbf1;
            }
            
            /* Printer Reset Button */
            QPushButton#printer_config_reset {
                background: white;
                color: #6b7280;
                font-weight: 500;
                padding: 10px 20px;
                border-radius: 6px;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton#printer_config_reset:hover {
                background: #f9fafb;
                border-color: #9ca3af;
                color: #374151;
            }
            
            /* ── List Widget ── */
            QListWidget {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background: #f9fafb;
                color: #374151;
                outline: none;
                padding: 4px;
                font-size: 13px;
                alternate-background-color: #f3f4f6;
            }
            QListWidget::item {
                padding: 10px 14px;
                border-radius: 4px;
                margin-bottom: 2px;
            }
            QListWidget::item:hover {
                background: #e5e7eb;
                color: #111827;
            }
            QListWidget::item:selected {
                background: #eff6ff;
                color: #2563eb;
                font-weight: 600;
            }
            
            /* ── Table Widget ── */
            QTableWidget {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background: white;
                gridline-color: #f3f4f6;
                color: #111827;
                font-size: 13px;
                alternate-background-color: #f9fafb;
                selection-background-color: #eff6ff;
                selection-color: #1e40af;
            }
            QHeaderView::section {
                background-color: #f9fafb;
                padding: 10px 12px;
                border: none;
                border-bottom: 2px solid #e5e7eb;
                border-right: 1px solid #f3f4f6;
                font-weight: 600;
                color: #6b7280;
                font-size: 11px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            
            /* ── Splitter ── */
            QSplitter::handle {
                background: #e5e7eb;
                width: 1px;
            }
            QSplitter::handle:horizontal {
                width: 6px;
                margin: 2px 0;
            }
            QSplitter::handle:horizontal:hover {
                background: #2563eb;
                width: 6px;
            }
            
            /* ── CheckBox ── */
            QCheckBox {
                background: transparent;
                spacing: 10px;
                font-size: 13px;
                color: #374151;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
                border: 2px solid #d1d5db;
                border-radius: 4px;
                background: white;
            }
            QCheckBox::indicator:checked {
                background-color: #2563eb;
                border-color: #2563eb;
            }
            QCheckBox::indicator:hover {
                border-color: #9ca3af;
            }
            QCheckBox::indicator:checked:hover {
                background-color: #1d4ed8;
                border-color: #1d4ed8;
            }
            
            /* ── Scroll Area ── */
            QScrollArea, QScrollArea QWidget {
                background: transparent;
                border: none;
            }
            
            /* ── Scroll Bar ── */
            QScrollBar:vertical {
                border: none;
                background: #f3f4f6;
                width: 8px;
                margin: 0;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #d1d5db;
                min-height: 24px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #9ca3af;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0;
            }
            
            /* ── Footer Frame ── */
            QFrame#footer_frame {
                background: #f9fafb;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
            }
            
            /* Footer version label */
            QLabel#lbl_version {
                font-weight: 600;
                color: #6b7280;
                font-size: 12px;
                background: transparent;
                padding: 2px 8px;
                border: 1px solid #e5e7eb;
                border-radius: 4px;
            }
            
            /* Footer labels */
            QLabel#lbl_footer_status, QLabel#lbl_last_sync {
                font-size: 12px;
                background: transparent;
            }
        """)

def show_settings():
    window = SettingsWindow()
    window.exec()
