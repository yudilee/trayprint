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
#  Theme Manager
# ─────────────────────────────────────────────

class ThemeManager:
    LIGHT = {
        'id': 'light', 'icon': '🌙', 'name': 'Light',
        'bg': '#f8fafc', 'surface': '#ffffff', 'surface2': '#f1f5f9',
        'border': '#e2e8f0', 'border_focus': '#3b82f6',
        'text_primary': '#0f172a', 'text_secondary': '#475569', 'text_muted': '#94a3b8',
        'accent': '#2563eb', 'accent_hover': '#1d4ed8', 'accent_pressed': '#1e40af',
        'accent_disabled': '#93c5fd', 'accent_text': '#ffffff',
        'accent_light': '#eff6ff', 'accent_light_text': '#1d4ed8',
        'teal': '#0d9488', 'teal_hover': '#0f766e', 'teal_disabled': '#99f6e4',
        'success': '#16a34a', 'error': '#dc2626', 'warning': '#d97706',
        'tab_bg': '#f1f5f9', 'tab_text': '#475569',
        'tab_active_bg': '#ffffff', 'tab_active_text': '#2563eb', 'tab_active_border': '#2563eb',
        'input_bg': '#ffffff', 'input_disabled_bg': '#f1f5f9', 'input_disabled_text': '#94a3b8',
        'list_bg': '#f8fafc', 'list_alt': '#f1f5f9', 'list_hover': '#e2e8f0',
        'list_sel_bg': '#eff6ff', 'list_sel_text': '#2563eb',
        'footer_bg': '#f1f5f9', 'scrollbar': '#cbd5e1', 'scrollbar_hover': '#94a3b8',
        'splitter': '#e2e8f0', 'table_header': '#f8fafc', 'table_header_text': '#64748b',
        'table_grid': '#f1f5f9', 'table_sel_bg': '#eff6ff', 'table_sel_text': '#1e40af',
        'group_border': '#e2e8f0', 'group_title': '#374151', 'group_bg': '#ffffff',
        'btn_bg': '#ffffff', 'btn_border': '#e2e8f0', 'btn_text': '#374151',
        'btn_hover': '#f8fafc', 'btn_hover_border': '#94a3b8', 'btn_pressed': '#f1f5f9',
        'checkbox_border': '#d1d5db', 'checkbox_bg': '#ffffff',
    }

    DARK = {
        'id': 'dark', 'icon': '☀️', 'name': 'Dark',
        'bg': '#0f172a', 'surface': '#1e293b', 'surface2': '#162032',
        'border': '#334155', 'border_focus': '#60a5fa',
        'text_primary': '#f1f5f9', 'text_secondary': '#94a3b8', 'text_muted': '#64748b',
        'accent': '#3b82f6', 'accent_hover': '#60a5fa', 'accent_pressed': '#2563eb',
        'accent_disabled': '#1e3a5f', 'accent_text': '#ffffff',
        'accent_light': '#1e3a5f', 'accent_light_text': '#93c5fd',
        'teal': '#14b8a6', 'teal_hover': '#2dd4bf', 'teal_disabled': '#134e4a',
        'success': '#22c55e', 'error': '#f87171', 'warning': '#fbbf24',
        'tab_bg': '#162032', 'tab_text': '#94a3b8',
        'tab_active_bg': '#1e293b', 'tab_active_text': '#60a5fa', 'tab_active_border': '#3b82f6',
        'input_bg': '#0f172a', 'input_disabled_bg': '#162032', 'input_disabled_text': '#475569',
        'list_bg': '#0f172a', 'list_alt': '#162032', 'list_hover': '#1e293b',
        'list_sel_bg': '#1e3a5f', 'list_sel_text': '#93c5fd',
        'footer_bg': '#162032', 'scrollbar': '#334155', 'scrollbar_hover': '#475569',
        'splitter': '#334155', 'table_header': '#162032', 'table_header_text': '#64748b',
        'table_grid': '#1e293b', 'table_sel_bg': '#1e3a5f', 'table_sel_text': '#93c5fd',
        'group_border': '#334155', 'group_title': '#94a3b8', 'group_bg': '#1e293b',
        'btn_bg': '#1e293b', 'btn_border': '#334155', 'btn_text': '#e2e8f0',
        'btn_hover': '#334155', 'btn_hover_border': '#475569', 'btn_pressed': '#162032',
        'checkbox_border': '#475569', 'checkbox_bg': '#0f172a',
    }

    @staticmethod
    def get(theme_id='light'):
        return ThemeManager.DARK if theme_id == 'dark' else ThemeManager.LIGHT

    @staticmethod
    def build_stylesheet(t):
        return f"""
            QDialog {{ background: {t['bg']}; color: {t['text_primary']}; }}
            QWidget {{ background: transparent; color: {t['text_primary']}; }}

            QTabWidget::pane {{
                border: 1px solid {t['border']}; border-radius: 10px;
                padding: 10px; background: {t['surface']}; top: -1px;
            }}
            QTabWidget::tab-bar {{ left: 8px; }}
            QTabBar::tab {{
                padding: 10px 24px; margin-right: 4px;
                border: 1px solid {t['border']}; border-bottom: none;
                border-top-left-radius: 8px; border-top-right-radius: 8px;
                background: {t['tab_bg']}; color: {t['tab_text']};
                font-weight: 500; font-size: 13px;
            }}
            QTabBar::tab:hover:!selected {{ background: {t['surface2']}; color: {t['text_primary']}; }}
            QTabBar::tab:selected {{
                background: {t['tab_active_bg']}; color: {t['tab_active_text']};
                font-weight: 700; border-color: {t['border']};
                border-bottom: 3px solid {t['tab_active_border']};
            }}

            QGroupBox {{
                font-weight: 700; font-size: 13px; color: {t['group_title']};
                border: 1px solid {t['group_border']}; border-radius: 10px;
                margin-top: 14px; padding: 18px 12px 12px 12px; background: {t['group_bg']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin; subcontrol-position: top left;
                padding: 4px 12px; background: {t['group_bg']}; color: {t['group_title']};
            }}

            QLabel {{ color: {t['text_secondary']}; background: transparent; font-size: 13px; }}

            QLineEdit, QSpinBox, QComboBox {{
                padding: 8px 12px; border: 1.5px solid {t['border']};
                border-radius: 7px; background: {t['input_bg']}; color: {t['text_primary']};
                font-size: 13px; min-height: 22px;
                selection-background-color: {t['accent']}; selection-color: {t['accent_text']};
            }}
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{
                border-color: {t['border_focus']}; background: {t['input_bg']};
            }}
            QLineEdit:hover, QSpinBox:hover, QComboBox:hover {{ border-color: {t['text_muted']}; }}
            QLineEdit:disabled, QSpinBox:disabled, QComboBox:disabled {{
                background: {t['input_disabled_bg']}; color: {t['input_disabled_text']};
            }}
            QComboBox::drop-down {{
                subcontrol-origin: padding; subcontrol-position: top right; width: 26px;
                border-left: 1px solid {t['border']}; border-top-right-radius: 7px;
                border-bottom-right-radius: 7px; background: transparent;
            }}
            QComboBox::down-arrow {{ width: 10px; height: 10px; }}

            /* ── CRITICAL: ComboBox popup dropdown ── */
            QComboBox QAbstractItemView {{
                background: {t['surface']}; color: {t['text_primary']};
                border: 1.5px solid {t['border']}; border-radius: 7px;
                padding: 4px; outline: none;
                selection-background-color: {t['accent_light']};
                selection-color: {t['accent_light_text']};
            }}
            QComboBox QAbstractItemView::item {{
                padding: 8px 12px; min-height: 28px; color: {t['text_primary']};
                border-radius: 4px;
            }}
            QComboBox QAbstractItemView::item:hover {{
                background: {t['surface2']}; color: {t['text_primary']};
            }}
            QComboBox QAbstractItemView::item:selected {{
                background: {t['accent_light']}; color: {t['accent_light_text']};
            }}

            QSpinBox::up-button, QSpinBox::down-button {{
                subcontrol-origin: border; border-left: 1px solid {t['border']};
                width: 22px; background: {t['input_bg']};
            }}
            QSpinBox::up-button {{
                subcontrol-position: top right; border-bottom: 1px solid {t['border']};
                border-top-right-radius: 7px;
            }}
            QSpinBox::down-button {{ subcontrol-position: bottom right; border-bottom-right-radius: 7px; }}
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {{ background: {t['surface2']}; }}

            QPushButton {{
                padding: 8px 20px; background: {t['btn_bg']}; border: 1.5px solid {t['btn_border']};
                border-radius: 7px; color: {t['btn_text']}; font-weight: 500;
                font-size: 13px; min-height: 22px;
            }}
            QPushButton:hover {{ background: {t['btn_hover']}; border-color: {t['btn_hover_border']}; color: {t['text_primary']}; }}
            QPushButton:pressed {{ background: {t['btn_pressed']}; }}
            QPushButton:disabled {{ background: {t['surface2']}; color: {t['text_muted']}; border-color: {t['border']}; }}

            QPushButton#save_btn {{
                background: {t['accent']}; color: {t['accent_text']}; font-weight: 700;
                font-size: 14px; padding: 10px 32px; border-radius: 8px; border: none;
            }}
            QPushButton#save_btn:hover {{ background: {t['accent_hover']}; }}
            QPushButton#save_btn:pressed {{ background: {t['accent_pressed']}; }}
            QPushButton#save_btn:disabled {{ background: {t['accent_disabled']}; }}

            QPushButton#cancel_btn {{
                background: {t['btn_bg']}; color: {t['btn_text']}; font-weight: 500;
                font-size: 14px; padding: 10px 32px; border-radius: 8px;
                border: 1.5px solid {t['btn_border']};
            }}
            QPushButton#cancel_btn:hover {{ background: {t['btn_hover']}; border-color: {t['btn_hover_border']}; }}

            QPushButton#theme_btn {{
                background: transparent; border: 1.5px solid {t['border']};
                border-radius: 7px; color: {t['text_secondary']};
                font-size: 14px; padding: 4px 12px; min-width: 80px;
            }}
            QPushButton#theme_btn:hover {{ background: {t['surface2']}; border-color: {t['border_focus']}; color: {t['accent']}; }}

            QPushButton#test_btn {{
                background: {t['accent']}; color: {t['accent_text']}; font-weight: 600;
                padding: 10px 24px; border-radius: 7px; border: none; font-size: 13px;
            }}
            QPushButton#test_btn:hover {{ background: {t['accent_hover']}; }}
            QPushButton#test_btn:pressed {{ background: {t['accent_pressed']}; }}
            QPushButton#test_btn:disabled {{ background: {t['accent_disabled']}; }}

            QPushButton#printer_config_save {{
                background: {t['teal']}; color: white; font-weight: 600;
                padding: 10px 24px; border-radius: 7px; border: none; font-size: 13px;
            }}
            QPushButton#printer_config_save:hover {{ background: {t['teal_hover']}; }}
            QPushButton#printer_config_save:disabled {{ background: {t['teal_disabled']}; color: rgba(255,255,255,0.5); }}

            QPushButton#printer_config_reset {{
                background: {t['btn_bg']}; color: {t['btn_text']}; font-weight: 500;
                padding: 10px 20px; border-radius: 7px; border: 1.5px solid {t['btn_border']}; font-size: 13px;
            }}
            QPushButton#printer_config_reset:hover {{ background: {t['btn_hover']}; border-color: {t['btn_hover_border']}; }}

            QListWidget {{
                border: 1.5px solid {t['border']}; border-radius: 8px;
                background: {t['list_bg']}; color: {t['text_primary']};
                outline: none; padding: 4px; font-size: 13px;
                alternate-background-color: {t['list_alt']};
            }}
            QListWidget::item {{ padding: 10px 14px; border-radius: 5px; margin-bottom: 2px; color: {t['text_primary']}; }}
            QListWidget::item:hover {{ background: {t['list_hover']}; color: {t['text_primary']}; }}
            QListWidget::item:selected {{ background: {t['list_sel_bg']}; color: {t['list_sel_text']}; font-weight: 600; }}

            QTableWidget {{
                border: 1.5px solid {t['border']}; border-radius: 8px;
                background: {t['surface']}; gridline-color: {t['table_grid']};
                color: {t['text_primary']}; font-size: 13px;
                alternate-background-color: {t['surface2']};
                selection-background-color: {t['table_sel_bg']};
                selection-color: {t['table_sel_text']};
            }}
            QHeaderView::section {{
                background-color: {t['table_header']}; padding: 10px 12px;
                border: none; border-bottom: 2px solid {t['border']};
                border-right: 1px solid {t['table_grid']};
                font-weight: 700; color: {t['table_header_text']}; font-size: 11px;
            }}

            QSplitter::handle {{ background: {t['splitter']}; width: 1px; }}
            QSplitter::handle:horizontal {{ width: 5px; margin: 2px 0; }}
            QSplitter::handle:horizontal:hover {{ background: {t['accent']}; }}

            QCheckBox {{
                background: transparent; spacing: 10px;
                font-size: 13px; color: {t['text_secondary']};
            }}
            QCheckBox::indicator {{
                width: 20px; height: 20px; border: 2px solid {t['checkbox_border']};
                border-radius: 5px; background: {t['checkbox_bg']};
            }}
            QCheckBox::indicator:checked {{
                background-color: {t['accent']}; border-color: {t['accent']};
            }}
            QCheckBox::indicator:hover {{ border-color: {t['text_muted']}; }}
            QCheckBox::indicator:checked:hover {{ background-color: {t['accent_hover']}; border-color: {t['accent_hover']}; }}

            QScrollArea, QScrollArea QWidget {{ background: transparent; border: none; }}

            QScrollBar:vertical {{
                border: none; background: {t['surface2']}; width: 8px;
                margin: 0; border-radius: 4px;
            }}
            QScrollBar::handle:vertical {{
                background: {t['scrollbar']}; min-height: 24px; border-radius: 4px;
            }}
            QScrollBar::handle:vertical:hover {{ background: {t['scrollbar_hover']}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}

            QFrame#footer_frame {{
                background: {t['footer_bg']}; border: 1px solid {t['border']}; border-radius: 8px;
            }}
            QLabel#lbl_version {{
                font-weight: 600; color: {t['text_muted']}; font-size: 12px;
                background: transparent; padding: 2px 8px;
                border: 1px solid {t['border']}; border-radius: 4px;
            }}
            QLabel#lbl_footer_status, QLabel#lbl_last_sync {{
                font-size: 12px; background: transparent;
            }}
        """


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
            "printer_configs": {},
            "theme": "light",
        }
        self.load_config()
        
        self._theme_id = self.config_data.get('theme', 'light')
        self._theme = ThemeManager.get(self._theme_id)
        
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

        self.btn_theme = QPushButton()
        self.btn_theme.setObjectName("theme_btn")
        self.btn_theme.setToolTip("Toggle Dark / Light theme")
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
                border: 1px solid {t['group_border']}; border-radius: 10px;
                margin-top: 16px; padding: 20px 12px 12px 12px;
                background: {t['group_bg']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin; subcontrol-position: top left;
                padding: 4px 12px; background: {t['group_bg']};
                color: {t['group_title']};
            }}
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
        url_label.setStyleSheet("font-weight: 600; font-size: 12px; color: #374151;")
        self.lbl_hub_url_display = QLabel("Not configured")
        self.lbl_hub_url_display.setStyleSheet("font-size: 12px; color: #6b7280;")
        url_layout.addWidget(url_label)
        url_layout.addWidget(self.lbl_hub_url_display)
        url_layout.addStretch()
        status_layout.addLayout(url_layout)

        # Status indicator row
        indicator_layout = QHBoxLayout()
        self.lbl_status_dot = QLabel("●")
        self.lbl_status_dot.setStyleSheet("font-size: 26px; color: #6b7280;")
        self.lbl_status_text = QLabel("Checking...")
        self.lbl_status_text.setStyleSheet("font-size: 14px; color: #374151; font-weight: 500;")
        indicator_layout.addWidget(self.lbl_status_dot)
        indicator_layout.addSpacing(6)
        indicator_layout.addWidget(self.lbl_status_text)
        indicator_layout.addStretch()

        self.lbl_conn_status = QLabel("")
        self.lbl_conn_status.setStyleSheet("color: #6b7280; font-size: 12px;")

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
        """Apply the current theme stylesheet to the window."""
        self.setStyleSheet(ThemeManager.build_stylesheet(self._theme))
        icon = self._theme.get('icon', '🌙')
        name = ThemeManager.get('light' if self._theme_id == 'dark' else 'dark').get('name', '')
        self.btn_theme.setText(f"{icon}  {name}")

    def toggle_theme(self):
        """Toggle between light and dark theme without restarting."""
        self._theme_id = 'dark' if self._theme_id == 'light' else 'light'
        self._theme = ThemeManager.get(self._theme_id)
        self.config_data['theme'] = self._theme_id
        self.save_config()
        self.apply_styles()
        # Re-apply group box per-widget styles (they use inline stylesheets)
        self.refresh_status()



def show_settings():
    window = SettingsWindow()
    window.exec()

