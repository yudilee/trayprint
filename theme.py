"""
Shared theme constants and stylesheet generator for TrayPrint UI.
Provides consistent dark/light theming across all dialogs with
WCAG AA compliant contrast ratios.
"""

# Color palette (WCAG AA compliant)
# Normal text contrast >= 4.5:1, large text >= 3:1
COLORS = {
    'dark': {
        'id': 'dark', 'icon': '\u2600\ufe0f', 'name': 'Light',
        # Backgrounds
        'bg': '#1e1e2e',              # bg_primary - Deep navy
        'bg_primary': '#1e1e2e',
        'bg_secondary': '#2d2d44',    # Slightly lighter
        'surface': '#2d2d44',         # Alias for bg_secondary
        'surface2': '#252538',        # Slightly darker than surface
        'bg_card': '#363650',         # Card / container background
        'bg_input': '#252538',        # Input fields
        'bg_hover': '#404060',        # Hover state
        
        # Text (all ratios verified against #1e1e2e bg)
        'text_primary': '#e0e0f0',    # Main text ~12:1 on dark bg ✓
        'text_secondary': '#a0a0c0',  # Secondary text ~7:1 ✓ (passes 4.5:1)
        'text_muted': '#707090',      # Muted/disabled text ~3.5:1 (passes 3:1 for large)
        'text_inverse': '#1a1a2e',    # Text on light backgrounds
        
        # Accent
        'accent': '#7c9bff',          # Primary accent (blue) ~8:1 on dark ✓
        'accent_hover': '#9bb3ff',    # Accent hover
        'accent_pressed': '#5a7ae0',  # Pressed state
        'accent_disabled': '#3a4a6a', # Disabled accent button bg
        'accent_text': '#1a1a2e',     # Text ON accent buttons (7:1 on #7c9bff ✓)
        'accent_light': '#2a3a5a',    # Light accent bg for selection
        'accent_light_text': '#9bb3ff', # Text on accent_light
        
        # Semantic
        'success': '#4caf88',         # Success green ~6:1 ✓
        'success_hover': '#5cc498',
        'success_disabled': '#2a5a4a',
        'error': '#e66565',           # Error red ~5.5:1 ✓
        'danger': '#e66565',          # Alias for error
        'warning': '#f0b34b',         # Warning amber ~7:1 ✓
        'info': '#5ba3d9',            # Info blue
        'teal': '#4caf88',            # Teal (matches success)
        'teal_hover': '#5cc498',
        'teal_disabled': '#2a5a4a',
        
        # Borders
        'border': '#404060',          # Borders ~3:1
        'border_focus': '#7c9bff',    # Focused border
        'checkbox_border': '#505070',
        
        # Tabs
        'tab_bg': '#252538',
        'tab_text': '#a0a0c0',
        'tab_active_bg': '#363650',
        'tab_active_text': '#7c9bff',
        'tab_active_border': '#7c9bff',
        
        # Input
        'input_bg': '#252538',
        'input_disabled_bg': '#1e1e2e',
        'input_disabled_text': '#707090',
        'checkbox_bg': '#252538',
        
        # Lists
        'list_bg': '#1e1e2e',
        'list_alt': '#252538',
        'list_hover': '#363650',
        'list_sel_bg': '#2a3a5a',
        'list_sel_text': '#9bb3ff',
        
        # Footer
        'footer_bg': '#252538',
        
        # Scrollbars
        'scrollbar_bg': '#2d2d44',
        'scrollbar': '#505070',
        'scrollbar_fg': '#505070',
        'scrollbar_hover': '#606080',
        
        # Splitters
        'splitter': '#404060',
        
        # Table
        'table_header': '#252538',
        'table_header_text': '#a0a0c0',
        'table_grid': '#404060',
        'table_sel_bg': '#2a3a5a',
        'table_sel_text': '#9bb3ff',
        
        # Group box
        'group_border': '#404060',
        'group_title': '#a0a0c0',
        'group_bg': '#363650',
        
        # Buttons
        'btn_bg': '#2d2d44',
        'btn_border': '#404060',
        'btn_text': '#e0e0f0',
        'btn_hover': '#363650',
        'btn_hover_border': '#505070',
        'btn_pressed': '#252538',
    },
    'light': {
        'id': 'light', 'icon': '\U0001f319', 'name': 'Dark',
        # Backgrounds
        'bg': '#f5f5fa',
        'bg_primary': '#f5f5fa',
        'bg_secondary': '#ffffff',
        'surface': '#ffffff',
        'surface2': '#e8e8f0',
        'bg_card': '#ffffff',
        'bg_input': '#ffffff',
        'bg_hover': '#e8e8f0',
        
        # Text
        'text_primary': '#1a1a2e',
        'text_secondary': '#50506e',
        'text_muted': '#8888a0',
        'text_inverse': '#ffffff',
        
        # Accent
        'accent': '#4a6fd9',
        'accent_hover': '#5a7fe9',
        'accent_pressed': '#3a5fc9',
        'accent_disabled': '#a0b0e0',
        'accent_text': '#ffffff',
        'accent_light': '#eef1ff',
        'accent_light_text': '#4a6fd9',
        
        # Semantic
        'success': '#2d8f6a',
        'success_hover': '#3a9f7a',
        'success_disabled': '#8fc0b0',
        'error': '#cc4a4a',
        'danger': '#cc4a4a',
        'warning': '#cc8f2e',
        'info': '#3a8abf',
        'teal': '#2d8f6a',
        'teal_hover': '#3a9f7a',
        'teal_disabled': '#8fc0b0',
        
        # Borders
        'border': '#d0d0e0',
        'border_focus': '#4a6fd9',
        'checkbox_border': '#b0b0c8',
        
        # Tabs
        'tab_bg': '#e8e8f0',
        'tab_text': '#50506e',
        'tab_active_bg': '#ffffff',
        'tab_active_text': '#4a6fd9',
        'tab_active_border': '#4a6fd9',
        
        # Input
        'input_bg': '#ffffff',
        'input_disabled_bg': '#f0f0f5',
        'input_disabled_text': '#8888a0',
        'checkbox_bg': '#ffffff',
        
        # Lists
        'list_bg': '#f5f5fa',
        'list_alt': '#e8e8f0',
        'list_hover': '#e8e8f0',
        'list_sel_bg': '#eef1ff',
        'list_sel_text': '#4a6fd9',
        
        # Footer
        'footer_bg': '#e8e8f0',
        
        # Scrollbars
        'scrollbar_bg': '#e0e0ea',
        'scrollbar': '#b0b0c8',
        'scrollbar_fg': '#b0b0c8',
        'scrollbar_hover': '#9090a8',
        
        # Splitters
        'splitter': '#d0d0e0',
        
        # Table
        'table_header': '#e8e8f0',
        'table_header_text': '#50506e',
        'table_grid': '#d0d0e0',
        'table_sel_bg': '#eef1ff',
        'table_sel_text': '#4a6fd9',
        
        # Group box
        'group_border': '#d0d0e0',
        'group_title': '#50506e',
        'group_bg': '#ffffff',
        
        # Buttons
        'btn_bg': '#ffffff',
        'btn_border': '#d0d0e0',
        'btn_text': '#1a1a2e',
        'btn_hover': '#f0f0f5',
        'btn_hover_border': '#b0b0c8',
        'btn_pressed': '#e8e8f0',
    }
}


def get_stylesheet(mode='dark'):
    """Generate complete Qt stylesheet for the given theme mode."""
    c = COLORS[mode]
    
    return f"""
    /* Global */
    QWidget {{
        background-color: {c['bg']};
        color: {c['text_primary']};
        font-family: 'Segoe UI', 'SF Pro Display', -apple-system, sans-serif;
        font-size: 13px;
    }}
    
    QDialog {{
        background-color: {c['bg']};
    }}
    
    /* Labels */
    QLabel {{
        color: {c['text_secondary']};
        background: transparent;
        font-size: 13px;
    }}
    
    QLabel[heading="true"] {{
        font-size: 16px;
        font-weight: 600;
        color: {c['text_primary']};
        padding: 4px 0;
    }}
    
    QLabel[subtitle="true"] {{
        font-size: 12px;
        color: {c['text_secondary']};
    }}
    
    QLabel[status="true"] {{
        font-size: 11px;
        color: {c['text_muted']};
    }}
    
    /* Group Box */
    QGroupBox {{
        background-color: {c['group_bg']};
        border: 1px solid {c['group_border']};
        border-radius: 8px;
        margin-top: 14px;
        padding: 16px 12px 12px 12px;
        font-weight: 600;
        font-size: 13px;
        color: {c['group_title']};
    }}
    
    QGroupBox::title {{
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 3px 10px;
        margin-left: 8px;
        background-color: {c['group_bg']};
        color: {c['group_title']};
        border: 1px solid {c['group_border']};
        border-radius: 4px;
    }}
    
    /* Push Buttons */
    QPushButton {{
        background-color: {c['btn_bg']};
        color: {c['btn_text']};
        border: 1px solid {c['btn_border']};
        border-radius: 6px;
        padding: 6px 16px;
        min-height: 24px;
        font-size: 13px;
    }}
    
    QPushButton:hover {{
        background-color: {c['btn_hover']};
        border-color: {c['btn_hover_border']};
    }}
    
    QPushButton:pressed {{
        background-color: {c['btn_pressed']};
        border-color: {c['accent']};
    }}
    
    QPushButton:disabled {{
        background-color: {c['bg']};
        color: {c['text_muted']};
        border-color: {c['border']};
    }}
    
    QPushButton[primary="true"] {{
        background-color: {c['accent']};
        color: {c['accent_text']};
        border: none;
        font-weight: 600;
    }}
    
    QPushButton[primary="true"]:hover {{
        background-color: {c['accent_hover']};
    }}
    
    QPushButton[danger="true"] {{
        background-color: {c['error']};
        color: white;
        border: none;
    }}
    
    QPushButton[success="true"] {{
        background-color: {c['success']};
        color: white;
        border: none;
    }}
    
    /* Line Edit / Text Edit */
    QLineEdit, QTextEdit, QPlainTextEdit {{
        background-color: {c['input_bg']};
        color: {c['text_primary']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        padding: 6px 10px;
        font-size: 13px;
        selection-background-color: {c['accent']};
        selection-color: {c['accent_text']};
    }}
    
    QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {{
        border-color: {c['border_focus']};
    }}
    
    QLineEdit:hover, QTextEdit:hover, QPlainTextEdit:hover {{
        border-color: {c['accent_hover']};
    }}
    
    QLineEdit:disabled, QTextEdit:disabled, QPlainTextEdit:disabled {{
        background-color: {c['input_disabled_bg']};
        color: {c['input_disabled_text']};
    }}
    
    /* Spin Box */
    QSpinBox {{
        background-color: {c['input_bg']};
        color: {c['text_primary']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        padding: 4px 8px;
        min-height: 24px;
        font-size: 13px;
        selection-background-color: {c['accent']};
        selection-color: {c['accent_text']};
    }}
    
    QSpinBox:focus {{
        border-color: {c['border_focus']};
    }}
    
    QSpinBox:hover {{
        border-color: {c['accent_hover']};
    }}
    
    QSpinBox:disabled {{
        background-color: {c['input_disabled_bg']};
        color: {c['input_disabled_text']};
    }}
    
    QSpinBox::up-button, QSpinBox::down-button {{
        subcontrol-origin: border;
        border-left: 1px solid {c['border']};
        width: 20px;
        background: {c['input_bg']};
    }}
    
    QSpinBox::up-button:hover, QSpinBox::down-button:hover {{
        background: {c['bg_hover']};
    }}
    
    /* Combo Box */
    QComboBox {{
        background-color: {c['input_bg']};
        color: {c['text_primary']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        padding: 4px 10px;
        min-height: 24px;
        font-size: 13px;
    }}
    
    QComboBox:hover {{
        border-color: {c['accent']};
    }}
    
    QComboBox::drop-down {{
        border: none;
        width: 24px;
    }}
    
    QComboBox::down-arrow {{
        width: 10px;
        height: 10px;
    }}
    
    QComboBox QAbstractItemView {{
        background-color: {c['surface']};
        color: {c['text_primary']};
        border: 1px solid {c['border']};
        border-radius: 4px;
        selection-background-color: {c['accent_light']};
        selection-color: {c['accent_light_text']};
        padding: 4px;
        outline: none;
    }}
    
    QComboBox QAbstractItemView::item {{
        padding: 6px 10px;
        min-height: 24px;
        border-radius: 3px;
    }}
    
    QComboBox QAbstractItemView::item:hover {{
        background: {c['bg_hover']};
    }}
    
    QComboBox QAbstractItemView::item:selected {{
        background: {c['accent_light']};
        color: {c['accent_light_text']};
    }}
    
    /* Check Box */
    QCheckBox {{
        spacing: 8px;
        color: {c['text_secondary']};
        font-size: 13px;
    }}
    
    QCheckBox::indicator {{
        width: 18px;
        height: 18px;
        border: 2px solid {c['checkbox_border']};
        border-radius: 4px;
        background-color: {c['checkbox_bg']};
    }}
    
    QCheckBox::indicator:checked {{
        background-color: {c['accent']};
        border-color: {c['accent']};
    }}
    
    QCheckBox::indicator:hover {{
        border-color: {c['accent']};
    }}
    
    QCheckBox::indicator:checked:hover {{
        background-color: {c['accent_hover']};
        border-color: {c['accent_hover']};
    }}
    
    /* Tab Widget */
    QTabWidget::pane {{
        border: 1px solid {c['border']};
        border-radius: 8px;
        background-color: {c['bg_card']};
        top: -1px;
    }}
    
    QTabBar::tab {{
        background-color: {c['tab_bg']};
        color: {c['tab_text']};
        border: 1px solid {c['border']};
        border-bottom: none;
        border-top-left-radius: 6px;
        border-top-right-radius: 6px;
        padding: 8px 18px;
        margin-right: 2px;
        font-size: 13px;
    }}
    
    QTabBar::tab:hover:!selected {{
        background-color: {c['bg_hover']};
        color: {c['accent']};
    }}
    
    QTabBar::tab:selected {{
        background-color: {c['tab_active_bg']};
        color: {c['tab_active_text']};
        font-weight: 700;
        border-bottom: 2px solid {c['tab_active_border']};
    }}
    
    /* Table Widget */
    QTableWidget {{
        background-color: {c['bg_card']};
        color: {c['text_primary']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        gridline-color: {c['table_grid']};
        selection-background-color: {c['table_sel_bg']};
        selection-color: {c['table_sel_text']};
        font-size: 13px;
        alternate-background-color: {c['list_alt']};
    }}
    
    QTableWidget::item {{
        padding: 6px 8px;
        border-bottom: 1px solid {c['table_grid']};
    }}
    
    QTableWidget::item:selected {{
        background-color: {c['accent_light']};
        color: {c['accent_light_text']};
    }}
    
    QHeaderView::section {{
        background-color: {c['table_header']};
        color: {c['table_header_text']};
        border: none;
        border-bottom: 2px solid {c['border']};
        border-right: 1px solid {c['table_grid']};
        padding: 8px 10px;
        font-weight: 600;
        font-size: 11px;
    }}
    
    /* List Widget */
    QListWidget {{
        background-color: {c['list_bg']};
        color: {c['text_primary']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        padding: 4px;
        font-size: 13px;
        outline: none;
        alternate-background-color: {c['list_alt']};
    }}
    
    QListWidget::item {{
        padding: 8px 12px;
        border-radius: 4px;
        color: {c['text_primary']};
    }}
    
    QListWidget::item:hover {{
        background-color: {c['list_hover']};
    }}
    
    QListWidget::item:selected {{
        background-color: {c['accent_light']};
        color: {c['accent_light_text']};
        font-weight: 600;
    }}
    
    /* Scroll Bars */
    QScrollBar:vertical {{
        background: {c['scrollbar_bg']};
        width: 8px;
        border-radius: 4px;
        margin: 0;
    }}
    
    QScrollBar::handle:vertical {{
        background: {c['scrollbar']};
        border-radius: 4px;
        min-height: 30px;
    }}
    
    QScrollBar::handle:vertical:hover {{
        background: {c['scrollbar_hover']};
    }}
    
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0;
    }}
    
    QScrollBar:horizontal {{
        background: {c['scrollbar_bg']};
        height: 8px;
        border-radius: 4px;
    }}
    
    QScrollBar::handle:horizontal {{
        background: {c['scrollbar']};
        border-radius: 4px;
        min-width: 30px;
    }}
    
    QScrollBar::handle:horizontal:hover {{
        background: {c['scrollbar_hover']};
    }}
    
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0;
    }}
    
    /* Progress Bar */
    QProgressBar {{
        background-color: {c['input_bg']};
        border: 1px solid {c['border']};
        border-radius: 4px;
        text-align: center;
        height: 16px;
        font-size: 11px;
        color: {c['text_primary']};
    }}
    
    QProgressBar::chunk {{
        background-color: {c['accent']};
        border-radius: 3px;
    }}
    
    /* Splitter */
    QSplitter::handle {{
        background-color: {c['splitter']};
    }}
    
    QSplitter::handle:horizontal {{
        width: 5px;
        margin: 2px 0;
    }}
    
    QSplitter::handle:horizontal:hover {{
        background-color: {c['accent']};
    }}
    
    QSplitter::handle:vertical {{
        height: 5px;
        margin: 0 2px;
    }}
    
    QSplitter::handle:vertical:hover {{
        background-color: {c['accent']};
    }}
    
    /* Tooltips */
    QToolTip {{
        background-color: {c['surface']};
        color: {c['text_primary']};
        border: 1px solid {c['border']};
        border-radius: 4px;
        padding: 4px 8px;
        font-size: 12px;
    }}
    
    /* Status Bar / Footer */
    QFrame#footer_frame {{
        background: {c['footer_bg']};
        border: 1px solid {c['border']};
        border-radius: 8px;
    }}
    
    /* Scroll Area */
    QScrollArea, QScrollArea > QWidget > QWidget {{
        background: transparent;
        border: none;
    }}
    """


def get_palette(mode='dark'):
    """Get color palette dict for use in code (inline styles, dynamic colors)."""
    return COLORS[mode]


def get_theme_icon(mode='dark'):
    """Return the toggle icon label for switching to the opposite theme."""
    return COLORS[mode]['icon']


def get_theme_name(mode='dark'):
    """Return the name of the opposite theme for toggle buttons."""
    return COLORS[mode]['name']
