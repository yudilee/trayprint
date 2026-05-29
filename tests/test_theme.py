"""
Tests for theme.py — color palette definitions and stylesheet generation.

All tests are pure (no mocking needed) since theme.py has no external
dependencies beyond Python built-ins.
"""

import pytest


# ===================================================================
# COLORS dictionary integrity
# ===================================================================

class TestColorsDict:
    """Verify the COLORS data structure is well-formed."""

    REQUIRED_KEYS = [
        "id", "icon", "name",
        "bg", "bg_primary", "bg_secondary", "surface", "surface2",
        "bg_card", "bg_input", "bg_hover",
        "text_primary", "text_secondary", "text_muted", "text_inverse",
        "accent", "accent_hover", "accent_pressed", "accent_disabled",
        "accent_text", "accent_light", "accent_light_text",
        "success", "success_hover", "success_disabled",
        "error", "danger", "warning", "info",
        "teal", "teal_hover", "teal_disabled",
        "border", "border_focus", "checkbox_border",
        "tab_bg", "tab_text", "tab_active_bg", "tab_active_text",
        "tab_active_border",
        "input_bg", "input_disabled_bg", "input_disabled_text", "checkbox_bg",
        "list_bg", "list_alt", "list_hover", "list_sel_bg", "list_sel_text",
        "footer_bg",
        "scrollbar_bg", "scrollbar", "scrollbar_fg", "scrollbar_hover",
        "splitter",
        "table_header", "table_header_text", "table_grid",
        "table_sel_bg", "table_sel_text",
        "group_border", "group_title", "group_bg",
        "btn_bg", "btn_border", "btn_text", "btn_hover",
        "btn_hover_border", "btn_pressed",
    ]

    def test_both_themes_present(self):
        """COLORS contains both 'dark' and 'light' themes."""
        from theme import COLORS
        assert "dark" in COLORS
        assert "light" in COLORS

    def test_same_keys_in_both_themes(self):
        """Both themes have identical key sets."""
        from theme import COLORS
        dark_keys = set(COLORS["dark"].keys())
        light_keys = set(COLORS["light"].keys())
        assert dark_keys == light_keys, (
            f"Key mismatch. Dark extra: {dark_keys - light_keys}. "
            f"Light extra: {light_keys - dark_keys}."
        )

    def test_dark_has_all_required_keys(self):
        """Dark theme contains all required keys."""
        from theme import COLORS
        missing = set(self.REQUIRED_KEYS) - set(COLORS["dark"].keys())
        assert not missing, f"Dark theme missing keys: {missing}"

    def test_light_has_all_required_keys(self):
        """Light theme contains all required keys."""
        from theme import COLORS
        missing = set(self.REQUIRED_KEYS) - set(COLORS["light"].keys())
        assert not missing, f"Light theme missing keys: {missing}"

    def test_all_color_values_are_strings(self):
        """All color values in both themes are strings."""
        from theme import COLORS
        for mode in ("dark", "light"):
            for key, value in COLORS[mode].items():
                # Skip non-color metadata keys
                if key in ("id", "icon", "name"):
                    continue
                assert isinstance(value, str), (
                    f"'{mode}.{key}' is {type(value).__name__}, expected str"
                )

    def test_dark_id_and_name(self):
        """Dark theme metadata is correct."""
        from theme import COLORS
        assert COLORS["dark"]["id"] == "dark"
        assert COLORS["dark"]["name"] == "Light"  # name of the OPPOSITE theme

    def test_light_id_and_name(self):
        """Light theme metadata is correct."""
        from theme import COLORS
        assert COLORS["light"]["id"] == "light"
        assert COLORS["light"]["name"] == "Dark"  # name of the OPPOSITE theme


# ===================================================================
# get_palette
# ===================================================================

class TestGetPalette:
    """Test the get_palette helper."""

    def test_returns_dark_palette(self):
        """get_palette('dark') returns the dark color dict."""
        from theme import get_palette, COLORS
        assert get_palette("dark") == COLORS["dark"]

    def test_returns_light_palette(self):
        """get_palette('light') returns the light color dict."""
        from theme import get_palette, COLORS
        assert get_palette("light") == COLORS["light"]

    def test_default_is_dark(self):
        """get_palette() defaults to dark."""
        from theme import get_palette, COLORS
        assert get_palette() == COLORS["dark"]

    def test_invalid_mode_raises_key_error(self):
        """An invalid mode raises KeyError."""
        from theme import get_palette
        with pytest.raises(KeyError):
            get_palette("nonexistent")


# ===================================================================
# get_stylesheet
# ===================================================================

class TestGetStylesheet:
    """Test the stylesheet generator."""

    def test_returns_string(self):
        """get_stylesheet returns a non-empty string."""
        from theme import get_stylesheet
        stylesheet = get_stylesheet("dark")
        assert isinstance(stylesheet, str)
        assert len(stylesheet) > 100

    def test_contains_dark_theme_colors(self):
        """Dark mode stylesheet contains dark theme color codes."""
        from theme import get_stylesheet
        stylesheet = get_stylesheet("dark")
        assert "#1e1e2e" in stylesheet  # dark bg
        assert "#7c9bff" in stylesheet  # dark accent

    def test_contains_light_theme_colors(self):
        """Light mode stylesheet contains light theme color codes."""
        from theme import get_stylesheet
        stylesheet = get_stylesheet("light")
        assert "#f5f5fa" in stylesheet  # light bg
        assert "#4a6fd9" in stylesheet  # light accent

    def test_default_is_dark(self):
        """get_stylesheet() defaults to dark."""
        from theme import get_stylesheet
        stylesheet = get_stylesheet()
        assert "#1e1e2e" in stylesheet

    def test_contains_qss_selectors(self):
        """Stylesheet contains expected Qt style selectors."""
        from theme import get_stylesheet
        stylesheet = get_stylesheet("dark")
        assert "QWidget" in stylesheet
        assert "QPushButton" in stylesheet
        assert "QLineEdit" in stylesheet
        assert "QComboBox" in stylesheet
        assert "QTabWidget" in stylesheet
        assert "QListWidget" in stylesheet
        assert "QTableWidget" in stylesheet
        assert "QScrollBar" in stylesheet
        assert "QGroupBox" in stylesheet
        assert "QCheckBox" in stylesheet
        assert "QSpinBox" in stylesheet
        assert "QProgressBar" in stylesheet
        assert "QSplitter" in stylesheet
        assert "QToolTip" in stylesheet
        assert "QScrollArea" in stylesheet


# ===================================================================
# get_theme_icon / get_theme_name
# ===================================================================

class TestThemeHelpers:
    """Test get_theme_icon and get_theme_name.

    Note: The icon/name stored for each theme mode is the label for the
    *opposite* theme (i.e. what you see when you're in dark mode is the
    icon/name to toggle TO light mode). So get_theme_icon('dark') returns
    COLORS['dark']['icon'] (the light/sun icon), NOT COLORS['light']['icon'].
    """

    def test_get_theme_icon_dark(self):
        """get_theme_icon('dark') returns the icon stored in COLORS['dark']['icon']."""
        from theme import get_theme_icon, COLORS
        assert get_theme_icon("dark") == COLORS["dark"]["icon"]

    def test_get_theme_icon_light(self):
        """get_theme_icon('light') returns the icon stored in COLORS['light']['icon']."""
        from theme import get_theme_icon, COLORS
        assert get_theme_icon("light") == COLORS["light"]["icon"]

    def test_get_theme_name_dark(self):
        """get_theme_name('dark') returns COLORS['dark']['name'] (the opposite name)."""
        from theme import get_theme_name, COLORS
        assert get_theme_name("dark") == COLORS["dark"]["name"]

    def test_get_theme_name_light(self):
        """get_theme_name('light') returns COLORS['light']['name'] (the opposite name)."""
        from theme import get_theme_name, COLORS
        assert get_theme_name("light") == COLORS["light"]["name"]
