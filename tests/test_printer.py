"""
Tests for printer.py — printer enumeration, option builders, and print helpers.

Since most printer functions interact with OS-level APIs (win32print, lp, GDI),
we use mocking to verify logic in isolation.
"""

import pytest


# ===================================================================
# _build_lp_options — CUPS lp command flag builder
# ===================================================================

class TestBuildLpOptions:
    """Test _build_lp_options which converts option dicts to lp flags."""

    def test_empty_options(self):
        """Empty options return an empty list."""
        from printer import _build_lp_options
        assert _build_lp_options({}) == []

    def test_none_options(self):
        """None options return an empty list."""
        from printer import _build_lp_options
        assert _build_lp_options(None) == []

    def test_copies(self):
        """copies > 1 adds -n flag."""
        from printer import _build_lp_options
        result = _build_lp_options({"copies": 3})
        assert "-n" in result
        assert "3" in result

    def test_copies_single(self):
        """copies == 1 does NOT add -n."""
        from printer import _build_lp_options
        result = _build_lp_options({"copies": 1})
        assert "-n" not in result

    def test_duplex_long(self):
        """duplex 'two-sided-long-edge' adds -o sides=two-sided-long-edge."""
        from printer import _build_lp_options
        result = _build_lp_options({"duplex": "two-sided-long-edge"})
        assert "-o" in result
        assert "sides=two-sided-long-edge" in result

    def test_duplex_short(self):
        """duplex 'two-sided-short-edge' adds -o sides=two-sided-short-edge."""
        from printer import _build_lp_options
        result = _build_lp_options({"duplex": "two-sided-short-edge"})
        assert "-o" in result
        assert "sides=two-sided-short-edge" in result

    def test_duplex_none(self):
        """duplex 'none' or 'one-sided' adds -o sides=one-sided."""
        from printer import _build_lp_options
        result = _build_lp_options({"duplex": "none"})
        assert "-o" in result
        assert "sides=one-sided" in result

    def test_orientation_landscape(self):
        """orientation 'landscape' adds -o landscape."""
        from printer import _build_lp_options
        result = _build_lp_options({"orientation": "landscape"})
        assert "-o" in result
        assert "landscape" in result

    def test_orientation_portrait(self):
        """orientation 'portrait' adds -o portrait."""
        from printer import _build_lp_options
        result = _build_lp_options({"orientation": "portrait"})
        assert "-o" in result
        assert "portrait" in result

    def test_paper_size(self):
        """paper_size adds -o media=PaperName."""
        from printer import _build_lp_options
        result = _build_lp_options({"paper_size": "A4"})
        assert "-o" in result
        assert "media=A4" in result

    def test_paper_size_custom(self):
        """paper_size 'custom' should not be added as media."""
        from printer import _build_lp_options
        result = _build_lp_options({"paper_size": "custom"})
        # Custom size might be handled via page-height/page-width
        media_flags = [r for r in result if r.startswith("media=")]
        assert len(media_flags) == 0 or "custom" not in media_flags[0].lower()

    def test_tray_source(self):
        """tray_source adds -o InputSlot=TrayName."""
        from printer import _build_lp_options
        result = _build_lp_options({"tray_source": "Tray2"})
        assert "-o" in result
        assert "InputSlot=Tray2" in result

    def test_color_mode_grayscale(self):
        """color_mode 'monochrome' adds -o ColorModel=Gray."""
        from printer import _build_lp_options
        result = _build_lp_options({"color_mode": "monochrome"})
        assert "-o" in result
        assert "ColorModel=Gray" in result or "ColorModel=KGray" in result

    def test_color_mode_color(self):
        """color_mode 'color' adds -o ColorModel=RGB."""
        from printer import _build_lp_options
        result = _build_lp_options({"color_mode": "color"})
        assert "-o" in result
        assert "ColorModel=RGB" in result or "ColorModel=CMYK" in result

    def test_all_options_together(self):
        """Multiple options produce multiple flags."""
        from printer import _build_lp_options
        result = _build_lp_options({
            "copies": 2,
            "duplex": "two-sided-long-edge",
            "orientation": "landscape",
            "paper_size": "A4",
            "tray_source": "Tray1",
            "color_mode": "monochrome",
        })
        # Should have multiple -o / -n flags
        assert "-n" in result
        assert "2" in result
        o_flags = [result[i+1] for i, v in enumerate(result) if v == "-o"]
        assert len(o_flags) >= 4  # duplex, orientation, media, InputSlot, ColorModel

    def test_fit_to_page(self):
        """fit_to_page adds -o fit-to-page."""
        from printer import _build_lp_options
        result = _build_lp_options({"fit_to_page": True})
        assert "-o" in result
        assert "fit-to-page" in result

    def test_reverse_order(self):
        """reverse_order adds -o outputorder=reverse."""
        from printer import _build_lp_options
        result = _build_lp_options({"reverse_order": True})
        assert "-o" in result
        assert "outputorder=reverse" in result

    def test_page_range(self):
        """page_range adds -o page-ranges=..."""
        from printer import _build_lp_options
        result = _build_lp_options({"page_range": "1-5"})
        assert "-o" in result
        assert "page-ranges=1-5" in result

    def test_collate(self):
        """collate adds -o Collate=True."""
        from printer import _build_lp_options
        result = _build_lp_options({"collate": True})
        assert "-o" in result
        assert "Collate=True" in result

    def test_duplex_saved_ignored(self):
        """eco tracking fields like duplex_saved are ignored by lp options."""
        from printer import _build_lp_options
        result = _build_lp_options({"duplex_saved": 10, "eco_mode": True})
        # These should not produce any lp flags
        assert len(result) == 0


# ===================================================================
# _build_sumatra_options — SumatraPDF CLI option builder
# ===================================================================

class TestBuildSumatraOptions:
    """Test _build_sumatra_options for SumatraPDF CLI flag construction."""

    def test_empty_options(self):
        """Empty options return an empty list."""
        from printer import _build_sumatra_options
        assert _build_sumatra_options({}) == []

    def test_none_options(self):
        """None options return an empty list."""
        from printer import _build_sumatra_options
        assert _build_sumatra_options(None) == []

    def test_copies(self):
        """copies > 1 adds 'Nx' to print-settings."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"copies": 3})
        assert "-print-settings" in result
        assert "3x" in result[result.index("-print-settings") + 1]

    def test_orientation_landscape(self):
        """orientation 'landscape' is reflected in print-settings."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"orientation": "landscape"})
        settings = result[result.index("-print-settings") + 1]
        assert "landscape" in settings

    def test_orientation_portrait(self):
        """orientation 'portrait' is reflected in print-settings."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"orientation": "portrait"})
        settings = result[result.index("-print-settings") + 1]
        assert "portrait" in settings

    def test_paper_size(self):
        """paper_size is reflected in print-settings."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"paper_size": "A4"})
        settings = result[result.index("-print-settings") + 1]
        assert "paper=A4" in settings

    def test_duplex(self):
        """duplex adds 'duplex' to print-settings."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"duplex": "two-sided-long-edge"})
        settings = result[result.index("-print-settings") + 1]
        assert "duplex" in settings

    def test_fit_to_page(self):
        """fit_to_page adds 'fit' to print-settings."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"fit_to_page": True})
        settings = result[result.index("-print-settings") + 1]
        assert "fit" in settings

    def test_reverse_order(self):
        """reverse_order adds 'rev' to print-settings."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"reverse_order": True})
        settings = result[result.index("-print-settings") + 1]
        assert "rev" in settings

    def test_color_mode_monochrome(self):
        """color_mode 'monochrome' adds -grayscale flag."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"color_mode": "monochrome"})
        assert "-grayscale" in result

    def test_color_mode_color(self):
        """color_mode 'color' adds -color flag."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"color_mode": "color"})
        assert "-color" in result

    def test_page_range(self):
        """page_range is reflected in print-settings."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"page_range": "2-5"})
        settings = result[result.index("-print-settings") + 1]
        assert "2-5" in settings

    def test_half_letter_mapped(self):
        """Half Letter is mapped to Statement for Windows drivers."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"paper_size": "Half Letter"})
        settings = result[result.index("-print-settings") + 1]
        assert "paper=Statement" in settings

    def test_no_print_settings_when_no_ui_options(self):
        """When no UI options are relevant, no -print-settings flag."""
        from printer import _build_sumatra_options
        result = _build_sumatra_options({"color_mode": "monochrome"})
        # Only -grayscale, no -print-settings
        assert "-print-settings" not in result

    def test_finishing_warning_logged(self, mocker):
        """Finishing options log a warning since Sumatra doesn't support them."""
        from printer import _build_sumatra_options
        log_warning = mocker.patch("printer.log.warning")
        result = _build_sumatra_options({"finishing_staple": "top_left"})
        assert log_warning.called
        assert "finishing" in log_warning.call_args[0][0].lower()


# ===================================================================
# Printer enumeration (mocked)
# ===================================================================

class TestGetPrinters:
    """Test get_printers with mocked OS backends."""

    def test_get_printers_linux(self, mock_subprocess_run):
        """On Linux, get_printers uses lpstat -a."""
        mock_subprocess_run.return_value.stdout = (
            "PrinterA accepting requests\n"
            "PrinterB accepting requests since Mon Jan 1\n"
        )
        mock_subprocess_run.return_value.returncode = 0

        import printer
        # Simulate Linux
        printer.sys.platform = "linux"
        result = printer.get_printers()
        assert len(result) >= 2
        names = [p["name"] for p in result]
        assert "PrinterA" in names
        assert "PrinterB" in names

    def test_get_printers_linux_empty(self, mock_subprocess_run):
        """When lpstat returns no printers, get_printers returns empty list."""
        mock_subprocess_run.return_value.stdout = ""
        mock_subprocess_run.return_value.returncode = 0

        import printer
        printer.sys.platform = "linux"
        result = printer.get_printers()
        assert result == []


# ===================================================================
# get_default_printer (mocked)
# ===================================================================

class TestGetDefaultPrinter:
    """Test get_default_printer with mocked OS backends."""

    def test_default_printer_linux(self, mock_subprocess_run):
        """On Linux, get_default_printer parses lpstat -d output."""
        mock_subprocess_run.return_value.stdout = (
            "system default destination: OfficePrinter\n"
        )
        mock_subprocess_run.return_value.returncode = 0
        # Mock sys.platform
        import printer
        printer.sys.platform = "linux"
        result = printer.get_default_printer()
        assert result == "OfficePrinter"

    def test_default_printer_linux_none(self, mock_subprocess_run):
        """When no default printer is set, get_default_printer returns empty string."""
        mock_subprocess_run.return_value.stdout = ""
        mock_subprocess_run.return_value.returncode = 1
        import printer
        printer.sys.platform = "linux"
        result = printer.get_default_printer()
        assert result == ""


# ===================================================================
# Paper mapping helper
# ===================================================================

class TestPaperMapping:
    """Test paper size resolution and mapping logic."""

    def test_find_windows_paper_name_returns_none_on_linux(self):
        """_find_windows_paper_name returns None when not on Windows."""
        import printer
        printer.sys.platform = "linux"
        # Should not crash — just return (None, None) or similar
        # This test just ensures it doesn't raise
        result = printer._find_windows_paper_name("AnyPrinter", 210, 297)
        assert result is None or result == (None, None)
