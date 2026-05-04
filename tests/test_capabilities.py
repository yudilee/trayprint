"""
Tests for capabilities.py — printer capability discovery and CUPS output parsing.
"""

import pytest


# ===================================================================
# _parse_cups_options — lpoptions output parser
# ===================================================================

class TestParseCupsOptions:
    """Test _parse_cups_options which parses lpoptions -l output."""

    def test_empty_output(self):
        """Empty string returns empty capability lists."""
        from capabilities import _parse_cups_options
        result = _parse_cups_options("")
        assert result["trays"] == []
        assert result["media_sizes"] == []
        assert result["color_modes"] == []
        assert result["duplex"] == []
        assert result["resolutions"] == []

    def test_typical_output(self):
        """Parse a realistic CUPS lpoptions dump."""
        output = (
            "InputSlot/Input Slot: *Auto Tray1 Tray2 ManualFeed\n"
            "PageSize/Media Size: *A4 Letter Legal A5 B5\n"
            "ColorModel/Color Model: *RGB Gray CMYK\n"
            "Duplex/Duplex: *None DuplexNoTumble DuplexTumble\n"
            "Resolution/Resolution: *600dpi 1200dpi\n"
        )
        from capabilities import _parse_cups_options
        result = _parse_cups_options(output)

        assert "Auto" in result["trays"]
        assert "Tray1" in result["trays"]
        assert "ManualFeed" in result["trays"]
        assert len(result["trays"]) == 4

        assert "A4" in result["media_sizes"]
        assert "Letter" in result["media_sizes"]
        assert "B5" in result["media_sizes"]

        assert "RGB" in result["color_modes"]
        assert "Gray" in result["color_modes"]
        assert "CMYK" in result["color_modes"]

        assert "None" in result["duplex"]
        assert "DuplexNoTumble" in result["duplex"]

        assert "600dpi" in result["resolutions"]
        assert "1200dpi" in result["resolutions"]

    def test_no_default_asterisk(self):
        """Options without an asterisk default marker are still parsed."""
        output = (
            "InputSlot/Input Slot: Tray1 Tray2\n"
            "PageSize/Media Size: A4 Letter\n"
        )
        from capabilities import _parse_cups_options
        result = _parse_cups_options(output)
        assert result["trays"] == ["Tray1", "Tray2"]
        assert result["media_sizes"] == ["A4", "Letter"]

    def test_unknown_keys_ignored(self):
        """Lines with unrecognised keywords are silently ignored."""
        output = (
            "InputSlot/Input Slot: *Auto Tray1\n"
            "JobHoldUntil/Job Hold Until: *NoHold Indefinite DayTime\n"
            "SomeRandomKey/Random: *Value1 Value2\n"
        )
        from capabilities import _parse_cups_options
        result = _parse_cups_options(output)
        # Only trays should be populated; others remain empty
        assert result["trays"] == ["Auto", "Tray1"]
        assert result["media_sizes"] == []
        assert result["color_modes"] == []
        assert result["duplex"] == []
        assert result["resolutions"] == []

    def test_sides_duplex_alias(self):
        """'sides' keyword maps to duplex capabilities."""
        output = "sides/Sides: *one-sided two-sided-long-edge two-sided-short-edge\n"
        from capabilities import _parse_cups_options
        result = _parse_cups_options(output)
        assert "one-sided" in result["duplex"]
        assert "two-sided-long-edge" in result["duplex"]

    def test_media_alias(self):
        """'media' keyword maps to media_sizes."""
        output = "media/Media: *iso_a4_210x297mm Letter na_letter_215x279mm\n"
        from capabilities import _parse_cups_options
        result = _parse_cups_options(output)
        assert "iso_a4_210x297mm" in result["media_sizes"]
        assert "Letter" in result["media_sizes"]

    def test_print_quality_resolution(self):
        """'print-quality' maps to resolutions."""
        output = "print-quality/Print Quality: *3 4 5\n"
        from capabilities import _parse_cups_options
        result = _parse_cups_options(output)
        assert "3" in result["resolutions"]
        assert "4" in result["resolutions"]

    def test_malformed_line(self):
        """Malformed lines that don't match the pattern are skipped."""
        output = (
            "Not a valid line\n"
            "InputSlot/Input Slot: *Auto\n"
            "AnotherBadLine\n"
        )
        from capabilities import _parse_cups_options
        result = _parse_cups_options(output)
        assert result["trays"] == ["Auto"]


# ===================================================================
# discover_capabilities — main entry point
# ===================================================================

class TestDiscoverCapabilities:
    """Test the discover_capabilities entry point."""

    def test_no_printer_name(self):
        """Empty printer name returns an error."""
        from capabilities import discover_capabilities
        result = discover_capabilities("")
        assert "error" in result

    def test_none_printer_name(self):
        """None printer name returns an error."""
        from capabilities import discover_capabilities
        result = discover_capabilities(None)
        assert "error" in result


# ===================================================================
# _cups_fallback_capabilities
# ===================================================================

class TestCupsFallback:
    """Test fallback capabilities when lpoptions is unavailable."""

    def test_fallback_contains_all_keys(self):
        """Fallback capabilities include trays, media, color, duplex, resolution."""
        from capabilities import _cups_fallback_capabilities
        result = _cups_fallback_capabilities("AnyPrinter")
        assert "trays" in result
        assert "media_sizes" in result
        assert "color_modes" in result
        assert "duplex" in result
        assert "resolutions" in result

    def test_fallback_has_reasonable_defaults(self):
        """Fallback includes common values like A4, Letter, Gray, None."""
        from capabilities import _cups_fallback_capabilities
        result = _cups_fallback_capabilities("AnyPrinter")
        assert "A4" in result["media_sizes"]
        assert "Letter" in result["media_sizes"]
        assert "Gray" in result["color_modes"]
        assert "None" in result["duplex"]
        assert "600dpi" in result["resolutions"]


# ===================================================================
# discover_all_printers_capabilities
# ===================================================================

class TestDiscoverAllPrinters:
    """Test bulk discovery of all printer capabilities."""

    def test_discover_all_returns_dict(self, mocker):
        """discover_all_printers_capabilities returns a dict."""
        # Mock get_printers to avoid actual enumeration
        mock_printers = [
            {"name": "PrinterA"},
            {"name": "PrinterB"},
        ]
        mocker.patch("capabilities.printer.get_printers", return_value=mock_printers)
        mocker.patch("capabilities.discover_capabilities", return_value={
            "trays": ["Auto"],
            "media_sizes": ["A4"],
            "color_modes": ["RGB"],
            "duplex": ["None"],
            "resolutions": ["600dpi"],
        })

        from capabilities import discover_all_printers_capabilities
        result = discover_all_printers_capabilities()
        assert isinstance(result, dict)
        assert "PrinterA" in result
        assert "PrinterB" in result

    def test_discover_all_handles_errors(self, mocker):
        """If discover_capabilities raises, the error is captured."""
        mocker.patch("capabilities.printer.get_printers", return_value=[
            {"name": "BrokenPrinter"},
        ])
        mocker.patch("capabilities.discover_capabilities",
                      side_effect=RuntimeError("No such printer"))

        from capabilities import discover_all_printers_capabilities
        result = discover_all_printers_capabilities()
        assert "BrokenPrinter" in result
        assert "error" in result["BrokenPrinter"]
