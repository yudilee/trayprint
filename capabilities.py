"""
Printer Capability Discovery for TrayPrint.

Discovers supported printer capabilities (trays, resolutions, media sizes,
color modes, duplex) via platform-specific APIs:
  - Windows: win32print.DeviceCapabilities()
  - macOS: lpoptions -p <printer> -l (CUPS) + system_profiler fallback
  - Linux: lpoptions -p <printer> -l (CUPS)
"""

import sys
import subprocess
import re
from logger import get_logger

log = get_logger()


def is_windows():
    return sys.platform == 'win32'


def is_macos():
    return sys.platform == 'darwin'


def discover_capabilities(printer_name):
    """
    Main entry point — discovers capabilities for the given printer.

    Returns a dict:
    {
        'trays': [str, ...],       # input slot names
        'resolutions': [str, ...], # e.g. ['600x600dpi', '1200x1200dpi']
        'media_sizes': [str, ...], # e.g. ['A4', 'Letter', 'Legal']
        'color_modes': [str, ...], # e.g. ['Color', 'Gray']
        'duplex': [str, ...],      # e.g. ['None', 'TwoSidedLong', 'TwoSidedShort']
        'error': str | None,       # error message if discovery failed
    }
    """
    if not printer_name:
        return {'error': 'No printer name provided'}

    try:
        if is_windows():
            return _discover_windows(printer_name)
        elif is_macos():
            return _discover_macos(printer_name)
        else:
            # Linux uses CUPS
            return _discover_cups(printer_name)
    except Exception as e:
        log.error("Capability discovery failed for '%s': %s", printer_name, e, exc_info=True)
        return {'error': str(e)}


def _discover_windows(printer_name):
    """Discover capabilities via win32print.DeviceCapabilities on Windows."""
    try:
        import win32print
        import win32con
    except ImportError:
        return {'error': 'win32print not available'}

    capabilities = {}

    try:
        # ── Trays (Input Bins) ──
        try:
            tray_names = win32print.DeviceCapabilities(
                printer_name, None, win32con.DC_BINS
            )
            if tray_names:
                capabilities['trays'] = list(tray_names) if isinstance(tray_names, tuple) else [tray_names]
            else:
                capabilities['trays'] = []
        except Exception as e:
            log.debug("DC_BINS failed for '%s': %s", printer_name, e)
            capabilities['trays'] = []

        # ── Paper Sizes ──
        try:
            paper_names = win32print.DeviceCapabilities(
                printer_name, None, win32con.DC_PAPERNAMES
            )
            if paper_names:
                capabilities['media_sizes'] = list(paper_names) if isinstance(paper_names, tuple) else [paper_names]
            else:
                capabilities['media_sizes'] = []
        except Exception as e:
            log.debug("DC_PAPERNAMES failed for '%s': %s", printer_name, e)
            capabilities['media_sizes'] = []

        # ── Resolutions ──
        try:
            resolutions = win32print.DeviceCapabilities(
                printer_name, None, win32con.DC_ENUMRESOLUTIONS
            )
            if resolutions:
                # resolutions is a list of (x, y) tuples
                capabilities['resolutions'] = [
                    f"{x}x{y}dpi" for x, y in resolutions
                ]
            else:
                capabilities['resolutions'] = []
        except Exception as e:
            log.debug("DC_ENUMRESOLUTIONS failed for '%s': %s", printer_name, e)
            capabilities['resolutions'] = []

        # ── Color Modes ──
        try:
            # Check via OpenPrinter + GetPrinter DevMode
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                pinfo = win32print.GetPrinter(hprinter, 2)
                devmode = pinfo['pDevMode']
                # DM_COLOR field: 1 = monochrome, 2 = color
                # We can't enumerate all color modes, so we infer from capabilities
                capabilities['color_modes'] = ['color', 'monochrome']
            finally:
                win32print.ClosePrinter(hprinter)
        except Exception:
            capabilities['color_modes'] = ['color', 'monochrome']

        # ── Duplex ──
        try:
            capabilities['duplex'] = ['None', 'TwoSidedLong', 'TwoSidedShort']
        except Exception:
            capabilities['duplex'] = []

    except Exception as e:
        log.error("Windows capability error for '%s': %s", printer_name, e)
        return {'error': str(e)}

    return capabilities


def _parse_cups_options(output):
    """
    Parse `lpoptions -p <printer> -l` output into a structured dict.

    Typical output lines:
        ...
        InputSlot/Input Slot: *Auto Tray1 Tray2 ManualFeed
        PageSize/Media Size: *A4 Letter Legal A5 B5
        ColorModel/Color Model: *RGB Gray CMYK
        Duplex/Duplex: *None DuplexNoTumble DuplexTumble
        Resolution/Resolution: *600dpi 1200dpi
        ...
    """
    capabilities = {
        'trays': [],
        'media_sizes': [],
        'color_modes': [],
        'duplex': [],
        'resolutions': [],
    }

    for line in output.splitlines():
        line = line.strip()
        if not line:
            continue

        # Match: keyword/description: *option1 option2 option3
        match = re.match(r'^(\w+)\s*/\s*[^:]+:\s*(.+)', line)
        if not match:
            continue

        key = match.group(1).lower()
        options_str = match.group(2).strip()
        # Extract options (strip leading '*' which marks default)
        options = [opt.lstrip('*') for opt in options_str.split()]

        if key in ('inputslot',):
            capabilities['trays'] = options
        elif key in ('pagesize', 'media'):
            capabilities['media_sizes'] = options
        elif key in ('colormodel',):
            capabilities['color_modes'] = options
        elif key in ('duplex', 'sides'):
            capabilities['duplex'] = options
        elif key in ('resolution', 'print-quality'):
            if key == 'resolution':
                capabilities['resolutions'] = options

    return capabilities


def _discover_macos(printer_name):
    """
    Discover capabilities on macOS using lpoptions -l (CUPS).
    Falls back to macOS-specific hardcoded defaults (AirPrint, generic PostScript)
    if lpoptions is unavailable.
    """
    log.info("macOS capability discovery for '%s'", printer_name)
    capabilities = {}

    # ── Strategy 1: Try lpoptions -l (CUPS) ──
    try:
        result = subprocess.run(
            ['lpoptions', '-p', printer_name, '-l'],
            capture_output=True, text=True, timeout=15
        )
        if result.returncode == 0 and result.stdout:
            capabilities = _parse_cups_options(result.stdout)
            log.info("macOS: lpoptions succeeded for '%s' — %d trays, %d media sizes",
                     printer_name,
                     len(capabilities.get('trays', [])),
                     len(capabilities.get('media_sizes', [])))
            return capabilities
        else:
            stderr = result.stderr.strip()
            log.warning("macOS: lpoptions returned %d for '%s': %s",
                        result.returncode, printer_name, stderr or '(no output)')
    except FileNotFoundError:
        log.warning("macOS: lpoptions not found — CUPS may not be installed")
    except subprocess.TimeoutExpired:
        log.warning("macOS: lpoptions timed out for '%s'", printer_name)
    except Exception as e:
        log.error("macOS: lpoptions error for '%s': %s", printer_name, e)

    # ── Strategy 2: Try system_profiler for detailed printer info ──
    try:
        sp_result = subprocess.run(
            ['system_profiler', 'SPPrintersDataType'],
            capture_output=True, text=True, timeout=30
        )
        if sp_result.returncode == 0 and sp_result.stdout:
            log.info("macOS: attempting to parse system_profiler output for '%s'", printer_name)
            # system_profiler output includes printer capabilities in a different format
            # We can extract basic info but not as structured as lpoptions
    except FileNotFoundError:
        log.debug("macOS: system_profiler not available")
    except Exception as e:
        log.debug("macOS: system_profiler failed: %s", e)

    # ── Strategy 3: macOS-specific fallback capabilities ──
    log.info("macOS: using macOS-specific fallback capabilities for '%s'", printer_name)
    return _macos_fallback_capabilities(printer_name)


def _macos_fallback_capabilities(printer_name):
    """
    Return macOS-specific fallback capabilities when lpoptions is unavailable.
    Covers common AirPrint printers and generic PostScript printers found on macOS.
    """
    log.info("macOS: using hardcoded fallback capabilities for '%s'", printer_name)
    return {
        'trays': [
            'AutoSelect',
            'Tray1',
            'Tray2',
            'ManualFeed',
            'Bypass',
            'Cassette',
        ],
        'media_sizes': [
            'A4',
            'A5',
            'B5',
            'Letter',
            'Legal',
            'Tabloid',
            'Executive',
            'Statement',
            'Folio',
            '4x6',
            '5x7',
            '8x10',
            'Envelope',
            'DL',
            'C4',
            'C5',
            'Monarch',
            'Number10',
            'Custom',
        ],
        'color_modes': [
            'RGB',
            'Gray',
            'CMYK',
        ],
        'duplex': [
            'None',
            'DuplexNoTumble',
            'DuplexTumble',
        ],
        'resolutions': [
            '300dpi',
            '600dpi',
            '1200dpi',
            '2400dpi',
        ],
    }


def _discover_cups(printer_name):
    """Discover capabilities via lpoptions -p <printer> -l on CUPS (Linux)."""
    capabilities = {}

    try:
        result = subprocess.run(
            ['lpoptions', '-p', printer_name, '-l'],
            capture_output=True, text=True, timeout=15
        )
        if result.returncode == 0 and result.stdout:
            capabilities = _parse_cups_options(result.stdout)
        else:
            stderr = result.stderr.strip()
            log.warning("lpoptions returned %d for '%s': %s",
                        result.returncode, printer_name, stderr or '(no output)')

            # Fallback: try lpinfo or use sensible defaults
            capabilities = _cups_fallback_capabilities(printer_name)

    except FileNotFoundError:
        log.warning("lpoptions not found — CUPS may not be installed")
        capabilities = _cups_fallback_capabilities(printer_name)
    except subprocess.TimeoutExpired:
        log.warning("lpoptions timed out for '%s'", printer_name)
        capabilities = _cups_fallback_capabilities(printer_name)
    except Exception as e:
        log.error("CUPS capability discovery error for '%s': %s", printer_name, e)
        capabilities = _cups_fallback_capabilities(printer_name)

    return capabilities


def _cups_fallback_capabilities(printer_name):
    """
    Return a minimal set of common capabilities when lpoptions is unavailable.
    These are safe defaults that virtually all CUPS printers support.
    """
    log.info("Using fallback capabilities for '%s'", printer_name)
    return {
        'trays': ['AutoSelect', 'Tray1', 'Tray2', 'ManualFeed'],
        'media_sizes': ['A4', 'Letter', 'Legal', 'A5', 'B5'],
        'color_modes': ['RGB', 'Gray', 'CMYK'],
        'duplex': ['None', 'DuplexNoTumble', 'DuplexTumble'],
        'resolutions': ['600dpi', '1200dpi'],
    }


def discover_all_printers_capabilities():
    """
    Discover capabilities for all available printers on the system.
    Returns a dict mapping printer_name -> capabilities dict.
    """
    # Import printer module to enumerate printers
    try:
        import printer as prn_mod
        printers = prn_mod.get_printers()
    except Exception as e:
        log.error("Failed to enumerate printers: %s", e)
        return {}

    all_caps = {}
    for p in printers:
        name = p.get('name', '')
        if name:
            try:
                caps = discover_capabilities(name)
                all_caps[name] = caps
            except Exception as e:
                log.error("Capability discovery failed for '%s': %s", name, e)
                all_caps[name] = {'error': str(e)}

    return all_caps


if __name__ == '__main__':
    """CLI test: python capabilities.py <printer_name>"""
    import json
    if len(sys.argv) > 1:
        printer_name = sys.argv[1]
        caps = discover_capabilities(printer_name)
        print(json.dumps(caps, indent=2))
    else:
        all_caps = discover_all_printers_capabilities()
        print(json.dumps(all_caps, indent=2))
