import sys
import os
import subprocess
import platform
import tempfile
import re
from logger import get_logger

# Windows-specific imports
try:
    import win32print
    import win32con
    import win32api
    import win32ui
    _win32_import_error = None
except Exception as e:
    win32print = None
    win32con = None
    win32api = None
    win32ui = None
    _win32_import_error = str(e)

log = get_logger()

def is_windows():
    return sys.platform == 'win32'

def is_macos():
    return sys.platform == 'darwin'


# ─────────────────────────────────────────────
#  macOS Paper Size Mapping
# ─────────────────────────────────────────────

MACOS_PAPER_MAP = {
    # Standard ISO sizes
    'A0': 'A0',
    'A1': 'A1',
    'A2': 'A2',
    'A3': 'A3',
    'A4': 'A4',
    'A5': 'A5',
    'A6': 'A6',
    'B0': 'B0',
    'B1': 'B1',
    'B2': 'B2',
    'B3': 'B3',
    'B4': 'B4',
    'B5': 'B5',
    # North American sizes
    'Letter': 'Letter',
    'Legal': 'Legal',
    'Tabloid': 'Tabloid',
    'Ledger': 'Ledger',
    'Executive': 'Executive',
    'Statement': 'Statement',
    'Folio': 'Folio',
    # Envelope sizes
    'C4': 'C4',
    'C5': 'C5',
    'C6': 'C6',
    'DL': 'DL',
    'Monarch': 'Monarch',
    'Number 10': 'Number10',
    'Number 9': 'Number9',
    # Japanese
    'B5 (JIS)': 'B5',
    'B4 (JIS)': 'B4',
    # Photo sizes
    '4x6': '4x6',
    '5x7': '5x7',
    '8x10': '8x10',
    # Custom / continuous
    'Custom': 'Custom',
}

# macOS-specific paper name aliases (CUPS names → macOS driver names)
MACOS_PAPER_ALIASES = {
    'halfletter': 'Statement',
    'half letter': 'Statement',
    'Half Letter': 'Statement',
    'F4': 'Folio',
    'f4': 'Folio',
    'A3': 'A3',
    'A4': 'A4',
    'A5': 'A5',
    'B4': 'B4',
    'B5': 'B5',
    'letter': 'Letter',
    'legal': 'Legal',
    'tabloid': 'Tabloid',
    'ledger': 'Ledger',
    'executive': 'Executive',
    'statement': 'Statement',
    'folio': 'Folio',
}


def _get_sumatra_path():
    """Find SumatraPDF.exe, handling PyInstaller --onefile bundles."""
    candidates = []
    
    # 1. PyInstaller bundle extraction directory
    if getattr(sys, 'frozen', False):
        candidates.append(os.path.join(sys._MEIPASS, 'SumatraPDF.exe'))
        # 2. Next to the exe itself (e.g., D:\trayprint\dist\SumatraPDF.exe)
        candidates.append(os.path.join(os.path.dirname(sys.executable), 'SumatraPDF.exe'))
    
    # 3. Next to the source file (development mode)
    candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'SumatraPDF.exe'))
    
    for path in candidates:
        if os.path.exists(path):
            log.info("Found SumatraPDF at: %s", path)
            return path
    
    log.warning("SumatraPDF.exe NOT FOUND in any of: %s", candidates)
    return None


# ─────────────────────────────────────────────
#  Printer Discovery
# ─────────────────────────────────────────────

def get_printers():
    """Returns a list of enriched printer dicts: name, is_default, status."""
    printers = []
    default_name = get_default_printer()

    if is_windows():
        if not win32print:
            log.error("Windows printer support missing: %s", _win32_import_error or "win32print not found")
            return []
            
        try:
            # Level 2 provides pPrinterName, pLocation, Status, etc.
            # We use a broad set of flags to find local, network, and shared printers.
            flags = (win32print.PRINTER_ENUM_LOCAL | 
                     win32print.PRINTER_ENUM_CONNECTIONS | 
                     win32print.PRINTER_ENUM_NETWORK)
            
            log.debug("Enumerating Windows printers with flags: %s", flags)
            printer_info = win32print.EnumPrinters(flags, None, 2)
            
            # Fallback for some environments (Level 5 is simpler/faster for local printers)
            if not printer_info:
                log.debug("EnumPrinters Level 2 returned 0, trying Level 5 fallback...")
                printer_info = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 5)
                # Level 5 has different keys, but we only need pPrinterName
                for info in printer_info:
                    name = info.get('pPrinterName') or info.get('PrinterName')
                    if name:
                        printers.append({
                            'name': name,
                            'is_default': (name == default_name),
                            'status': 'Ready (L5)',
                            'location': '',
                        })
                if printers:
                    log.info("Found %d printer(s) via Level 5 fallback", len(printers))
                    return printers

            for info in printer_info:
                name = info.get('pPrinterName')
                if not name: continue
                
                status_code = info.get('Attributes', 0) # Use Attributes as a secondary source
                status_bits = info.get('Status', 0)
                
                status_list = []
                if status_bits == 0:
                    status = 'Ready'
                else:
                    if status_bits & 128: status_list.append("Offline")
                    if status_bits & 8: status_list.append("Paper Jam")
                    if status_bits & 16: status_list.append("Out of Paper")
                    if status_bits & 2: status_list.append("Error")
                    if status_bits & 1: status_list.append("Paused")
                    if status_bits & 131072: status_list.append("Toner Low")
                    if status_bits & 1024: status_list.append("Printing")
                    status = ", ".join(status_list) if status_list else f"Code {status_bits}"

                printers.append({
                    'name': name,
                    'is_default': (name == default_name),
                    'status': status,
                    'location': info.get('pLocation', ''),
                })
        except Exception as e:
            log.error("Error enumerating windows printers: %s", e, exc_info=True)
    elif is_macos():
        try:
            import platform_darwin
            printers = platform_darwin.get_printers_macos()
            log.info("macOS: found %d printer(s) via platform_darwin", len(printers))
        except Exception as e:
            log.warning("platform_darwin printer enumeration failed: %s", e)
            log.info("macOS: falling back to lpstat -a for printer discovery")

        # If platform_darwin returned nothing, try lpstat directly
        if not printers:
            try:
                result = subprocess.run(
                    ['lpstat', '-a'],
                    capture_output=True, text=True, timeout=10
                )
                if result.returncode == 0:
                    for line in result.stdout.splitlines():
                        if line.strip():
                            parts = line.split()
                            if len(parts) > 0:
                                name = parts[0]
                                status = 'accepting' if 'accepting' in line.lower() else 'unknown'
                                printers.append({
                                    'name': name,
                                    'is_default': (name == default_name),
                                    'status': status,
                                    'location': '',
                                })
                    log.info("macOS: found %d printer(s) via lpstat -a", len(printers))
            except Exception as e2:
                log.error("macOS lpstat -a failed: %s", e2)

        # Try system_profiler for detailed printer info (macOS-specific)
        if printers:
            try:
                sp_result = subprocess.run(
                    ['system_profiler', 'SPPrintersDataType'],
                    capture_output=True, text=True, timeout=30
                )
                if sp_result.returncode == 0:
                    # Parse system_profiler output to enrich printer info
                    current_name = None
                    for line in sp_result.stdout.splitlines():
                        line_stripped = line.strip()
                        # Match: "Name: PrinterName"
                        if line_stripped.startswith('Name:'):
                            current_name = line_stripped.split(':', 1)[1].strip()
                        elif line_stripped.startswith('Location:') and current_name:
                            location = line_stripped.split(':', 1)[1].strip()
                            for p in printers:
                                if p['name'] == current_name:
                                    p['location'] = location
                                    break
                            current_name = None
                    log.info("macOS: enriched printer info via system_profiler")
            except FileNotFoundError:
                log.debug("macOS: system_profiler not available (non-macOS or restricted)")
            except Exception as e3:
                log.debug("macOS: system_profiler enrichment failed: %s", e3)
    else:
        try:
            result = subprocess.run(['lpstat', '-a'], capture_output=True, text=True, check=True)
            for line in result.stdout.splitlines():
                if line.strip():
                    parts = line.split()
                    if len(parts) > 0:
                        name = parts[0]
                        status = 'accepting' if 'accepting' in line.lower() else 'unknown'
                        printers.append({
                            'name': name,
                            'is_default': (name == default_name),
                            'status': status,
                            'location': '',
                        })
        except Exception as e:
            log.error("Error enumerating unix printers: %s", e)

    log.info("Found %d printer(s), default=%s", len(printers), default_name)
    return printers


def get_default_printer():
    """Returns the name of the OS default printer."""
    if is_windows():
        if not win32print:
            log.error("Cannot get default printer: win32print import error: %s", _win32_import_error)
            return ''
        try:
            return win32print.GetDefaultPrinter()
        except Exception:
            return ''
    elif is_macos():
        try:
            import platform_darwin
            return platform_darwin.get_default_printer_macos()
        except Exception:
            return ''
    else:
        try:
            result = subprocess.run(['lpstat', '-d'], capture_output=True, text=True)
            # Output: "system default destination: PRINTER_NAME"
            if result.returncode == 0 and ':' in result.stdout:
                return result.stdout.split(':')[-1].strip()
        except Exception:
            pass
    return ''


# ─────────────────────────────────────────────
#  Build OS-specific option flags
# ─────────────────────────────────────────────

def _build_lp_options(options):
    """Converts an options dict into lp command-line flags for CUPS."""
    args = []
    if not options:
        return args

    # 1. Basic Options
    copies = options.get('copies')
    if copies and int(copies) > 1:
        args += ['-n', str(int(copies))]

    # 2. Paper Size & Custom Dimensions
    paper = options.get('paper_size')
    width_mm = options.get('paper_width_mm')
    height_mm = options.get('paper_height_mm')
    
    if width_mm and height_mm:
        # Custom size is the best for continuous forms / labels
        args += ['-o', f'media=Custom.{width_mm}x{height_mm}mm']
    elif paper:
        # Standard size name
        args += ['-o', f'media={paper}']

    # 3. Orientation (3=Portrait, 4=Landscape)
    orientation = options.get('orientation')
    if orientation == 'landscape':
        args += ['-o', 'orientation-requested=4']
    else:
        args += ['-o', 'orientation-requested=3']

    # 4. Margins (Convert mm to Points: 1mm = 2.83465 pts)
    m_top = options.get('margin_top', 0)
    m_bottom = options.get('margin_bottom', 0)
    m_left = options.get('margin_left', 0)
    m_right = options.get('margin_right', 0)
    
    if any([m_top, m_bottom, m_left, m_right]):
        args += [
            '-o', f'page-top={int(float(m_top) * 2.83465)}',
            '-o', f'page-bottom={int(float(m_bottom) * 2.83465)}',
            '-o', f'page-left={int(float(m_left) * 2.83465)}',
            '-o', f'page-right={int(float(m_right) * 2.83465)}'
        ]

    # 4b. Fit to Page
    if options.get('fit_to_page') or any([m_top, m_bottom, m_left, m_right]):
        args += ['-o', 'fit-to-page']

    # 5. Duplex
    duplex = options.get('duplex')
    if duplex == 'two-sided-long':
        args += ['-o', 'sides=two-sided-long-edge']
    elif duplex == 'two-sided-short':
        args += ['-o', 'sides=two-sided-short-edge']

    # 6. Page Range
    page_range = options.get('page_range')
    if page_range:
        args += ['-o', f'page-ranges={page_range}']

    # ─────────────────────────────────────────────
    # 7. Tray Source (InputSlot)
    # ─────────────────────────────────────────────
    tray_source = options.get('tray_source')
    if tray_source:
        # CUPS InputSlot values typically match the tray name
        # e.g., "AutoSelect", "Tray1", "Tray2", "ManualFeed"
        args += ['-o', f'InputSlot={tray_source}']
        log.debug("CUPS: InputSlot=%s", tray_source)

    # ─────────────────────────────────────────────
    # 8. Color Mode
    # ─────────────────────────────────────────────
    color_mode = options.get('color_mode')
    if color_mode == 'monochrome':
        args += ['-o', 'ColorModel=Gray']
        log.debug("CUPS: ColorModel=Gray")
    elif color_mode == 'color':
        args += ['-o', 'ColorModel=RGB']
        log.debug("CUPS: ColorModel=RGB")

    # ─────────────────────────────────────────────
    # 9. Print Quality (IPP values)
    # ─────────────────────────────────────────────
    # IPP print-quality: 3=draft, 4=normal, 5=high
    quality_map = {'draft': '3', 'normal': '4', 'high': '5'}
    print_quality = options.get('print_quality')
    if print_quality and print_quality in quality_map:
        args += ['-o', f'print-quality={quality_map[print_quality]}']
        log.debug("CUPS: print-quality=%s", quality_map[print_quality])

    # ─────────────────────────────────────────────
    # 10. Media Type
    # ─────────────────────────────────────────────
    media_type = options.get('media_type')
    if media_type:
        args += ['-o', f'media-type={media_type}']
        log.debug("CUPS: media-type=%s", media_type)

    # ─────────────────────────────────────────────
    # 11. Collate
    # ─────────────────────────────────────────────
    collate = options.get('collate')
    if collate is not None:
        args += ['-o', f'Collate={str(collate).lower()}']
        log.debug("CUPS: Collate=%s", str(collate).lower())

    # ─────────────────────────────────────────────
    # 12. Reverse Order
    # ─────────────────────────────────────────────
    reverse_order = options.get('reverse_order')
    if reverse_order:
        args += ['-o', 'OutputOrder=reverse']
        log.debug("CUPS: OutputOrder=reverse")

    # ─────────────────────────────────────────────
    # 13. Scaling Percentage
    # ─────────────────────────────────────────────
    scaling = options.get('scaling_percentage')
    if scaling is not None:
        try:
            scaling_val = int(scaling)
            if 1 <= scaling_val <= 1000:
                args += ['-o', f'scaling={scaling_val}']
                log.debug("CUPS: scaling=%d", scaling_val)
        except (ValueError, TypeError):
            log.warning("CUPS: invalid scaling_percentage value: %s", scaling)

    # ─────────────────────────────────────────────
    # 14. Finishing Options (Staple, Punch, Booklet, Fold)
    # ─────────────────────────────────────────────
    # Staple — IPP "finishings" collection
    staple = options.get('finishing_staple')
    if staple:
        if staple == 'single':
            args += ['-o', 'StapleLocation=SinglePortrait']
            log.debug("CUPS: StapleLocation=SinglePortrait")
        elif staple == 'dual':
            args += ['-o', 'StapleLocation=DualPortrait']
            log.debug("CUPS: StapleLocation=DualPortrait")
        elif staple == 'saddle':
            args += ['-o', 'StapleLocation=SaddleStitch']
            log.debug("CUPS: StapleLocation=SaddleStitch")

    # Punch — IPP "punch" finishings
    punch = options.get('finishing_punch')
    if punch:
        args += ['-o', f'Punch={punch}']
        log.debug("CUPS: Punch=%s", punch)

    # Booklet — 2-up + reverse stack via number-up
    booklet = options.get('finishing_booklet')
    if booklet:
        args += ['-o', 'number-up=2', '-o', 'page-set=all']
        log.debug("CUPS: booklet mode (number-up=2)")

    # Fold — IPP "fold" finishings
    fold = options.get('finishing_fold')
    if fold:
        if fold == 'half':
            args += ['-o', 'Fold=Half']
        elif fold == 'tri-fold':
            args += ['-o', 'Fold=TriFold']
        elif fold == 'z-fold':
            args += ['-o', 'Fold=ZFold']
        log.debug("CUPS: Fold=%s", fold)

    return args


def _find_windows_paper_name(printer_name, w_mm, h_mm):
    """
    Queries the Windows printer driver for all supported paper sizes
    and returns (form_name, paper_id) matching the given width/height.
    
    Uses multiple strategies:
    1. DeviceCapabilities DC_PAPERSIZE (works for most drivers)
    2. EnumForms API (fallback for drivers like Epson LQ that return 0x0 sizes)
    """
    if not is_windows() or not printer_name or not win32print:
        return None, None
    try:
        
        # Get paper names and IDs supported by this specific printer
        names = win32print.DeviceCapabilities(printer_name, "", win32con.DC_PAPERNAMES)
        sizes = win32print.DeviceCapabilities(printer_name, "", win32con.DC_PAPERSIZE)
        ids = win32print.DeviceCapabilities(printer_name, "", win32con.DC_PAPERS)
        
        if not names or not ids:
            log.debug("DeviceCapabilities returned empty for %s", printer_name)
            return None, None
        
        # Build a name→ID lookup for this printer
        printer_papers = {}
        for i, name in enumerate(names):
            if i < len(ids):
                clean_name = name.strip() if isinstance(name, str) else str(name)
                printer_papers[clean_name.lower()] = (clean_name, int(ids[i]))
        
        log.info("Printer '%s' has %d paper sizes available", printer_name, len(names))
        for i, name in enumerate(names):
            if i < len(ids):
                clean = name.strip() if isinstance(name, str) else str(name)
                log.debug("  Paper[%d]: '%s' ID=%s", i, clean, ids[i])
        
        target_w = int(float(w_mm) * 10)  # 0.1mm units
        target_h = int(float(h_mm) * 10)
        
        # ── Strategy 1: Match via DC_PAPERSIZE dimensions ──
        if sizes:
            has_real_sizes = False
            for s in sizes:
                if isinstance(s, (list, tuple)) and (int(s[0]) > 0 or int(s[1]) > 0):
                    has_real_sizes = True
                    break
            
            if has_real_sizes:
                for i, s in enumerate(sizes):
                    if i >= len(names) or i >= len(ids):
                        break
                    if isinstance(s, (list, tuple)):
                        w, h = int(s[0]), int(s[1])
                    else:
                        continue
                    if abs(w - target_w) <= 5 and abs(h - target_h) <= 5:
                        name = names[i].strip() if isinstance(names[i], str) else str(names[i])
                        log.info("Matched via DC_PAPERSIZE: '%s' (ID:%d) (%dx%d)", name, int(ids[i]), w, h)
                        return name, int(ids[i])
                    # Try swapped
                    if abs(h - target_w) <= 5 and abs(w - target_h) <= 5:
                        name = names[i].strip() if isinstance(names[i], str) else str(names[i])
                        log.info("Matched via DC_PAPERSIZE (swapped): '%s' (ID:%d) (%dx%d)", name, int(ids[i]), w, h)
                        return name, int(ids[i])
                log.debug("DC_PAPERSIZE: no dimension match found")
            else:
                log.info("DC_PAPERSIZE returned all zeros — using EnumForms fallback")
        
        # ── Strategy 2: Match via EnumForms API ──
        # EnumForms returns ALL Windows forms with actual dimensions
        # Then we cross-reference with the printer's supported papers
        try:
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                forms = win32print.EnumForms(hprinter)
                log.debug("EnumForms returned %d forms", len(forms))
                
                for form in forms:
                    # form is a dict with keys: Name, Flags, Size, ImageableArea
                    form_name = form.get('Name', '')
                    form_size = form.get('Size', {})
                    # Size is in 0.001mm (thousandths of mm)
                    fw = form_size.get('cx', 0) // 100  # convert to 0.1mm
                    fh = form_size.get('cy', 0) // 100
                    
                    if abs(fw - target_w) <= 5 and abs(fh - target_h) <= 5:
                        # Found a matching form! Now find it in the printer's paper list
                        key = form_name.strip().lower()
                        if key in printer_papers:
                            matched_name, matched_id = printer_papers[key]
                            log.info("Matched via EnumForms: '%s' (ID:%d) form_size=%dx%d (0.1mm)",
                                     matched_name, matched_id, fw, fh)
                            return matched_name, matched_id
                        else:
                            log.debug("Form '%s' matches dimensions but not in printer's paper list", form_name)
                    
                    # Try swapped
                    if abs(fh - target_w) <= 5 and abs(fw - target_h) <= 5:
                        key = form_name.strip().lower()
                        if key in printer_papers:
                            matched_name, matched_id = printer_papers[key]
                            log.info("Matched via EnumForms (swapped): '%s' (ID:%d) form_size=%dx%d (0.1mm)",
                                     matched_name, matched_id, fw, fh)
                            return matched_name, matched_id
            finally:
                win32print.ClosePrinter(hprinter)
        except Exception as ef:
            log.warning("EnumForms fallback failed: %s", ef)
        
        log.info("No paper form matched dimensions %.1f x %.1f mm", w_mm, h_mm)
                
    except Exception as e:
        log.warning("Error finding Windows paper name: %s", e, exc_info=True)
    return None, None



from contextlib import contextmanager

@contextmanager
def windows_printer_override(printer_name, options):
    """
    Context manager that temporarily overrides the printer's DEFAULT DevMode 
    at the OS level to force paper size, then restores it.
    """
    if not is_windows() or not printer_name or not options or not win32print:
        yield
        return

    try:
        
        # Open printer with administrative access to change settings
        # Use PRINTER_ALL_ACCESS if possible, or fall back to PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE
        try:
            hprinter = win32print.OpenPrinter(printer_name, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
        except Exception:
            hprinter = win32print.OpenPrinter(printer_name, {"DesiredAccess": win32con.PRINTER_ACCESS_ADMINISTER | win32con.PRINTER_ACCESS_USE})
            
        try:
            # 1. Backup original settings
            pinfo = win32print.GetPrinter(hprinter, 2)
            original_devmode = pinfo['pDevMode']
            
            # 2. Find matching paper index/name
            w_mm = options.get('paper_width_mm')
            h_mm = options.get('paper_height_mm')
            paper_name, paper_id = _find_windows_paper_name(printer_name, w_mm, h_mm)
            
            # 3. Create modified DevMode
            # We must use DocumentProperties to correctly modify a DevMode object
            new_devmode = win32print.DocumentProperties(0, hprinter, printer_name, original_devmode, original_devmode, 0)
            
            modified = False
            if paper_id:
                new_devmode.PaperSize = paper_id
                new_devmode.Fields |= win32con.DM_PAPERSIZE
                modified = True
            
            if w_mm and h_mm:
                new_devmode.PaperWidth = int(float(w_mm) * 10)
                new_devmode.PaperLength = int(float(h_mm) * 10)
                new_devmode.Fields |= (win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH)
                modified = True
                
            orientation = options.get('orientation')
            if orientation == 'landscape':
                new_devmode.Orientation = win32con.DMORIENT_LANDSCAPE
                new_devmode.Fields |= win32con.DM_ORIENTATION
                modified = True
            elif orientation == 'portrait':
                new_devmode.Orientation = win32con.DMORIENT_PORTRAIT
                new_devmode.Fields |= win32con.DM_ORIENTATION
                modified = True

            if modified:
                log.info("Applying temporary Windows DevMode override: paper=%s, id=%s, orient=%s", paper_name, paper_id, orientation)
                pinfo['pDevMode'] = new_devmode
                win32print.SetPrinter(hprinter, 2, pinfo, 0)
                
            yield # Run the SumatraPDF command now
            
        finally:
            if modified:
                log.info("Restoring original Windows printer settings.")
                pinfo['pDevMode'] = original_devmode
                win32print.SetPrinter(hprinter, 2, pinfo, 0)
            win32print.ClosePrinter(hprinter)
            
    except Exception as e:
        log.warning("Windows printer override failed: %s", e)
        yield


def _build_sumatra_options(options, printer_name=None):
    """Builds SumatraPDF -print-settings string and color flags."""
    parts = []
    extra_args = []
    
    if not options:
        return extra_args

    copies = options.get('copies')
    if copies and int(copies) > 1:
        parts.append(f'{int(copies)}x')

    orientation = options.get('orientation')
    if orientation == 'landscape':
        parts.append('landscape')
    else:
        parts.append('portrait')

    # paper size
    paper = options.get('paper_size')
    w_mm = options.get('paper_width_mm')
    h_mm = options.get('paper_height_mm')

    # Priority 1: Try to match exact dimensions to a Windows Paper Form (critical for Dot-Matrix)
    if is_windows() and printer_name and w_mm and h_mm:
        matched_name, matched_id = _find_windows_paper_name(printer_name, w_mm, h_mm)
        if matched_name:
            paper = matched_name

    if paper and paper.upper() != 'CUSTOM':
        # Fallback mappings for common names that vary between Linux/Windows
        # Example: Hub/Linux says "Half Letter", Windows driver says "Statement"
        mappings = {
            'Half Letter': 'Statement',
            'halfletter': 'Statement',
            'F4': 'Folio',
        }
        paper = mappings.get(paper, paper)
        parts.append(f'paper={paper}')

    duplex = options.get('duplex')
    if duplex and duplex.startswith('two-sided'):
        parts.append('duplex')

    page_range = options.get('page_range')
    if page_range:
        parts.append(page_range)

    # Scaling / Fit to page (Sumatra uses 'shrink' or 'fit')
    if options.get('fit_to_page'):
        parts.append('fit')

    # Reverse order — Sumatra uses 'rev' in print-settings
    if options.get('reverse_order'):
        parts.append('rev')
        log.debug("SumatraPDF: reverse order enabled")

    # Color mode — use -color or -grayscale flag (separate from -print-settings)
    color_mode = options.get('color_mode')
    if color_mode == 'monochrome':
        extra_args.append('-grayscale')
        log.debug("SumatraPDF: grayscale mode")
    elif color_mode == 'color':
        extra_args.append('-color')
        log.debug("SumatraPDF: color mode")

    # ── Finishing Options (SumatraPDF does not support staple/punch/fold natively) ──
    finishing_opts = any([
        options.get('finishing_staple'),
        options.get('finishing_punch'),
        options.get('finishing_booklet'),
        options.get('finishing_fold'),
        options.get('finishing_bind'),
    ])
    if finishing_opts:
        log.warning("SumatraPDF: finishing options (%s) not supported natively; using driver defaults",
                    ', '.join(k for k in ['finishing_staple', 'finishing_punch', 'finishing_booklet',
                                          'finishing_fold', 'finishing_bind'] if options.get(k)))

    if parts:
        return ['-print-settings', ','.join(parts)] + extra_args
    return extra_args


# ─────────────────────────────────────────────
#  Raw Printing
# ─────────────────────────────────────────────

def print_raw(printer_name, data_str, options=None):
    """Sends raw data bypassing the printer driver."""
    log.info("RAW print → printer=%s, data_len=%d, options=%s",
             printer_name, len(data_str) if data_str else 0, options)
    success = False
    error_msg = ""

    if isinstance(data_str, str):
        try:
            raw_bytes = data_str.encode('utf-8')
        except Exception:
            raw_bytes = data_str.encode('latin-1', errors='replace')
    else:
        raw_bytes = data_str

    if is_windows():
        if not win32print:
            return False, f"Windows printer support missing: { _win32_import_error }"
        try:
            
            # Open printer with write access
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                # If we have custom options, try to set the DevMode
                has_devmode_opts = any([
                    options.get('paper_width_mm'),
                    options.get('paper_height_mm'),
                    options.get('orientation'),
                    options.get('tray_source'),
                    options.get('color_mode'),
                    options.get('print_quality'),
                    options.get('collate') is not None,
                    options.get('copies'),
                    options.get('media_type'),
                    options.get('finishing_staple'),
                    options.get('finishing_punch'),
                    options.get('finishing_booklet'),
                    options.get('finishing_fold'),
                    options.get('finishing_bind'),
                ])
                if options and has_devmode_opts:
                    try:
                        # Get default DevMode
                        pinfo = win32print.GetPrinter(hprinter, 2)
                        devmode = pinfo['pDevMode']
                        
                        modified = False
                        
                        # Orientation (1=Portrait, 2=Landscape)
                        if options.get('orientation') == 'landscape':
                            devmode.Orientation = win32con.DMORIENT_LANDSCAPE
                            devmode.Fields |= win32con.DM_ORIENTATION
                            modified = True
                        elif options.get('orientation') == 'portrait':
                            devmode.Orientation = win32con.DMORIENT_PORTRAIT
                            devmode.Fields |= win32con.DM_ORIENTATION
                            modified = True

                        # Paper Size (Width/Height in 0.1mm units)
                        w = options.get('paper_width_mm')
                        h = options.get('paper_height_mm')
                        if w and h:
                            devmode.PaperSize = 0 # Custom
                            devmode.PaperWidth = int(float(w) * 10)
                            devmode.PaperLength = int(float(h) * 10)
                            devmode.Fields |= (win32con.DM_PAPERSIZE | win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH)
                            modified = True

                        # ── Printer Control Fields ──
                        
                        # Default Source (Tray)
                        tray_source = options.get('tray_source')
                        if tray_source:
                            devmode.DefaultSource = _tray_source_to_dmbin(tray_source)
                            devmode.Fields |= win32con.DM_DEFAULTSOURCE
                            modified = True
                        
                        # Color Mode
                        color_mode = options.get('color_mode')
                        if color_mode == 'monochrome':
                            devmode.Color = 1  # DMCOLOR_MONOCHROME
                            devmode.Fields |= win32con.DM_COLOR
                            modified = True
                        elif color_mode == 'color':
                            devmode.Color = 2  # DMCOLOR_COLOR
                            devmode.Fields |= win32con.DM_COLOR
                            modified = True
                        
                        # Print Quality
                        print_quality = options.get('print_quality')
                        if print_quality:
                            quality_map = {'draft': -1, 'low': -2, 'normal': -3, 'high': -4}
                            pq = print_quality.lower()
                            if pq in quality_map:
                                devmode.PrintQuality = quality_map[pq]
                                devmode.Fields |= win32con.DM_PRINTQUALITY
                                modified = True
                        
                        # Collate
                        collate = options.get('collate')
                        if collate is not None:
                            devmode.Collate = 1 if collate else 0
                            devmode.Fields |= win32con.DM_COLLATE
                            modified = True
                        
                        # Copies
                        copies = options.get('copies')
                        if copies:
                            devmode.Copies = int(copies)
                            devmode.Fields |= win32con.DM_COPIES
                            modified = True
                        
                        # Media Type
                        media_type = options.get('media_type')
                        if media_type:
                            media_map = {
                                'plain': 1, 'transparency': 2, 'glossy': 3,
                                'envelope': 4, 'labels': 6,
                            }
                            mt = media_type.lower()
                            if mt in media_map:
                                devmode.MediaType = media_map[mt]
                                devmode.Fields |= win32con.DM_MEDIATYPE
                                modified = True
                        
                        if modified:
                            # Update printer settings for this session
                            win32print.DocumentProperties(0, hprinter, printer_name, devmode, devmode, win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER)
                            log.info("Raw print DevMode: Fields=%s, DefaultSource=%s, Color=%s, PrintQuality=%s, Collate=%s, Copies=%s, MediaType=%s",
                                     devmode.Fields, getattr(devmode, 'DefaultSource', 'N/A'),
                                     getattr(devmode, 'Color', 'N/A'), getattr(devmode, 'PrintQuality', 'N/A'),
                                     getattr(devmode, 'Collate', 'N/A'), getattr(devmode, 'Copies', 'N/A'),
                                     getattr(devmode, 'MediaType', 'N/A'))
                    except Exception as de:
                        log.warning("Could not set Windows DevMode: %s", de)

                # Send the print job
                hjob = win32print.StartDocPrinter(hprinter, 1, ("Raw Web Print Job", None, "RAW"))
                try:
                    win32print.StartPagePrinter(hprinter)
                    win32print.WritePrinter(hprinter, raw_bytes)
                    win32print.EndPagePrinter(hprinter)
                finally:
                    win32print.EndDocPrinter(hprinter)
                success = True
            finally:
                win32print.ClosePrinter(hprinter)
        except Exception as e:
            error_msg = str(e)
            log.error("Windows raw print error: %s", e)
    elif is_macos():
        try:
            cmd = ['lp', '-d', printer_name, '-o', 'raw']
            cmd += _build_lp_options(options)
            subprocess.run(cmd, input=raw_bytes, capture_output=True, check=True)
            success = True
        except subprocess.CalledProcessError as e:
            error_msg = e.stderr.decode('utf-8') if e.stderr else str(e)
            log.error("macOS raw print error: %s", error_msg)
        except Exception as e:
            error_msg = str(e)
            log.error("macOS raw print exception: %s", e)
    else:
        try:
            cmd = ['lp', '-d', printer_name, '-o', 'raw']
            cmd += _build_lp_options(options)
            subprocess.run(cmd, input=raw_bytes, capture_output=True, check=True)
            success = True
        except subprocess.CalledProcessError as e:
            error_msg = e.stderr.decode('utf-8') if e.stderr else str(e)
            log.error("Unix raw print error: %s", error_msg)
        except Exception as e:
            error_msg = str(e)
            log.error("Unix raw print exception: %s", e)

    log.info("RAW print result: success=%s", success)
    return success, error_msg


# ─────────────────────────────────────────────
#  PDF Printing
# ─────────────────────────────────────────────

def _tray_source_to_dmbin(tray_source_name):
    """Map tray source display name to Windows DMBIN constant integer."""
    tray_map = {
        'AutoSelect': 7,    # DMBIN_AUTO
        'Upper': 1,         # DMBIN_UPPER
        'Tray1': 1,
        'Lower': 2,         # DMBIN_LOWER
        'Tray2': 2,
        'Middle': 3,        # DMBIN_MIDDLE
        'Tray3': 3,
        'Manual': 4,        # DMBIN_MANUAL
        'ManualFeed': 4,
        'Envelope': 5,      # DMBIN_ENVELOPE
        'EnvManual': 6,     # DMBIN_ENVMANUAL
        'Auto': 7,          # DMBIN_AUTO
        'Tractor': 8,       # DMBIN_TRACTOR
        'SmallFmt': 9,      # DMBIN_SMALLFMT
        'LargeFmt': 10,     # DMBIN_LARGEFMT
        'LargeCapacity': 11,
        'Cassette': 14,     # DMBIN_CASSETTE
        'FormSource': 15,   # DMBIN_FORMSOURCE
    }
    # Try exact match, then case-insensitive fallback
    if tray_source_name in tray_map:
        return tray_map[tray_source_name]
    lower_name = tray_source_name.lower()
    for key, value in tray_map.items():
        if key.lower() == lower_name:
            return value
    log.warning("Unknown tray source '%s', defaulting to AutoSelect (7)", tray_source_name)
    return 7  # DMBIN_AUTO


def _create_devmode_for_options(printer_name, options):
    """
    Creates a DEVMODE structure with paper size AND printer control fields for win32print.
    Returns (devmode, paper_name) or (None, None) on failure.
    """
    if not is_windows() or not options or not win32print:
        return None, None
    
    w_mm = options.get('paper_width_mm')
    h_mm = options.get('paper_height_mm')
    orientation = options.get('orientation')
    
    # Check for printer control fields
    tray_source = options.get('tray_source')
    color_mode = options.get('color_mode')
    print_quality = options.get('print_quality')
    collate = options.get('collate')
    copies = options.get('copies')
    media_type = options.get('media_type')
    
    # Nothing to customize
    has_paper = bool(w_mm or h_mm or orientation)
    has_controls = any([
        tray_source,
        color_mode,
        print_quality,
        collate is not None,
        copies,
        media_type,
    ])
    if not has_paper and not has_controls:
        return None, None
    
    try:
        
        hprinter = win32print.OpenPrinter(printer_name)
        try:
            # Get the current default DevMode from the printer
            pinfo = win32print.GetPrinter(hprinter, 2)
            devmode = pinfo['pDevMode']
            log.info("Got default DevMode: PaperSize=%s, W=%s, H=%s, Orient=%s, Fields=%s",
                     devmode.PaperSize, devmode.PaperWidth, devmode.PaperLength,
                     devmode.Orientation, devmode.Fields)
            
            paper_name = None
            paper_id = None
            
            if w_mm and h_mm:
                paper_name, paper_id = _find_windows_paper_name(printer_name, w_mm, h_mm)
            
            modified = False
            
            # ── Paper Size & Orientation ──
            
            # Set paper size by ID if found (e.g., custom "kuitansi" form)
            if paper_id:
                devmode.PaperSize = paper_id
                devmode.Fields |= win32con.DM_PAPERSIZE

                if w_mm and h_mm:
                    devmode.PaperWidth  = int(float(w_mm) * 10)
                    devmode.PaperLength = int(float(h_mm) * 10)
                    devmode.Fields |= (win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH)

                if orientation == 'landscape':
                    devmode.Orientation = win32con.DMORIENT_LANDSCAPE
                    devmode.Fields |= win32con.DM_ORIENTATION
                    log.info("DevMode: PaperSize=%d (%s), W=%d, H=%d, Orientation=LANDSCAPE",
                             paper_id, paper_name, devmode.PaperWidth, devmode.PaperLength)
                elif orientation == 'portrait':
                    devmode.Orientation = win32con.DMORIENT_PORTRAIT
                    devmode.Fields |= win32con.DM_ORIENTATION
                    log.info("DevMode: PaperSize=%d (%s), W=%d, H=%d, Orientation=PORTRAIT",
                             paper_id, paper_name, devmode.PaperWidth, devmode.PaperLength)
                else:
                    # CLEAR orientation flag — let the form shape control layout
                    devmode.Fields &= ~win32con.DM_ORIENTATION
                    log.info("DevMode: PaperSize=%d (%s), W=%d, H=%d — orientation flag cleared",
                             paper_id, paper_name, devmode.PaperWidth, devmode.PaperLength)

                modified = True
            else:
                if w_mm and h_mm:
                    # No matching form found — use DMPAPER_USER with explicit dimensions
                    devmode.PaperSize = 256  # DMPAPER_USER
                    devmode.Fields |= win32con.DM_PAPERSIZE
                    devmode.PaperWidth = int(float(w_mm) * 10)
                    devmode.PaperLength = int(float(h_mm) * 10)
                    devmode.Fields |= (win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH)
                    modified = True
                    log.info("DevMode: DMPAPER_USER W=%d, H=%d (0.1mm units)",
                             devmode.PaperWidth, devmode.PaperLength)
                
                # Only set orientation when no matched form
                if orientation == 'landscape':
                    devmode.Orientation = win32con.DMORIENT_LANDSCAPE
                    devmode.Fields |= win32con.DM_ORIENTATION
                    modified = True
                elif orientation == 'portrait':
                    devmode.Orientation = win32con.DMORIENT_PORTRAIT
                    devmode.Fields |= win32con.DM_ORIENTATION
                    modified = True

            # ── Printer Control Fields (DEVMODE) ──
            
            # 1. Default Source (Tray)
            if tray_source:
                devmode.DefaultSource = _tray_source_to_dmbin(tray_source)
                devmode.Fields |= win32con.DM_DEFAULTSOURCE
                modified = True
                log.info("DevMode: DefaultSource=%d (%s)", devmode.DefaultSource, tray_source)

            # 2. Color Mode
            if color_mode == 'monochrome':
                devmode.Color = 1  # DMCOLOR_MONOCHROME
                devmode.Fields |= win32con.DM_COLOR
                modified = True
                log.info("DevMode: Color=MONOCHROME (1)")
            elif color_mode == 'color':
                devmode.Color = 2  # DMCOLOR_COLOR
                devmode.Fields |= win32con.DM_COLOR
                modified = True
                log.info("DevMode: Color=COLOR (2)")

            # 3. Print Quality
            if print_quality:
                quality_map = {'draft': -1, 'low': -2, 'normal': -3, 'high': -4}
                pq = print_quality.lower()
                if pq in quality_map:
                    devmode.PrintQuality = quality_map[pq]
                    devmode.Fields |= win32con.DM_PRINTQUALITY
                    modified = True
                    log.info("DevMode: PrintQuality=%d (%s)", devmode.PrintQuality, pq)

            # 4. Collate
            if collate is not None:
                devmode.Collate = 1 if collate else 0
                devmode.Fields |= win32con.DM_COLLATE
                modified = True
                log.info("DevMode: Collate=%d", devmode.Collate)

            # 5. Copies
            if copies:
                devmode.Copies = int(copies)
                devmode.Fields |= win32con.DM_COPIES
                modified = True
                log.info("DevMode: Copies=%d", devmode.Copies)

            # 6. Media Type
            if media_type:
                media_map = {
                    'plain': 1,          # DMMEDIA_STANDARD
                    'transparency': 2,   # DMMEDIA_TRANSPARENCY
                    'glossy': 3,         # DMMEDIA_GLOSSY
                    'envelope': 4,       # DMMEDIA_ENVELOPE
                    'labels': 6,
                }
                mt = media_type.lower()
                if mt in media_map:
                    devmode.MediaType = media_map[mt]
                    devmode.Fields |= win32con.DM_MEDIATYPE
                    modified = True
                    log.info("DevMode: MediaType=%d (%s)", devmode.MediaType, mt)

            # ── 7. Finishing Options (DEVMODE does not support staple/punch/fold directly) ──
            # DEVMODE has limited finishing support. We log the request and use
            # DM_PRINTQUALITY-adjacent fields if available, but most staple/punch/fold
            # options are driver-specific extended fields not covered by DEVMODE.
            finishing_booklet = options.get('finishing_booklet')
            if finishing_booklet:
                # Booklet can be approximated via duplex + 2-up in some drivers
                log.info("DevMode: booklet mode requested (driver-dependent)")
            
            if modified:
                if paper_id:
                    log.info("DevMode ready (no validation). PaperSize=%s, W=%s, H=%s, Orient=%s, "
                             "DefaultSource=%s, Color=%s, PrintQuality=%s, Collate=%s, Copies=%s, MediaType=%s, Fields=%s",
                             devmode.PaperSize, devmode.PaperWidth, devmode.PaperLength,
                             devmode.Orientation, getattr(devmode, 'DefaultSource', 'N/A'),
                             getattr(devmode, 'Color', 'N/A'), getattr(devmode, 'PrintQuality', 'N/A'),
                             getattr(devmode, 'Collate', 'N/A'), getattr(devmode, 'Copies', 'N/A'),
                             getattr(devmode, 'MediaType', 'N/A'), devmode.Fields)
                else:
                    # For generic custom sizes, validate through DocumentProperties
                    result = win32print.DocumentProperties(
                        0, hprinter, printer_name, devmode, devmode,
                        win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER
                    )
                    log.info("DevMode validated (result=%s). Final: PaperSize=%s, W=%s, H=%s, Orient=%s, "
                             "DefaultSource=%s, Color=%s, PrintQuality=%s, Collate=%s, Copies=%s, MediaType=%s",
                             result, devmode.PaperSize, devmode.PaperWidth,
                             devmode.PaperLength, devmode.Orientation,
                             getattr(devmode, 'DefaultSource', 'N/A'), getattr(devmode, 'Color', 'N/A'),
                             getattr(devmode, 'PrintQuality', 'N/A'), getattr(devmode, 'Collate', 'N/A'),
                             getattr(devmode, 'Copies', 'N/A'), getattr(devmode, 'MediaType', 'N/A'))
                return devmode, paper_name
        finally:
            win32print.ClosePrinter(hprinter)
    except Exception as e:
        log.warning("Failed to create DevMode: %s", e, exc_info=True)
    return None, None


def _parse_page_range(page_range_str):
    """
    Parse a page range string into a sorted list of 0-based page indices.

    Supported formats:
      - "1-5"       → [0, 1, 2, 3, 4]
      - "1,3,5"     → [0, 2, 4]
      - "1-5,7,9-11" → [0,1,2,3,4,6,8,9,10]
      - None / ""   → None (all pages)
    """
    if not page_range_str:
        return None
    pages = set()
    parts = str(page_range_str).split(',')
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            try:
                start_s, end_s = part.split('-', 1)
                start = int(start_s.strip())
                end = int(end_s.strip())
                if start < 1 or end < 1:
                    log.warning("Invalid page range part '%s': page numbers must be >= 1", part)
                    continue
                pages.update(range(start - 1, end))  # Convert to 0-based
            except (ValueError, TypeError):
                log.warning("Invalid page range part '%s'", part)
        else:
            try:
                p = int(part)
                if p < 1:
                    log.warning("Invalid page number '%s': must be >= 1", part)
                    continue
                pages.add(p - 1)  # Convert to 0-based
            except (ValueError, TypeError):
                log.warning("Invalid page number '%s'", part)
    if not pages:
        return None
    return sorted(pages)


def _print_pdf_windows(printer_name, pdf_path, options):
    """
    Print a PDF directly via Windows GDI + PyMuPDF — no SumatraPDF.

    Strategy:
      1. Build DevMode from options (paper ID=207, W/H dimensions)
      2. Temporarily set printer default to our DevMode via SetPrinter
      3. Create a printer DC (which inherits the default DevMode we just set)
      4. PyMuPDF renders each page at printer DPI → BitBlt onto DC
      5. Restore original printer default

    Supports:
      - Page range (page_range option: "1-5", "1,3,5", "1-5,7,9-11")
      - Copies (copies option)
      - Collate (collate option)
      - GDI mode selection via gdi_mode option
    """
    try:
        import fitz  # PyMuPDF
        import win32print
        import win32ui
        import win32con
        from PIL import Image
        import io

        log.info("GDI print (PyMuPDF): printer=%s, path=%s, options=%s",
                 printer_name, pdf_path, options)

        devmode, paper_name = _create_devmode_for_options(printer_name, options)
        log.info("GDI print: printer=%s paper=%s", printer_name, paper_name)

        # ── Step 1: Temporarily set printer default to our DevMode ──
        original_devmode = None
        hprinter = None
        if devmode:
            # Set copies to 1 in DevMode since GDI printing manually loops to render copies
            devmode.Copies = 1
            devmode.Fields |= win32con.DM_COPIES
            
            hprinter = win32print.OpenPrinter(printer_name,
                {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
            pinfo = win32print.GetPrinter(hprinter, 2)
            original_devmode = pinfo['pDevMode']
            pinfo['pDevMode'] = devmode
            win32print.SetPrinter(hprinter, 2, pinfo, 0)
            log.info("GDI: printer default set to paper=%s (ID=%s), copies set to 1 in devmode", paper_name,
                     devmode.PaperSize if devmode else None)

        try:
            # ── Step 2: Create printer DC (uses our DevMode default) ──
            dc = win32ui.CreateDC()
            dc.CreatePrinterDC(printer_name)

            try:
                pwidth_px  = dc.GetDeviceCaps(win32con.PHYSICALWIDTH)
                pheight_px = dc.GetDeviceCaps(win32con.PHYSICALHEIGHT)
                dpi_x      = dc.GetDeviceCaps(win32con.LOGPIXELSX)
                dpi_y      = dc.GetDeviceCaps(win32con.LOGPIXELSY)
                offset_x   = dc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
                offset_y   = dc.GetDeviceCaps(win32con.PHYSICALOFFSETY)

                log.info("Printer DC: %dx%d px @ %dx%d dpi, offset=%d,%d",
                         pwidth_px, pheight_px, dpi_x, dpi_y, offset_x, offset_y)

                # ── Step 3: Render PDF with PyMuPDF ──
                doc = fitz.open(pdf_path)
                total_pages = len(doc)
                log.info("GDI: opened PDF with %d pages", total_pages)

                # ── Parse page range ──
                page_range_str = (options or {}).get('page_range')
                selected_pages = _parse_page_range(page_range_str)
                if selected_pages is not None:
                    # Filter to valid page numbers within document range
                    valid_pages = [p for p in selected_pages if 0 <= p < total_pages]
                    log.info("GDI: page_range='%s' → %d selected pages (out of %d)",
                             page_range_str, len(valid_pages), total_pages)
                    if not valid_pages:
                        log.warning("GDI: page_range '%s' yielded no valid pages, printing all", page_range_str)
                        valid_pages = list(range(total_pages))
                else:
                    valid_pages = list(range(total_pages))
                    log.info("GDI: no page_range, printing all %d pages", total_pages)

                # ── Copies and collate ──
                copies = max(1, int((options or {}).get('copies', 1)))
                collate = (options or {}).get('collate', False)
                log.info("GDI: copies=%d, collate=%s, pages_to_print=%d",
                         copies, collate, len(valid_pages))

                dc.StartDoc('PrintHub Job')

                if collate:
                    # Collated: print all pages for copy 1, then all pages for copy 2, etc.
                    for copy_idx in range(copies):
                        log.info("GDI: collated copy %d/%d", copy_idx + 1, copies)
                        for page_num in valid_pages:
                            _render_and_print_page(
                                dc, doc, page_num, options,
                                pwidth_px, pheight_px, dpi_x, dpi_y,
                                offset_x, offset_y, copy_idx, copies
                            )
                else:
                    # Non-collated: print all copies of page 1, then all copies of page 2, etc.
                    for page_num in valid_pages:
                        for copy_idx in range(copies):
                            _render_and_print_page(
                                dc, doc, page_num, options,
                                pwidth_px, pheight_px, dpi_x, dpi_y,
                                offset_x, offset_y, copy_idx, copies
                            )

                dc.EndDoc()
                doc.close()
                log.info("GDI print complete: %d page(s) x %d copy(ies), collate=%s",
                         len(valid_pages), copies, collate)
                return True, ""

            finally:
                dc.DeleteDC()

        finally:
            # ── Step 4: Restore original printer default ──
            if hprinter and original_devmode is not None:
                import time as _t
                _t.sleep(2)  # let spooler pick up the job first
                pinfo['pDevMode'] = original_devmode
                win32print.SetPrinter(hprinter, 2, pinfo, 0)
                win32print.ClosePrinter(hprinter)
                log.info("GDI: printer default restored.")

    except ImportError as e:
        log.warning("PyMuPDF not available, using SumatraPDF fallback: %s", e)
        return _print_pdf_sumatra(printer_name, pdf_path, options)
    except Exception as e:
        log.error("GDI print failed: %s", e, exc_info=True)
        log.info("Falling back to SumatraPDF...")
        return _print_pdf_sumatra(printer_name, pdf_path, options)


def _render_and_print_page(dc, doc, page_num, options,
                           pwidth_px, pheight_px, dpi_x, dpi_y,
                           offset_x, offset_y, copy_idx, total_copies):
    """
    Render a single PDF page and send it to the printer DC via StretchDIBits.
    Extracted as a helper to avoid code duplication in collated/non-collated loops.
    """
    import ctypes
    from PIL import Image

    page = doc[page_num]

    # ── Margins: convert mm → pixels ──
    mm_to_px_x = dpi_x / 25.4
    mm_to_px_y = dpi_y / 25.4
    margin_l = int((options.get('margin_left',   0) or 0) * mm_to_px_x) if options else 0
    margin_r = int((options.get('margin_right',  0) or 0) * mm_to_px_x) if options else 0
    margin_t = int((options.get('margin_top',    0) or 0) * mm_to_px_y) if options else 0
    margin_b = int((options.get('margin_bottom', 0) or 0) * mm_to_px_y) if options else 0
    avail_w = max((pwidth_px  - 2 * offset_x) - margin_l - margin_r, 1)
    avail_h = max((pheight_px - 2 * offset_y) - margin_t - margin_b, 1)

    # ── Render at 720dpi (2x oversampling for sharper text) ──
    RENDER_DPI = 720
    import fitz
    mat = fitz.Matrix(RENDER_DPI / 72.0, RENDER_DPI / 72.0)
    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)

    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    img_w, img_h = img.size
    log.info("Copy %d/%d Page %d: rendered %dx%d @ %ddpi, DC=%dx%d dpi, avail=%dx%d px",
             copy_idx + 1, total_copies, page_num + 1, img_w, img_h, RENDER_DPI,
             dpi_x, dpi_y, avail_w, avail_h)

    # ── Orientation: rotate page to match DC shape ──
    dc_is_landscape   = pwidth_px > pheight_px
    page_is_landscape = img_w > img_h
    if dc_is_landscape and not page_is_landscape:
        img = img.rotate(-90, expand=True)
        img_w, img_h = img.size
        log.info("Rotated CW 90° (DC=landscape, page=portrait)")
    elif not dc_is_landscape and page_is_landscape:
        img = img.rotate(90, expand=True)
        img_w, img_h = img.size
        log.info("Rotated CCW 90° (DC=portrait, page=landscape)")
    else:
        log.info("No rotation (DC=%s, page=%s)",
                 'landscape' if dc_is_landscape else 'portrait',
                 'landscape' if page_is_landscape else 'portrait')

    # ── Send via StretchDIBits ──
    # Convert RGB → BGR (Windows GDI 24-bit expects BGR)
    r, g, b = img.split()
    img_bgr = Image.merge('RGB', (b, g, r))

    # Build DWORD-aligned pixel buffer
    row_bytes = img_w * 3
    pad_bytes = (4 - (row_bytes % 4)) % 4
    raw_rgb   = img_bgr.tobytes()
    if pad_bytes:
        padded  = bytearray()
        padding = b'\x00' * pad_bytes
        for row in range(img_h):
            padded += raw_rgb[row * row_bytes:(row + 1) * row_bytes]
            padded += padding
        pixel_data = ctypes.create_string_buffer(bytes(padded))
    else:
        pixel_data = ctypes.create_string_buffer(raw_rgb)
    stride = row_bytes + pad_bytes

    # BITMAPINFOHEADER
    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ('biSize',          ctypes.c_uint32),
            ('biWidth',         ctypes.c_int32),
            ('biHeight',        ctypes.c_int32),
            ('biPlanes',        ctypes.c_uint16),
            ('biBitCount',      ctypes.c_uint16),
            ('biCompression',   ctypes.c_uint32),
            ('biSizeImage',     ctypes.c_uint32),
            ('biXPelsPerMeter', ctypes.c_int32),
            ('biYPelsPerMeter', ctypes.c_int32),
            ('biClrUsed',       ctypes.c_uint32),
            ('biClrImportant',  ctypes.c_uint32),
        ]

    bmi = BITMAPINFOHEADER()
    bmi.biSize          = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.biWidth         = img_w
    bmi.biHeight        = -img_h
    bmi.biPlanes        = 1
    bmi.biBitCount      = 24
    bmi.biCompression   = 0
    bmi.biSizeImage     = stride * img_h
    bmi.biXPelsPerMeter = int(RENDER_DPI / 0.0254)
    bmi.biYPelsPerMeter = int(RENDER_DPI / 0.0254)
    bmi.biClrUsed       = 0
    bmi.biClrImportant  = 0

    gdi32 = ctypes.windll.gdi32
    gdi32.StretchDIBits.argtypes = [
        ctypes.c_void_p,
        ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        ctypes.c_void_p, ctypes.c_void_p,
        ctypes.c_uint, ctypes.c_ulong,
    ]
    gdi32.StretchDIBits.restype = ctypes.c_int

    hdc = dc.GetSafeHdc()

    HALFTONE = 4
    gdi32.SetStretchBltMode(hdc, HALFTONE)
    gdi32.SetBrushOrgEx(hdc, 0, 0, None)

    dc.StartPage()
    result = gdi32.StretchDIBits(
        hdc,
        margin_l, margin_t, avail_w, avail_h,
        0, 0, img_w, img_h,
        pixel_data,
        ctypes.byref(bmi),
        0,          # DIB_RGB_COLORS
        0x00CC0020, # SRCCOPY
    )
    dc.EndPage()
    log.info("StretchDIBits page %d: %dx%d src → %dx%d dst, result=%s",
             page_num + 1, img_w, img_h, avail_w, avail_h, result)


def _print_pdf_sumatra(printer_name, pdf_path, options):
    """Fallback: print via SumatraPDF (used only if PyMuPDF is unavailable)."""
    try:
        sumatra_path = _get_sumatra_path()
        if not sumatra_path:
            import win32api
            win32api.ShellExecute(0, "printto", pdf_path, f'"{printer_name}"', ".", 0)
            return True, ""

        devmode, paper_name = _create_devmode_for_options(printer_name, options)
        if devmode:
            import win32print
            hprinter = win32print.OpenPrinter(printer_name,
                {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
            try:
                pinfo = win32print.GetPrinter(hprinter, 2)
                original_devmode = pinfo['pDevMode']
                pinfo['pDevMode'] = devmode
                win32print.SetPrinter(hprinter, 2, pinfo, 0)
                log.info("SumatraPDF fallback: printer default set to paper=%s", paper_name)
                try:
                    cmd = [sumatra_path, "-print-to", printer_name, "-silent", pdf_path]
                    log.info("SumatraPDF cmd: %s", ' '.join(cmd))
                    subprocess.run(cmd, check=True, timeout=60)
                finally:
                    import time as _time
                    _time.sleep(2)
                    pinfo['pDevMode'] = original_devmode
                    win32print.SetPrinter(hprinter, 2, pinfo, 0)
                    log.info("SumatraPDF fallback: printer default restored.")
            finally:
                win32print.ClosePrinter(hprinter)
        else:
            cmd = [sumatra_path, "-print-to", printer_name, "-silent"]
            cmd += _build_sumatra_options(options, printer_name)
            cmd.append(pdf_path)
            log.info("SumatraPDF cmd (no DevMode): %s", ' '.join(cmd))
            subprocess.run(cmd, check=True, timeout=60)
        return True, ""
    except Exception as e:
        log.error("SumatraPDF fallback failed: %s", e, exc_info=True)
        return False, str(e)


# ─────────────────────────────────────────────
#  macOS PDF Printing
# ─────────────────────────────────────────────

def _build_macos_lp_options(options):
    """
    Converts an options dict into lp command-line flags for macOS CUPS.
    Includes macOS-specific paper name handling and color management.
    """
    args = []
    if not options:
        return args

    # 1. Copies
    copies = options.get('copies')
    if copies and int(copies) > 1:
        args += ['-n', str(int(copies))]

    # 2. Paper Size — map to macOS-compatible names
    paper = options.get('paper_size')
    width_mm = options.get('paper_width_mm')
    height_mm = options.get('paper_height_mm')

    if width_mm and height_mm:
        # Custom size
        args += ['-o', f'media=Custom.{width_mm}x{height_mm}mm']
        log.debug("macOS: custom media=Custom.%sx%smm", width_mm, height_mm)
    elif paper:
        # Map paper name through macOS aliases
        mapped = MACOS_PAPER_ALIASES.get(paper, paper)
        args += ['-o', f'media={mapped}']
        log.debug("macOS: media=%s (original=%s)", mapped, paper)

    # 3. Orientation
    orientation = options.get('orientation')
    if orientation == 'landscape':
        args += ['-o', 'orientation-requested=4']
    else:
        args += ['-o', 'orientation-requested=3']

    # 4. Margins (mm → points: 1mm = 2.83465 pts)
    m_top = options.get('margin_top', 0)
    m_bottom = options.get('margin_bottom', 0)
    m_left = options.get('margin_left', 0)
    m_right = options.get('margin_right', 0)
    if any([m_top, m_bottom, m_left, m_right]):
        args += [
            '-o', f'page-top={int(float(m_top) * 2.83465)}',
            '-o', f'page-bottom={int(float(m_bottom) * 2.83465)}',
            '-o', f'page-left={int(float(m_left) * 2.83465)}',
            '-o', f'page-right={int(float(m_right) * 2.83465)}',
        ]

    # 5. Fit to Page
    if options.get('fit_to_page') or any([m_top, m_bottom, m_left, m_right]):
        args += ['-o', 'fit-to-page']

    # 6. Duplex
    duplex = options.get('duplex')
    if duplex == 'two-sided-long':
        args += ['-o', 'sides=two-sided-long-edge']
    elif duplex == 'two-sided-short':
        args += ['-o', 'sides=two-sided-short-edge']

    # 7. Page Range
    page_range = options.get('page_range')
    if page_range:
        args += ['-o', f'page-ranges={page_range}']

    # 8. Tray Source (InputSlot)
    tray_source = options.get('tray_source')
    if tray_source:
        args += ['-o', f'InputSlot={tray_source}']
        log.debug("macOS: InputSlot=%s", tray_source)

    # 9. Color Mode — macOS-specific ColorModel values
    color_mode = options.get('color_mode')
    if color_mode == 'monochrome':
        args += ['-o', 'ColorModel=Gray']
        log.debug("macOS: ColorModel=Gray")
    elif color_mode == 'color':
        args += ['-o', 'ColorModel=RGB']
        log.debug("macOS: ColorModel=RGB")

    # 10. Print Quality (IPP values)
    quality_map = {'draft': '3', 'normal': '4', 'high': '5'}
    print_quality = options.get('print_quality')
    if print_quality and print_quality in quality_map:
        args += ['-o', f'print-quality={quality_map[print_quality]}']
        log.debug("macOS: print-quality=%s", quality_map[print_quality])

    # 11. Media Type
    media_type = options.get('media_type')
    if media_type:
        args += ['-o', f'media-type={media_type}']
        log.debug("macOS: media-type=%s", media_type)

    # 12. Collate
    collate = options.get('collate')
    if collate is not None:
        args += ['-o', f'Collate={str(collate).lower()}']
        log.debug("macOS: Collate=%s", str(collate).lower())

    # 13. Reverse Order
    reverse_order = options.get('reverse_order')
    if reverse_order:
        args += ['-o', 'OutputOrder=reverse']
        log.debug("macOS: OutputOrder=reverse")

    # 14. Scaling Percentage
    scaling = options.get('scaling_percentage')
    if scaling is not None:
        try:
            scaling_val = int(scaling)
            if 1 <= scaling_val <= 1000:
                args += ['-o', f'scaling={scaling_val}']
                log.debug("macOS: scaling=%d", scaling_val)
        except (ValueError, TypeError):
            log.warning("macOS: invalid scaling_percentage value: %s", scaling)

    # 15. Finishing Options (Staple, Punch, Booklet, Fold)
    staple = options.get('finishing_staple')
    if staple:
        if staple == 'single':
            args += ['-o', 'StapleLocation=SinglePortrait']
        elif staple == 'dual':
            args += ['-o', 'StapleLocation=DualPortrait']
        elif staple == 'saddle':
            args += ['-o', 'StapleLocation=SaddleStitch']
        log.debug("macOS: StapleLocation=%s", staple)

    punch = options.get('finishing_punch')
    if punch:
        args += ['-o', f'Punch={punch}']
        log.debug("macOS: Punch=%s", punch)

    booklet = options.get('finishing_booklet')
    if booklet:
        args += ['-o', 'number-up=2', '-o', 'page-set=all']
        log.debug("macOS: booklet mode (number-up=2)")

    fold = options.get('finishing_fold')
    if fold:
        if fold == 'half':
            args += ['-o', 'Fold=Half']
        elif fold == 'tri-fold':
            args += ['-o', 'Fold=TriFold']
        elif fold == 'z-fold':
            args += ['-o', 'Fold=ZFold']
        log.debug("macOS: Fold=%s", fold)

    return args


def _print_pdf_macos(printer_name, pdf_path, options):
    """
    Print a PDF on macOS using CUPS/lp with macOS-specific options.
    Falls back to 'open' command (Preview.app) if lp is unavailable.
    """
    log.info("macOS PDF print: printer=%s, path=%s, options=%s",
             printer_name, pdf_path, options)

    # Determine temp directory — macOS uses $TMPDIR which points to
    # a per-user temp dir under /var/folders/...
    macos_temp = os.environ.get('TMPDIR', '/tmp')
    log.debug("macOS: using TMPDIR=%s", macos_temp)

    # ── Strategy 1: Use lp (CUPS) ──
    try:
        # Check if lp is available
        subprocess.run(['which', 'lp'], capture_output=True, check=True)

        cmd = ['lp', '-d', printer_name]
        cmd += _build_macos_lp_options(options)
        cmd.append(pdf_path)

        log.info("macOS lp cmd: %s", ' '.join(cmd))
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

        if result.returncode == 0:
            log.info("macOS lp print succeeded: %s", result.stdout.strip())
            return True, ""
        else:
            error_msg = result.stderr.strip() if result.stderr else f"lp returned {result.returncode}"
            log.warning("macOS lp failed: %s", error_msg)
            # Fall through to 'open' fallback

    except FileNotFoundError:
        log.warning("macOS: lp command not found — CUPS may not be installed")
    except subprocess.TimeoutExpired:
        log.warning("macOS: lp timed out after 120s")
    except Exception as e:
        log.warning("macOS: lp exception: %s", e)

    # ── Strategy 2: Fallback to 'open' command (opens in Preview.app) ──
    try:
        log.info("macOS: falling back to 'open' command (Preview.app)")
        cmd = ['open', pdf_path]
        subprocess.run(cmd, capture_output=True, check=True, timeout=30)
        log.info("macOS: opened PDF in Preview.app via 'open' command")
        return True, ""
    except FileNotFoundError:
        error_msg = "Neither 'lp' nor 'open' commands found on macOS"
        log.error(error_msg)
        return False, error_msg
    except subprocess.TimeoutExpired:
        error_msg = "'open' command timed out"
        log.error(error_msg)
        return False, error_msg
    except Exception as e:
        error_msg = f"macOS fallback failed: {e}"
        log.error(error_msg)
        return False, error_msg


def print_pdf(printer_name, pdf_base64, options=None):
    """Decodes a base64 PDF and prints it silently using OS handlers."""
    import base64
    import tempfile
    import os
    import threading
    import time

    log.info("PDF print → printer=%s, b64_len=%d, options=%s",
             printer_name, len(pdf_base64) if pdf_base64 else 0, options)
    success = False
    error_msg = ""
    temp_path = None

    try:
        pdf_bytes = base64.b64decode(pdf_base64)
        fd, temp_path = tempfile.mkstemp(suffix=".pdf")
        with os.fdopen(fd, 'wb') as f:
            f.write(pdf_bytes)

        if is_windows():
            # Check for GDI mode selection
            gdi_mode = (options or {}).get('gdi_mode', 'pymupdf')
            log.info("GDI mode selected: %s", gdi_mode)
            if gdi_mode == 'sumatra':
                log.info("Using SumatraPDF mode (gdi_mode=sumatra)")
                success, error_msg = _print_pdf_sumatra(printer_name, temp_path, options)
            else:
                success, error_msg = _print_pdf_windows(printer_name, temp_path, options)
        elif is_macos():
            success, error_msg = _print_pdf_macos(printer_name, temp_path, options)
        else:
            try:
                cmd = ['lp', '-d', printer_name]
                cmd += _build_lp_options(options)
                cmd.append(temp_path)
                subprocess.run(cmd, capture_output=True, check=True)
                success = True
            except subprocess.CalledProcessError as e:
                error_msg = e.stderr.decode('utf-8') if e.stderr else str(e)
                log.error("Unix PDF Print Error: %s", error_msg)
    except Exception as e:
        error_msg = str(e)
        log.error("PDF Parsing Error: %s", e)

    # Cleanup temp file after 30s (longer wait for Windows spooler)
    if temp_path:
        def cleanup():
            time.sleep(30)
            try:
                if os.path.exists(temp_path):
                    os.unlink(temp_path)
            except Exception:
                pass
        threading.Thread(target=cleanup, daemon=True).start()

    log.info("PDF print result: success=%s", success)
    return success, error_msg


if __name__ == '__main__':
    printers = get_printers()
    print("Available Printers:", printers)


def get_printer_hardware_status(printer_name: str) -> dict:
    """
    Returns normalized hardware status dictionary for telemetry:
    {'state': 'ready'|'paper_out'|'paper_jam'|'door_open'|'no_toner'|'low_toner'|'offline'|'error', 'message': str}
    """
    if not is_windows() or not win32print:
        return {'state': 'ready', 'message': 'Online'}

    try:
        h = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(h, 2)
        win32print.ClosePrinter(h)
        status_bits = info.get('Status', 0)

        if status_bits == 0:
            return {'state': 'ready', 'message': 'Ready'}
        if status_bits & 128:
            return {'state': 'offline', 'message': 'Printer offline'}
        if status_bits & 4194304:
            return {'state': 'door_open', 'message': 'Door or cover open'}
        if status_bits & 16:
            return {'state': 'paper_out', 'message': 'Out of paper'}
        if status_bits & 8:
            return {'state': 'paper_jam', 'message': 'Paper jam'}
        if status_bits & 262144:
            return {'state': 'no_toner', 'message': 'No toner'}
        if status_bits & 131072:
            return {'state': 'low_toner', 'message': 'Toner low'}
        if status_bits & 2:
            return {'state': 'error', 'message': 'General printer error'}
        return {'state': 'ready', 'message': f'Status code {status_bits}'}
    except Exception as e:
        return {'state': 'error', 'message': str(e)}
