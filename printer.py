import sys
import os
import subprocess
from logger import get_logger

log = get_logger()

def is_windows():
    return sys.platform == 'win32'

def is_macos():
    return sys.platform == 'darwin'


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
        try:
            import win32print
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            printer_info = win32print.EnumPrinters(flags, None, 2)
            for info in printer_info:
                name = info['pPrinterName']
                status_code = info.get('Status', 0)
                status = 'idle' if status_code == 0 else f'status:{status_code}'
                printers.append({
                    'name': name,
                    'is_default': (name == default_name),
                    'status': status,
                    'location': info.get('pLocation', ''),
                })
        except Exception as e:
            log.error("Error enumerating windows printers: %s", e)
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
        try:
            import win32print
            return win32print.GetDefaultPrinter()
        except:
            return ''
    else:
        try:
            result = subprocess.run(['lpstat', '-d'], capture_output=True, text=True)
            # Output: "system default destination: PRINTER_NAME"
            if result.returncode == 0 and ':' in result.stdout:
                return result.stdout.split(':')[-1].strip()
        except:
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

    return args


def _find_windows_paper_name(printer_name, w_mm, h_mm):
    """
    Queries the Windows printer driver for all supported paper sizes
    and returns (form_name, paper_id) matching the given width/height.
    
    Uses multiple strategies:
    1. DeviceCapabilities DC_PAPERSIZE (works for most drivers)
    2. EnumForms API (fallback for drivers like Epson LQ that return 0x0 sizes)
    """
    if not is_windows() or not printer_name:
        return None, None
    try:
        import win32print
        import win32con
        
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
    if not is_windows() or not printer_name or not options:
        yield
        return

    try:
        import win32print
        import win32con
        
        # Open printer with administrative access to change settings
        # Use PRINTER_ALL_ACCESS if possible, or fall back to PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE
        try:
            hprinter = win32print.OpenPrinter(printer_name, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
        except:
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
    """Builds SumatraPDF -print-settings string."""
    parts = []
    if not options:
        return parts

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

    if parts:
        return ['-print-settings', ','.join(parts)]
    return []


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
        except:
            raw_bytes = data_str.encode('latin-1', errors='replace')
    else:
        raw_bytes = data_str

    if is_windows():
        try:
            import win32print
            import win32con
            
            # Open printer with write access
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                # If we have custom size options, try to set the DevMode
                if options and (options.get('paper_width_mm') or options.get('paper_height_mm') or options.get('orientation')):
                    try:
                        # Get default DevMode
                        pinfo = win32print.GetPrinter(hprinter, 2)
                        devmode = pinfo['pDevMode']
                        
                        modified = False
                        # Orientation (1=Portrait, 2=Landscape)
                        if options.get('orientation') == 'landscape':
                            devmode.Orientation = win32con.DMORIENT_LANDSCAPE
                            modified = True
                        elif options.get('orientation') == 'portrait':
                            devmode.Orientation = win32con.DMORIENT_PORTRAIT
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
                        
                        if modified:
                            # Update printer settings for this session
                            win32print.DocumentProperties(0, hprinter, printer_name, devmode, devmode, win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER)
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

def _create_devmode_for_options(printer_name, options):
    """
    Creates a DEVMODE structure with the correct paper size for win32print.
    Returns (devmode, paper_name) or (None, None) on failure.
    """
    if not is_windows() or not options:
        return None, None
    
    w_mm = options.get('paper_width_mm')
    h_mm = options.get('paper_height_mm')
    orientation = options.get('orientation')
    
    # Nothing to customize
    if not w_mm and not h_mm and not orientation:
        return None, None
    
    try:
        import win32print
        import win32con
        import copy
        
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
            
            # Set paper size by ID if found (e.g., custom "kuitansi" form)
            if paper_id:
                # Set form ID + explicit dimensions matching the form.
                # We MUST set W/H because the default DevMode has A4 values (2100x2970)
                # and SumatraPDF uses those DC dimensions for auto-rotation decisions.
                # By setting W=2413, H=1397 the DC is landscape-shaped (wider than tall)
                # which matches the landscape PDF → SumatraPDF won't auto-rotate.
                devmode.PaperSize = paper_id
                devmode.Fields |= win32con.DM_PAPERSIZE
                
                if w_mm and h_mm:
                    devmode.PaperWidth = int(float(w_mm) * 10)
                    devmode.PaperLength = int(float(h_mm) * 10)
                    devmode.Fields |= (win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH)
                
                # CLEAR orientation flag — let the form shape control layout,  
                # not an explicit portrait/landscape flag that SumatraPDF uses to decide rotation
                devmode.Fields &= ~win32con.DM_ORIENTATION
                modified = True
                log.info("DevMode: PaperSize=%d (%s), W=%d, H=%d — orientation flag cleared",
                         paper_id, paper_name, devmode.PaperWidth, devmode.PaperLength)
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
            
            if modified:
                if paper_id:
                    log.info("DevMode ready (no validation). PaperSize=%s, W=%s, H=%s, Orient=%s, Fields=%s",
                             devmode.PaperSize, devmode.PaperWidth, devmode.PaperLength, 
                             devmode.Orientation, devmode.Fields)
                else:
                    # For generic custom sizes, validate through DocumentProperties
                    result = win32print.DocumentProperties(
                        0, hprinter, printer_name, devmode, devmode,
                        win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER
                    )
                    log.info("DevMode validated (result=%s). Final: PaperSize=%s, W=%s, H=%s, Orient=%s",
                             result, devmode.PaperSize, devmode.PaperWidth, 
                             devmode.PaperLength, devmode.Orientation)
                return devmode, paper_name
        finally:
            win32print.ClosePrinter(hprinter)
    except Exception as e:
        log.warning("Failed to create DevMode: %s", e, exc_info=True)
    return None, None


def _print_pdf_windows(printer_name, pdf_path, options):
    """
    Print a PDF file directly via Windows GDI using PyMuPDF for rendering.
    
    This bypasses SumatraPDF entirely, giving us full control over the DevMode
    (paper size, orientation) without any auto-rotation interference.
    
    Flow:
      1. Build DevMode from options (paper ID, dimensions)
      2. Open a printer DC with that DevMode using win32print + win32ui
      3. Use PyMuPDF (fitz) to render each page at the printer's DPI
      4. BitBlt the rendered bitmap onto the printer DC
      5. Eject the page
    """
    try:
        import fitz  # PyMuPDF
        import win32print
        import win32ui
        import win32con
        from PIL import Image
        import io
        import struct

        devmode, paper_name = _create_devmode_for_options(printer_name, options)
        log.info("GDI print: printer=%s paper=%s devmode=%s", printer_name, paper_name, devmode is not None)

        # Open a printer DC with our exact DevMode.
        # win32ui.CreateDC() wraps the Windows CreateDC() GDI call.
        # "WINSPOOL" is the print driver type; device=printer_name; devmode from our options.
        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(printer_name)
        # Apply our DevMode if available (ResetDC rebuilds the DC with a new DevMode)
        if devmode:
            dc.ResetDC(devmode)

        try:
            # Get printer physical dimensions in pixels
            pwidth_px  = dc.GetDeviceCaps(win32con.PHYSICALWIDTH)
            pheight_px = dc.GetDeviceCaps(win32con.PHYSICALHEIGHT)
            dpi_x      = dc.GetDeviceCaps(win32con.LOGPIXELSX)
            dpi_y      = dc.GetDeviceCaps(win32con.LOGPIXELSY)
            offset_x   = dc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
            offset_y   = dc.GetDeviceCaps(win32con.PHYSICALOFFSETY)

            log.info("Printer DC: %dx%d px @ %d×%d dpi, offset=%d,%d",
                     pwidth_px, pheight_px, dpi_x, dpi_y, offset_x, offset_y)

            # Open PDF
            doc = fitz.open(pdf_path)
            copies = int(options.get('copies', 1)) if options else 1

            doc_info = win32ui.DOCINFO()
            doc_info['szDocName'] = 'PrintHub Job'

            dc.StartDoc(doc_info)

            for copy in range(copies):
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    
                    # Render page at printer DPI — use the page's natural size
                    # fitz matrix: 1 unit = 1 point (1/72 inch)
                    # scale to printer DPI
                    mat = fitz.Matrix(dpi_x / 72.0, dpi_y / 72.0)
                    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
                    
                    img_data = pix.tobytes("ppm")
                    img = Image.open(io.BytesIO(img_data))
                    img = img.convert("RGB")

                    img_w, img_h = img.size
                    log.info("Page %d: rendered %dx%d px for printer %dx%d px",
                             page_num + 1, img_w, img_h, pwidth_px, pheight_px)

                    # Scale image to fill the printable area if needed
                    printable_w = pwidth_px - 2 * offset_x
                    printable_h = pheight_px - 2 * offset_y
                    if printable_w > 0 and printable_h > 0:
                        scale = min(printable_w / img_w, printable_h / img_h)
                        new_w = int(img_w * scale)
                        new_h = int(img_h * scale)
                        if scale != 1.0:
                            img = img.resize((new_w, new_h), Image.LANCZOS)
                            img_w, img_h = new_w, new_h

                    # Convert PIL image to a Windows HBITMAP
                    # Write BMP in memory for win32ui
                    bmp_buf = io.BytesIO()
                    img.save(bmp_buf, format="BMP")
                    bmp_buf.seek(0)
                    bmp_bytes = bmp_buf.read()

                    # Create a memory DC and bitmap
                    mem_dc = dc.CreateCompatibleDC()
                    bmp = win32ui.CreateBitmap()
                    bmp.CreateBitmapFromBmpData(bmp_bytes)
                    old_bmp = mem_dc.SelectObject(bmp)

                    dc.StartPage()
                    # BitBlt: destination (x, y, w, h) from source (0, 0)
                    dc.BitBlt((0, 0), (img_w, img_h), mem_dc, (0, 0), win32con.SRCCOPY)
                    dc.EndPage()

                    mem_dc.SelectObject(old_bmp)
                    mem_dc.DeleteDC()
                    bmp.DeleteObject()

            dc.EndDoc()
            total_pages = len(doc)
            doc.close()
            log.info("GDI print complete: %d page(s), %d copy(ies)", total_pages, copies)
            return True, ""

        finally:
            dc.DeleteDC()

    except ImportError as e:
        log.warning("PyMuPDF not available, falling back to SumatraPDF: %s", e)
        return _print_pdf_sumatra(printer_name, pdf_path, options)
    except Exception as e:
        log.error("GDI print failed: %s", e, exc_info=True)
        log.info("Falling back to SumatraPDF...")
        return _print_pdf_sumatra(printer_name, pdf_path, options)


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
            success, error_msg = _print_pdf_windows(printer_name, temp_path, options)
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
            except:
                pass
        threading.Thread(target=cleanup, daemon=True).start()

    log.info("PDF print result: success=%s", success)
    return success, error_msg


if __name__ == '__main__':
    printers = get_printers()
    print("Available Printers:", printers)
