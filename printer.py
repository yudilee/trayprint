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
    and returns the name of the one matching the given width/height.
    """
    if not is_windows() or not printer_name:
        return None
    try:
        import win32print
        import win32con
        # DeviceCapabilities returns paper names and sizes
        names = win32print.DeviceCapabilities(printer_name, "", win32con.DC_PAPERNAMES)
        sizes = win32print.DeviceCapabilities(printer_name, "", win32con.DC_PAPERSIZE)
        ids = win32print.DeviceCapabilities(printer_name, "", win32con.DC_PAPERS)
        
        # sizes are in 0.1mm units. 
        # Example: 210mm = 2100 units
        target_w = int(float(w_mm) * 10)
        target_h = int(float(h_mm) * 10)
        
        for i, (w, h) in enumerate(sizes):
            # Use a tolerance of 2 units (0.2mm)
            if abs(w - target_w) <= 2 and abs(h - target_h) <= 2:
                log.info("Matched Windows paper size by dimensions: %s (ID:%d) (%dx%d)", names[i], ids[i], w, h)
                return names[i], ids[i]
        
        # Try swapped dimensions for orientation-agnostic drivers
        for i, (w, h) in enumerate(sizes):
            if abs(h - target_w) <= 2 and abs(w - target_h) <= 2:
                log.info("Matched Windows paper size by swapped dimensions: %s (ID:%d) (%dx%d)", names[i], ids[i], w, h)
                return names[i], ids[i]
                
    except Exception as e:
        log.debug("Error finding Windows paper name: %s", e)
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

    if paper:
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
                devmode.PaperSize = paper_id
                devmode.Fields |= win32con.DM_PAPERSIZE
                modified = True
                log.info("DevMode: Set PaperSize ID=%d (%s)", paper_id, paper_name)
            
            # Always set explicit dimensions for custom/dot-matrix
            if w_mm and h_mm:
                devmode.PaperWidth = int(float(w_mm) * 10)
                devmode.PaperLength = int(float(h_mm) * 10)
                devmode.Fields |= (win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH)
                modified = True
                log.info("DevMode: Set PaperWidth=%d, PaperLength=%d (0.1mm units)", 
                         devmode.PaperWidth, devmode.PaperLength)
            
            # Orientation
            if orientation == 'landscape':
                devmode.Orientation = win32con.DMORIENT_LANDSCAPE
                devmode.Fields |= win32con.DM_ORIENTATION
                modified = True
            elif orientation == 'portrait':
                devmode.Orientation = win32con.DMORIENT_PORTRAIT
                devmode.Fields |= win32con.DM_ORIENTATION
                modified = True
            
            if modified:
                # Validate the DevMode through DocumentProperties
                # DM_IN_BUFFER: read from devmode input
                # DM_OUT_BUFFER: write validated result to devmode output
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
            sumatra_path = _get_sumatra_path()
            if sumatra_path:
                try:
                    # Build the DevMode with correct paper settings
                    devmode, paper_name = _create_devmode_for_options(printer_name, options)
                    
                    if devmode:
                        # Strategy: Temporarily set the printer default, then print
                        log.info("Using DevMode override strategy for printer '%s' (paper=%s)", 
                                 printer_name, paper_name)
                        
                        import win32print
                        hprinter = win32print.OpenPrinter(printer_name, 
                            {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
                        try:
                            # Save original
                            pinfo = win32print.GetPrinter(hprinter, 2)
                            original_devmode = pinfo['pDevMode']
                            
                            # Apply our DevMode as the printer default
                            pinfo['pDevMode'] = devmode
                            win32print.SetPrinter(hprinter, 2, pinfo, 0)
                            log.info("Printer default temporarily changed to custom paper.")
                            
                            try:
                                # SumatraPDF will now use the printer's default (which we just set)
                                cmd = [sumatra_path, "-print-to", printer_name, "-silent"]
                                cmd += _build_sumatra_options(options, printer_name)
                                cmd.append(temp_path)
                                log.info("SumatraPDF cmd: %s", ' '.join(cmd))
                                subprocess.run(cmd, check=True, timeout=60)
                            finally:
                                # Restore original printer settings
                                import time as _time
                                _time.sleep(2)  # Give spooler time to pick up the job
                                pinfo['pDevMode'] = original_devmode
                                win32print.SetPrinter(hprinter, 2, pinfo, 0)
                                log.info("Printer default restored.")
                        finally:
                            win32print.ClosePrinter(hprinter)
                    else:
                        # No custom DevMode needed, just use SumatraPDF normally
                        cmd = [sumatra_path, "-print-to", printer_name, "-silent"]
                        cmd += _build_sumatra_options(options, printer_name)
                        cmd.append(temp_path)
                        log.info("SumatraPDF cmd (no DevMode): %s", ' '.join(cmd))
                        subprocess.run(cmd, check=True, timeout=60)
                    
                    success = True
                except Exception as e:
                    error_msg = str(e)
                    log.error("SumatraPDF Print Error: %s", e, exc_info=True)
            else:
                import win32api
                try:
                    win32api.ShellExecute(0, "printto", temp_path, f'"{printer_name}"', ".", 0)
                    success = True
                except Exception as e:
                    error_msg = str(e)
                    log.error("Windows PDF Print Error: %s", e)
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
