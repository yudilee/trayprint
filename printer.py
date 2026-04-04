import sys
import subprocess
from logger import get_logger

log = get_logger()

def is_windows():
    return sys.platform == 'win32'

def is_macos():
    return sys.platform == 'darwin'


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

    copies = options.get('copies')
    if copies and int(copies) > 1:
        args += ['-n', str(int(copies))]

    paper = options.get('paper_size')
    if paper:
        args += ['-o', f'media={paper}']

    orientation = options.get('orientation')
    if orientation == 'landscape':
        args += ['-o', 'orientation-requested=4']

    duplex = options.get('duplex')
    if duplex == 'two-sided-long':
        args += ['-o', 'sides=two-sided-long-edge']
    elif duplex == 'two-sided-short':
        args += ['-o', 'sides=two-sided-short-edge']

    page_range = options.get('page_range')
    if page_range:
        args += ['-o', f'page-ranges={page_range}']

    return args


def _build_sumatra_options(options):
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

    duplex = options.get('duplex')
    if duplex and duplex.startswith('two-sided'):
        parts.append('duplex')

    page_range = options.get('page_range')
    if page_range:
        parts.append(page_range)

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
            hprinter = win32print.OpenPrinter(printer_name)
            try:
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
            sumatra_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "SumatraPDF.exe")
            if os.path.exists(sumatra_path):
                try:
                    cmd = [sumatra_path, "-print-to", printer_name, "-silent"]
                    cmd += _build_sumatra_options(options)
                    cmd.append(temp_path)
                    subprocess.run(cmd, check=True)
                    success = True
                except Exception as e:
                    error_msg = str(e)
                    log.error("SumatraPDF Print Error: %s", e)
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

    # Cleanup temp file after 15s
    if temp_path:
        def cleanup():
            time.sleep(15)
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
