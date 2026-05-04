"""
macOS (Darwin) platform-specific helpers for TrayPrint.
"""
import os
import sys
import platform
import subprocess
import logging

logger = logging.getLogger(__name__)


def is_macos():
    """Return True if running on macOS."""
    return sys.platform == 'darwin'


def get_macos_version():
    """Return the macOS version string (e.g. '10.15.7', '11.6', '12.3')."""
    try:
        ver = platform.mac_ver()[0]
        return ver if ver else 'Unknown'
    except Exception:
        return 'Unknown'


def request_notification_permission():
    """
    macOS 10.14+ requires user notification permission for UNUserNotificationCenter.
    This is a placeholder — on modern macOS, the system handles the permission
    prompt automatically when the app first tries to show a notification.
    
    For PySide6/Qt-based notifications, this is typically handled by the
    QSystemTrayIcon.showMessage() call automatically.
    """
    if not is_macos():
        return False
    
    version = get_macos_version()
    logger.debug("macOS version: %s — notification permission assumed granted via Qt", version)
    return True


def get_app_support_dir():
    """
    Return the standard application support directory for TrayPrint.
    ~/Library/Application Support/TrayPrint/
    """
    home = os.path.expanduser('~')
    support_dir = os.path.join(home, 'Library', 'Application Support', 'TrayPrint')
    try:
        os.makedirs(support_dir, exist_ok=True)
    except OSError as e:
        logger.warning("Could not create app support dir %s: %s", support_dir, e)
    return support_dir


def get_app_cache_dir():
    """
    Return the standard application cache directory for TrayPrint.
    ~/Library/Caches/TrayPrint/
    """
    home = os.path.expanduser('~')
    cache_dir = os.path.join(home, 'Library', 'Caches', 'TrayPrint')
    try:
        os.makedirs(cache_dir, exist_ok=True)
    except OSError as e:
        logger.warning("Could not create cache dir %s: %s", cache_dir, e)
    return cache_dir


def get_default_printer_macos():
    """
    Get the default printer name on macOS using lpstat.
    """
    try:
        result = subprocess.run(
            ['lpstat', '-d'],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0 and ':' in result.stdout:
            return result.stdout.split(':')[-1].strip()
    except Exception as e:
        logger.error("Failed to get macOS default printer: %s", e)
    return ''


def get_printers_macos():
    """
    Enumerate printers on macOS using lpstat.
    macOS uses CUPS under the hood, so this is similar to Linux.
    """
    printers = []
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
                            'is_default': (name == get_default_printer_macos()),
                            'status': status,
                            'location': '',
                        })
    except Exception as e:
        logger.error("Failed to enumerate macOS printers: %s", e)
    return printers


def print_file_macos(printer_name, file_path, options=None):
    """
    Print a file on macOS using the lp command.
    macOS uses CUPS, so this is similar to Linux but may need
    additional options for driverless printing.
    """
    try:
        cmd = ['lp', '-d', printer_name]
        
        # Build CUPS options
        if options:
            copies = options.get('copies')
            if copies and int(copies) > 1:
                cmd += ['-n', str(int(copies))]
            
            paper = options.get('paper_size')
            if paper:
                cmd += ['-o', f'media={paper}']
            
            orientation = options.get('orientation')
            if orientation == 'landscape':
                cmd += ['-o', 'orientation-requested=4']
            elif orientation == 'portrait':
                cmd += ['-o', 'orientation-requested=3']
            
            duplex = options.get('duplex')
            if duplex == 'two-sided-long':
                cmd += ['-o', 'sides=two-sided-long-edge']
            elif duplex == 'two-sided-short':
                cmd += ['-o', 'sides=two-sided-short-edge']
            
            tray_source = options.get('tray_source')
            if tray_source:
                cmd += ['-o', f'InputSlot={tray_source}']
            
            color_mode = options.get('color_mode')
            if color_mode == 'monochrome':
                cmd += ['-o', 'ColorModel=Gray']
            elif color_mode == 'color':
                cmd += ['-o', 'ColorModel=RGB']
            
            print_quality = options.get('print_quality')
            quality_map = {'draft': '3', 'normal': '4', 'high': '5'}
            if print_quality and print_quality in quality_map:
                cmd += ['-o', f'print-quality={quality_map[print_quality]}']
        
        cmd.append(file_path)
        
        logger.info("macOS print cmd: %s", ' '.join(cmd))
        subprocess.run(cmd, capture_output=True, check=True, timeout=60)
        return True, ""
    except subprocess.CalledProcessError as e:
        error_msg = e.stderr.decode('utf-8') if e.stderr else str(e)
        logger.error("macOS print failed: %s", error_msg)
        return False, error_msg
    except Exception as e:
        logger.error("macOS print exception: %s", e)
        return False, str(e)
