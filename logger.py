import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from path_utils import get_root_dir

_logger = None

def get_logger():
    """Returns the singleton application logger."""
    global _logger
    if _logger is not None:
        return _logger

    _logger = logging.getLogger('trayprint')
    _logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        '[%(asctime)s] %(levelname)-7s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Console handler
    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.INFO)
    console.setFormatter(formatter)
    _logger.addHandler(console)

    # File handler with rotation (5MB max, keep 3 backups)
    log_dir = get_root_dir()
    log_path = os.path.join(log_dir, 'trayprint.log')
    file_handler = RotatingFileHandler(
        log_path, maxBytes=5 * 1024 * 1024, backupCount=3, encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    _logger.addHandler(file_handler)

    _logger.info("Logger initialized — log file: %s", log_path)
    return _logger

def get_log_path():
    """Returns the path to the log file."""
    return os.path.join(get_root_dir(), 'trayprint.log')
