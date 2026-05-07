"""
TrayPrint — Logging Utilities

Provides a centralized logging setup with rotating file handler support.
Configurable via config.json keys:
  - log_max_bytes:  Maximum log file size before rotation (default: 5MB)
  - log_backup_count: Number of backup files to keep (default: 3)

Logs are written to logs/app.log in the application root directory.
"""

import logging
import os
import sys
from logging.handlers import RotatingFileHandler

from path_utils import get_root_dir

# Module-level logger cache
_loggers = {}


def setup_logging(
    name="trayprint",
    log_dir=None,
    log_file="app.log",
    max_bytes=5 * 1024 * 1024,
    backup_count=3,
    console_level=logging.INFO,
    file_level=logging.DEBUG,
    config=None,
):
    """
    Configure and return a logger with rotating file handler.

    Parameters
    ----------
    name : str
        Logger name (default: "trayprint").
    log_dir : str or None
        Directory for log files. If None, uses get_root_dir() / "logs".
    log_file : str
        Log filename (default: "app.log").
    max_bytes : int
        Maximum log file size before rotation (default: 5MB).
    backup_count : int
        Number of backup files to retain (default: 3).
    console_level : int
        Logging level for console handler (default: logging.INFO).
    file_level : int
        Logging level for file handler (default: logging.DEBUG).
    config : dict or None
        Optional config dict to override max_bytes/backup_count from config.json.

    Returns
    -------
    logging.Logger
        Configured logger instance.
    """
    # Check if already configured
    if name in _loggers:
        return _loggers[name]

    # Apply config overrides if provided
    if config:
        max_bytes = config.get("log_max_bytes", max_bytes)
        backup_count = config.get("log_backup_count", backup_count)

    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # Avoid duplicate handlers if setup_logging is called multiple times
    if logger.handlers:
        _loggers[name] = logger
        return logger

    formatter = logging.Formatter(
        "[%(asctime)s] %(levelname)-7s %(name)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # File handler with rotation
    if log_dir is None:
        log_dir = os.path.join(get_root_dir(), "logs")
    os.makedirs(log_dir, exist_ok=True)

    log_path = os.path.join(log_dir, log_file)
    file_handler = RotatingFileHandler(
        log_path,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8",
    )
    file_handler.setLevel(file_level)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    logger.info(
        "Logging initialized — file=%s, maxBytes=%d, backupCount=%d",
        log_path,
        max_bytes,
        backup_count,
    )

    _loggers[name] = logger
    return logger


def get_log_path(name="trayprint", log_dir=None, log_file="app.log"):
    """
    Return the path to the log file for the given logger.

    Parameters
    ----------
    name : str
        Logger name (unused, for API compatibility).
    log_dir : str or None
        Log directory. If None, uses get_root_dir() / "logs".
    log_file : str
        Log filename (default: "app.log").

    Returns
    -------
    str
        Full path to the log file.
    """
    if log_dir is None:
        log_dir = os.path.join(get_root_dir(), "logs")
    return os.path.join(log_dir, log_file)


def shutdown_logging(name="trayprint"):
    """
    Flush and remove all handlers for the given logger.

    Parameters
    ----------
    name : str
        Logger name to shut down.
    """
    logger = _loggers.pop(name, None)
    if logger:
        for handler in logger.handlers[:]:
            handler.flush()
            handler.close()
            logger.removeHandler(handler)
        logger.handlers.clear()
