"""
Tests for logger.py — singleton logger factory and log path helper.

All tests use the temp_dir fixture to avoid polluting real log directories.
"""

import logging
import os

import pytest


# ===================================================================
# get_logger
# ===================================================================

class TestGetLogger:
    """Test the singleton get_logger factory."""

    def test_returns_logger_instance(self, temp_dir):
        """get_logger returns a logging.Logger instance."""
        import logger as logger_mod
        logger_mod._logger = None

        # Override data dir to our temp dir
        logger_mod.get_data_dir = lambda: temp_dir

        log = logger_mod.get_logger()
        assert isinstance(log, logging.Logger)
        assert log.name == "trayprint"

        # Clean up handlers to avoid side effects
        for h in log.handlers[:]:
            h.close()
            log.removeHandler(h)
        logger_mod._logger = None

    def test_singleton_behavior(self, temp_dir):
        """get_logger returns the same instance on repeated calls."""
        import logger as logger_mod
        logger_mod._logger = None
        logger_mod.get_data_dir = lambda: temp_dir

        log1 = logger_mod.get_logger()
        log2 = logger_mod.get_logger()
        assert log1 is log2

        # Clean up
        for h in log1.handlers[:]:
            h.close()
            log1.removeHandler(h)
        logger_mod._logger = None

    def test_sets_debug_level(self, temp_dir):
        """The logger is set to DEBUG level."""
        import logger as logger_mod
        logger_mod._logger = None
        logger_mod.get_data_dir = lambda: temp_dir

        log = logger_mod.get_logger()
        assert log.level == logging.DEBUG

        # Clean up
        for h in log.handlers[:]:
            h.close()
            log.removeHandler(h)
        logger_mod._logger = None

    def test_adds_console_and_file_handlers(self, temp_dir):
        """get_logger adds both a StreamHandler and a RotatingFileHandler."""
        import logger as logger_mod
        logger_mod._logger = None
        logger_mod.get_data_dir = lambda: temp_dir

        log = logger_mod.get_logger()

        handler_types = [type(h).__name__ for h in log.handlers]
        assert "StreamHandler" in handler_types
        assert "RotatingFileHandler" in handler_types

        # Clean up
        for h in log.handlers[:]:
            h.close()
            log.removeHandler(h)
        logger_mod._logger = None

    def test_console_handler_info_level(self, temp_dir):
        """The console handler is set to INFO level."""
        import logger as logger_mod
        logger_mod._logger = None
        logger_mod.get_data_dir = lambda: temp_dir

        log = logger_mod.get_logger()
        console = [h for h in log.handlers
                   if isinstance(h, logging.StreamHandler)]
        assert console[0].level == logging.INFO

        # Clean up
        for h in log.handlers[:]:
            h.close()
            log.removeHandler(h)
        logger_mod._logger = None

    def test_file_handler_debug_level(self, temp_dir):
        """The rotating file handler is set to DEBUG level."""
        import logger as logger_mod
        logger_mod._logger = None
        logger_mod.get_data_dir = lambda: temp_dir

        log = logger_mod.get_logger()
        file_handler = [h for h in log.handlers
                        if hasattr(h, 'baseFilename')]
        assert file_handler[0].level == logging.DEBUG

        # Clean up
        for h in log.handlers[:]:
            h.close()
            log.removeHandler(h)
        logger_mod._logger = None

    def test_file_handler_uses_get_data_dir(self, temp_dir):
        """The file handler's base path comes from get_data_dir."""
        import logger as logger_mod
        logger_mod._logger = None
        logger_mod.get_data_dir = lambda: temp_dir

        log = logger_mod.get_logger()
        file_handler = [h for h in log.handlers
                        if hasattr(h, 'baseFilename')]
        assert file_handler[0].baseFilename == os.path.join(temp_dir, "trayprint.log")

        # Clean up
        for h in log.handlers[:]:
            h.close()
            log.removeHandler(h)
        logger_mod._logger = None

    def test_rotation_config(self, temp_dir):
        """File handler has 5MB maxBytes and 3 backup files."""
        import logger as logger_mod
        logger_mod._logger = None
        logger_mod.get_data_dir = lambda: temp_dir

        log = logger_mod.get_logger()
        file_handler = [h for h in log.handlers
                        if hasattr(h, 'baseFilename')][0]
        assert file_handler.maxBytes == 5 * 1024 * 1024
        assert file_handler.backupCount == 3

        # Clean up
        for h in log.handlers[:]:
            h.close()
            log.removeHandler(h)
        logger_mod._logger = None


# ===================================================================
# get_log_path
# ===================================================================

class TestGetLogPath:
    """Test the get_log_path helper."""

    def test_returns_path_using_get_data_dir(self, mocker):
        """get_log_path joins get_data_dir() with 'trayprint.log'."""
        import os
        mocker.patch("logger.get_data_dir", return_value="/data/path")

        from logger import get_log_path
        result = get_log_path()
        expected = os.path.join("/data/path", "trayprint.log")
        assert result == expected
