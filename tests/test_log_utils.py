"""
Tests for log_utils.py — logging configuration, rotation, and lifecycle.

Covers setup_logging, get_log_path, and shutdown_logging with mocked
filesystem operations to avoid polluting real log directories.
"""

import logging
import os

import pytest


# ===================================================================
# setup_logging
# ===================================================================

class TestSetupLogging:
    """Test the setup_logging factory function."""

    def test_basic_initialization(self, temp_dir):
        """setup_logging creates a logger with expected defaults."""
        from log_utils import setup_logging, shutdown_logging

        logger = setup_logging(name="test-basic", log_dir=temp_dir)
        assert logger.name == "test-basic"
        assert logger.level == logging.DEBUG
        assert len(logger.handlers) == 2  # console + file

        # Verify file handler exists
        file_handlers = [h for h in logger.handlers
                         if isinstance(h, logging.handlers.RotatingFileHandler)]
        assert len(file_handlers) == 1

        # Verify console handler exists
        console_handlers = [h for h in logger.handlers
                            if isinstance(h, logging.StreamHandler)
                            and not isinstance(h, logging.handlers.RotatingFileHandler)]
        assert len(console_handlers) == 1

        shutdown_logging("test-basic")

    def test_cached_logger_returned(self, temp_dir):
        """Calling setup_logging twice with the same name returns cached logger."""
        from log_utils import setup_logging, shutdown_logging

        logger1 = setup_logging(name="test-cached", log_dir=temp_dir)
        logger2 = setup_logging(name="test-cached", log_dir=temp_dir)
        assert logger1 is logger2

        shutdown_logging("test-cached")

    def test_config_overrides_max_bytes(self, temp_dir):
        """Config dict overrides max_bytes from default."""
        from log_utils import setup_logging, shutdown_logging

        config = {"log_max_bytes": 1024, "log_backup_count": 1}
        logger = setup_logging(
            name="test-override", log_dir=temp_dir, config=config
        )
        file_handlers = [h for h in logger.handlers
                         if isinstance(h, logging.handlers.RotatingFileHandler)]
        assert len(file_handlers) == 1
        assert file_handlers[0].maxBytes == 1024
        assert file_handlers[0].backupCount == 1

        shutdown_logging("test-override")

    def test_config_overrides_backup_count(self, temp_dir):
        """Config dict overrides backup_count from default."""
        from log_utils import setup_logging, shutdown_logging

        config = {"log_backup_count": 5}
        logger = setup_logging(
            name="test-backup", log_dir=temp_dir, config=config
        )
        file_handlers = [h for h in logger.handlers
                         if isinstance(h, logging.handlers.RotatingFileHandler)]
        assert file_handlers[0].backupCount == 5

        shutdown_logging("test-backup")

    def test_custom_log_file(self, temp_dir):
        """A custom log_file name is used instead of the default."""
        from log_utils import setup_logging, shutdown_logging, get_log_path

        logger = setup_logging(
            name="test-custom-file", log_dir=temp_dir, log_file="custom.log"
        )
        file_handlers = [h for h in logger.handlers
                         if isinstance(h, logging.handlers.RotatingFileHandler)]
        log_path = file_handlers[0].baseFilename
        assert log_path.endswith("custom.log")
        assert os.path.exists(log_path)

        shutdown_logging("test-custom-file")

    def test_custom_console_level(self, temp_dir):
        """Console handler respects custom console_level."""
        from log_utils import setup_logging, shutdown_logging

        logger = setup_logging(
            name="test-console-level", log_dir=temp_dir,
            console_level=logging.WARNING,
        )
        console_handlers = [h for h in logger.handlers
                            if isinstance(h, logging.StreamHandler)
                            and not isinstance(h, logging.handlers.RotatingFileHandler)]
        assert console_handlers[0].level == logging.WARNING

        shutdown_logging("test-console-level")

    def test_avoids_duplicate_handlers(self, temp_dir):
        """If logger already has handlers, no new ones are added."""
        from log_utils import setup_logging, shutdown_logging

        # First call — adds 2 handlers
        logger = setup_logging(name="test-dup-check", log_dir=temp_dir)
        count_after_first = len(logger.handlers)

        # Simulate someone calling setup_logging again after a cache miss
        # by clearing the cache
        import log_utils as lu
        lu._loggers.pop("test-dup-check", None)

        # Second call — should detect existing handlers and not add more
        logger2 = setup_logging(name="test-dup-check", log_dir=temp_dir)
        # Should be the same logger object with same handlers
        assert len(logger2.handlers) == count_after_first

        shutdown_logging("test-dup-check")

    def test_default_log_dir(self, mocker, temp_dir):
        """When log_dir is None, it falls back to get_data_dir()/logs."""
        from log_utils import setup_logging, shutdown_logging

        # Mock get_data_dir to return temp_dir
        mocker.patch("log_utils.get_data_dir", return_value=temp_dir)

        logger = setup_logging(name="test-default-dir")
        file_handlers = [h for h in logger.handlers
                         if isinstance(h, logging.handlers.RotatingFileHandler)]
        log_dir = os.path.dirname(file_handlers[0].baseFilename)
        assert log_dir == os.path.join(temp_dir, "logs")
        assert os.path.exists(log_dir)

        shutdown_logging("test-default-dir")


# ===================================================================
# get_log_path
# ===================================================================

class TestGetLogPath:
    """Test get_log_path helper."""

    def test_default_path(self, mocker):
        """Default path uses get_data_dir()/logs/app.log."""
        import os
        from log_utils import get_log_path

        mocker.patch("log_utils.get_data_dir", return_value="/fake/data")
        result = get_log_path()
        # Use os.path.join to be platform-agnostic (Windows uses backslashes)
        expected = os.path.join("/fake/data", "logs", "app.log")
        assert result == expected

    def test_custom_log_dir(self):
        """A custom log_dir is reflected in the result."""
        import os
        from log_utils import get_log_path

        result = get_log_path(log_dir="/custom/path")
        expected = os.path.join("/custom/path", "app.log")
        assert result == expected

    def test_custom_log_file(self):
        """A custom log_file is reflected in the result."""
        from log_utils import get_log_path

        result = get_log_path(log_file="debug.log")
        assert "debug.log" in result


# ===================================================================
# shutdown_logging
# ===================================================================

class TestShutdownLogging:
    """Test shutdown_logging cleanup."""

    def test_removes_all_handlers(self, temp_dir):
        """shutdown_logging flushes, closes, and removes all handlers."""
        from log_utils import setup_logging, shutdown_logging

        logger = setup_logging(name="test-shutdown", log_dir=temp_dir)
        assert len(logger.handlers) > 0

        shutdown_logging("test-shutdown")
        assert len(logger.handlers) == 0

    def test_unknown_logger_is_noop(self):
        """shutdown_logging for a non-existent logger does not raise."""
        from log_utils import shutdown_logging

        shutdown_logging("this-logger-does-not-exist")  # should not raise

    def test_logger_removed_from_cache(self, temp_dir):
        """After shutdown, the logger is removed from the cache."""
        from log_utils import setup_logging, shutdown_logging, _loggers

        setup_logging(name="test-removed", log_dir=temp_dir)
        assert "test-removed" in _loggers

        shutdown_logging("test-removed")
        assert "test-removed" not in _loggers
