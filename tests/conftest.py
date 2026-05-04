"""
Shared fixtures for the TrayPrint test suite.

All tests run against a temporary directory to avoid polluting the real
config.json, jobs.db, or autostart registry.
"""

import os
import sys
import json
import tempfile
import uuid

import pytest

# Ensure the project root is on sys.path so `import server` etc. work
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def temp_dir():
    """Provide a temporary directory and chdir into it for the test scope."""
    with tempfile.TemporaryDirectory() as tmp:
        old_cwd = os.getcwd()
        os.chdir(tmp)
        yield tmp
        os.chdir(old_cwd)


@pytest.fixture
def sample_config(temp_dir):
    """Write a minimal config.json into temp_dir and return its parsed content."""
    cfg = {
        "hub_url": "https://hub.example.com",
        "agent_key": "test-agent-key-1234",
        "port": 9120,
        "interval": 10,
        "theme": "light",
        "windows_startup": False,
        "profiles": [
            {
                "printer": "TestPrinter",
                "name": "Default",
                "options": {"copies": 1, "duplex": "none"},
            }
        ],
    }
    path = os.path.join(temp_dir, "config.json")
    with open(path, "w") as f:
        json.dump(cfg, f)
    return cfg


@pytest.fixture
def job_queue(temp_dir):
    """Return a JobQueue that uses an in-memory or isolated SQLite db."""
    # Use a unique filename so parallel tests do not collide
    db_name = f"test_jobs_{uuid.uuid4().hex[:8]}.db"
    from server import JobQueue
    q = JobQueue(db_name=db_name)
    yield q
    # Cleanup: remove the db file
    db_path = os.path.join(temp_dir, db_name)
    if os.path.exists(db_path):
        os.remove(db_path)


@pytest.fixture
def flask_app():
    """Return the Flask test client for the TrayPrint REST API."""
    from server import create_app
    app = create_app()
    app.config["TESTING"] = True
    with app.test_client() as client:
        yield client


@pytest.fixture
def mock_win32print(mocker):
    """Mock win32print functions to simulate Windows printer environment."""
    mock = mocker.patch("printer.win32print", autospec=True)
    return mock


@pytest.fixture
def mock_subprocess_run(mocker):
    """Mock subprocess.run for CUPS/lp command tests."""
    mock = mocker.patch("subprocess.run", autospec=True)
    return mock
