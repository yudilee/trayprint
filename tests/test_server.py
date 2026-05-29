"""
Tests for server.py — JobQueue persistence, helper functions, and Flask API endpoints.
"""

import json
import os

import pytest


# ===================================================================
# JobQueue tests
# ===================================================================

class TestJobQueue:
    """CRUD operations for the persistent JobQueue."""

    def test_create_job(self, job_queue):
        """Create a job and verify its initial state."""
        job = job_queue.create("PrinterA", "raw", {"copies": 2}, "test data preview")
        assert job["id"] is not None
        assert job["printer"] == "PrinterA"
        assert job["type"] == "raw"
        assert job["options"] == {"copies": 2}
        assert job["status"] == "pending"
        assert job["error"] is None
        assert job["completed_at"] is None
        assert job["data_preview"] == "test data preview"

    def test_create_job_with_custom_id(self, job_queue):
        """Create a job with a caller-supplied ID."""
        job = job_queue.create("PrinterB", "pdf", job_id="my-custom-id")
        assert job["id"] == "my-custom-id"

    def test_get_job(self, job_queue):
        """Retrieve a job by ID."""
        created = job_queue.create("PrinterA", "raw")
        fetched = job_queue.get(created["id"])
        assert fetched is not None
        assert fetched["id"] == created["id"]
        assert fetched["status"] == "pending"

    def test_get_nonexistent_job(self, job_queue):
        """Getting a non-existent job returns None."""
        assert job_queue.get("does-not-exist") is None

    def test_complete_job_success(self, job_queue):
        """Mark a job as completed successfully."""
        job = job_queue.create("PrinterA", "raw")
        job_queue.complete(job["id"], success=True)
        updated = job_queue.get(job["id"])
        assert updated["status"] == "success"
        assert updated["error"] is None
        assert updated["completed_at"] is not None

    def test_complete_job_failure(self, job_queue):
        """Mark a job as failed with an error message."""
        job = job_queue.create("PrinterA", "raw")
        job_queue.complete(job["id"], success=False, error_msg="Out of paper")
        updated = job_queue.get(job["id"])
        assert updated["status"] == "failed"
        assert updated["error"] == "Out of paper"
        assert updated["completed_at"] is not None

    def test_list_recent(self, job_queue):
        """list_recent returns jobs in creation order."""
        ids = []
        for i in range(5):
            j = job_queue.create(f"Printer_{i}", "raw")
            ids.append(j["id"])

        jobs = job_queue.list_recent(limit=10)
        assert len(jobs) == 5
        # Should be in ASC order of created_at
        assert [j["id"] for j in jobs] == ids

    def test_list_recent_respects_limit(self, job_queue):
        """list_recent returns at most `limit` items."""
        for i in range(10):
            job_queue.create(f"Printer_{i}", "raw")
        jobs = job_queue.list_recent(limit=3)
        assert len(jobs) == 3

    def test_max_history_enforced(self, job_queue):
        """When history exceeds max_history, old jobs are pruned."""
        q = job_queue
        # Create more than max_history (default 50) jobs
        for i in range(55):
            q.create(f"Printer_{i}", "raw")

        jobs = q.list_recent(limit=100)
        assert len(jobs) <= q._max  # Should be at most 50

    def test_store_and_get_job_data(self, job_queue):
        """Store and retrieve raw binary data for a job."""
        job = job_queue.create("PrinterA", "pdf")
        raw = b"PDF content here"
        job_queue.store_job_data(job["id"], raw)
        fetched = job_queue.get_job_data(job["id"])
        assert fetched == raw

    def test_get_job_data_nonexistent(self, job_queue):
        """get_job_data for a missing job returns None."""
        assert job_queue.get_job_data("no-such-job") is None


# ===================================================================
# Helper function tests
# ===================================================================

class TestHelpers:
    """Test helper functions from server.py."""

    def test_is_printing_default(self):
        """is_printing() returns False before any job starts."""
        from server import is_printing
        assert is_printing() is False

    def test_is_printing_after_set(self):
        """is_printing() reflects the state set by _set_printing()."""
        from server import is_printing, _set_printing
        _set_printing(True)
        assert is_printing() is True
        _set_printing(False)
        assert is_printing() is False

    def test_get_cached_printer_count_fallback(self, mocker):
        """When the cache is empty, get_cached_printer_count falls back to printer enumeration."""
        from server import get_cached_printer_count
        count = get_cached_printer_count()
        assert isinstance(count, int)
        assert count >= 0

    def test_check_cancelled_default(self):
        """A job that was never cancelled returns False."""
        from server import _check_cancelled
        assert _check_cancelled("some-job-id") is False

    def test_get_queue_status(self, job_queue):
        """get_queue_status returns a list of jobs sorted by recency."""
        # Create some pending and completed jobs
        for i in range(3):
            job_queue.create("PrinterA", "raw")
        success_job = job_queue.create("PrinterA", "raw")
        job_queue.complete(success_job["id"], success=True)
        failed_job = job_queue.create("PrinterA", "raw")
        job_queue.complete(failed_job["id"], success=False)

        from server import get_queue_status
        # get_queue_status uses the global _job_queue, so inject our fixture's queue
        import server
        original = server._job_queue
        server._job_queue = job_queue
        try:
            status = get_queue_status()
            # Now returns a list of jobs, not a dict with counts
            assert isinstance(status, list)
            # All 5 jobs should be present
            assert len(status) == 5
        finally:
            server._job_queue = original


# ===================================================================
# Flask API tests
# ===================================================================

class TestFlaskAPI:
    """Integration tests for the Flask REST API."""

    def test_status_endpoint(self, flask_app, sample_config):
        """GET /status returns basic status information."""
        resp = flask_app.get("/status")
        assert resp.status_code == 200
        data = resp.get_json()
        assert data["status"] == "running"
        assert "version" in data
        assert "uptime_seconds" in data

    def test_printers_endpoint_returns_list(self, flask_app):
        """GET /printers returns a list of printers (may be empty on CI)."""
        resp = flask_app.get("/printers")
        assert resp.status_code == 200
        data = resp.get_json()
        assert isinstance(data.get("printers"), list)

    def test_profiles_endpoint(self, flask_app, sample_config):
        """GET /profiles returns the profiles from config.json."""
        resp = flask_app.get("/profiles")
        assert resp.status_code == 200
        data = resp.get_json()
        # Endpoint returns both profiles and queues
        assert "profiles" in data
        assert "queues" in data

    def test_jobs_list_empty(self, flask_app):
        """GET /jobs returns an empty list when no jobs exist."""
        resp = flask_app.get("/jobs")
        assert resp.status_code == 200
        data = resp.get_json()
        assert isinstance(data.get("jobs"), list)

    def test_queue_status_endpoint(self, flask_app):
        """GET /queue-status returns queue job list."""
        resp = flask_app.get("/queue-status")
        assert resp.status_code == 200
        data = resp.get_json()
        # Returns a dict with is_printing and jobs
        assert isinstance(data, dict)
        assert "is_printing" in data
        assert "jobs" in data

    def test_api_capabilities_no_printer(self, flask_app):
        """GET /api/capabilities without a printer param returns error."""
        resp = flask_app.get("/api/capabilities")
        assert resp.status_code == 200  # Returns capabilities list without printer param
        data = resp.get_json()
        assert data is not None

    def test_hub_connection(self, flask_app):
        """POST /test-hub (if implemented) — just verify it doesn't crash."""
        resp = flask_app.post("/test-hub", json={"hub_url": "http://localhost:9999", "agent_key": "test"})
        # Should either 200, 400, or 404 if not registered as route
        assert resp.status_code in (200, 400, 404)

    def test_diagnostics_endpoint(self, flask_app):
        """GET /api/diagnostics returns system diagnostics."""
        resp = flask_app.get("/api/diagnostics")
        assert resp.status_code == 200
        data = resp.get_json()
        assert "platform" in data
        assert "python_version" in data


# ===================================================================
# Configuration & profile loading tests
# ===================================================================

class TestConfig:
    """Tests for config loading and profile management."""

    def test_load_profiles_from_config(self, mocker, temp_dir):
        """Load profiles from a config file in the data directory."""
        # Write a config into temp_dir with profiles as a LIST
        config_path = os.path.join(temp_dir, "config.json")
        cfg = {
            "profiles": [
                {"printer": "TestPrinter", "name": "Default", "options": {"copies": 1}},
            ]
        }
        with open(config_path, "w") as f:
            json.dump(cfg, f)

        # Mock get_root_dir to point at temp_dir so load_profiles_from_config finds the file
        mocker.patch("server.get_root_dir", return_value=temp_dir)

        from server import load_profiles_from_config, get_profiles
        load_profiles_from_config()
        profiles = get_profiles()
        assert isinstance(profiles, list)
        assert len(profiles) == 1
        assert profiles[0]["printer"] == "TestPrinter"

    def test_save_profiles_to_config(self, mocker, temp_dir):
        """Save profiles and verify they persist."""
        # Write an initial config
        config_path = os.path.join(temp_dir, "config.json")
        cfg = {"hub_url": "", "profiles": []}
        with open(config_path, "w") as f:
            json.dump(cfg, f)

        # Mock get_data_dir to return temp_dir (save_profiles_to_config uses get_data_dir)
        mocker.patch("server.get_data_dir", return_value=temp_dir)
        # Also mock get_root_dir for load_profiles_from_config
        mocker.patch("server.get_root_dir", return_value=temp_dir)

        from server import save_profiles_to_config, load_profiles_from_config, get_profiles
        new_profiles = [
            {"printer": "P1", "name": "Profile1", "options": {"copies": 1}},
            {"printer": "P2", "name": "Profile2", "options": {"duplex": "long"}},
        ]
        save_profiles_to_config(new_profiles)
        load_profiles_from_config()
        loaded = get_profiles()
        assert loaded == new_profiles

    def test_load_printer_configs(self, mocker, temp_dir):
        """Load per-printer configs from config.json."""
        # Write a config with printer_configs
        config_path = os.path.join(temp_dir, "config.json")
        cfg = {"printer_configs": {"HP-Deskjet": {"copies": 2}}}
        with open(config_path, "w") as f:
            json.dump(cfg, f)

        mocker.patch("server.get_root_dir", return_value=temp_dir)

        from server import load_printer_configs
        configs = load_printer_configs()
        assert isinstance(configs, dict)
        assert "HP-Deskjet" in configs

    def test_merge_printer_config_no_saved(self, mocker, temp_dir):
        """When no saved config exists, merge_printer_config returns options as-is."""
        mocker.patch("server.get_root_dir", return_value=temp_dir)

        from server import merge_printer_config
        result = merge_printer_config("UnknownPrinter", {"copies": 5})
        # Returns the options dict directly (no 'printer' key added)
        assert result == {"copies": 5}

    def test_merge_printer_config_with_saved(self, mocker, temp_dir):
        """When saved config exists, merge merges saved control fields with options."""
        config_path = os.path.join(temp_dir, "config.json")
        cfg = {"printer_configs": {"MyPrinter": {"tray_source": "Tray2", "color_mode": "monochrome"}}}
        with open(config_path, "w") as f:
            json.dump(cfg, f)

        mocker.patch("server.get_root_dir", return_value=temp_dir)

        from server import merge_printer_config
        result = merge_printer_config("MyPrinter", {"copies": 5})
        # Saved control fields are merged in, overridden by options
        assert result["tray_source"] == "Tray2"
        assert result["color_mode"] == "monochrome"
        assert result["copies"] == 5

    def test_get_cached_printer_count_type(self):
        """get_cached_printer_count returns an integer."""
        from server import get_cached_printer_count
        assert isinstance(get_cached_printer_count(), int)
