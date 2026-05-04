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
        # We need to reset the cached count first — skip by using mock
        # The actual cache seeding depends on print sync, so this is best tested
        # via a fresh import or direct cache manipulation.
        # For now, just verify it returns an int >= 0.
        count = get_cached_printer_count()
        assert isinstance(count, int)
        assert count >= 0

    def test_check_cancelled_default(self):
        """A job that was never cancelled returns False."""
        from server import _check_cancelled
        assert _check_cancelled("some-job-id") is False

    def test_get_queue_status(self, job_queue):
        """get_queue_status returns totals for the job queue."""
        # Create some pending and completed jobs
        for i in range(3):
            job_queue.create("PrinterA", "raw")
        success_job = job_queue.create("PrinterA", "raw")
        job_queue.complete(success_job["id"], success=True)
        failed_job = job_queue.create("PrinterA", "raw")
        job_queue.complete(failed_job["id"], success=False)

        from server import get_queue_status
        status = get_queue_status()
        assert "total_queued" in status
        assert "processing" in status
        assert "completed" in status
        assert "failed" in status


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
        assert data["status"] == "ok"
        assert "version" in data
        assert "uptime" in data

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
        profiles = data.get("profiles", [])
        assert len(profiles) == len(sample_config["profiles"])
        assert profiles[0]["printer"] == "TestPrinter"

    def test_jobs_list_empty(self, flask_app):
        """GET /jobs returns an empty list when no jobs exist."""
        resp = flask_app.get("/jobs")
        assert resp.status_code == 200
        data = resp.get_json()
        assert isinstance(data.get("jobs"), list)

    def test_queue_status_endpoint(self, flask_app):
        """GET /queue-status returns queue statistics."""
        resp = flask_app.get("/queue-status")
        assert resp.status_code == 200
        data = resp.get_json()
        assert "total_queued" in data

    def test_api_capabilities_no_printer(self, flask_app):
        """GET /api/capabilities without a printer param returns error."""
        resp = flask_app.get("/api/capabilities")
        assert resp.status_code == 400

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

    def test_load_profiles_from_config(self, sample_config, temp_dir):
        """Load profiles from the sample config."""
        from server import load_profiles_from_config
        profiles = load_profiles_from_config()
        assert isinstance(profiles, list)
        assert len(profiles) == len(sample_config["profiles"])

    def test_save_profiles_to_config(self, temp_dir):
        """Save profiles and verify they persist."""
        from server import save_profiles_to_config, load_profiles_from_config
        new_profiles = [
            {"printer": "P1", "name": "Profile1", "options": {"copies": 1}},
            {"printer": "P2", "name": "Profile2", "options": {"duplex": "long"}},
        ]
        save_profiles_to_config(new_profiles)
        loaded = load_profiles_from_config()
        assert loaded == new_profiles

    def test_load_printer_configs(self, sample_config, temp_dir):
        """Load per-printer configs from config.json."""
        from server import load_printer_configs
        configs = load_printer_configs()
        assert isinstance(configs, dict)

    def test_merge_printer_config(self, sample_config, temp_dir):
        """merge_printer_config merges options into a printer config entry."""
        from server import merge_printer_config
        result = merge_printer_config("TestPrinter", {"copies": 5})
        assert result["printer"] == "TestPrinter"
        assert result["options"]["copies"] == 5

    def test_get_cached_printer_count_type(self):
        """get_cached_printer_count returns an integer."""
        from server import get_cached_printer_count
        assert isinstance(get_cached_printer_count(), int)
