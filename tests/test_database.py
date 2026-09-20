import os
import sys
import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from documint.database import DatabaseManager

def test_database_initialization_and_logging(tmp_path):
    db_file = tmp_path / "test_history.db"
    db = DatabaseManager(str(db_file))

    # Log first job with items
    items = [
        {"Name": "Alice", "Email": "alice@test.com", "Status": "Success", "PdfPath": "/tmp/a.pdf"},
        {"Name": "Bob", "Email": "invalid", "Status": "Failed: Invalid Email", "Error": "Invalid Email"}
    ]
    job_id = db.log_job(total=2, success=1, fail=1, mode="SMTP (Parallel)", records=items)
    assert job_id == 1

    # Log second job
    db.log_job(total=1, success=1, fail=0, mode="Outlook (Serial)")

    # Fetch Stats
    stats = db.get_stats()
    summary = stats["summary"]
    assert summary["total_jobs"] == 2
    assert summary["total_emails"] == 3
    assert summary["total_success"] == 2
    assert summary["total_failed"] == 1
    assert summary["success_rate"] == 66.7

    # Check Recent Jobs
    assert len(stats["recent"]) == 2
    assert stats["recent"][0]["mode"] == "Outlook (Serial)"

    # Check Job Details
    details = db.get_job_details(1)
    assert details is not None
    assert details["id"] == 1
    assert details["total_items"] == 2
    assert len(details["items"]) == 2
    assert details["items"][0]["recipient_name"] == "Alice"
    assert details["items"][0]["status"] == "Success"
    assert details["items"][1]["recipient_name"] == "Bob"
    assert details["items"][1]["status"] == "Failed: Invalid Email"
