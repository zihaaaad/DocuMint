import sqlite3
import os
from datetime import datetime
from typing import Dict, Any, List, Optional

class DatabaseManager:
    """ACID SQLite Audit Ledger for tracking DocuMint job runs and itemized recipient execution histories."""
    
    def __init__(self, db_path: str):
        self.db_path = os.path.abspath(db_path)
        db_dir = os.path.dirname(self.db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir, exist_ok=True)
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path, timeout=10.0)
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA busy_timeout=5000;")
        return conn

    def _init_db(self) -> None:
        conn = self._get_connection()
        try:
            c = conn.cursor()
            c.execute('''
                CREATE TABLE IF NOT EXISTS jobs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT,
                    total_items INTEGER,
                    success_count INTEGER,
                    fail_count INTEGER,
                    mode TEXT
                )
            ''')
            c.execute('''
                CREATE TABLE IF NOT EXISTS job_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_id INTEGER,
                    recipient_name TEXT,
                    recipient_email TEXT,
                    status TEXT,
                    error_details TEXT,
                    pdf_path TEXT,
                    timestamp TEXT,
                    FOREIGN KEY(job_id) REFERENCES jobs(id) ON DELETE CASCADE
                )
            ''')
            conn.commit()
        finally:
            conn.close()

    def log_job(
        self, 
        total: int, 
        success: int, 
        fail: int, 
        mode: str = "Standard", 
        records: Optional[List[Dict[str, Any]]] = None
    ) -> int:
        """Logs a completed batch execution job along with optional individual recipient item records."""
        conn = self._get_connection()
        try:
            c = conn.cursor()
            now_iso = datetime.now().isoformat()
            c.execute('''
                INSERT INTO jobs (timestamp, total_items, success_count, fail_count, mode)
                VALUES (?, ?, ?, ?, ?)
            ''', (now_iso, int(total), int(success), int(fail), str(mode)))
            job_id = c.lastrowid
            
            if records and job_id:
                for rec in records:
                    name = rec.get("Name", "")
                    email = rec.get("Email", "")
                    status = rec.get("Status", "")
                    error_details = rec.get("Error", "")
                    pdf_path = rec.get("PdfPath", "")
                    ts = rec.get("Timestamp", now_iso)
                    if isinstance(ts, datetime):
                        ts = ts.isoformat()
                    
                    c.execute('''
                        INSERT INTO job_items (job_id, recipient_name, recipient_email, status, error_details, pdf_path, timestamp)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (job_id, str(name), str(email), str(status), str(error_details), str(pdf_path), str(ts)))
                    
            conn.commit()
            return job_id
        finally:
            conn.close()

    def get_stats(self) -> Dict[str, Any]:
        """Returns aggregate execution statistics and recent batch histories."""
        conn = self._get_connection()
        conn.row_factory = sqlite3.Row
        try:
            c = conn.cursor()
            
            # Aggregate stats
            c.execute("SELECT COUNT(*) as total_jobs, COALESCE(SUM(total_items), 0) as total_emails, COALESCE(SUM(success_count), 0) as total_success, COALESCE(SUM(fail_count), 0) as total_failed FROM jobs")
            row = c.fetchone()
            summary = dict(row) if row else {"total_jobs": 0, "total_emails": 0, "total_success": 0, "total_failed": 0}
            
            # Success Rate Calculation
            total_processed = summary.get("total_success", 0) + summary.get("total_failed", 0)
            summary["success_rate"] = round((summary.get("total_success", 0) / total_processed * 100), 1) if total_processed > 0 else 100.0
            
            # Recent jobs
            c.execute("SELECT * FROM jobs ORDER BY id DESC LIMIT 10")
            recent = [dict(r) for r in c.fetchall()]

            # Recent item logs (latest 25 items across jobs)
            c.execute("SELECT * FROM job_items ORDER BY id DESC LIMIT 25")
            recent_items = [dict(r) for r in c.fetchall()]
            
            return {
                "summary": summary,
                "recent": recent,
                "recent_items": recent_items
            }
        finally:
            conn.close()

    def get_job_details(self, job_id: int) -> Optional[Dict[str, Any]]:
        """Returns full metadata and recipient items for a specific job."""
        conn = self._get_connection()
        conn.row_factory = sqlite3.Row
        try:
            c = conn.cursor()
            c.execute("SELECT * FROM jobs WHERE id = ?", (job_id,))
            job_row = c.fetchone()
            if not job_row:
                return None
            
            c.execute("SELECT * FROM job_items WHERE job_id = ? ORDER BY id ASC", (job_id,))
            items = [dict(r) for r in c.fetchall()]
            
            result = dict(job_row)
            result["items"] = items
            return result
        finally:
            conn.close()
