import sqlite3
from datetime import datetime
from ruta import DB_PATH


class ProcessedEmails:
    """Track incremental execution state so each run only processes new emails."""

    def __init__(self):
        self.conn = sqlite3.connect(str(DB_PATH))
        self._ensure_control_table()

    def _ensure_control_table(self):
        cur = self.conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS etl_control (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                last_execution TEXT,
                last_email_received TEXT,
                created_at TEXT DEFAULT (datetime('now','localtime'))
            )
        """)
        self.conn.commit()

    def get_last_processed_time(self) -> datetime | None:
        """Return the most recent processed email timestamp, or None."""
        cur = self.conn.cursor()
        cur.execute(
            "SELECT last_email_received FROM etl_control ORDER BY id DESC LIMIT 1"
        )
        row = cur.fetchone()
        if row and row[0]:
            return datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S")
        return None

    def update_control(self, last_received: datetime):
        """Record a new execution with the latest processed email timestamp."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        last_str = last_received.strftime("%Y-%m-%d %H:%M:%S")
        cur = self.conn.cursor()
        cur.execute(
            "INSERT INTO etl_control (last_execution, last_email_received) VALUES (?, ?)",
            (now, last_str),
        )
        self.conn.commit()

    def close(self):
        self.conn.close()
