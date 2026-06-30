import re
import sqlite3
from datetime import datetime
from ruta import DB_PATH


class SQLiteManager:
    """Persist parsed metric rows to a local SQLite database."""

    def __init__(self):
        self.conn = sqlite3.connect(str(DB_PATH))
        self._create_tables()

    def _create_tables(self):
        cur = self.conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS bes_status (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entry_id TEXT NOT NULL,
                received_time TEXT NOT NULL,
                overall_status TEXT,
                dimension TEXT,
                scenario TEXT,
                status TEXT,
                peak_hours TEXT,
                standard_value TEXT,
                current_value TEXT,
                inserted_at TEXT DEFAULT (datetime('now','localtime')),
                UNIQUE(entry_id, dimension, scenario, received_time)
            )
        """)
        self.conn.commit()

    @staticmethod
    def _es_entry_id_valido(entry_id: str) -> bool:
        """Reject entry IDs that look like timestamps rather than real Outlook EntryIDs.

        Real Outlook EntryIDs are 72-character hex strings. Timestamp-like
        IDs are purely numeric and start with a 4-digit year.
        """
        if (
            len(entry_id) >= 12
            and re.match(r"^\d{4}", entry_id)
            and entry_id.replace("-", "").isdigit()
        ):
            return False
        return True

    def insert_status_batch(
        self, entry_id: str, received_time: str, rows_list: list
    ) -> bool:
        """Insert a batch of parsed rows, skipping duplicates.

        Returns True if at least one new row was inserted.
        """
        if not self._es_entry_id_valido(entry_id):
            print(
                f" Skipping batch: entry_id looks like a timestamp "
                f"({entry_id[:30]}...)"
            )
            return False

        cur = self.conn.cursor()
        inserted_count = 0

        for row in rows_list:
            cur.execute(
                """
                SELECT 1 FROM bes_status
                WHERE entry_id = ? AND dimension = ? AND scenario = ? AND received_time = ?
                """,
                (entry_id, row.get("dimension"), row.get("scenario"), received_time),
            )

            if cur.fetchone() is not None:
                continue

            cur.execute(
                """
                INSERT INTO bes_status
                    (entry_id, received_time, overall_status,
                     dimension, scenario, status, peak_hours, standard_value, current_value)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    entry_id,
                    received_time,
                    row.get("overall_status"),
                    row.get("dimension"),
                    row.get("scenario"),
                    row.get("status"),
                    row.get("peak_hours"),
                    row.get("standard_value"),
                    row.get("current_value"),
                ),
            )
            inserted_count += 1

        self.conn.commit()
        return inserted_count > 0

    def get_records_by_date(self, date_str: str) -> list:
        """Return all records for a given date (YYYY-MM-DD)."""
        cur = self.conn.cursor()
        cur.execute(
            "SELECT * FROM bes_status WHERE date(received_time) = ? ORDER BY received_time ASC",
            (date_str,),
        )
        columns = [desc[0] for desc in cur.description]
        return [dict(zip(columns, row)) for row in cur.fetchall()]

    def close(self):
        self.conn.close()
