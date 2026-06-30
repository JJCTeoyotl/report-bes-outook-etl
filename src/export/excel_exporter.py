from datetime import datetime
import pandas as pd
from ruta import PROCESSED_DIR, SYNC_DIR
from src.database.sqlite_manager import SQLiteManager


class ExcelExporter:
    """Export database records to Excel files, organized by date."""

    def __init__(self):
        self.db = SQLiteManager()

    def export_today(self):
        """Convenience: export records for the current date."""
        today = datetime.now().strftime("%Y-%m-%d")
        return self.export_by_date(today)

    def export_by_date(self, date_str: str):
        """Generate an Excel file for all records matching the given date."""
        records = self.db.get_records_by_date(date_str)
        if not records:
            print(f"No database records found for {date_str}")
            return None

        df = pd.DataFrame(records)

        column_mapping = {
            "id": "Row ID",
            "received_time": "Received Time (Outlook)",
            "inserted_at": "Extraction Time (System)",
            "overall_status": "Overall Status",
            "dimension": "Dimension",
            "scenario": "Scenario",
            "status": "Status",
            "peak_hours": "Peak Hours",
            "standard_value": "Standard Value",
            "current_value": "Current Value",
            "entry_id": "Outlook Entry ID",
        }

        df = df.rename(columns=column_mapping)

        ordered_columns = [
            "Row ID",
            "Received Time (Outlook)",
            "Extraction Time (System)",
            "Overall Status",
            "Dimension",
            "Scenario",
            "Status",
            "Peak Hours",
            "Standard Value",
            "Current Value",
            "Outlook Entry ID",
        ]

        df = df[[col for col in ordered_columns if col in df.columns]]

        filename = f"report_{date_str.replace('-', '_')}.xlsx"
        local_path = PROCESSED_DIR / filename
        paths_to_write = [local_path]

        if SYNC_DIR:
            sync_path = SYNC_DIR / filename
            SYNC_DIR.mkdir(parents=True, exist_ok=True)
            paths_to_write.append(sync_path)

        PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

        for path in paths_to_write:
            df.to_excel(path, index=False)
            print(f"Exported: {path}")

        return str(local_path)

    def close(self):
        self.db.close()
