from pathlib import Path

# ---------------------------------------------------------------------------
# CONFIGURATION — Edit these paths to match your environment
# ---------------------------------------------------------------------------

# Path to an Outlook PST data file (optional). Set to None if using the
# default Outlook profile instead of a standalone PST archive.
# Example: DATA_OUTLOOK = Path.home() / "Documents" / "Outlook Files" / "archive.pst"
DATA_OUTLOOK = None

# Project root (resolved automatically — do not change)
BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = BASE_DIR / "data"
RAW_HTML_DIR = DATA_DIR / "raw" / "html"
PROCESSED_DIR = DATA_DIR / "processed"
DATABASE_DIR = DATA_DIR / "database"

DB_PATH = DATABASE_DIR / "bes.db"

# Directory for syncing output Excel files to cloud storage (optional).
# Set to None to skip the cloud copy.
# Example: SYNC_DIR = Path.home() / "OneDrive" / "bes_extraction"
SYNC_DIR = Path.home() / "OneDrive" / "bes_extraction"

for d in [RAW_HTML_DIR, PROCESSED_DIR, DATABASE_DIR]:
    d.mkdir(parents=True, exist_ok=True)

if SYNC_DIR:
    SYNC_DIR.mkdir(parents=True, exist_ok=True)


if __name__ == "__main__":
    print(f"DATA_OUTLOOK configured: {DATA_OUTLOOK.exists() if DATA_OUTLOOK else 'No PST'}")
    print(f"RAW_HTML_DIR exists: {RAW_HTML_DIR.exists()}")
    print(f"PROCESSED_DIR exists: {PROCESSED_DIR.exists()}")
    print(f"DATABASE_DIR exists: {DATABASE_DIR.exists()}")
    print(f"SYNC_DIR exists: {SYNC_DIR.exists() if SYNC_DIR else 'Not configured'}")
