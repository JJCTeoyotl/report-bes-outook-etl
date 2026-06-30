from datetime import datetime
from pathlib import Path

from ruta import DATA_OUTLOOK


class OutlookReader:
    # ------------------------------------------------------------------
    # CONFIGURATION
    # ------------------------------------------------------------------
    # Folder path within your Outlook mailbox or PST file where the
    # target emails are stored. Example: ["Inbox", "Reports", "System"]
    FOLDER_PATH = ["Inbox", "Reports"]

    # Keywords that identify the target email by its HTML body content.
    # The email must contain ALL specified strings to be processed.
    REPORT_KEYWORDS = ["System Brief Report", "Overall System Operation"]

    # ------------------------------------------------------------------

    def __init__(self):
        self._data_file = DATA_OUTLOOK if DATA_OUTLOOK else None
        self._use_data_file = False

        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        self._outlook = win32com.client.Dispatch("Outlook.Application")
        self._namespace = self._outlook.GetNamespace("MAPI")

        self._loaded_store = None

        if self._data_file and self._data_file.exists():
            self._use_data_file = True
            self._namespace.AddStore(str(self._data_file))

    def _find_folder_in_store(self):
        """Navigate the PST file store to locate the target folder."""
        for store in self._namespace.Stores:
            if not getattr(store, "IsDataFileStore", False):
                continue
            root = store.GetRootFolder()
            if root.Name.lower() != self.FOLDER_PATH[0].lower():
                continue
            try:
                folder = root
                for name in self.FOLDER_PATH[1:]:
                    folder = folder.Folders[name]
                self._loaded_store = store
                return folder
            except Exception:
                continue
        return None

    def _navigate_to_folder(self):
        """Locate the target folder in either the PST file or the default MAPI profile."""
        if self._use_data_file:
            folder = self._find_folder_in_store()
            if folder:
                return folder
            raise ValueError(
                f"Folder path {self.FOLDER_PATH} not found inside the PST file."
            )
        folder = self._namespace.Folders.Item(1)
        for name in self.FOLDER_PATH:
            folder = folder.Folders[name]
        return folder

    def _read_from_folder(self, since_datetime: datetime | None = None):
        """Iterate emails and yield those whose HTML body matches REPORT_KEYWORDS."""
        folder = self._navigate_to_folder()
        messages = folder.Items

        orden_descendente = (__name__ == "__main__")
        messages.Sort("[ReceivedTime]", Descending=orden_descendente)

        emails = []
        for msg in messages:
            if not hasattr(msg, "ReceivedTime"):
                continue

            received = msg.ReceivedTime
            if received.tzinfo:
                received = received.replace(tzinfo=None)

            if since_datetime and received < since_datetime:
                continue

            html = msg.HTMLBody or ""
            if all(kw in html for kw in self.REPORT_KEYWORDS):
                emails.append({
                    "entry_id": msg.EntryID,
                    "subject": msg.Subject,
                    "received_time": received,
                    "html_body": html,
                })

                if __name__ == "__main__":
                    break

        return emails

    def get_emails_since(self, since_datetime: datetime):
        """Return matching emails received at or after the given datetime."""
        return self._read_from_folder(since_datetime)

    def get_all_emails(self):
        """Return all matching emails in the folder (no temporal filter)."""
        return self._read_from_folder()

    def close(self):
        """Release COM resources and unload the PST file if one was loaded."""
        if self._loaded_store:
            try:
                self._namespace.RemoveStore(self._loaded_store)
            except Exception:
                pass
        try:
            import pythoncom

            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    print("Starting reader self-test...")
    reader = OutlookReader()
    print(f"Using external PST: {reader._use_data_file}")
    print(f"Searching folder: {reader.FOLDER_PATH}")
    print(f"Looking for keywords: {reader.REPORT_KEYWORDS}")
    test_emails = reader.get_all_emails()
    print(f"Emails found: {len(test_emails)}")
    if test_emails:
        print(f"Sample subject: {test_emails[0]['subject']}")
        print(f"Sample date: {test_emails[0]['received_time']}")
        print(f"HTML size: {len(test_emails[0]['html_body'])} chars")
    else:
        print("No matching emails found.")
    reader.close()
