from bs4 import BeautifulSoup
import re


class TableParser:
    """
    Parse HTML email bodies to extract structured metric rows.

    The HTML typically contains multiple <table> elements. The parser:
    1. Finds the correct table by matching its header row
    2. Extracts the overall system status from the surrounding text
    3. Validates and filters each data row before returning it
    """

    # Dimension values that indicate non-metric rows (headers, info, other tables)
    INVALID_DIMENSIONS = {
        "SupplementInfo", "TransferOwner",
        "ID", "REGION", "BUSINESS_CODE", "ALARMID", "ActiveSubs", "READY",
        "TITLE", "BIZCODE", "GLOBALNAME", "RETCODE", "RETDESC",
    }

    # Expected header row for the metrics table
    EXPECTED_HEADERS = [
        "Dimension", "Scenario", "Status",
        "PeakHours", "StandardValue", "CurrentValue",
    ]

    def _es_fila_valida(self, texts: list) -> bool:
        """Return True if the row contains valid metric data."""
        if len(texts) != 6:
            return False
        col0, col1 = texts[0], texts[1]

        if col0 == "Dimension" or col1 == "Scenario":
            return False

        if col0 in self.INVALID_DIMENSIONS:
            return False

        limpio_num = (
            col1.replace(".", "")
            .replace("-", "")
            .replace(" ", "")
            .replace(">", "")
            .replace("<", "")
            .replace("=", "")
        )
        if limpio_num.isdigit():
            return False

        limpio_dim = col0.replace(".", "").replace("-", "").replace(" ", "")
        if limpio_dim.isdigit():
            return False

        return True

    def parse(self, html_body: str) -> list:
        """Parse the HTML body and return a list of metric dictionaries."""
        soup = BeautifulSoup(html_body, "lxml")
        overall_status = self._extract_overall_status(soup)
        table = self._find_metrics_table(soup)
        if not table:
            raise ValueError("Metrics table not found in the HTML body.")

        rows = table.find_all("tr")
        results = []

        for row in rows:
            cols = row.find_all(["td", "th"])
            texts = [c.get_text(strip=True) for c in cols]

            if not self._es_fila_valida(texts):
                continue

            results.append({
                "overall_status": overall_status,
                "dimension": texts[0],
                "scenario": texts[1],
                "status": texts[2],
                "peak_hours": texts[3],
                "standard_value": texts[4],
                "current_value": texts[5],
            })

        return results

    def _find_metrics_table(self, soup):
        """Locate the first <table> whose header row matches EXPECTED_HEADERS."""
        tables = soup.find_all("table")
        for table in tables:
            first_row = table.find("tr")
            if not first_row:
                continue
            cells = first_row.find_all(["td", "th"])
            header_texts = [c.get_text(strip=True) for c in cells]
            if header_texts[:6] == self.EXPECTED_HEADERS:
                return table
        return None

    def _extract_overall_status(self, soup) -> str:
        """Extract the overall system status from the full HTML text."""
        text = soup.get_text(separator=" ", strip=True)
        if "Overall System Operation" in text:
            m = re.search(
                r"Overall System Operation\s*:\s*(Normal|Warning|Critical)",
                text,
                re.IGNORECASE,
            )
            if m:
                return m.group(1)
        return "Unknown"


if __name__ == "__main__":
    from src.extract.outlook_reader import OutlookReader

    reader = OutlookReader()
    emails = reader.get_all_emails()
    reader.close()

    if not emails:
        print("No test emails available.")
    else:
        sample = emails[0]
        print(f"Email: {sample['subject']}")

        parser = TableParser()
        try:
            rows = parser.parse(sample["html_body"])
            print(f"Rows extracted: {len(rows)}")
            for i, row in enumerate(rows[:3], 1):
                print(f"  [{i}] {row['dimension']} | {row['scenario']} | {row['status']}")
        except Exception as e:
            print(f"Parse error: {e}")
