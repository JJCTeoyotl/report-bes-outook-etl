import argparse
from datetime import datetime, timedelta
import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent))

from src.extract.outlook_reader import OutlookReader
from src.discover.table_parser import TableParser
from src.database.sqlite_manager import SQLiteManager
from src.control.processed_emails import ProcessedEmails
from src.export.excel_exporter import ExcelExporter


def main():
    parser = argparse.ArgumentParser(
        description="Extract structured metric data from Outlook email reports."
    )
    parser.add_argument(
        "--mode",
        choices=["PRODUCTION", "HISTORICAL", "TEST"],
        default="PRODUCTION",
        help="Execution mode (default: PRODUCTION)",
    )
    parser.add_argument(
        "--days",
        type=int,
        default=2,
        help="Days to look back in TEST mode (default: 2)",
    )
    parser.add_argument(
        "--hours",
        type=int,
        default=0,
        help="Hours to look back in TEST mode (default: 0)",
    )
    args = parser.parse_args()

    print("=" * 60)
    print("Outlook Email ETL Pipeline")
    print("=" * 60)

    reader = OutlookReader()
    parser_engine = TableParser()
    db = SQLiteManager()
    control = ProcessedEmails()
    exporter = ExcelExporter()

    try:
        MODO_EJECUCION = args.mode
        TEST_HOURS = args.hours
        TEST_DAYS = args.days

        print(f"\n[STEP 1] Determining search range (mode: {MODO_EJECUCION})")

        if MODO_EJECUCION == "PRODUCTION":
            last_time = control.get_last_processed_time()
            if last_time:
                fecha_busqueda = last_time
                print(
                    f" PRODUCTION mode active. Searching from: "
                    f"{fecha_busqueda.strftime('%Y-%m-%d %H:%M:%S')}"
                )
            else:
                fecha_busqueda = datetime.now() - timedelta(days=90)
                print(
                    f" First PRODUCTION run detected. Starting 90-day lookback: "
                    f"{fecha_busqueda.strftime('%Y-%m-%d %H:%M:%S')}"
                )

        elif MODO_EJECUCION == "HISTORICAL":
            fecha_busqueda = datetime.now() - timedelta(days=90)
            print(
                f" HISTORICAL mode. Loading last 90 days from: "
                f"{fecha_busqueda.strftime('%Y-%m-%d %H:%M:%S')}"
            )

        elif MODO_EJECUCION == "TEST":
            fecha_busqueda = datetime.now() - timedelta(
                days=TEST_DAYS, hours=TEST_HOURS
            )
            print(
                f" TEST mode. Narrow range from: "
                f"{fecha_busqueda.strftime('%Y-%m-%d %H:%M:%S')}"
            )

        correos = reader.get_emails_since(fecha_busqueda)
        total_correos = len(correos)
        print(f" Emails found in Outlook: {total_correos}")

        if total_correos > 0:
            print(f"\n[STEP 2] Processing {total_correos} emails...")

            correos.reverse()

            success_count = 0
            total_rows_inserted = 0
            days_in_batch = set()
            max_success_timestamp = None

            for idx, correo in enumerate(correos, start=1):
                try:
                    rows = parser_engine.parse(correo["html_body"])
                    correo_id = correo.get("entry_id", f"AUTO_ID_{idx}")
                    received_str = correo["received_time"].strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

                    fecha_corta = correo["received_time"].strftime("%Y-%m-%d")
                    days_in_batch.add(fecha_corta)

                    inserted_new = db.insert_status_batch(
                        entry_id=correo_id,
                        received_time=received_str,
                        rows_list=rows,
                    )

                    if inserted_new:
                        success_count += 1
                        total_rows_inserted += len(rows)
                        if (
                            max_success_timestamp is None
                            or correo["received_time"] > max_success_timestamp
                        ):
                            max_success_timestamp = correo["received_time"]
                        print(
                            f" [{idx}/{total_correos}] "
                            f"{correo['subject'][:35]}... "
                            f"-> {len(rows)} rows inserted ({fecha_corta})"
                        )
                    else:
                        print(
                            f" [{idx}/{total_correos}] "
                            f"{correo['subject'][:35]}... "
                            f"(skipped — all rows already in DB)"
                        )

                except Exception as e:
                    print(
                        f" [{idx}/{total_correos}] "
                        f"Error processing email: {e}"
                    )
                    continue

            if MODO_EJECUCION == "PRODUCTION":
                if max_success_timestamp is not None:
                    control.update_control(max_success_timestamp)
                    print(
                        f"\n[STEP 3] Control timestamp updated to: "
                        f"{max_success_timestamp}"
                    )
                else:
                    print(
                        "\n[STEP 3] Control NOT updated — "
                        "no emails were successfully processed."
                    )
            else:
                print(
                    f"\n[STEP 3] Control update skipped "
                    f"(mode: {MODO_EJECUCION})."
                )

            print(
                f"\n[STEP 4] Exporting Excel files for days: "
                f"{sorted(days_in_batch)}"
            )

            for day in sorted(days_in_batch):
                print(f" Generating report for {day}...")
                exporter.export_by_date(day)

            print(
                f"\n SUMMARY: {success_count} emails processed, "
                f"{total_rows_inserted} rows inserted, "
                f"{len(days_in_batch)} Excel files generated."
            )

        else:
            print("\nNo new emails in the specified range. System is up to date.")
            print("[STEP 4] Updating today's Excel report...")
            exporter.export_today()

    except Exception as e:
        print(f"\n[CRITICAL ERROR] ETL pipeline failed: {e}")

    finally:
        reader.close()
        db.close()
        control.close()
        exporter.close()
        print("\nPipeline finished.")


if __name__ == "__main__":
    main()
