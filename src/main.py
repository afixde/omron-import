from pathlib import Path
from csv_reader import OmronCsvReader
from excel_service import ExcelService
from config import CSV_DIR, EXCEL_FILE, BACKUP_DIR, ARCHIVE_DIR
from file_utils import archive_file

import sys

def main() -> None:

    print("=" * 50)
    print(" Omron Import")
    print("=" * 50)

    try:
        reader = OmronCsvReader()

        csv_file = reader.find_latest_csv(CSV_DIR)

        print(f"CSV gefunden: {csv_file.name}")

        measurements = reader.read(csv_file)

        print(f"{len(measurements)} Messungen eingelesen")

#        if measurements:
#            print()
#            print("Erste Messung:")
#            print(measurements[0])
            
        excel = ExcelService(EXCEL_FILE)

        backup = excel.backup_workbook(BACKUP_DIR)
        print(f"Backup erstellt: {backup.name}")
        
        excel.open()

        existing = excel.get_existing_keys()

        new_measurements = [
            m for m in measurements
            if m.key not in existing
        ]
        
        print(f"{len(new_measurements)} neue Messungen gefunden")
        
        if new_measurements:
            print(f"{len(new_measurements)} neue Messungen werden importiert...")
        
            for measurement in new_measurements:
                excel.append_measurement(measurement)

            excel.sort_table()
            excel.save()
            print("Import abgeschlossen.")
            print("Excel gespeichert.")
            archived = archive_file(
                csv_file,
                ARCHIVE_DIR
            )
            print()
            print(f"CSV archiviert: {archived.name}")        
#        print()
#        print(f"{len(existing)} vorhandene Messungen in Excel")

        excel.close()

    except FileNotFoundError:
        print()
        print("Keine OMRON-CSV gefunden.")
        print()
        print("Bitte kopieren Sie eine OMRON-CSV nach:")
        print("data\\csv")
        print()
        print("Anschließend starten Sie das Programm erneut.")
        sys.exit(1)
        
    except Exception as ex:
        print()
        print("Unerwarteter Fehler:")
        print(ex)
        sys.exit(1)

    input("\nDrücken Sie ENTER zum Beenden...")


if __name__ == "__main__":
    main()