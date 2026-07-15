from pathlib import Path

from csv_reader import OmronCsvReader

from excel_service import ExcelService

from config import CSV_DIR, EXCEL_FILE

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

        if measurements:
            print()
            print("Erste Messung:")
            print(measurements[0])
            
        excel = ExcelService(EXCEL_FILE)
        excel.open()

        print(excel)
        print(excel.table)

        existing = excel.get_existing_keys()

        new_measurements = [
            m for m in measurements
            if m.key not in existing
        ]
        
        print(f"{len(new_measurements)} neue Messungen gefunden")
        
        if new_measurements:
            print("Schreibe erste neue Messung...")
            excel.append_measurement(new_measurements[0])
            excel.save()
        
        print()
        print(f"{len(existing)} vorhandene Messungen in Excel")


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