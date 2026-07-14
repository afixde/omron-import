from pathlib import Path

from csv_reader import OmronCsvReader

from excel_service import ExcelService

import sys

def main() -> None:

    print("=" * 50)
    print(" Omron Import")
    print("=" * 50)

    try:
        reader = OmronCsvReader()

        csv_dir = Path("data") / "csv"

        csv_file = reader.find_latest_csv(csv_dir)

        print(f"CSV gefunden: {csv_file.name}")

        measurements = reader.read(csv_file)

        print(f"{len(measurements)} Messungen eingelesen")

        if measurements:
            print()
            print("Erste Messung:")
            print(measurements[0])
            
        excel = ExcelService(
            Path("data") / "excel" / "Omron.xlsx"
)
        excel.open()

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