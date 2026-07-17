from pathlib import Path
from csv_reader import OmronCsvReader
from excel_service import ExcelService
from config import CSV_DIR, EXCEL_FILE, BACKUP_DIR, ARCHIVE_DIR
from file_utils import archive_file
from logger_config import get_logger
from version import VERSION
from statistics_service import StatisticsService
from console import info

import sys

logger = get_logger()

def show_statistics(excel: ExcelService) -> None:
    """
    Zeigt die Statistik aller Messungen an.
    """

    stats = StatisticsService()

    all_measurements = excel.get_all_measurements()

    summary = stats.summarize(all_measurements)

    print()
    print("Statistik")
    print("-" * 30)

    print(f"Messungen : {summary['count']}")

    print()

    print(f"SYS Ø      : {summary['avg_sys']} mmHg")
    print(f"SYS Median : {summary['median_sys']} mmHg")
    print(f"SYS Min    : {summary['min_sys']} mmHg")
    print(f"SYS Max    : {summary['max_sys']} mmHg")

    print()

    print(f"DIA Ø      : {summary['avg_dia']} mmHg")
    print(f"DIA Median : {summary['median_dia']} mmHg")
    print(f"DIA Min    : {summary['min_dia']} mmHg")
    print(f"DIA Max    : {summary['max_dia']} mmHg")
    
    print()
    
    print(f"Puls Ø     : {summary['avg_pulse']} bpm")
    print(f"Puls Median: {summary['median_pulse']} bpm")
    print(f"Puls Min   : {summary['min_pulse']} bpm")
    print(f"Puls Max   : {summary['max_pulse']} bpm")

def main() -> None:

    print("=" * 50)
    print(f" Omron Import v{VERSION}")
    print("=" * 50)
    
    logger.info("Programmstart")

    try:
        reader = OmronCsvReader()

        csv_file = reader.find_latest_csv(CSV_DIR)
        logger.info("CSV gefunden: %s", csv_file.name)

        measurements = reader.read(csv_file)
        logger.info("%s Messungen eingelesen", len(measurements))        

        print(f"{len(measurements)} Messungen eingelesen")
        stats = StatisticsService()
        summary = stats.summarize(measurements)
           
        excel = ExcelService(EXCEL_FILE)
        logger.info("Excel geöffnet")

        backup = excel.backup_workbook(BACKUP_DIR)
        print(f"Backup erstellt: {backup.name}")
        
        excel.open()

        show_statistics(excel)

        all_measurements = excel.get_all_measurements()
        summary = stats.summarize(all_measurements)
       
        print()
        print(f"Excel enthält {len(all_measurements)} Messungen")
        if all_measurements:
            print()
            print("Erste Messung aus Excel:")
            print(all_measurements[0])
    
        existing = excel.get_existing_keys()
        logger.info("%d vorhandene Messungen", len(existing))

        new_measurements = [
            m for m in measurements
            if m.key not in existing
        ]
        
        logger.info("%d neue Messungen", len(new_measurements))
        print(f"{len(new_measurements)} neue Messungen gefunden")
        
        if new_measurements:
            print(f"{len(new_measurements)} neue Messungen werden importiert...")
        
            for measurement in new_measurements:
                excel.append_measurement(measurement)

            excel.sort_table()
            excel.create_chart()
            excel.save()
            print("Import abgeschlossen.")
            print("Excel gespeichert.")

            archived = archive_file(
                csv_file,
                ARCHIVE_DIR
            )
            print()
            print(f"CSV archiviert: {archived.name}")        
            logger.info("Import erfolgreich abgeschlossen")#        print()
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
        logger.exception("Unerwarteter Fehler")
        sys.exit(1)

    input("\nDrücken Sie ENTER zum Beenden...")


if __name__ == "__main__":
    main()