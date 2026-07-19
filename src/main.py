from pathlib import Path
from csv_reader import OmronCsvReader
from excel_service import ExcelService
from config import CSV_DIR, EXCEL_FILE, BACKUP_DIR, ARCHIVE_DIR
from file_utils import archive_file
from logger_config import get_logger
from version import VERSION
from statistics_service import StatisticsService
from console import info
from chart_service import ChartService
from import_service import ImportService

import sys

logger = get_logger()

def show_statistics(excel: ExcelService) -> None:
    """
    Zeigt die Statistik aller Messungen an.
    """

    stats = StatisticsService()
    all_measurements = excel.get_all_measurements()
    stats.report(all_measurements)

def main() -> None:
    service = ImportService()
    service.run()

def main1() -> None:

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
        excel.open()
        logger.info("Excel geöffnet")

        backup = excel.backup_workbook(BACKUP_DIR)
        print(f"Backup erstellt: {backup.name}")
        
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
            print("Import abgeschlossen.")
            excel.create_chart()
#            chart = ChartService(
#                excel.worksheet,
#                excel.table
#            )
#            chart.update_chart()
            print("Chart abgeschlossen")
            excel.save()
            print("Excel gespeichert.")

            archived = archive_file(
                csv_file,
                ARCHIVE_DIR
            )
            print()
            print(f"CSV archiviert: {archived.name}")        
            logger.info("Import erfolgreich abgeschlossen")#        print()
#        print(f"{len(existing)} vorhandene Messungen in Excel")

#        excel.close()

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

    finally:
        excel.close()

    input("\nDrücken Sie ENTER zum Beenden...")


if __name__ == "__main__":
    main()