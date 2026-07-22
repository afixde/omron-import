from csv_reader import OmronCsvReader
from excel_service import ExcelService
from statistics_service import StatisticsService
from chart_service import ChartService
from helper_columns_service import HelperColumnsService
from datetime import date, time, datetime, timedelta

from config import (
    CSV_DIR,
    EXCEL_FILE,
    BACKUP_DIR,
    ARCHIVE_DIR,
)
from file_utils import archive_file
from logger_config import get_logger

logger = get_logger()


class ImportService:

    def __init__(self):
        self.reader = OmronCsvReader()
        self.stats = StatisticsService()
        self.excel = ExcelService(EXCEL_FILE)
        self.helper_columns_service = HelperColumnsService()        
        self.chart = None
        self.csv_file = None
        self.backup_file = None
        self.archive_file = None
        
        self.measurements = []
        self.excel_measurements = []
        self.new_measurements = []
        
    def _read_csv(self):
        self.csv_file = self.reader.find_latest_csv(CSV_DIR)
        logger.info("CSV gefunden: %s", self.csv_file.name)
        self.csv_measurements = self.reader.read(self.csv_file)
        logger.info("%s Messungen eingelesen", len(self.csv_measurements))
        print(f"{len(self.csv_measurements)} Messungen eingelesen")
    
    def _find_new_measurements(self):
        existing_keys = self.excel.get_existing_keys()

        self.new_measurements = [
            m
            for m in self.csv_measurements
            if (
                f"{m.date.isoformat()} {m.time.strftime('%H:%M')}",
                m.systolic,
                m.diastolic,
                m.pulse,
            ) not in existing_keys
        ]
                
        print(f"{len(self.new_measurements)} neue Messungen gefunden")
        logger.info(
            "%s neue Messungen gefunden",
            len(self.new_measurements)
        )
    
    def _import_measurements(self) -> bool:
        if not self.new_measurements:
            print("Keine neuen Messungen vorhanden.")
            return False
        print(
            f"{len(self.new_measurements)} neue Messungen werden eingefügt..."
        )
        helpers = self.helper_columns_service.build(self.new_measurements)
    
        for measurement, helper in zip(self.new_measurements, helpers):
            self.excel.append_measurement(
                measurement,
                helper,
            )
        return True
    
    def _open_excel(self):
        self.excel.open()
        logger.info("Excel geöffnet")
        self.chart = ChartService(
            self.excel.worksheet,
            self.excel.table
        )
        self.backup_file = self.excel.backup_workbook(BACKUP_DIR)
        print(f"Backup erstellt: {self.backup_file.name}")

    def _sort_table(self):
        self.excel.sort_table()

    def _update_chart(self):
#        self.chart.create_or_update()
        self.chart.update_chart()
        print("Chart abgeschlossen")

    def _save_excel(self):
        self.excel.save()
        print("Excel gespeichert.")

    def _archive_csv(self):
        self.archive_file = archive_file(
            self.csv_file,
            ARCHIVE_DIR
        )
        print(f"\nCSV archiviert: {self.archive_file.name}")

    def _close_excel(self):
        self.excel.close()

    def _show_statistics(self):
        self.excel_measurements = self.excel.get_all_measurements()
        self.stats.report(self.excel_measurements)

    def run(self):
        """
        Führt den kompletten Import durch.
        """
        try:
            self._read_csv()
            self._open_excel()
            self._show_statistics()
            self._find_new_measurements()
            if self._import_measurements():
                self._sort_table()
                self._update_chart()
                self._save_excel()   
                self._archive_csv()  
    
        except Exception as ex:
            print()
            print("Unerwarteter Fehler:")
            print(ex)
            logger.exception("Unerwarteter Fehler")
             
        finally:
            self._close_excel()

        print("weiter...")
       