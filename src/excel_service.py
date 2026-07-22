from pathlib import Path
from datetime import date, time, datetime, timedelta
from win32com.client import constants
from models import Measurement, HelperRow

import win32com.client
import shutil

class ExcelService:
    """Zugriff auf die Omron-Exceldatei."""

    def __init__(self, workbook_path: Path):
        self.workbook_path = workbook_path

        self.excel = None
        self.workbook = None
        self.worksheet = None
        self.table = None
        self._formula_source_row = None
        self._first_new_row = None
        self._last_new_row = None

# ============================================================================
# Workbook
# ============================================================================
    def open(self) -> None:
        """Öffnet die Arbeitsmappe."""

        self.excel = win32com.client.Dispatch("Excel.Application")

        # TODO Während der Entwicklung sichtbar
        self.excel.Visible = True

        self.workbook = self.excel.Workbooks.Open(
            str(self.workbook_path.resolve())
        )

        if self.workbook is None:
            return

        self.worksheet = self.workbook.Worksheets("Omron")

        self.table = self.worksheet.ListObjects(1)

    def close(self) -> None:
        """Schließt Excel."""

        if self.workbook:
            self.workbook.Close(SaveChanges=False)

        if self.excel:
            self.excel.Quit()

    def save(self) -> None:
        """Speichert die Arbeitsmappe."""
        self.workbook.Save()
        
# ============================================================================
# Backup
# ============================================================================
    def backup_workbook(self, backup_dir: Path) -> Path:
        """
        Erstellt eine Sicherung der Excel-Datei.
        """
    
        backup_dir.mkdir(parents=True, exist_ok=True)
    
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
        backup_file = backup_dir / f"Omron_{timestamp}.xlsx"
    
        shutil.copy2(self.workbook_path, backup_file)
    
        return backup_file         
 
# ============================================================================
# Measurement Import
# ============================================================================
    def append_measurement(
        self,
        measurement: Measurement,
        helper: HelperRow,
    ) -> None:        
        """
        Fügt eine Messung am Ende der Excel-Tabelle hinzu.
        """
       
        new_row = self.table.ListRows.Add()
    
        row = new_row.Range

        row.Cells(1, 1).Value = self._date_to_excel(measurement.date)
        row.Cells(1, 1).NumberFormatLocal = "TT.MM.JJJJ"        
 
        row.Cells(1, 2).Value = self._time_to_excel(measurement.time)                
        row.Cells(1, 2).NumberFormatLocal = "hh:mm"        
        
        row.Cells(1, 3).Value = measurement.systolic
#        print("SYS:", row.Cells(1, 3).Value)
        
        row.Cells(1, 4).Value = measurement.diastolic
#        print("DIA:", row.Cells(1, 4).Value)
        
        row.Cells(1, 5).Value = measurement.pulse
#        print("Puls:", row.Cells(1, 5).Value)    

        row.Cells(1, 6).Value = helper.marker
        row.Cells(1,10).Value = helper.daytime
        
        if helper.systolic_avg is not None:
            row.Cells(1, 7).Value = helper.systolic_avg
        
        if helper.diastolic_avg is not None:
            row.Cells(1, 8).Value = helper.diastolic_avg
        
        if helper.pulse_avg is not None:
            row.Cells(1, 9).Value = helper.pulse_avg        
        
#        self._copy_calculated_formulas(row)

            
    def get_existing_keys(self) -> set[tuple[str, int, int, int]]:
        """
        Liest alle vorhandenen Messungen aus der Excel-Tabelle
        und liefert (Datum, Uhrzeit)-Schlüssel zurück.
        """
        keys = set()
    
        data = self.table.DataBodyRange.Value
   
        if not data:
            return keys
    
        # Bei nur einer Datenzeile liefert Excel kein Tupel von Tupeln
        if not isinstance(data[0], tuple):
            data = (data,)
    
        for row in data:
    
            datum = row[0]
            uhrzeit = row[1]

            if datum is None:
#                print("-> Leere Zeile übersprungen")
                continue
              
            # Excel liefert Datum/Uhrzeit normalerweise als datetime
            if hasattr(datum, "date"):
                datum = datum.date()
        
            if isinstance(uhrzeit, float):
                uhrzeit = self._excel_time_to_time(uhrzeit)
            elif isinstance(uhrzeit, datetime):
                uhrzeit = uhrzeit.time()
        
            key = (
                f"{datum.isoformat()} {uhrzeit.strftime('%H:%M')}",
                int(row[2]),
                int(row[3]),
                int(row[4]),
            )
            keys.add(key)
            
#            print(f"Excel-Key: {key}")
#            print(f"Keys gelesen: {len(keys)}")
#            print("Verarbeitete Zeilen:", count)
#            print(f"Tabellenzeilen: {self.table.ListRows.Count}")    
        
        return keys

    def get_all_measurements(self) -> list[Measurement]:
        """
        Liest alle Messungen aus der Excel-Tabelle.
        """
        measurements = []
    
        data = self.table.DataBodyRange.Value
    
        if not data:
            return measurements
    
        if not isinstance(data[0], tuple):
            data = (data,)

        row = data[0]
        
        for row in data:

            datum = row[0]
            uhrzeit = row[1]
        
            if datum is None or uhrzeit is None:
                continue
        
            if isinstance(datum, datetime):
                datum = datum.date()
        
            if isinstance(uhrzeit, float):
                uhrzeit = self._excel_time_to_time(uhrzeit)
            elif isinstance(uhrzeit, datetime):
                uhrzeit = uhrzeit.time()
        
            measurement = Measurement(
                date=datum,
                time=uhrzeit,
                systolic=int(row[2]),
                diastolic=int(row[3]),
                pulse=int(row[4])
            )
        
            measurements.append(measurement)
    
        return measurements
 

# ============================================================================
# Helpers
# ============================================================================
    @staticmethod
    def _date_to_excel(d: date) -> int:
        excel_epoch = date(1899, 12, 30)
        return (d - excel_epoch).days

    @staticmethod
    def _time_to_excel(t: time) -> float:
        return (
            t.hour * 3600
            + t.minute * 60
            + t.second
        ) / 86400

    @staticmethod
    def _excel_time_to_time(value: float) -> time:
        sekunden = round(value * 24 * 60 * 60)
        return (datetime.min + timedelta(seconds=sekunden)).time()
            
# ============================================================================
# Table
# ============================================================================
    def sort_table(self) -> None:
        """Sortiert die Tabelle nach Datum und Uhrzeit."""
    
        sort = self.table.Sort
    
        sort.SortFields.Clear()
        
        sort.SortFields.Add(
            Key=self.table.ListColumns(1).Range,
            SortOn=0,
            Order=1,
            DataOption=0
        ) 

        sort.SortFields.Add(
            Key=self.table.ListColumns(2).Range,
            SortOn=0,
            Order=1,
            DataOption=0
        )

        sort.Header = 1          # xlYes
        sort.MatchCase = False
        sort.Orientation = 1     # xlTopToBottom
 
        sort.Apply()

# ============================================================================
# Charts
# ============================================================================

    def create_chart(self) -> None:

        ws = self.worksheet
    
        chart_name = "Blutdruck"
    
        # Vorhandenes Diagramm löschen
        for chart in ws.ChartObjects():
            if chart.Name == chart_name:
                chart.Delete()
                break
    
#        print("Vorhandenes Diagramm entfernt.")
    
        chart = ws.ChartObjects().Add(
            Left=750,
            Top=20,
            Width=650,
            Height=350
        )
    
        chart.Name = chart_name
    
#        chart.Chart.ChartType = constants.xlLine
        chart.Chart.ChartType = 65
        date_range = self.table.ListColumns(1).DataBodyRange
        sys_range = self.table.ListColumns(3).DataBodyRange

        series = chart.Chart.SeriesCollection().NewSeries()
        series.Name = "Systolisch"
        series.XValues = date_range
        series.Values = sys_range

        dia_range = self.table.ListColumns(4).DataBodyRange
        series2 = chart.Chart.SeriesCollection().NewSeries()
        series2.Name = "Diastolisch"
        series2.XValues = date_range
        series2.Values = dia_range
    
        chart.Chart.HasTitle = True
        chart.Chart.ChartTitle.Text = "Blutdruckverlauf"
    
#        print("Diagramm erzeugt.")
 


    

        
   