from pathlib import Path
from models import Measurement
from datetime import date, time, datetime, timedelta
from win32com.client import constants

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
        
    @staticmethod
    def _date_to_excel(d: date) -> int:
        excel_epoch = date(1899, 12, 30)
        return (d - excel_epoch).days

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

#        print("Excel geöffnet.")
#        print(f"Arbeitsblatt : {self.worksheet.Name}")
#        print(f"Tabelle      : {self.table.Name}")

    def close(self) -> None:
        """Schließt Excel."""

        if self.workbook:
            self.workbook.Close(SaveChanges=False)

        if self.excel:
            self.excel.Quit()

    @staticmethod
    def _excel_time_to_time(value: float) -> time:
        sekunden = round(value * 24 * 60 * 60)
        return (datetime.min + timedelta(seconds=sekunden)).time()

    def _copy_calculated_formulas(self, row) -> None:
        if self._formula_source_row is None:
            return
        source = self.worksheet.Rows(self._formula_source_row)
        for col in (6, 7, 8, 9, 10):
            row.Cells(1, col).FormulaLocal = source.Cells(col).FormulaLocal
        self._last_new_row = row.Row

    def fill_calculated_columns(self) -> None:
        """
        Überträgt die Formelspalten F:J
        mittels Excel Copy/PasteSpecial.
        """
        if self._formula_source_row is None:
            return
        if self._last_new_row is None:
            return
        source = self.worksheet.Range(
            f"F{self._formula_source_row}:J{self._formula_source_row}"
        )
        target = self.worksheet.Range(
            f"F{self._first_new_row}:J{self._last_new_row}"
        )
        source.Copy()
        target.PasteSpecial(constants.xlPasteFormulas)
        self.excel.CutCopyMode = False


    def prepare_formula_copy(self):
    
        self._formula_source_row = (
            self.table.DataBodyRange.Rows(
                self.table.DataBodyRange.Rows.Count
            ).Row
        )
    
        self._first_new_row = self._formula_source_row + 1
            
    def get_existing_keys(self) -> set[str]:
        """
        Liest alle vorhandenen Messungen aus der Excel-Tabelle
        und liefert (Datum, Uhrzeit)-Schlüssel zurück.
        """
    
        keys = set()
    
        data = self.table.DataBodyRange.Value

#        print(type(data))
#        print(len(data))
        
#        print()
#        print("Erste Zeile:")
#        print(data[0])
#        print(type(data[0]))
    
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
                f"{datum.isoformat()} "
                f"{uhrzeit.strftime('%H:%M')}"
            )

            
#            print(f"{key}")
            
            keys.add(key)
            
#            print(f"Keys gelesen: {len(keys)}")
#            print(f"Tabellenzeilen: {self.table.ListRows.Count}")    
        
        return keys

    def append_measurement(self, measurement: Measurement) -> None:
        """
        Fügt eine Messung am Ende der Excel-Tabelle hinzu.
        """
       
        new_row = self.table.ListRows.Add()
    
        row = new_row.Range

#        # ===== Mini-Test =====
#        print()
#        print("Neue Tabellenzeile")
#        print("------------------")
#        print("Excel-Zeile        :", row.Row)
#        print("Tabellenzeilen     :", self.table.ListRows.Count)
#        print("Tabellenbereich    :", self.table.DataBodyRange.Address)
#        print("Zeilenadresse      :", row.Address)
#        print("===================")
#
#        print("Marker neu        :", row.Cells(1, 6).FormulaLocal)
#        print("SYS Avg neu       :", row.Cells(1, 7).FormulaLocal)
#        print("DIA Avg neu       :", row.Cells(1, 8).FormulaLocal)
#        print("Puls Avg neu      :", row.Cells(1, 9).FormulaLocal)
#        print("Tageszeit neu     :", row.Cells(1,10).FormulaLocal)
#
#        prev = self.worksheet.Rows(row.Row - 1)
#        
#        print()
#        print("Vorherige Zeile")
#        print("------------------")
#        print("Marker alt        :", prev.Cells(6).FormulaLocal)
#        print("SYS Avg alt       :", prev.Cells(7).FormulaLocal)
#        print("DIA Avg alt       :", prev.Cells(8).FormulaLocal)
#        print("Puls Avg alt      :", prev.Cells(9).FormulaLocal)
#        print("Tageszeit alt     :", prev.Cells(10).FormulaLocal)
#    
        row.Cells(1,1).Value = datetime.combine(
            measurement.date,
            time.min
        )
        
        row.Cells(1,1).NumberFormatLocal = "TT.MM.JJJJ"
        
        row.Cells(1,2).Value = datetime.combine(
            date.today(),
            measurement.time
        )
        
        row.Cells(1,2).NumberFormatLocal = "hh:mm"
        
#        print("Zeit Excel :", row.Cells(1, 2).Value)
        
        row.Cells(1, 3).Value = measurement.systolic
#        print("SYS:", row.Cells(1, 3).Value)
        
        row.Cells(1, 4).Value = measurement.diastolic
#        print("DIA:", row.Cells(1, 4).Value)
        
        row.Cells(1, 5).Value = measurement.pulse
#        print("Puls:", row.Cells(1, 5).Value)    
        self._copy_calculated_formulas(row)

            
    def save(self) -> None:
        """Speichert die Arbeitsmappe."""
        self.workbook.Save()
        
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

#        sort.SetRange(self.table.Range)
    
        sort.Header = 1          # xlYes
        sort.MatchCase = False
        sort.Orientation = 1     # xlTopToBottom
 
#        print(self.table.Name)
#        print(self.table.Range.Address)
#        print(self.table.DataBodyRange.Address)  
         
        sort.Apply()

    def backup_workbook(self, backup_dir: Path) -> Path:
        """
        Erstellt eine Sicherung der Excel-Datei.
        """
    
        backup_dir.mkdir(parents=True, exist_ok=True)
    
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
        backup_file = backup_dir / f"Omron_{timestamp}.xlsx"
    
        shutil.copy2(self.workbook_path, backup_file)
    
        return backup_file         
    
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
 
    def create_chart(self) -> None:

        ws = self.worksheet
    
        chart_name = "Blutdruck"
    
        # Vorhandenes Diagramm löschen
        for chart in ws.ChartObjects():
            if chart.Name == chart_name:
                chart.Delete()
                break
    
        print("Vorhandenes Diagramm entfernt.")
    
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
    
        print("Diagramm erzeugt.")
    