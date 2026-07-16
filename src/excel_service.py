from pathlib import Path
from models import Measurement
from datetime import date, time, datetime, timedelta

import win32com.client
import shutil

class ExcelService:
    """Zugriff auf die Omron-Exceldatei."""

    def __init__(self, workbook_path: Path):
        self.workbook_path = workbook_path

        self.excel = None
        self.workbook = None
        self.sheet = None
        self.table = None

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

        self.sheet = self.workbook.Worksheets("Omron")

        self.table = self.sheet.ListObjects(1)

#        print("Excel geöffnet.")
#        print(f"Arbeitsblatt : {self.sheet.Name}")
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
    
        row.Cells(1, 1).Value = self._date_to_excel(measurement.date)
        row.Cells(1, 1).NumberFormatLocal = "TT.MM.JJJJ"
        
        row.Cells(1, 2).Value = (
            measurement.time.hour / 24
            + measurement.time.minute / 1440
            + measurement.time.second / 86400
        )
        
#        print("Zeit Excel :", row.Cells(1, 2).Value)
        
        row.Cells(1, 3).Value = measurement.systolic
#        print("SYS:", row.Cells(1, 3).Value)
        
        row.Cells(1, 4).Value = measurement.diastolic
#        print("DIA:", row.Cells(1, 4).Value)
        
        row.Cells(1, 5).Value = measurement.pulse
#        print("Puls:", row.Cells(1, 5).Value)    
            
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
    