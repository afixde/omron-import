from pathlib import Path

# Projektverzeichnis
PROJECT_ROOT = Path(__file__).resolve().parent.parent
# Datenverzeichnisse
DATA_DIR = PROJECT_ROOT / "data"
CSV_DIR = DATA_DIR / "csv"
EXCEL_DIR = DATA_DIR / "excel"
BACKUP_DIR = DATA_DIR / "backup"
BACKUP_DIR.mkdir(parents=True, exist_ok=True)
ARCHIVE_DIR = PROJECT_ROOT / "data" / "archive"
# Dateien
EXCEL_FILE = EXCEL_DIR / "Omron.xlsx"

# Excel
TABLE_NAME = "Omron"

# Diagramm
CHART_NAME = "Blutdruck"


# Tabellenstruktur Excel
TABLE_COLUMNS = {
    "date": "Datum",
    "time": "Zeit",
    "sys": "Systolisch (mmHg)",
    "dia": "Diastolisch (mmHg)",
    "pulse": "Puls (bpm)",
    "marker": "Marker",
    "sys_avg": "Systolisch",
    "dia_avg": "Diastolisch",
    "pulse_avg": "Puls",
    "daytime": "Tageszeit",
    "cuff": "Manschettensitzkontrolle",
    "position": "Positionierungsanzeige",
    "mode": "Messmodus",
    "error": "Fehler-Nr.",
    "afib": "Mögliches AFib",
    "device": "Gerät",
    "notes": "Notizen",
}
