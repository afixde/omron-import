from pathlib import Path

# Projektverzeichnis
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Datenverzeichnisse
DATA_DIR = PROJECT_ROOT / "data"

CSV_DIR = DATA_DIR / "csv"

EXCEL_DIR = DATA_DIR / "excel"

BACKUP_DIR = DATA_DIR / "backup"

# Dateien
EXCEL_FILE = EXCEL_DIR / "Omron.xlsx"

BACKUP_DIR.mkdir(parents=True, exist_ok=True)
