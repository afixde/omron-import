# Omron Import

Importiert Blutdruckmessungen aus der OMRON Connect App in eine Microsoft-Excel-Arbeitsmappe. Neue Messungen werden automatisch erkannt, importiert und ausgewertet.

## Funktionen

- CSV automatisch erkennen
- Dublettenerkennung
- Excel automatisch aktualisieren
- automatische Sortierung
- Backup der Excel-Datei
- Archivierung importierter CSV-Dateien
- Logging

## Voraussetzungen

- Windows 11
- Python 3.14
- Microsoft Excel
- pywin32

## Installation

```bash
git clone ...
cd omron-import
pip install -r requirements.txt
```
## Verwendung

### Vorbereitung

1. Exportiere die gewünschten Blutdruckmessungen aus der OMRON Connect App als CSV-Datei.
2. Kopiere die CSV-Datei in das konfigurierte Importverzeichnis.
3. Stelle sicher, dass die Excel-Arbeitsmappe nicht geöffnet ist.

### Import starten

Der Import wird über die Batch-Datei gestartet:

```bat
start_import.bat
```

Während des Imports werden automatisch folgende Schritte ausgeführt:

1. Die neueste OMRON-CSV-Datei wird gesucht.
2. Von der Excel-Arbeitsmappe wird ein Backup erstellt.
3. Vorhandene Messungen werden eingelesen.
4. Neue Messungen werden erkannt (Dublettenerkennung).
5. Neue Messungen werden in die Excel-Tabelle eingefügt.
6. Hilfsspalten werden berechnet.
7. Die Tabelle wird chronologisch sortiert.
8. Das Diagramm wird aktualisiert.
9. Die Excel-Datei wird gespeichert.
10. Die importierte CSV-Datei wird archiviert.

### Dublettenerkennung

Bereits importierte Messungen werden automatisch erkannt und nicht erneut übernommen. Dadurch kann dieselbe CSV-Datei beliebig oft importiert werden, ohne doppelte Einträge zu erzeugen.

### Backup

Vor jedem Import wird automatisch eine Sicherungskopie der Excel-Dateits erstellt.

### Archivierung

Nach einem erfolgreichen Import wird die verarbeitete CSV-Datei in das Archivverzeichnis verschoben. Dadurch bleibt das Importverzeichnis übersichtlich und bereits verarbeitete Dateien werden nicht erneut berücksichtigt.

## Tests

Alle Tests ausführen:

```bash
python -m pytest
```

Coverage anzeigen:

```bash
python -m pytest --cov=src --cov-report=term-missing
```
