# Changelog

## [0.8.0] - 2026-07-22

### Added
- Unit tests for CsvReader
- Unit tests for HelperColumnsService
- Unit tests for ImportService
- Test data (`test/data/sample.csv`)
- Coverage reporting using pytest-cov

### Changed
- ImportService supports dependency injection
- Improved project test structure
- Improved maintainability and testability

### Quality
- 22 automated tests
- ImportService coverage increased to 76%
- HelperColumnsService fully covered by unit tests

## [0.7.0] - 2026-07-21

### Added
- Modularer ChartService zur automatischen Erstellung und Aktualisierung des Blutdruckdiagramms
- Unterstützung einer Sekundärachse für die Pulswerte
- Zentrale Tabellenkonfiguration über `TABLE_COLUMNS`

### Changed
- ChartService vollständig refaktoriert
- Diagrammerstellung in Aufbau (`_build_chart`) und Formatierung (`_format_chart`) getrennt
- Mehrfache Serien durch generische `_add_series()`-Methode ersetzt
- Codebereinigung und Vereinheitlichung der Typannotationen und Docstrings
- Hilfsspalten werden jetzt vollständig in Python berechnet
- Excel-Datum und -Uhrzeit werden als Excel-Serienwerte gespeichert
- Konsolenausgabe bereinigt und Debug-Ausgaben entfernt

### Fixed
- Fehler bei der Zuordnung der Primär- und Sekundärachsen
- Verbesserung der Diagrammaktualisierung nach dem Import
- Entfernen redundanter und veralteter Implementierungen
- Wiederholter Import derselben CSV-Datei erzeugt keine Dubletten mehr
- Dublettenerkennung anhand von Datum, Uhrzeit und Messwerten stabilisiert

## [0.6.0] - 2026-07-20

### Added
- Backup der Excel-Arbeitsmappe vor jedem Import
- Automatische Archivierung importierter CSV-Dateien
- Statistikservice mit Mittelwert, Median sowie Min-/Max-Werten
- Verbesserte Dublettenerkennung
- Strukturierter ImportService

### Changed
- Projektarchitektur weiter in Service-Klassen aufgeteilt
- Logging erweitert
- Sortierung der Excel-Tabelle verbessert

### Fixed
- Fehler bei der Datums- und Zeitformatierung in Excel
- COM-Probleme beim Schreiben neuer Messwerte
- Erkennung vorhandener Messungen verbessert

# Version 0.5.0

## Neu

- Automatische Backup-Erstellung der Excel-Datei
- Archivierung importierter CSV-Dateien
- Erweiterte Statistik
  - Durchschnitt
  - Median
  - Minimum
  - Maximum
- Excel-Diagramm (Grundgerüst)
- Verbesserte Protokollierung

## Verbessert

- Stabilerer CSV-Import
- Zuverlässige Dublettenerkennung
- Korrekte Datum-/Zeitbehandlung
- Automatische Sortierung der Excel-Tabelle

## Behoben

- Fehler bei Excel-Datumswerten
- Fehler bei Zeitwerten
- COM-Probleme beim Schreiben in Excel
- Sortierfehler

## v0.4.2

- Logging eingeführt
- Backup erstellt
- CSV-Archivierung
- Versionierung eingebaut

## v0.4.1

- Excel-Import stabilisiert

## v0.4.0

- erster funktionsfähiger Import