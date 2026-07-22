# Architektur

## Ziel

Das Projekt automatisiert den Import von Blutdruckmessungen aus der OMRON Connect App in eine bestehende Excel-Arbeitsmappe. Neben dem eigentlichen Import werden Hilfsspalten berechnet, Statistiken aktualisiert, Diagramme gepflegt und die Quelldateien archiviert.

---

## Architekturübersicht

                    main.py
                       │
                       ▼
                ImportService
                       │
 ┌───────────────┬─────┴─────────────┬──────────────┐
 ▼               ▼                   ▼              ▼
OmronCsvReader ExcelService StatisticsService ChartService
                       │
                       ▼
             HelperColumnsService

Der `ImportService` koordiniert den kompletten Ablauf und delegiert die einzelnen Aufgaben an spezialisierte Service-Klassen.

---

## Komponenten

### main.py

Programmeinstieg.

Erzeugt den `ImportService` und startet den Import.

---

### ImportService

Zentrale Steuerklasse.

Verantwortlich für

- CSV einlesen
- Excel öffnen
- Backup
- Dublettenerkennung
- Import
- Sortierung
- Diagramm
- Speichern
- Archivierung

---

### OmronCsvReader

- sucht die neueste CSV
- liest Messungen ein
- liefert `Measurement`-Objekte

---

### ExcelService

Kapselt sämtliche COM-Zugriffe auf Excel.

Verantwortlich für

- Arbeitsmappe öffnen/schließen
- Tabelle lesen
- Messungen einfügen
- Sortieren
- Speichern
- Backup

---

### HelperColumnsService

Berechnet

- Marker
- Tageszeit
- Durchschnittswerte

für neue Messungen.

---

### StatisticsService

Berechnet

- Mittelwert
- Median
- Minimum
- Maximum

für SYS, DIA und Puls.

---

### ChartService

Aktualisiert das Blutdruckdiagramm nach einem erfolgreichen Import.

---

## Datenfluss

CSV
 │
 ▼
OmronCsvReader
 │
 ▼
Measurement
 │
 ▼
ImportService
 │
 ├── Dublettenerkennung
 ├── HelperColumnsService
 ├── ExcelService
 ├── StatisticsService
 └── ChartService

---

## Designprinzipien

- Eine Klasse übernimmt genau eine Hauptaufgabe.
- Business-Logik ist von Excel-spezifischem Code getrennt.
- COM-Zugriffe sind im `ExcelService` gekapselt.
- Der `ImportService` enthält keine Berechnungen, sondern steuert ausschließlich den Ablauf.