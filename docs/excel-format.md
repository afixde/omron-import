# Excel-Dateiformat

## Zweck

Die Excel-Arbeitsmappe dient als zentrale Datenbasis für alle importierten Blutdruckmessungen. Der Import aktualisiert ausschließlich die Tabelle **Omron** und das zugehörige Diagramm.

---

# Arbeitsblatt

Name:

```
Omron
```

Enthält:

- Messwerte
- Hilfsspalten
- Diagramm

---

# Tabelle

Die Messwerte befinden sich in einer Excel-Tabelle (ListObject).

Die erste Zeile enthält die Spaltenüberschriften.

Neue Messungen werden immer am Tabellenende eingefügt und anschließend chronologisch sortiert.

---

# Spalten

| Nr. | Spalte | Beschreibung | Quelle |
|----:|---------|--------------|---------|
| 1 | Datum | Messdatum | OMRON CSV |
| 2 | Zeit | Messzeit | OMRON CSV |
| 3 | SYS | Systolischer Blutdruck | OMRON CSV |
| 4 | DIA | Diastolischer Blutdruck | OMRON CSV |
| 5 | Puls | Puls | OMRON CSV |
| 6 | Marker | Benutzerdefiniert | Python |
| 7 | SYS Ø | Durchschnitt | Python |
| 8 | DIA Ø | Durchschnitt | Python |
| 9 | Puls Ø | Durchschnitt | Python |
|10 | Tageszeit | Morgen / Mittag / Abend | Python |

---

# Datums- und Zeitformat

## Datum

Wird als Excel-Serienwert gespeichert.

Anzeigeformat:

```
TT.MM.JJJJ
```

## Uhrzeit

Wird als Excel-Zeitwert gespeichert.

Anzeigeformat:

```
hh:mm
```

Dadurch bleiben Sortierung und Darstellung in Excel zuverlässig.

---

# Dublettenerkennung

Eine Messung gilt als vorhanden, wenn folgende Werte übereinstimmen:

- Datum
- Uhrzeit
- SYS
- DIA
- Puls

---

# Sortierung

Nach jedem erfolgreichen Import wird die Tabelle sortiert nach

1. Datum
2. Uhrzeit

aufsteigend.

---

# Diagramm

Nach einem erfolgreichen Import wird das Diagramm aktualisiert.

Dargestellt werden:

- SYS
- DIA
- Puls

Die X-Achse verwendet das Messdatum.

---

# Backup

Vor jedem Import wird eine Sicherungskopie der Arbeitsmappe erstellt.

---

# CSV-Archiv

Nach erfolgreichem Import wird die CSV-Datei archiviert.