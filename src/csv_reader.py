from pathlib import Path

from datetime import date, time

from models import Measurement

import csv

GERMAN_MONTHS = {
    "Januar": 1,
    "Februar": 2,
    "März": 3,
    "April": 4,
    "Mai": 5,
    "Juni": 6,
    "Juli": 7,
    "August": 8,
    "September": 9,
    "Oktober": 10,
    "November": 11,
    "Dezember": 12,
}


def parse_date(text: str) -> date:
    """Parst ein OMRON-Datum, z.B. '13 Juli 2026'."""

    text = text.strip()
    parts = text.split()

    if len(parts) != 3:
        raise ValueError(f"Ungültiges Datum: {text}")

    day = int(parts[0])
    month = GERMAN_MONTHS[parts[1]]
    year = int(parts[2])

    return date(year, month, day)


def parse_time(text: str) -> time:
    """Parst eine Uhrzeit im Format HH:MM."""

    hour, minute = text.strip().split(":")
    return time(int(hour), int(minute))


class OmronCsvReader:
    """Liest OMRON-CSV-Dateien."""

    def find_latest_csv(self, directory: Path) -> Path:
        """Sucht die neueste OMRON-CSV im Verzeichnis."""

        files = sorted(
            directory.glob("Ihr angeforderter OMRON-Bericht*.csv"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        
        for file in files:
            print (file)

        if not files:
            raise FileNotFoundError(
                f"Keine OMRON-CSV in {directory} gefunden."
            )

        return files[0]

    def read_header(self, csv_file: Path) -> tuple[list[str], int]:
        """Liest die Kopfzeile und zählt die Datensätze."""

        with csv_file.open(
            "r",
            encoding="utf-8-sig",
            newline=""
        ) as f:

            sample = f.read(2048)
            f.seek(0)

            dialect = csv.Sniffer().sniff(sample)

            reader = csv.reader(f, dialect)

            header = next(reader)
            rows = sum(1 for _ in reader)

        return header, rows

    def _parse_row(self, row: dict[str, str]) -> Measurement:
        """Wandelt eine CSV-Zeile in ein Measurement um."""

        return Measurement(
            date=parse_date(row["Datum"]),
            time=parse_time(row["Zeit"]),
            systolic=int(row["Systolisch (mmHg)"]),
            diastolic=int(row["Diastolisch (mmHg)"]),
            pulse=int(row["Puls (bpm)"]),
        )

    def read(self, csv_file: Path) -> list[Measurement]:
        """Liest eine komplette OMRON-CSV-Datei."""

        with csv_file.open(
            "r",
            encoding="utf-8-sig",
            newline=""
        ) as f:

            sample = f.read(2048)
            f.seek(0)

            dialect = csv.Sniffer().sniff(sample)

            reader = csv.DictReader(f, dialect=dialect)

            measurements = []

            for row in reader:
                measurements.append(self._parse_row(row))

        return measurements