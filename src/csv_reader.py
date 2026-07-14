from pathlib import Path

from models import Measurement


class OmronCsvReader:
    """Liest eine OMRON-CSV-Datei."""

    def read(self, csv_file: Path) -> list[Measurement]:
        raise NotImplementedError()