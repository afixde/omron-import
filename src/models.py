from dataclasses import dataclass
from datetime import date, time, datetime

@dataclass(frozen=True, slots=True)
class Measurement:
    """Eine einzelne Blutdruckmessung."""

    date: date
    time: time
    systolic: int
    diastolic: int
    pulse: int

    @property
    def key(self) -> tuple[date, time]:
        """Eindeutiger Schlüssel für die Duplikaterkennung."""
        return self.date, self.time

    @property
    def timestamp(self) -> datetime:
        """Datum und Uhrzeit als datetime."""
        return datetime.combine(self.date, self.time)