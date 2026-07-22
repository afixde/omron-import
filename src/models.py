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
    def key(self) -> str:
        """Eindeutiger Schlüssel einer Messung."""
    
        return (
            f"{self.date.isoformat()} "
            f"{self.time.strftime('%H:%M')}"
        )
    
    @property
    def timestamp(self) -> datetime:
        """Datum und Uhrzeit als datetime."""
        return datetime.combine(self.date, self.time)
    
@dataclass(slots=True)
class HelperRow:
    marker: str
    daytime: str
    systolic_avg: float | None = None
    diastolic_avg: float | None = None
    pulse_avg: float | None = None
    
