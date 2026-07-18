from models import Measurement
from statistics import median

class StatisticsService:
    """
    Ermittelt statistische Kennzahlen
    aus einer Liste von Blutdruckmessungen.
    """

    def summarize(self, measurements: list[Measurement]) -> dict:

        if not measurements:
            return {}

        sys_values = [m.systolic for m in measurements]
        dia_values = [m.diastolic for m in measurements]
        pulse_values = [m.pulse for m in measurements]

        return {
            "count": len(measurements),
        
            "avg_sys": round(sum(sys_values) / len(sys_values), 1),
            "avg_dia": round(sum(dia_values) / len(dia_values), 1),
            "avg_pulse": round(sum(pulse_values) / len(pulse_values), 1),
        
            "median_sys": median(sys_values),
            "median_dia": median(dia_values),
            "median_pulse": median(pulse_values),
        
            "min_sys": min(sys_values),
            "max_sys": max(sys_values),
        
            "min_dia": min(dia_values),
            "max_dia": max(dia_values),
        
            "min_pulse": min(pulse_values),
            "max_pulse": max(pulse_values),
        }        

    def report(self, measurements):
        summary = self.summarize(measurements)
    
        # <-- Hier den bisherigen Statistik-Ausgabecode aus main.py
        #     unverändert hineinkopieren.
    
        return summary      
 