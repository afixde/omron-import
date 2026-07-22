from models import Measurement
from statistics import median

class StatisticsService:
    """
    Ermittelt statistische Kennzahlen
    aus einer Liste von Blutdruckmessungen.
    """
    def _print_header(self, summary):
    
        print()
        print("Statistik")
        print("-" * 30)
        print()
        print(f"Messungen : {summary['count']}")
        print()

    def _print_blood_pressure(self, summary):
    
        print(f"SYS Ø      : {summary['avg_sys']} mmHg")
        print(f"SYS Median : {summary['median_sys']} mmHg")
        print(f"SYS Min    : {summary['min_sys']} mmHg")
        print(f"SYS Max    : {summary['max_sys']} mmHg")
    
        print()
    
        print(f"DIA Ø      : {summary['avg_dia']} mmHg")
        print(f"DIA Median : {summary['median_dia']} mmHg")
        print(f"DIA Min    : {summary['min_dia']} mmHg")
        print(f"DIA Max    : {summary['max_dia']} mmHg")
    
        print()

    def _print_pulse(self, summary):
    
        print(f"Puls Ø     : {summary['avg_pulse']} bpm")
        print(f"Puls Median: {summary['median_pulse']} bpm")
        print(f"Puls Min   : {summary['min_pulse']} bpm")
        print(f"Puls Max   : {summary['max_pulse']} bpm")
    
        print()

    def _print_measurement_info(self, measurements):
    
        print(f"Excel enthält {len(measurements)} Messungen")
    
#        if measurements:
#            print()
#            print("Erste Messung aus Excel:")
#            print(measurements[0])
#        
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
        self._print_header(summary)
        self._print_blood_pressure(summary)
        self._print_pulse(summary)
        self._print_measurement_info(measurements)
        
        return summary      
 