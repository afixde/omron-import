from models import Measurement, HelperRow


class HelperColumnsService:
    """Erzeugt und berechnet die Hilfsspalten für den Excel-Export."""

    def build(
        self,
        measurements: list[Measurement]
    ) -> list[HelperRow]:
        """Erzeugt alle Hilfsspalten."""

        helpers = self._create_helper_rows(measurements)
        
        self._calculate_averages(measurements, helpers)
        
        return helpers

    def _create_helper_rows(
        self,
        measurements: list[Measurement]
    ) -> list[HelperRow]:
        """Erzeugt Marker und Tageszeit."""

        helpers: list[HelperRow] = []

        for measurement in measurements:
            marker = self._build_marker(measurement)
            daytime = self._get_daytime(measurement)

            helpers.append(
                HelperRow(
                    marker=marker,
                    daytime=daytime
                )
            )

        return helpers

    @staticmethod
    def _build_marker(measurement: Measurement) -> str:
        """Erzeugt den Marker YYYYMMDDHH."""
        return (
            measurement.date.strftime("%Y%m%d")
            + f"{measurement.time.hour:02d}"
        )

    @staticmethod
    def _get_daytime(measurement: Measurement) -> str:
        """Bestimmt die Tageszeit."""
        return "Morgen" if measurement.time.hour < 12 else "Abend"
    
    def _calculate_averages(
        self,
        measurements: list[Measurement],
        helpers: list[HelperRow],
    ) -> None:
        if not measurements:
            return
    
        current_marker = helpers[0].marker
        last_index = 0
    
        sum_sys = measurements[0].systolic
        sum_dia = measurements[0].diastolic
        sum_pulse = measurements[0].pulse
        count = 1
    
        for i in range(1, len(measurements)):
    
            if helpers[i].marker == current_marker:
                sum_sys += measurements[i].systolic
                sum_dia += measurements[i].diastolic
                sum_pulse += measurements[i].pulse
                count += 1
                last_index = i
    
            else:
                self._finish_group(
                    helpers,
                    last_index,
                    sum_sys,
                    sum_dia,
                    sum_pulse,
                    count,
                )
    
                current_marker = helpers[i].marker
                last_index = i
    
                sum_sys = measurements[i].systolic
                sum_dia = measurements[i].diastolic
                sum_pulse = measurements[i].pulse
                count = 1
    
        self._finish_group(
            helpers,
            last_index,
            sum_sys,
            sum_dia,
            sum_pulse,
            count,
        )

    def _finish_group(
        self,
        helpers: list[HelperRow],
        last_index: int,
        sum_sys: int,
        sum_dia: int,
        sum_pulse: int,
        count: int,
    ) -> None:
    
        helpers[last_index].systolic_avg = round(sum_sys / count, 1)
        helpers[last_index].diastolic_avg = round(sum_dia / count, 1)
        helpers[last_index].pulse_avg = round(sum_pulse / count, 1)
