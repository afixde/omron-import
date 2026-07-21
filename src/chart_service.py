import win32com.client

from win32com.client import constants
from config import TABLE_COLUMNS

class ChartService:

    CHART_NAME = "Blutdruck"

    def __init__(self, worksheet, table):
        self.ws = worksheet
        self.table = table
        self.chart = None
#        excel = gencache.EnsureDispatch("Excel.Application")
        
    def update_chart(self):
        """
        Erstellt oder aktualisiert das Blutdruckdiagramm.
        """
        self.chart = self._get_chart()
        
        self._clear_series()
        self._add_sys_series()
        self._add_dia_series()
        self._add_pulse_series()    
        
        self._format_title()
        self._format_legend()
        self._format_axes()
        self._format_series()    

    def _get_chart(self):
        """
        Liefert das Diagramm. Existiert es nicht, wird es erzeugt.
        """
        for chart_object in self.ws.ChartObjects():
            if chart_object.Name == self.CHART_NAME:
                return chart_object.Chart
    
        self._create_chart()
    
        for chart_object in self.ws.ChartObjects():
            if chart_object.Name == self.CHART_NAME:
                return chart_object.Chart
    
        raise RuntimeError(
            f"Diagramm '{self.CHART_NAME}' konnte nicht erzeugt werden."
        )
    
    def _get_series(self, name: str):
        """
        Liefert eine Datenreihe anhand ihres Namens.
        """
        for series in self.chart.SeriesCollection():
            if series.Name == name:
                return series
    
        return None
    
    def _create_chart(self):
        """Erzeugt ein leeres Diagramm."""
        chart_object = self.ws.ChartObjects().Add(
            Left=420,
            Top=20,
            Width=700,
            Height=350
        )

        chart_object.Name = self.CHART_NAME

        chart = chart_object.Chart
        chart.ChartType = constants.xlLine
##        chart.ChartType = 4
        chart.SetSourceData(self.table.Range)

        chart.HasTitle = True
        chart.ChartTitle.Text = "Blutdruckverlauf"

        chart.HasLegend = True

    def _get_data_range(self, column_key: str):
        """
        Liefert den Datenbereich einer Tabellenspalte.
        """
        column_name = TABLE_COLUMNS[column_key]
        return self.table.ListColumns(column_name).DataBodyRange

    def _clear_series(self):
        """
        Entfernt alle vorhandenen Datenreihen.
        """
        while self.chart.SeriesCollection().Count > 0:
            self.chart.SeriesCollection(1).Delete()

    def _add_sys_series(self):
        series = self.chart.SeriesCollection().NewSeries()
        series.Name = TABLE_COLUMNS["sys"]
        series.Values = self._get_data_range("sys")
        series.XValues = self._get_data_range("date")
        series.AxisGroup = constants.xlPrimary
        print("SYS-Serie angelegt:", series.Name)   

    def _add_dia_series(self):
        """
        Fügt die diastolische Messreihe hinzu.
        """
        series = self.chart.SeriesCollection().NewSeries()
        series.Name = TABLE_COLUMNS["dia"]
        series.Values = self._get_data_range("dia")
        series.XValues = self._get_data_range("date")
        series.AxisGroup = constants.xlPrimary
                 
    def _add_pulse_series(self):
        """
        Fügt die Puls-Messreihe hinzu.
        """
        series = self.chart.SeriesCollection().NewSeries()
        series.Name = TABLE_COLUMNS["pulse"]
        series.Values = self._get_data_range("pulse")
        series.XValues = self._get_data_range("date")
        series.AxisGroup = constants.xlSecondary
    
    def _format_title(self):
        self.chart.HasTitle = True
        self.chart.ChartTitle.Text = "Blutdruckverlauf"
        self.chart.ChartTitle.Font.Bold = True
        self.chart.ChartTitle.Font.Size = 14

    def _format_legend(self):
        self.chart.HasLegend = True
        self.chart.Legend.Position = (
            constants.xlLegendPositionTop
##            -4160
        )
        self.chart.Legend.Font.Size = 10    

    def _format_axes(self):
    
        value_axis = self.chart.Axes(constants.xlValue)
##        value_axis = self.chart.Axes(2)
    
        # Achsentitel
        value_axis.HasTitle = True
        value_axis.AxisTitle.Text = "mmHg"
    
        # Feste Skalierung
        value_axis.MinimumScale = 60
        value_axis.MaximumScale = 160
        value_axis.MajorUnit = 10
    
        # Gitternetzlinien
        value_axis.HasMajorGridlines = True
    
        # Schrift
        value_axis.TickLabels.Font.Size = 10    

        category_axis = self.chart.Axes(constants.xlCategory)
##        category_axis = self.chart.Axes(1)
    
        category_axis.TickLabels.Font.Size = 9
    
        # Excel entscheidet automatisch,
        # welche Datumsbeschriftungen angezeigt werden.
        category_axis.TickLabelSpacing = 5
        secondary_axis = self.chart.Axes(
            constants.xlValue,
            constants.xlSecondary
        )
        
        secondary_axis.HasTitle = True
        secondary_axis.AxisTitle.Text = "bpm"
        
        secondary_axis.MinimumScale = 50
        secondary_axis.MaximumScale = 100
        secondary_axis.MajorUnit = 10
        
    def _format_series(self):
    
        for series in self.chart.SeriesCollection():

            series.Format.Line.Weight = 2
    
            # Marker ausschalten
            series.MarkerStyle = constants.xlMarkerStyleNone

#        for series in self.chart.SeriesCollection():
#            print(series.Name)
#