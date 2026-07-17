from win32com.client import constants


class ChartService:

    def __init__(self, worksheet, table):
        self.worksheet = worksheet
        self.table = table

    def update_chart(self):
        """
        Erstellt oder aktualisiert das Blutdruckdiagramm.
        """
        pass