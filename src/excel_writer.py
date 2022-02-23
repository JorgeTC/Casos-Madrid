import openpyxl


class Excel_writer():
    def __init__(self):
        self.excel_name = "Casos Comunidad de Madrid.xlsx"
        print("Abriendo Excel…")
        self.workbook = openpyxl.load_workbook(self.excel_name)

        # Guardo la hoja donde voy a escribir los datos.
        self.sheet = self.workbook[self.workbook.sheetnames[0]]
        # Las celdas se numeran a partir de 1.
        # No puedo escribir en la primera fila, es para los encabezados.
        self.index_line = 2

    def __del__(self):
        print("Guardando Excel…")
        # Guardo el xlsx.
        # Le doy el mismo nombre que tenía.
        self.workbook.save(self.excel_name)
        # Cierro el xlsx
        self.workbook.close()

    def write_data(self, data):
        # Data debe ser una lista de pares.
        # El primer elemento debe ser una fecha en formato fecha.
        # El segundo elemento, la cantidad de positivos en formato entero.
        print("Escribiendo Excel…")
        for value in data:
            # Fecha escrita con formato de fecha.
            self.__write_date(value)
            # Positivos escritos con formato de entero.
            self.__write_cases(value)

            # Avanzo a la siguiente linea.
            self.index_line = self.index_line + 1

    def __write_date(self, value):
        # Dado el par value, escribo la fecha (1º elemento) en la casilla adecuada de excel.

        cell = self.sheet.cell(row=self.index_line, column=1)
        # Fecha escrita con formato de fecha.
        cell.value = value[0]
        cell.number_format = 'mm-dd-yy'

    def __write_cases(self, value):
        # Dado el par value, escribo los positivos (2º elemento) en la casilla adecuada de excel.

        cell = self.sheet.cell(row=self.index_line, column=2)
        # Positivos escritos con formato de entero.
        cell.value = value[1]
        cell.number_format = '#,##0'
