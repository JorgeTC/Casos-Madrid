import datetime
import requests
import os
import pdfplumber
from openpyxl import load_workbook

class PDF_Reader():
    def __init__(self, date=datetime.date.today()):
        # Fecha que me interesa leer
        self.date = date
        # Nombre con el que guardo el PDF temporal
        self.pdf_name = 'tmp.pdf'

        # Descargo el último pdf disponible
        self.download_pdf()

        # Abro el archivo pdf en modo lectura
        self.pdf_file = open(self.pdf_name, 'rb')
        self.fileText = ""
        self.data = []
        return

    def __del__(self):
        # Cierro el archivo PDF
        self.pdf_file.close()
        # Elimino el PDF
        os.remove(self.pdf_name)

    def get_map_url(self, attempt=1):
        # Aplico los formatos de cadena a año, mes y día
        date_str = str(self.date.year)[-2:] + "{:02d}".format(self.date.month) + "{:02d}".format(self.date.day)
        # Sabiendo la fecha, compongo la dirección de su pdf
        if attempt == 1:
            prefix = "https://www.comunidad.madrid/sites/default/files/doc/sanidad/"
            sufix = "_cam_covid19.pdf"
        elif attempt == 2:
            prefix = "https://www.comunidad.madrid/sites/default/files/doc/sanidad/prev/"
            sufix = "_cam_covid19.pdf"

        url = prefix + date_str + sufix

        return url

    def download_pdf(self):
        # Estoy intentando acceder al informe de hoy
        while True:
            # Pruebo las distintas modulaciones de la dirección
            response = requests.get(self.get_map_url(1))
            if response.status_code == 200:
                break

            response = requests.get(self.get_map_url(2))
            if response.status_code == 200:
                break

            # Ninguna de las modulaciones es válida.
            # Asumo que aún no se ha publicado el informa diario.
            self.date = self.date - datetime.timedelta(days=1)

        with open(self.pdf_name, 'wb') as f:
            print("Último informe del día " + str(self.date.day) + "-" + str(self.date.month) + "-" + str(self.date.year))
            # Descargo el PDF
            f.write(response.content)

    def read_file(self):

        print("Abriendo PDF…")

        # creating a pdf reader object
        fileReader = pdfplumber.open(self.pdf_file)

        # Bool que me indica si he encontrado alguna tabla
        table_found = False

        print("Leyendo PDF…")
        for page in fileReader.pages:
            # Extraigo el texto de la página actual
            self.fileText = page.extract_text().replace('\n','')

            # Compruebo si la página actual tiene tablas.
            if self.__has_tables():
                # Estoy en una página con tablas, leo su contenido.
                self.__get_clear_data()
                table_found = True
            # No estoy en una página con tablas.
            elif table_found:
                # Ya he leído todas las tablas.
                # Podemos dejar de leer.
                break

        # Ordeno los datos por fecha
        print("Ordenando datos…")
        self.data.sort(key=lambda tup: tup[0])

        return self.data

    def __has_tables(self):
        # Leo el encabezado de la página.
        # Compruebo si es el encabezado de una página con tablas.
        return self.fileText.find("Se realiza una actualización diaria ") >= 0

    def __get_clear_data(self):
        # Elimino la cabecera del texto
        self.fileText = self.fileText.split("Diario  Acumulado")[-1]
        # Elimino la cola del texto
        # La "F" es el inicio del pie de página:
        # "Fuente: Dirección General de Salud Pública"
        self.fileText = self.fileText.split("F")[0]
        # Guardo en una lista todos los datos, no están ordenados
        data = self.fileText.split(" ")
        # Elimino las cadenas vacías
        data = [i for i in data if i]
        # Por si ha entrado algún texto que no quería
        data = self.__check_header(data)
        # Obtengo los pares eliminando los agregados
        data = list(zip(data[::3], data[1::3]))
        # Convierto los datos a fecha y enteros
        data = [[datetime.datetime.strptime(i[0], "%d/%m/%Y"), int(i[1])] for i in data]
        # Actualizo la variable miembro con lo leído en la página actual
        self.data = self.data + data

    def __check_header(self, list_page):
        # Los encabezados de la tabla siempre estarán al principio de mi lista.
        # Leo su primer elemento hasta que se afectivamente una fecha.
        while list_page:
            try:
                # Compruebo que el primer dato sea una fecha…
                datetime.datetime.strptime(list_page[0], "%d/%m/%Y")
            except:
                # …en caso contrario elimino este elemento
                list_page.pop(0)
            else:
                # Si en efecto es una fecha, no hay nada más que modificar en el encabezado.
                break

        return list_page

class Excel_writer():
    def __init__(self):
        self.excel_name = "Casos Comunidad de Madrid.xlsx"
        print("Abriendo Excel…")
        self.workbook = load_workbook(self.excel_name)

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


if __name__ == "__main__":

    last_data = PDF_Reader()
    list_data = last_data.read_file()
    del last_data

    excel = Excel_writer()
    excel.write_data(list_data)
    del excel
