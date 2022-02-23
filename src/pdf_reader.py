import datetime
import os
import re

import pdfplumber

from src.downloader import Downloader


class PDF_Reader():
    def __init__(self):
        # Descargo el último pdf disponible
        download = Downloader()
        download.download_pdf()
        self.pdf_name = download.pdf_name

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

    def read_file(self):

        print("Abriendo PDF…")

        # creating a pdf reader object
        fileReader = pdfplumber.open(self.pdf_file)

        # Bool que me indica si he encontrado alguna tabla
        table_found = False

        print("Leyendo PDF…")
        for page in fileReader.pages:
            # Extraigo el texto de la página actual
            self.fileText = page.extract_text().replace('\n', '')

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
        # Me quedo solo con los datos
        # Empiezan con una fecha y terminan con un número
        reSearch = re.search("(\d{2}/\d{2}/\d{4}.+\d+)", self.fileText)
        self.fileText = reSearch.group(1)
        # Guardo en una lista todos los datos, no están ordenados
        data = self.fileText.split(" ")
        # Elimino las cadenas vacías
        data = [i for i in data if i]
        # Por si ha entrado algún texto que no quería
        data = self.__check_header(data)
        # Obtengo los pares eliminando los agregados
        data = list(zip(data[::3], data[1::3]))
        # Convierto los datos a fecha y enteros
        data = [[datetime.datetime.strptime(
            i[0], "%d/%m/%Y"), int(i[1])] for i in data]
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
