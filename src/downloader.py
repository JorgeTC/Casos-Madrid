import concurrent.futures
import datetime

import requests
from bs4 import BeautifulSoup


class URLFormat():

    def __init__(self, prefix, sufix, extension, date_format):
        self.prefix = prefix
        self.sufix = sufix
        self.extension = extension
        self.date_format = date_format

    def get_url(self, date):
        return self.prefix + self.date_format(date) + self.sufix + self.extension


class Downloader():

    SZ_CAM_URL = 'https://www.comunidad.madrid'
    SZ_ACTUAL_SITUATION = SZ_CAM_URL + \
        '/servicios/salud/coronavirus#datos-situacion-actual'
    SZ_CAM_FILES = SZ_CAM_URL + "/sites/default/files"

    def __init__(self, date=datetime.date.today()):
        # Fecha que me interesa leer
        self.date = date

        # Nombre con el que guardo el PDF temporal
        self.pdf_name = 'tmp.pdf'

    def download_pdf(self):
        # Intento obtener el link desde la página de la CAM
        response = self.__get_pdf_response_fromCAM()

        # Si no lo he conseguido, pruebo las combinaciones de links
        if not response:
            response = self.__try_combinations()

        with open(self.pdf_name, 'wb') as f:
            # Descargo el PDF
            f.write(response.content)

    def __get_pdf_response_fromCAM(self):
        # Accedo a la página de la CAM donde está el enlace a la situación actual
        response = requests.get(self.SZ_ACTUAL_SITUATION)
        parsed = BeautifulSoup(response.text, 'html.parser')

        # Accedo al hipervínculo.
        # Hay varios h2 en la página per el primero es el que contiene la información que necesito
        hipervinculo = parsed.find('h2', {'class': "rtecenter"})
        try:
            # Obtengo la dirección a la que condice el enlace
            a_link = hipervinculo.find('a', href=True)
            sz_link = a_link.attrs['href']
        except:
            return None

        # Descargo la página
        response = requests.get(self.SZ_CAM_URL + sz_link)
        # Comrpuebo que la desarga haya ido bien
        if response.status_code == 200:
            # En este punto ya sé seguro que voy a devolver una página correcta.
            # Imprimo por pantalla el texto para decir qué fecha voy a extraer
            print(hipervinculo.text)
            return response
        else:
            return None

    def __try_combinations(self):
        # Defino la lista de prefijos
        self.__list_of_prefix()

        # Estoy intentando acceder al informe de hoy
        while True:
            # Pido el informe de hoy
            response = self.__get_date_response()
            if response.status_code == 200:
                # Aviso por pantalla de cuál es el último informe
                print("Último informe del día " + str(self.date.day) +
                      "-" + str(self.date.month) + "-" + str(self.date.year))
                # Devuelvo el contenido de la página a la que he conseguido acceder
                return response

            # Ninguna de las modulaciones es válida.
            # Asumo que aún no se ha publicado el informa diario.
            self.date = self.date - datetime.timedelta(days=1)

    def __list_of_prefix(self):
        # Lista de todas las variantes de prefijo hasta ahora
        prefix_list = [self.SZ_CAM_FILES + "/doc/sanidad/",
                       self.SZ_CAM_FILES + "/doc/sanidad/prev/",
                       self.SZ_CAM_FILES + "/aud/sanidad/prev/",
                       self.SZ_CAM_FILES + "/aud/sanidad/",
                       self.SZ_CAM_FILES + "/doc/presidencia/"]
        # Lista de todas las variantes de sufijo hasta ahora
        sufix_list = ["_cam_covid19",
                      "_cam_covid"]
        # Lista de todas las extensiones hasta ahora
        extension_list = [".pdf", ".pdf.pdf"]
        # Lista de cómo se ha formateado la fecha
        date_format_list = [lambda x: str(x.year)[-2:] + "{:02d}".format(x.month) + "{:02d}".format(x.day),
                            lambda x: str(x.year) + "{:02d}".format(x.month) + "{:02d}".format(x.day)]

        # Creo una lista donde guardo todas sus posibles combinaciones
        self.pre_sufix_list = []
        # Iteraciones ordenadas de más cambiante a más estable
        for extension in extension_list:
            for sufix in sufix_list:
                for prefix in prefix_list:
                    for date in date_format_list:
                        self.pre_sufix_list.append(URLFormat(prefix,
                                                             sufix,
                                                             extension,
                                                             date))
        return

    def __get_date_response(self):

        # Obtengo todas las variantes que puedo para el link de hoy
        links = [i.get_url(self.date) for i in self.pre_sufix_list]

        # De forma paralelizada descargo el contenido de todos los links
        with concurrent.futures.ThreadPoolExecutor() as executor:
            responses = list(executor.map(requests.get, links))

        # Miro cuántos de ellos me han devuelto realmente una página.
        # Me espero que sólo uno de ellos tengo response_code 200
        valid_resposes = [
            respose for respose in responses if respose.status_code == 200]
        if len(valid_resposes):
            # Devuelvo una página válida
            return valid_resposes[0]
        else:
            # Devuelvo una página no válida
            return responses[0]
