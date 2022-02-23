import requests
from bs4 import BeautifulSoup


class Downloader():

    def __init__(self):

        # Nombre con el que guardo el PDF temporal
        self.pdf_name = 'tmp.pdf'

    def download_pdf(self):
        # Intento obtener el link desde la página de la CAM
        response = self.__get_pdf_response_fromCAM()

        # Si no lo he conseguido, pruebo las combinaciones de links
        if not response:
            return False

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
        # Compruebo que la descarga haya ido bien
        if response.status_code == 200:
            # En este punto ya sé seguro que voy a devolver una página correcta.
            # Imprimo por pantalla el texto para decir qué fecha voy a extraer
            print(hipervinculo.text)
            return response
        else:
            return None
