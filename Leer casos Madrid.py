import datetime
import requests
import os
import pdfplumber
from openpyxl import load_workbook

def get_map_url(date):
    prefix = "https://www.comunidad.madrid/sites/default/files/doc/sanidad/"
    sufix = "_cam_covid19.pdf"
    date_str = str(date.year)[-2:] + "{:02d}".format(date.month) + "{:02d}".format(date.day)
    url = prefix + date_str + sufix
    return url

def download_pdf(date):
    response = requests.get(get_map_url(date))
    # Estoy intentando acceder al PDF de hoy y quizás aún no se ha publicado.
    while not response.status_code == 200:
        # Me muevo al día anterior
        date = date - datetime.timedelta(days=1)
        response = requests.get(get_map_url(date))

    with open("tmp.pdf", 'wb') as f:
        print("Último informe del día " + str(date.day) + "-" + str(date.month) + "-" + str(date.year))
        # Descargo el PDF
        f.write(response.content)

def check_header(data):
    while data:
        try:
            # Compruebo que el primer dato sea una fecha
            datetime.datetime.strptime(data[0], "%d/%m/%Y")
        except:
            # en caso contrario elimino este elemento
            data.pop(0)
        else:
            # Si en efecto es una fecha, no hay nada más que modificar en el encabezado
            break

    return data

def get_clear_data(text):
    # Elimino la cabecera del texto
    text = text.split("Diario  Acumulado")[-1]
    # Elimino la cola del texto
    # La "F" es el inicio del pie de página:
    # "Fuente: Dirección General de Salud Pública"
    text = text.split("F")[0]
    # Guardo en una lista todos los datos, no están ordenados
    data = text.split(" ")
    # Elimino las cadenas vacías
    data = [i for i in data if i]
    # Por si ha entrado algún texto que no quería
    data = check_header(data)
    # Obtengo los pares eliminando los agregados
    data = list(zip(data[::3], data[1::3]))
    # Convierto los datos a fecha y enteros
    data = [[datetime.datetime.strptime(i[0], "%d/%m/%Y"), int(i[1])] for i in data]
    # Ordeno los datos por fecha
    data.sort(key=lambda tup: tup[0])
    return data


if __name__ == "__main__":
    # Descargo el último PDF
    download_pdf(datetime.date.today())

    # creating an object
    file = open('tmp.pdf', 'rb')
    # Fila donde escribo los datos
    indexLine = 2
    ExcelName = "Casos Comunidad de Madrid.xlsx"
    workbook = load_workbook(ExcelName)
    worksheet = workbook[workbook.sheetnames[0]]


    # creating a pdf reader object
    fileReader = pdfplumber.open(file)
    for page in fileReader.pages:
        fileText = page.extract_text().replace('\n','')
        # Busco el encabezado de las hojas con tablas
        if fileText.find("Se realiza una actualización diaria ") < 0:
            if fileText.find("Evolución casos positivos de") >= 0:
                # Encabezado de la página posterior a las tablas. Podemos dejar de leer.
                break
            # Avanzo a la siguiente página. Aún no he encontrado las tablas.
            continue
        data = get_clear_data(fileText)
        for value in data:
            # Fecha escrita con formato de fecha
            worksheet.cell(row=indexLine, column=1).value = value[0]
            worksheet.cell(row=indexLine, column=1).number_format = 'mm-dd-yy'
            # Positivos escritos con formato de entero
            worksheet.cell(row=indexLine, column=2).value = value[1]
            worksheet.cell(row=indexLine, column=2).number_format = '#,##0'
            indexLine = indexLine + 1

    file.close()  # Cierro el archivo PDF
    os.remove("tmp.pdf")  # Elimino el PDF
    workbook.save(ExcelName)  # Guardo el xlsx
    workbook.close()  # Cierro el xlsx
