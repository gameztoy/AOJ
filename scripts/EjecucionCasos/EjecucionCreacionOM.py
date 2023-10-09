import os
import openpyxl

"""
Created on Fri Nov 25 2022
@author: macalderontsoft

Ejecución de los casos de creación de Oficios Manuales.
En la variable cells indicar inicio y fin de datos en hoja.

Edited by:
"""

path = os.getcwd()
path = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\datosCreacionOm.xlsx"

book = openpyxl.load_workbook(path)
sheet = book.active
cells = sheet['A2':'I15']

for row in cells:
    datos = [cell.value for cell in row]
    caso = "".join(datos[0])
    cuit = datos[1]
    oficio = "".join(datos[2])
    vigencia = "".join(datos[3])
    transferencia = "".join(datos[4])
    monto = "".join(datos[5])
    moneda = "".join(datos[6])
    capital = datos[7]
    interes = datos[8]

    os.system("python ..\\CreacionOM\\CPMB09_CreacionOM.py -n" + caso +
              " -p " + cuit +
              " -t " + oficio +
              " -v " + vigencia +
              " -r " + transferencia +
              " -m " + monto +
              " -a " + moneda +
              " -e " + capital +
              " -s " + interes)
