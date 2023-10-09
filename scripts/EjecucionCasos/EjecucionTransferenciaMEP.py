import os
import openpyxl

"""
Created on Fri Dic 2 2022
@author: macalderontsoft

Ejecución de los casos de Embargo por Transferencia MEP.
En la variable cells indicar inicio y fin de datos en hoja.

Edited by:
"""

path = os.getcwd()
path = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\transferenciaMEP.xlsx"

book = openpyxl.load_workbook(path)
sheet = book.active
cells = sheet['A2':'H3']

for row in cells:
    datos = [cell.value for cell in row]
    caso = "".join(datos[0])
    cuit = datos[1]
    imprrre = "".join(datos[2])
    bloqueo = "".join(datos[3])
    impmep = "".join(datos[4])
    moneda = "".join(datos[5])
    capital = datos[6]
    interes = datos[7]

    os.system("python ..\\CreacionOM\\CPMB09_26_OMEmbargoPorTransferenciaMEP.py -n" + caso +
              " -p " + cuit +
              " -i " + imprrre +
              " -b " + bloqueo +
              " -m " + impmep +
              " -a " + moneda +
              " -e " + capital +
              " -s " + interes)
