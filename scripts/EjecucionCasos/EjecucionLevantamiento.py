import csv
import os

"""
    Ejecución de los casos:
        - CPMB09_28 Creación OM Levantamiento de Embargo General de Fondos y Valores

    Parametros:
        :param -p Cuit de la persona
        :param -o Numero del oficio 
        :param -a Año del oficio
"""

path = os.getcwd()
path = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\cuitOficioAnio.csv"

with open(path) as csv_file:
    cantFilas = sum(1 for line in csv_file)

    csv_file.seek(0)

    csv_reader = csv.reader(csv_file)
    rows = list(csv_reader)

    for index in range(cantFilas):
        os.system("python ..\\CreacionOM\\CPMB09_28_CreacionOM_Levantamiento.py -p " + str(rows[index][0]) +
                  " -o " + str(rows[index][1]) +
                  " -a " + str(rows[index][2]))
