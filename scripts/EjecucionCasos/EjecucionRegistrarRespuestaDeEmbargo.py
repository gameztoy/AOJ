import os
import csv

"""
    Ejecución de los casos:
        - CPMB09_20 Retención por registrar respuesta de embargo

    Parametros:
        :param -p Cuit de la persona 
        :param -o Numero del oficio a enlazar
        :param -a Año del oficio a enlazar
        :param -t Tipo de cuenta - Dolar / Pesos
        :param -i Importe 
"""

path = os.getcwd()
path = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\registrarRespuestaDeEmbargo.csv"

with open(path) as csv_file:
    cantFilas = sum(1 for line in csv_file)

    csv_file.seek(0)

    csv_reader = csv.reader(csv_file)
    rows = list(csv_reader)

    for index in range(cantFilas):
        os.system("python ..\\RetencionRegistrarRespuestaEmbargo\\CPMB09_20_RetencionPorRegistrarRespuestaDeEmbargo.py -p " + str(rows[index][0]) +
                  " -o " + str(rows[index][1]) +
                  " -a " + str(rows[index][2]) +
                  " -t " + str(rows[index][3]) +
                  " -i " + str(rows[index][4]))
