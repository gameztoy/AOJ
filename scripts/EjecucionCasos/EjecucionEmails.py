import os
import csv

"""
    Ejecución de los casos:
        - CPMB09_67_EnviarEmail
        - CPMB09_47_EnviarEmailPorLote
        
    Parametros:
        :param -o Numero del oficio a Enviar Email
        :param -a Año del oficio a Enviar Email
"""

path = os.getcwd()
path = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\oficioAnio.csv"

with open(path) as csv_file:
    cantFilas = sum(1 for line in csv_file)

    csv_file.seek(0)

    csv_reader = csv.reader(csv_file)
    rows = list(csv_reader)

    for index in range(cantFilas):
        os.system("python ..\\Email\\CPMB09_67_EnviarEmail.py -o " + str(rows[index][0]) +
                  " -a " + str(rows[index][1]))

# Enviar Email Por Lote (PARA SOJ)
#os.system("python ..\\Email\\CPMB09_47_EnviarEmailPorLote.py")
