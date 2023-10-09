import os
import csv

"""
    Ejecución de los casos:
        - CPMB09.35 Emitir Carta
        - CPMB09.36 Confirmar Carta
        - CPMB09.37 Editar Carta
        - CPMB09.38 Volver a Confirmar Carta
        - CPMB09.39 Registrar Recepcion de Carta
        
    Parametros:
        :param -o Numero del oficio 
        :param -a Año del oficio 
"""

path = os.getcwd()
path = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\oficioAnio.csv"

with open(path) as csv_file:
    cantFilas = sum(1 for line in csv_file)

    csv_file.seek(0)

    csv_reader = csv.reader(csv_file)
    rows = list(csv_reader)

    for index in range(cantFilas):
        os.system("python ..\\Cartas\\CPMB09_35_EmitirCarta.py -o " + str(rows[index][0]) +
                  " -a " + str(rows[index][1]))
        # os.system("python ..\\Cartas\\CPMB09_36_ConfirmarCarta.py -o " + str(rows[index][0]) +
        #           " -a " + str(rows[index][1]))
        # os.system("python ..\\Cartas\\CPMB09_37_EditarCarta.py -o " + str(rows[index][0]) +
        #           " -a " + str(rows[index][1]))
        # os.system("python ..\\Cartas\\CPMB09_38_VolverAConfirmarCarta.py -o " + str(rows[index][0]) +
        #           " -a " + str(rows[index][1]))
        # os.system("python ..\\Cartas\\CPMB09_39_RegistrarRecepcionCarta.py -o " + str(rows[index][0]) +
        #           " -a " + str(rows[index][1]))
