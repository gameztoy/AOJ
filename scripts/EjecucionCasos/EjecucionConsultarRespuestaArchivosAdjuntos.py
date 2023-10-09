import os
import csv

"""
    Ejecución de los casos:
        -CPMB 09.54 - Consultar Respuesta del Oficio_OM adjuntar Archivo Word
        -CPMB 09.55 - Consultar Respuesta del Oficio_OM adjuntar Archivo Excel
        -CPMB 09.53 - Consultar Respuesta del Oficio_OM adjuntar Archivo PDF
        -CPMB 09.56 - Consultar Respuesta del Oficio_OM adjuntar Archivo JPG
        -CPMB 09.57 - Consultar Respuesta del Oficio_OM Eliminar Archivo PDF
        -CPMB 09.58 - Consultar Respuesta del Oficio_OM Eliminar Archivo Word
        -CPMB 09.59 - Consultar Respuesta del Oficio_OM Eliminar Archivo Excel
        -CPMB 09.60 - Consultar Respuesta del Oficio_OM Eliminar Archivo JPG
        -CPMB 09.61 - Consultar Respuesta del Oficio_OM Ver Archivo PDF
        -CPMB 09.62 - Consultar Respuesta del Oficio_OM Ver Archivo Word
        -CPMB 09.63 - Consultar Respuesta del Oficio_OM Ver Archivo Excel
        -CPMB 09.64 - Consultar Respuesta del Oficio_OM Ver Archivo JPG

    Parametros:
        :param -o Numero del oficio 
        :param -a Año del oficio
"""

# PDF
pdf = "\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\pdf.pdf"
# WORD
word = "\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\word.doc"
# EXCEL
excel = "\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\excel.xls"
# JPG
jpg = "\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\imagen.jpg"

path = os.getcwd()
path = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\oficioAnio.csv"

with open(path) as csv_file:
    cantFilas = sum(1 for line in csv_file)

    csv_file.seek(0)

    csv_reader = csv.reader(csv_file)
    rows = list(csv_reader)

    for index in range(cantFilas):
        # PDF
        os.system("python ..\\ConsultaRespuestaArchivosAdjuntos\\ConsultarRespuestaArchivosAdjuntos.py -a " + pdf +
                  " -o " + str(rows[index][0]) +
                  " -y " + str(rows[index][1]))
        # WORD
        os.system("python ..\\ConsultaRespuestaArchivosAdjuntos\\ConsultarRespuestaArchivosAdjuntos.py -a " + word +
                  " -o " + str(rows[index][0]) +
                  " -y " + str(rows[index][1]))
        # EXCEL
        os.system("python ..\\ConsultaRespuestaArchivosAdjuntos\\ConsultarRespuestaArchivosAdjuntos.py -a " + excel +
                  " -o " + str(rows[index][0]) +
                  " -y " + str(rows[index][1]))
        # JPG
        os.system("python ..\\ConsultaRespuestaArchivosAdjuntos\\ConsultarRespuestaArchivosAdjuntos.py -a " + jpg +
                  " -o " + str(rows[index][0]) +
                  " -y " + str(rows[index][1]))
