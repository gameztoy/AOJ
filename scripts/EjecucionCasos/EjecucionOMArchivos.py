import csv
import os

"""
    Ejecución de los casos de Creación de Oficios Manuales con Archivo Adjuntado:
        -CPMB 09.41 - Creacion OM Adjuntar Archivo PDF
        -CPMB 09.42 - Creacion OM Adjuntar Archivo Word
        -CPMB 09.43 - Creacion OM Adjuntar Archivo Excel
        -CPMB 09.44 - Creacion OM Adjuntar Archivo JPG
        -CPMB 09.45 - Creacion OM Eliminar Archivo PDF
        -CPMB 09.46 - Creacion OM Eliminar Archivo Word
        -CPMB 09.47 - Creacion OM Eliminar Archivo Excel
        -CPMB 09.48 - Creacion OM Eliminar Archivo JPG
        -CPMB 09.49 - Creacion OM Ver Archivo PDF
        -CPMB 09.50 - Creacion OM Ver Archivo Word
        -CPMB 09.51 - Creacion OM Ver Archivo Excel
        -CPMB 09.52 - Creacion OM Ver Archivo JPG
    
    Parametros:
        :param -a Path del archivo a adjuntar
        :param -p Cuit persona
"""

path = os.getcwd()
path = path.split('automatizacionaoj')[0]+"automatizacionaoj\\recursos\\cuitsPersonas.csv"

with open(path) as csv_file:
    csv_reader = csv.reader(csv_file)
    rows = list(csv_reader)

    # PDF
    pdf = "\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\pdf.pdf"
    os.system("python ..\\CreacionOMArchivosAdjuntos\\CreacionOMArchivosAdjuntos.py -a " + pdf + " -p " + str(rows[1]))

    # WORD
    word = "\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\word.doc"
    os.system("python ..\\CreacionOMArchivosAdjuntos\\CreacionOMArchivosAdjuntos.py -a " + word + " -p " + str(rows[2]))

    # EXCEL
    excel = "\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\excel.xls"
    os.system("python ..\\CreacionOMArchivosAdjuntos\\CreacionOMArchivosAdjuntos.py -a " + excel + " -p " + str(rows[3]))

    # JPG
    jpg = "\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\imagen.jpg"
    os.system("python ..\\CreacionOMArchivosAdjuntos\\CreacionOMArchivosAdjuntos.py -a " + jpg + " -p " + str(rows[4]))
