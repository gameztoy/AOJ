import time
from modulos.AOJApp import AOJ
from modulos.BarraMenu import irA
from modulos.ConsultaRespuesta import CR
import argparse
from modulos.Reporte import Reporte

"""
Realiza las operaciones (adjuntar, ver, eliminar) sobre los archivos adjuntos en Consulta Respuesta
"""

parser = argparse.ArgumentParser()
parser.add_argument("-a", "--archivo")
parser.add_argument("-o", "--oficio")
parser.add_argument("-y", "--year")
args = parser.parse_args()

reporte = Reporte("Ejecucion CR Archivos Adjuntos", "Operaciones sobre archivos en CR")

newInstance = AOJ()
app = newInstance.retornarAOJApp()

# ingreso a consulta respuesta
newConsultaRespuesta = CR(app, reporte)
irA(app, "Oficios->Consulta respuestas", reporte)

if args.archivo:
    oficio = args.oficio
    year = args.year
else:
    oficio = 179
    year = 2020

# selecciona un oficio
newConsultaRespuesta.seleccionarOficio(year, oficio)

# va a la pestaña detalle
newConsultaRespuesta.irADetalle()

# Ingresamos la denominación
if args.archivo:
    archivo = args.archivo
else:
    archivo = '\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\word.doc'

# Agrego el archivo
newConsultaRespuesta.agregarArchivoAdjunto(archivo)

# Veo el archivo agregado
index = newConsultaRespuesta.buscarArchivoAdjunto(archivo)
newConsultaRespuesta.verAchivoAdjunto(index)
time.sleep(4)

# Elimina el archivo
newConsultaRespuesta.eliminarAchivoAdjunto()

# Agrego el archivo
newConsultaRespuesta.agregarArchivoAdjunto(archivo)

newInstance.closeAOJApp()

reporte.terminarReporte()
