import argparse
import time

from modulos.AOJApp import AOJ
from modulos.BarraMenu import irA
from modulos.OficioManualPartes import OficioManualPartes
from modulos.OficioManual import OficioManual, agregarArchivoAdjunto, OMModificar
from modulos.Reporte import Reporte

# Obtenemos el parametro
parser = argparse.ArgumentParser()
parser.add_argument("-a", "--archivo")
parser.add_argument("-p", "--persona")
args = parser.parse_args()

reporte = Reporte("Ejecucion Creación OM Adjuntar Archivo", "Operaciones sobre archivos en Creacion OM")

# Si no se le pasa un archivo por parametro toma por defecto un pdf
if args.archivo:
    pathArchivo = args.archivo
else:
    pathArchivo = "\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\pdf.pdf"

if args.persona:
    denominacion = args.persona
else:
    denominacion = 27203123995

# Abrimos la app
newInstance = AOJ()
app=newInstance.retornarAOJApp()

# Abrimos la ventana de oficios manuales
oficioManual = OficioManual(app, reporte)
irA(app,"Novedades -> Oficio Manual", reporte)

# Presionamos el boton agregar
oficioManual.presionarAgregarOficioManual()

# Cargamos los datos generales (expediente, caratula, juzgado)
oficioManual.cargarDatosGenerales(12345)

# Presionamos agregar en Partes
oficioManual.presionarAgregarParte()

oficioManualPartes = OficioManualPartes(app, reporte)

# Ingresamos la denominación
oficioManualPartes.ingresarDenominacion(denominacion)

# Realizamos el cruce individual
oficioManualPartes.realizarCruceIndividual()

# Seleccionamos el tipo de oficio
oficioManualPartes.seleccionarTipoOficio('Embargos General de Fondos y  Valores')

# Seleccionamos la vigencia
oficioManualPartes.seleccionarVigencia('Actuales y Futuras')

# Seleccionar transferencia
oficioManualPartes.seleccionarTransferencia('No')

# Seleccionamos la moneda
oficioManualPartes.seleccionarMoneda('Pesos')

oficioManualPartes.ingresarCapitalYIntereses(5000, 100)

# Presionamos aceptar para agregar las partes
oficioManualPartes.presionarAceptar()

# Agregamos archivo adjunto
agregarArchivoAdjunto(app, pathArchivo, reporte)

# Presionamos aceptar para crear el oficio manual
numero = oficioManual.aceptarOficioManual()

anio = numero.split(",")[1]
numero = numero.split(",")[0]

oficioManual.presionarSalir()

omModificar = OMModificar(app, reporte)

irA(app, 'Novedades -> Oficios Manuales', reporte)
newOMM = OMModificar(app, reporte)
newOMM.modificarOM(numero, anio)
time.sleep(2)

#ME MUEVO A LA PESTAÑA ARCHIVOS ADJUNTOS
newOMM.irArchivoAdjunto()
time.sleep(2)

# Veo el archivo
newOMM.verAchivoAdjunto(pathArchivo)
time.sleep(2)

# Elimino el archivo
newOMM.eliminarAchivoAdjunto(pathArchivo)

# Vuelvo a agregar el archivo
agregarArchivoAdjunto(app, pathArchivo, reporte)

# Finalizo las modificaciones
newOMM.finalizarModificaciones()

# Cerramos la aplicación
newInstance.closeAOJApp()
