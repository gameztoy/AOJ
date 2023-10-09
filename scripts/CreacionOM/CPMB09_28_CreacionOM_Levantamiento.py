import argparse
from modulos.AOJApp import AOJ
from modulos.BarraMenu import irA
from modulos.OficioManualPartes import OficioManualPartes
from modulos.OficioManual import OficioManual
from modulos.Reporte import Reporte

"""
CREA UN OFICIO MANUAL CPMB09_28
- Persona (default 27203123995[Betty] | argumento --persona)
- Tipo de oficio: Levantamiento de Embargo General de Fondos y Valores
- Vigencia: Actuales
- Sin Transferencia (Transferencia NO)
- Moneda: Pesos
- Sin montos
"""

parser = argparse.ArgumentParser()
parser.add_argument("-p", "--persona")
parser.add_argument("-o", "--oficio")
parser.add_argument("-a", "--anio")
args = parser.parse_args()
denominacion = args.persona
oficio = args.oficio
anio = args.anio

"""Comienza la ejecucion del caso"""
newInstance = AOJ()
app = newInstance.retornarAOJApp()
reporte = Reporte("Ejecucion Levantamiento", "CPMB 09.28 - Creacion OM  de Levantamiento de Embargo General de Fondos y Valores")

# Abrimos la ventana de oficios manuales
irA(app, "Novedades -> Oficio Manual", reporte)
oficioManual = OficioManual(app, reporte)

# Presionamos el boton agregar
oficioManual.presionarAgregarOficioManual()

# Cargamos los datos generales (expediente, caratula, juzgado)
oficioManual.cargarDatosGenerales(12345)

# Presionamos agregar en Partes
oficioManual.presionarAgregarParte()

oficioManualPartes = OficioManualPartes(app, reporte)

# Ingresamos el cuit de la persona
oficioManualPartes.ingresarDenominacion(denominacion)

# Realizamos el cruce individual
oficioManualPartes.realizarCruceIndividual()

# Seleccionamos el tipo de oficio
oficioManualPartes.seleccionarTipoOficio('Levantamiento de Embargo General de Fondos y Valores')

# Enlazamos un oficio del embargo a levantar
oficioManualPartes.enlazarOficio(oficio, anio)

# Seleccionamos la vigencia
oficioManualPartes.seleccionarVigencia('Actuales')

# Seleccionar transferencia
oficioManualPartes.seleccionarTransferencia('No')

# Tildamos la opción sin montos
oficioManualPartes.seleccionarSinMontos()

# Seleccionamos la moneda
oficioManualPartes.seleccionarMoneda('Pesos')

# Presionamos aceptar para agregar las partes
oficioManualPartes.presionarAceptar()

# Presionamos aceptar para crear el oficio manual
numero = oficioManual.aceptarOficioManual()

# Cerramos la aplicación
newInstance.closeAOJApp()

#Terminamos el reporte
reporte.terminarReporte()
