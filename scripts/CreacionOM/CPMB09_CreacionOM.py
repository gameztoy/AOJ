import argparse
from modulos.AOJApp import AOJ
from modulos.BarraMenu import irA
from modulos.ConsultaRespuesta import CR
from modulos.OficioManual import OficioManual
from modulos.OficioManualPartes import OficioManualPartes
from modulos.Reporte import Reporte

"""
Created on Fri Nov 25 2022
@author: macalderontsoft

Creacion de distintos tipos de Oficio Manuales - Casos desde el CPMB09_01 a CPMB09_14
Datos obtenidos desde .xlsx: 
- Caso: nombre de caso - caso
- Cuit: cuit persona - cuit
- Oficio: Tipo de oficio - oficio
- Vigencia: Tipo de vigencia - vigencia
- Transferencia: Con o Sin - transferencia
- Moneda: Pesos o Dolar - moneda
- Capital: Monto a definir - capital
- Interes: Monto a definir - interes

Edited by:
"""

parser = argparse.ArgumentParser()
parser.add_argument("-n", "--caso", nargs='+')
parser.add_argument("-p", "--cuit", nargs='+')
parser.add_argument("-t", "--oficio", nargs='+')
parser.add_argument("-v", "--vigencia", nargs='+')
parser.add_argument("-r", "--transferencia", nargs='+')
parser.add_argument("-m", "--monto", nargs='+')
parser.add_argument("-a", "--moneda", nargs='+')
parser.add_argument("-e", "--capital", nargs='+')
parser.add_argument("-s", "--interes", nargs='+')
args = parser.parse_args()
caso = args.caso[0]
cuit = args.cuit[0]
oficio = args.oficio[0]
vigencia = args.vigencia[0]
transferencia = args.transferencia[0]
monto = args.monto[0]
moneda = args.moneda[0]
capital = args.capital[0]
interes = args.interes[0]

"""Comienzo del caso en AOJ"""
newInstance = AOJ()
app = newInstance.retornarAOJApp()
reporte = Reporte("Oficios Creados", "CPMB09_"+caso)

# Abrimos la ventana de oficios manuales
irA(app, "Novedades -> Oficio Manual", reporte)
oficioManual = OficioManual(app, reporte)

# Presionamos el boton agregar
oficioManual.presionarAgregarOficioManual()

# Cargamos los datos generales (expediente, caratula, juzgado)
oficioManual.cargarDatosGenerales(12346)

# Presionamos agregar en Partes
oficioManual.presionarAgregarParte()
oficioManualPartes = OficioManualPartes(app, reporte)

# Ingresamos la persona
if cuit == '':
    # Presionamos sin persona
    oficioManualPartes.presionarSinPersona()
else:
    # Ingresamos el cuit de la persona
    oficioManualPartes.ingresarDenominacion(cuit)

# Realizamos el cruce individual
oficioManualPartes.realizarCruceIndividual()

# Seleccionamos el tipo de oficio
oficioManualPartes.seleccionarTipoOficio(oficio)

# Seleccionamos la vigencia
oficioManualPartes.seleccionarVigencia(vigencia)

# Seleccionar transferencia
oficioManualPartes.seleccionarTransferencia(transferencia)

# Seleccionar monto
if monto == 'Sin montos':
    oficioManualPartes.seleccionarSinMontos()

# Seleccionamos la moneda
oficioManualPartes.seleccionarMoneda(moneda)

# Ingresamos capital e intereses
if capital and interes != '':
    oficioManualPartes.ingresarCapitalYIntereses(capital, interes)

# Presionamos aceptar para agregar las partes
oficioManualPartes.presionarAceptar()

# Presionamos aceptar para crear el oficio manual
numero=oficioManual.aceptarOficioManual()

anio = numero.split(",")[1]
numero = numero.split(",")[0]

oficioManual.presionarSalir()

# Vamos a Consulta Respuesta a ver el oficio creado
irA(app, "Oficios -> Consulta Respuestas", reporte)

cr = CR(app, reporte)
cr.seleccionarOficio(anio, numero)
cr.verOficio()
cr.verDatosOficioCreado()

#Cerramos la aplicación
newInstance.closeAOJApp()

#Terminamos el reporte
reporte.terminarReporte()

