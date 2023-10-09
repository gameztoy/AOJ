import argparse
import time
from modulos.AOJApp import AOJ
from modulos.Acciones import Acciones
from modulos.BarraMenu import irA
from modulos.Reporte import Reporte
from modulos.RespuestaDeEmbargos import RegistrarRespuestasEmbargos
from modulos.ValidacionEmbargoCrecer import CuentasPorCuentas

parser = argparse.ArgumentParser()
parser.add_argument("-o", "--oficio")
parser.add_argument("-a", "--anio")
parser.add_argument("-p", "--persona")
parser.add_argument("-t", "--tipo")
parser.add_argument("-i", "--importe")
args = parser.parse_args()
cuit = args.persona
oficio = args.oficio
anio = args.anio
importe = args.importe
tipo = args.tipo

"""Comienza la ejecucion del caso"""
newInstance = AOJ()
app = newInstance.retornarAOJApp()
reporte = Reporte("Ejecucion RRRE", "CPMB 09.20 - Retención por registrar respuesta de embargo")

#Abrimos la ventana de Acciones
irA(app, "Oficio -> Acciones", reporte)
acciones = Acciones(app, reporte)
acciones.desplegarMenu("Respuestas")
acciones.abrirAccion("Registrar Respuestas de Embargos")

#Se abre la ventana 'Registrar Respuestas de Embargos'
registrarRE = RegistrarRespuestasEmbargos(app, reporte)

#Se selecciona el oficio con el años correspondiente
registrarRE.seleccionarOficio(oficio, anio)

#Se agregan las cuentas a embargar
registrarRE.agregarCuenta(tipo, 1, importe, cuit)

#Presiona aceptar y avanzar
registrarRE.presionarAceptarYAvanzar()

#Confirma las cuentas a embargar
registrarRE.aceptarSeleccionarCuentas()

#Salir de Resgistrar Respuesta de Embargo
registrarRE.presionarSalir()

#Cierra la app
newInstance.closeAOJApp()

#Validacion en Crecer - Cuentas por Cuenta - Mov. por Fecha.
time.sleep(2)
cxc = CuentasPorCuentas(app, reporte)
cxc.buscarCuentasPorCuenta()
cxc.consultaMovPorFecha(importe)

#Terminamos el reporte
reporte.terminarReporte()

