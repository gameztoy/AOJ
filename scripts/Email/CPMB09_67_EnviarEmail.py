import argparse

from modulos.AOJApp import AOJ
from modulos.Acciones import Acciones
from modulos.BarraMenu import irA
from modulos.EnviarEmail import EnviarEmail
from modulos.Reporte import Reporte
from modulos.SeleccionarOficio import SelectOficio

parser = argparse.ArgumentParser()
parser.add_argument("-o", "--oficio")
parser.add_argument("-a", "--anio")
args = parser.parse_args()

reporte = Reporte("Ejecucion Emails", "CPMB09.67 - Enviar Email Individual")

if args.oficio and args.anio:
    numOficio = args.oficio
    anioOficio = args.anio
else:
    numOficio = 215
    anioOficio = 2020

#Iniciamos la aplicación
newInstance = AOJ()
app = newInstance.retornarAOJApp()

#Abrimos el menu acciones
acciones = Acciones(app, reporte)
irA(app, "Oficios->Acciones", reporte)

#Entramos a enviar emails individuales
acciones.desplegarMenu("Emails")
acciones.abrirAccion("Enviar Email Individual")

enviarEmail = EnviarEmail(app, reporte)

#Seleccionamos un oficio existente
enviarEmail.presionarSeleccionarOficio()
seleccion = SelectOficio(app, reporte)
seleccion.seleccionarOficioBase(numOficio, anioOficio)

#Aceptamos el envio
enviarEmail.aceptarEnvio()

reporte.terminarReporte()

#Cerramos la aplicación
newInstance.closeAOJApp()
