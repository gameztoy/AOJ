import argparse
from modulos.AOJApp import AOJ
from modulos.Acciones import Acciones
from modulos.BarraMenu import irA
from modulos.Cartas import EmitirCarta, ConfirmarCarta
from modulos.ConsultaRespuesta import CR
from modulos.Reporte import Reporte
from modulos.SeleccionarOficio import SelectOficio

newInstance = AOJ()
app = newInstance.retornarAOJApp()
reporte = Reporte("Smoke de Cartas", "CPMB09.36 Confirmar Carta")

parser = argparse.ArgumentParser()
parser.add_argument("-o", "--oficio")
parser.add_argument("-a", "--anio")
args = parser.parse_args()

if args.oficio and args.anio:
    numOficio = args.oficio
    anioOficio = args.anio
else:
    numOficio = 215
    anioOficio = 2020

# Abrimos el menu acciones y entramos a enviar emails individuales
acciones = Acciones(app, reporte)
irA(app, "Oficios->Acciones", reporte)

acciones.desplegarMenu("Cartas")
acciones.abrirAccion("Confirmar Carta")

# Presionamos Seleccionar Oficio dentro de confirmar carta
confirmarCarta = ConfirmarCarta(app, reporte)
confirmarCarta.presionarSeleccionarOficio()

# Seleccionamos el oficio a confirmar carta
seleccion = SelectOficio(app, reporte)
seleccion.seleccionarOficioBase(numOficio,anioOficio)

# Presionamos aceptar para Confirmar la carta
confirmarCarta.presionarAceptar()

# Presionamos salir de Confirmar Carta
confirmarCarta.presionarSalir()

# Vamos a Consulta Respuestas
irA(app, "Oficios->Consulta respuestas", reporte)

# Selecciona el oficio
newConsultaRespuesta = CR(app, reporte)
newConsultaRespuesta.seleccionarOficio(anioOficio, numOficio)

# Validamos el estado bloqueado del oficio
emitirCarta = EmitirCarta(app, reporte)
emitirCarta.obtenerEstadoBloqueado("Si")

# Presionamos en editar carta
emitirCarta.clickearEmitirCarta()

# Verificamos que el archivo esta en modo solo lectura
emitirCarta.verificarSoloLectura(True)

# Terminamos el reporte
reporte.terminarReporte()

# Cerramos la app
newInstance.closeAOJApp()