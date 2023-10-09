import argparse
from modulos.AOJApp import AOJ
from modulos.Acciones import Acciones
from modulos.BarraMenu import irA
from modulos.Cartas import EmitirCarta
from modulos.ConsultaRespuesta import CR
from modulos.Reporte import Reporte
from modulos.SeleccionarOficio import SelectOficio

newInstance = AOJ()
app = newInstance.retornarAOJApp()

reporte = Reporte("Smoke de Cartas", "CPMB09.35 Emitir Carta")

parser = argparse.ArgumentParser()
parser.add_argument("-o", "--oficio")
parser.add_argument("-a", "--anio")
args = parser.parse_args()

if args.oficio and args.anio:
    numOficio = args.oficio
    anioOficio = args.anio
else:
    numOficio = 25
    anioOficio = 20201

# Abrimos el menu acciones
acciones = Acciones(app, reporte)
irA(app, "Oficios->Acciones", reporte)

# Entramos a enviar emails individuales
acciones.desplegarMenu("Cartas")
acciones.abrirAccion("Emitir Carta Individual")

emitirCarta = EmitirCarta(app, reporte)

# Seleccionamos un oficio existente
emitirCarta.presionarSeleccionarOficio()
seleccion = SelectOficio(app, reporte)
seleccion.seleccionarOficioBase(numOficio,anioOficio)

# Seleccionamos el recorrido
emitirCarta.seleccionarRecorrido()

# Aceptamos la emisión de cartas
emitirCarta.aceptarEmisionCarta()

# Cerramos la ventana de Emitir Carta
emitirCarta.presionarSalir()

# Ingreso a consulta respuesta
newConsultaRespuesta = CR(app, reporte)
irA(app, "Oficios->Consulta respuestas", reporte)

# Selecciona un oficio
newConsultaRespuesta.seleccionarOficio(anioOficio, numOficio)

# Validamos el estado bloqueado del oficio
emitirCarta.obtenerEstadoBloqueado("NO")

# Presionamos en editar carta
emitirCarta.clickearEmitirCarta()

# Editamos la carta
emitirCarta.editarCarta()

# Cerramos la ventana de consulta de respuesta
newConsultaRespuesta.presionarSalir()

# Terminamos el reporte
reporte.terminarReporte()

# Cerramos la app
newInstance.closeAOJApp()
