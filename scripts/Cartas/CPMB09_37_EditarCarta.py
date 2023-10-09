import argparse
import time
from modulos.AOJApp import AOJ
from modulos.Acciones import Acciones
from modulos.BarraMenu import irA
from modulos.Cartas import EmitirCarta, BlanquearCarta
from modulos.ConsultaRespuesta import CR
from modulos.Reporte import Reporte

newInstance = AOJ()
app = newInstance.retornarAOJApp()

reporte = Reporte("Smoke de Cartas", "CPMB09.37 Editar Carta")

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

# Ingreso a consulta respuesta
newConsultaRespuesta = CR(app, reporte)
irA(app, "Oficios->Consulta respuestas", reporte)

# Selecciona un oficio
newConsultaRespuesta.seleccionarOficio(anioOficio, numOficio)

# Clickea sobre el boton Bloq/DbloqRta
newConsultaRespuesta.presionarBloqDbloq()
time.sleep(1)

# Clickea sobre el boton Actualizar
newConsultaRespuesta.presionarActualizar()
time.sleep(1)

# Validamos el estado bloqueado del oficio
emitirCarta = EmitirCarta(app, reporte)
time.sleep(1)
emitirCarta.obtenerEstadoBloqueado("NO")
time.sleep(1)

# Presionamos en editar carta
emitirCarta.clickearEmitirCarta()

# Editamos la carta
editarCartaBlanqueada = BlanquearCarta(app,reporte)
editarCartaBlanqueada.editarCartaBlanqueada()

# Terminamos el reporte
reporte.terminarReporte()

# Cerramos la ventana de consulta de respuesta
newConsultaRespuesta.presionarSalir()
# Cerramos la app
newInstance.closeAOJApp()