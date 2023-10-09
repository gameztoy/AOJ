import argparse
from modulos.AOJApp import AOJ
from modulos.Reporte import Reporte
from modulos.BarraMenu import irA
from modulos.ConsultaRespuesta import CR
from modulos.OficioManual import OficioManual
from modulos.OficioManualPartes import OficioManualPartes
from modulos.Acciones import Acciones
from modulos.RespuestaDeEmbargos import RegistrarRespuestasEmbargos
from modulos.ValidacionEmbargoCrecer import CuentasPorCuentas
from modulos.EnviarEmail import EnviarEmail
from modulos.SeleccionarOficio import SelectOficio
from modulos.Cartas import EmitirCarta, ConfirmarCarta
from modulos.RegistrarTransferenciaFondos import RegistrarTransferenciaFondos
from modulos.ConfirmarTransferenciaFondos import ConfirmarTransferenciaFondos

"""
Creacion de Embargo por Transferencia MEP - CPMB09_26_01 y 02
Datos obtenidos desde .xlsx:
- Caso: nombre de caso - caso
- Cuit: cuit persona - cuit
- Importe RRRE: importe de registrar respuesta de embargo - imprrre
- Bloqueo: con o sin - bloqueo
- Importe MEP: importe de transferencia MEP - impmep
- Moneda: Pesos o Dolar - moneda
- Capital: Monto a definir - capital
- Interes: Monto a definir - interes
"""

# Inciar parametros del script
parser = argparse.ArgumentParser()
parser.add_argument("-n", "--caso", nargs='+')
parser.add_argument("-p", "--cuit", nargs='+')
parser.add_argument("-i", "--imprrre", nargs='+')
parser.add_argument("-b", "--bloqueo", nargs='+')
parser.add_argument("-m", "--impmep", nargs='+')
parser.add_argument("-a", "--moneda", nargs='+')
parser.add_argument("-e", "--capital", nargs='+')
parser.add_argument("-s", "--interes", nargs='+')
args = parser.parse_args()
caso = args.caso[0]
cuit = args.cuit[0]
imprrre = args.imprrre[0]
bloqueo = args.bloqueo[0]
impmep = args.impmep[0]
moneda = args.moneda[0]
capital = args.capital[0]
interes = args.interes[0]

"""Comienza la ejecucion del caso"""
reporte = Reporte("Ejecucion Transferencia MEP", "CPMB26_" + caso)
newInstance = AOJ()
app = newInstance.retornarAOJApp()

"""Creación OM Embargo General con Vigencia Actuales y Futuras - Pesos sin transferencia"""
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
# Ingresamos la denominacion
oficioManualPartes.ingresarDenominacion(cuit)
# Realizamos el cruce individual
oficioManualPartes.realizarCruceIndividual()
# Seleccionamos el tipo de oficio
oficioManualPartes.seleccionarTipoOficio('Embargos General de Fondos y Valores')
# Seleccionamos la vigencia
oficioManualPartes.seleccionarVigencia('Actuales y Futuras')
# Seleccionar transferencia
oficioManualPartes.seleccionarTransferencia('No')
# Seleccionamos la moneda
oficioManualPartes.seleccionarMoneda(moneda)
oficioManualPartes.ingresarCapitalYIntereses(5000, 100)
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
cr.presionarSalir()

""" Retención por registrar respuesta de embargo """
#Abrimos la ventana de Acciones
irA(app, "Oficio -> Acciones", reporte)
acciones = Acciones(app, reporte)
acciones.desplegarMenu("Respuestas")
acciones.abrirAccion("Registrar Respuestas de Embargos")
#Se abre la ventana 'Registrar Respuestas de Embargos'
registrarRE = RegistrarRespuestasEmbargos(app, reporte)
#Se selecciona el oficio con el años correspondiente
registrarRE.seleccionarOficio(numero, anio)
#Se agregan las cuentas a embargar
registrarRE.agregarCuenta(moneda, 1, imprrre, cuit)
#Presiona aceptar y avanzar
registrarRE.presionarAceptarYAvanzar()
#Confirma las cuentas a embargar
registrarRE.aceptarSeleccionarCuentas()
#Salir de Resgistrar Respuesta de Embargo
registrarRE.presionarSalir()
#Se cierra AOJ, sino no es posible acceder a la siguiente accion
newInstance.closeAOJApp()

""" Validacion en Crecer"""
#Validacion en Crecer - Cuentas por Cuenta - Nro Cuenta
cxc = CuentasPorCuentas(app, reporte)
if bloqueo == 'Sin bloqueo':
    cxc.buscarCuentasPorCuentaSBloqueo()
elif bloqueo == 'Bloqueo parcial':
    cxc.buscarCuentasPorCuentaSBloqueo()

"""Enviar Email Individual """
# Abrimos la app nuevamente
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
seleccion.seleccionarOficioBase(numero, anio)
#Aceptamos el envio
enviarEmail.aceptarEnvio()
#Salir de Enviar Email
enviarEmail.presionarSalir()
#Se cierra AOJ, sino no es posible acceder a la siguiente accion
newInstance.closeAOJApp()

"""Emitir Carta"""
# Abrimos la app nuevamente
newInstance = AOJ()
app = newInstance.retornarAOJApp()
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
seleccion.seleccionarOficioBase(numero, anio)
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
newConsultaRespuesta.seleccionarOficio(anio, numero)
# Validamos el estado bloqueado del oficio
emitirCarta.obtenerEstadoBloqueado("NO")
# Presionamos en editar carta
emitirCarta.clickearEmitirCarta()
# Editamos la carta
emitirCarta.editarCarta()
# Cerramos la ventana de consulta de respuesta
newConsultaRespuesta.presionarSalir()
#Se cierra AOJ, sino no es posible acceder a la siguiente accion
newInstance.closeAOJApp()

"""Confirmar Carta"""
# Abrimos la app nuevamente
newInstance = AOJ()
app = newInstance.retornarAOJApp()
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
seleccion.seleccionarOficioBase(numero, anio)
# Presionamos aceptar para Confirmar la carta
confirmarCarta.presionarAceptar()
# Presionamos salir de Confirmar Carta
confirmarCarta.presionarSalir()
# Vamos a Consulta Respuestas
irA(app, "Oficios->Consulta respuestas", reporte)
# Selecciona el oficio
newConsultaRespuesta = CR(app, reporte)
newConsultaRespuesta.seleccionarOficio(anio, numero)
# Validamos el estado bloqueado del oficio
emitirCarta = EmitirCarta(app, reporte)
emitirCarta.obtenerEstadoBloqueado("Si")
# Presionamos en editar carta
emitirCarta.clickearEmitirCarta()
# Verificamos que el archivo esta en modo solo lectura
emitirCarta.verificarSoloLectura(True)
# Cerramos la ventana de consulta de respuesta
newConsultaRespuesta.presionarSalir()
#Se cierra AOJ, sino no es posible acceder a la siguiente accion
newInstance.closeAOJApp()

"""" Creacion de OM Embargo de Transferencia MEP """
# Abrimos la app nuevamente
newInstance = AOJ()
app = newInstance.retornarAOJApp()
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
# Ingresamos la denominación
oficioManualPartes.ingresarDenominacion(cuit)
# Realizamos el cruce individual
oficioManualPartes.realizarCruceIndividual()
# Seleccionamos el tipo de oficio
oficioManualPartes.seleccionarTipoOficio('Transferencia de Fondos Embargados a la Orden del Juzgado')
# Enlazamos un oficio del tipo: Embargos General de Fondos y  Valores
oficioManualPartes.enlazarOficio(numero, anio)
# Seleccionamos la vigencia
oficioManualPartes.seleccionarVigencia('Actuales')
# Seleccionar transferencia
oficioManualPartes.seleccionarTransferencia('Si')
# Seleccionamos la moneda
oficioManualPartes.seleccionarMoneda(moneda)
# Ingresamos capital e intereses - EL MONTO DEBE SER IGUAL O INFERIOR AL MONTO EMBARGADO
oficioManualPartes.ingresarCapitalYIntereses(capital, interes)
# Presionamos aceptar para agregar las partes
oficioManualPartes.presionarAceptar()
# Presionamos aceptar para crear el oficio manual
numero = oficioManual.aceptarOficioManual()
oficioManual.presionarSalir()
# Obtenemos el OM Transferencia
anio = numero.split(",")[1]
numero = numero.split(",")[0]
# Vamos a Consulta Respuesta a ver el oficio creado
irA(app, "Oficios -> Consulta Respuestas", reporte)
cr = CR(app, reporte)
cr.seleccionarOficio(anio, numero)
cr.verOficio()
cr.verDatosOficioCreado()
# Cerramos la aplicación
newInstance.closeAOJApp()

"""" Tratamiento de OM Embargo de Transferencia MEP """
# Abrimos la app nuevamente
newInstance = AOJ()
app = newInstance.retornarAOJApp()
# Abrimos el menu acciones
acciones = Acciones(app, reporte)
irA(app, "Oficios->Acciones", reporte)
# Entramos a Transferencias - Registrar Transferencia de Fondos
acciones.desplegarMenu("Transferencias")
acciones.abrirAccion("Registrar Transferencia de Fondos")
# Se abre la ventana Registrar Transferencia de Fondos
registrarTF = RegistrarTransferenciaFondos(app, reporte)
# Se selecciona el oficio de Transferencia
registrarTF.seleccionarOficio(numero, anio)
# Se agregan la cuenta a transferir
registrarTF.agregarCuenta(impmep)
# Presiona aceptar y avanzar
registrarTF.presionarAceptarYAvanzar()
# Confirma las cuentas a transferir
registrarTF.presionarSiMovimientosNoConfirmados()
# Salir de Registrar Transferencia de Fondos
registrarTF.presionarSalir()
# Cerramos la aplicación
newInstance.closeAOJApp()

""" Transferencias - Confirmar Transferencia de Fondos """
# Abrimos la app nuevamente
newInstance = AOJ()
app = newInstance.retornarAOJApp()
# Abrimos el menu acciones
acciones = Acciones(app, reporte)
irA(app, "Oficios->Acciones", reporte)
# Entramos a Transferencias - Confirmar Transferencia de Fondos
acciones.desplegarMenu("Transferencias")
acciones.abrirAccion("Confirmar Transferencia de Fondos")
# Se abre la ventana Registrar Transferencia de Fondos
confirmarTF = ConfirmarTransferenciaFondos(app, reporte)
# Se selecciona el oficio de Confirmacion de Transferencia
confirmarTF.seleccionarOficio(numero, anio)
# Se completa el campo Nro MEP
confirmarTF.completarNroMEP(numero)
# Aceptar la Transferencia de Fondos
confirmarTF.aceptarConfirmarTransferencia()
# Finalizar Transferencia de Fondos
confirmarTF.finalizarOperacion()
# Salir de Registrar Transferencia de Fondos
confirmarTF.presionarSalir()
# Cerramos la aplicación
newInstance.closeAOJApp()

"""Emitir Carta"""
# Abrimos la app nuevamente
newInstance = AOJ()
app = newInstance.retornarAOJApp()
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
seleccion.seleccionarOficioBase(numero, anio)
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
newConsultaRespuesta.seleccionarOficio(anio, numero)
# Validamos el estado bloqueado del oficio
emitirCarta.obtenerEstadoBloqueado("NO")
# Presionamos en editar carta
emitirCarta.clickearEmitirCarta()
# Editamos la carta
emitirCarta.editarCarta()
# Cerramos la ventana de consulta de respuesta
newConsultaRespuesta.presionarSalir()
#Se cierra AOJ, sino no es posible acceder a la siguiente accion
newInstance.closeAOJApp()

"""Confirmar Carta"""
# Abrimos la app nuevamente
newInstance = AOJ()
app = newInstance.retornarAOJApp()
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
seleccion.seleccionarOficioBase(numero, anio)
# Presionamos aceptar para Confirmar la carta
confirmarCarta.presionarAceptar()
# Presionamos salir de Confirmar Carta
confirmarCarta.presionarSalir()
# Vamos a Consulta Respuestas
irA(app, "Oficios->Consulta respuestas", reporte)
# Selecciona el oficio
newConsultaRespuesta = CR(app, reporte)
newConsultaRespuesta.seleccionarOficio(anio, numero)
# Validamos el estado bloqueado del oficio
emitirCarta = EmitirCarta(app, reporte)
emitirCarta.obtenerEstadoBloqueado("Si")
# Presionamos en editar carta
emitirCarta.clickearEmitirCarta()
# Verificamos que el archivo esta en modo solo lectura
emitirCarta.verificarSoloLectura(True)
# Cerramos la ventana de consulta de respuesta
newConsultaRespuesta.presionarSalir()

""" Consulta de respuesta del oficio de Transferencia - Cuenta Corriente"""
irA(app, "Oficios -> Consulta Respuestas", reporte)
cr = CR(app, reporte)
cr.seleccionarOficio(anio, numero)
cr.verCuentaCorriente()
cr.verDatosCuentaCorriente()
cr.presionarSalir()
# Cerramos la aplicación
newInstance.closeAOJApp()

""" Validacion en Crecer"""
#Validacion en Crecer - Cuentas por Cuenta - Nro Cuenta
cxc = CuentasPorCuentas(app, reporte)
if bloqueo == 'Sin bloqueo':
    cxc.buscarCuentasPorCuentaSBloqueo()
elif bloqueo == 'Bloqueo parcial':
    cxc.buscarCuentasPorCuentaSBloqueo()

#Terminamos el reporte
reporte.terminarReporte()
