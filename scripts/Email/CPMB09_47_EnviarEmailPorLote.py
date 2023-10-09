from modulos.AOJApp import AOJ
from modulos.Acciones import Acciones
from modulos.BarraMenu import irA
from modulos.EnviarEmail import EnviarEmailsPorLote
from modulos.Reporte import Reporte

#Iniciamos la aplicación
newInstance = AOJ()
app = newInstance.retornarAOJApp()

reporte = Reporte("Enviar email por lote","Envio de emails por lote")

# Abrimos el menu acciones
acciones = Acciones(app, reporte)
irA(app, "Oficios->Acciones", reporte)

# Entramos a enviar emails individuales
acciones.desplegarMenu("Emails")
acciones.abrirAccion("Enviar Email por Lote")

enviarEmailsPorLote = EnviarEmailsPorLote(app, reporte)

# Seleccionamos el tipo de oficio
enviarEmailsPorLote.seleccionarOficioParaEnviarMails(tipoOficio='SOJ')

# Aceptamos envío de emails
enviarEmailsPorLote.aceptarEnvioDeEmails()

# Cerramos la aplicación
newInstance.closeAOJApp()



