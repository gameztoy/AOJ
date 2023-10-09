from modulos.AOJApp import AOJ
from modulos.BarraMenu import irA
from modulos.ConsultaRespuesta import CR
from modulos.Reporte import Reporte

'''
Realiza las operaciones sobre los archivos adjuntos en Consulta Respuesta
 Consultar Respuesta del Oficio_OM adjuntar Archivo 
 Consultar Respuesta del Oficio_OM Eliminar Archivo 
 Consultar Respuesta del Oficio_OM Ver Archivo
'''

newInstance = AOJ()
app = newInstance.retornarAOJApp()
reporte = Reporte("Consulta de respuesta","Consulta de un oficio manual")

#ingreso a consulta respuesta
newConsultaRespuesta = CR(app, reporte)

irA(app, "Oficios->Consulta respuestas", reporte)

#selecciona un oficio
newConsultaRespuesta.seleccionarOficio(2020, 179)

reporte.terminarReporte()
newInstance.closeAOJApp()