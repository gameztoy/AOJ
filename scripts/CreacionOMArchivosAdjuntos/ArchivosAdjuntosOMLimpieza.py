from modulos.AOJApp import AOJ
from modulos.BarraMenu import irA
from modulos.OficioManual import OMModificar
from modulos.Reporte import Reporte

newInstance = AOJ()
app = newInstance.retornarAOJApp()

reporte = Reporte("Archivos Adjuntos OM Limpieza","Se borran los archivos de un oficio manual")

#INGRESO A OFICIO MANUAL [MODIFICAR]
irA(app,'Novedades -> Oficios Manuales', reporte)
newOMM = OMModificar(app, reporte)

#ELIMINO EL ARCHIVO .DOC
newOMM.modificarOM('179', '2020')
newOMM.irArchivoAdjunto()
newOMM.eliminarTodoslosAdjuntos()
newOMM.finalizarModificaciones()

#CIERRO APP
newInstance.closeAOJApp()
