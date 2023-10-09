import time
from modulos.AOJApp import AOJ
from modulos.BarraMenu import irA
from modulos.OficioManual import agregarArchivoAdjunto, OMModificar
from modulos.Reporte import Reporte

newInstance = AOJ()
app = newInstance.retornarAOJApp()

reporte = Reporte("Archivos Adjuntos","Se modifica un oficio manual y se agrega/lee/elimina archivos")

oficio='179'
anio = '2020'

#INGRESO A OFICIO MANUAL [MODIFICAR]
irA(app,'Novedades -> Oficios Manuales', reporte)
newOMM = OMModificar(app, reporte)
newOMM.modificarOM(oficio, anio)

#ME MUEVO A LA PESTAÑA ARCHIVOS ADJUNTOS
newOMM.irArchivoAdjunto()

#AGREGO ARCHIVOS ADJUNTOS
agregarArchivoAdjunto(app, '\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\word.doc', archivoReporte)
agregarArchivoAdjunto(app, '\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\imagen.jpg', archivoReporte)
agregarArchivoAdjunto(app, '\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\pdf.pdf', archivoReporte)
agregarArchivoAdjunto(app, '\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\excel.xls', archivoReporte)

#VEO LOS ARCHIVOS AGREGADOS
index = newOMM.buscarArchivoAdjunto('\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\imagen.jpg')
newOMM.verAchivoAdjunto(index)
index = newOMM.buscarArchivoAdjunto('\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\word.doc')
newOMM.verAchivoAdjunto(index)
index = newOMM.buscarArchivoAdjunto('\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\pdf.pdf')
newOMM.verAchivoAdjunto(index)
index = newOMM.buscarArchivoAdjunto('\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\excel.xls')
newOMM.verAchivoAdjunto(index)
time.sleep(4)

#CONFIRMO LOS CAMBIOS EN EL OFICIO
newOMM.finalizarModificaciones()

#ELIMINO EL ARCHIVO .DOC
irA(app,'Novedades -> Oficios Manuales', reporte)
newOMM = OMModificar(app, reporte)
newOMM.modificarOM(oficio, anio)
newOMM.irArchivoAdjunto()
index = newOMM.buscarArchivoAdjunto('\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\word.doc')
newOMM.eliminarAchivoAdjunto(index)
newOMM.finalizarModificaciones()

#COMPRUEBO QUE SE ELIMINO EL ARCHIVO .DOC
#INGRESO A OFICIO MANUAL [MODIFICAR]
irA(app,'Novedades -> Oficios Manuales', reporte)
newOMM = OMModificar(app, reporte)
newOMM.modificarOM(oficio, '2020')
newOMM.irArchivoAdjunto()
index = newOMM.buscarArchivoAdjunto('\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\word.doc')

#CONFIRMO MODIFICACIONES
newOMM.finalizarModificaciones()

#CIERRO APP
newInstance.closeAOJApp()
