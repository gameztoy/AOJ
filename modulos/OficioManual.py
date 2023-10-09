import time
from pywinauto import Application
from modulos.SeleccionarOficio import SelectOficio


class OficioManual:

    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def presionarAgregarOficioManual(self):
        main_win = self.app.top_window()

        # Click en el boton 'Agregar'
        main_win.ThunderRT6FormDC.ThunderRT6UserControlDC10.ThunderRT6PictureBoxDC2.msvb_lib_toolbar4.click()
        time.sleep(1)

        self.reporte.agregarLogInformativo("Presionamos Agregar Oficio Manual")

    def cargarDatosGenerales(self, exped):
        main_win = self.app.top_window()
        time.sleep(1)

        # Ingresa un valor en el input 'Exped.'
        main_win.ThunderRT6FormDC.ThunderRT6PictureBoxDC2.ThunderRT6UserControlDC3.Edit.type_keys(exped)

        # Campo desplegable 'Juzgado', selecciona la primera
        juzgado = '1º CIRCUNSCRIPCIÓN JUDICIAL 2º JUZGADO DE PROCESOS CONCURSALES Y DE REGISTROS DE MENDOZA'
        main_win.ThunderRT6FormDC.ThunderRT6PictureBoxDC2.ThunderRT6UserControlDC.ThunderRT6ComboBox.select(juzgado)
        time.sleep(1)

        self.reporte.agregarLogInformativo("Ingresamos juzgado: " + juzgado)

        # Campo 'Caratula'
        caratula = 'Juzgado de Primera Instancia en lo Contencioso Administrativo y Tributario y de Relaciones de Consumo de la Ciudad Autónoma de Buenos Aires nº 15 - Secretaria 29'
        main_win.ThunderRT6FormDC.ThunderRT6PictureBoxDC2.ThunderRT6UserControlDC5.Edit.type_keys(caratula, with_spaces=True)
        time.sleep(1)

        self.reporte.agregarLogInformativo("Ingresamos caratula: " + caratula)

    def presionarAgregarParte(self):
        main_win = self.app.top_window()

        # Click en Agregar parte
        main_win.ThunderRT6FormDC.ThunderRT6UserControlDC8.ThunderRT6PictureBoxDC.msvb_lib_toolbar.click()
        time.sleep(1)

        self.reporte.agregarLogInformativo("Presionamos Agregar Parte")

    def aceptarOficioManual(self):
        main_win = self.app.top_window()

        # Boton Aceptar en 'Oficio Manual [Agregar]'
        main_win.ThunderRT6FormDCwindow.ThunderRT6UserControlDC10.msvb_lib_toolbar2.click()
        time.sleep(1)

        # Confirmar dialogo de operacion: Confirma la operacion: si-no
        try:
            dlg_handles = self.app.window(title="Oficio Manual [Agregar]")
            if dlg_handles:
                time.sleep(5)
                dlg_handles.Button.click()
        except:
            print("")

        time.sleep(3)
        numero = ""
        anio = ""

        # Aceptar dialogo de oficio creado: El Nro de Oficio asignado es xxx para el Año xxxx
        try:
            dlg_handles = self.app.window(class_name="#32770")
            if dlg_handles:

                # Agregamos una imagen del dialogo al reporte
                self.reporte.agregarLogInformativoConScreen("Oficio Manual Creado", dlg_handles)

                print(dlg_handles.Static2.texts())
                time.sleep(3)
                numero = (dlg_handles.Static2.texts()[0]).split(' ')[7]
                anio = (dlg_handles.Static2.texts()[0]).split(' ')[11]
                time.sleep(3)
                dlg_handles.Button.click()
        except:
            print("")

        numeroAnio = str(numero) + "," + str(anio)
        print(numeroAnio)
        print("--Oficio Manual Creado--")
        self.reporte.agregarLogInformativo("Se genero el oficio " + numero + " en el año " + anio)

        # Vuelva a la pantalla 'Oficio Manual [Agregar]', donde los campos estan limpios
        return numeroAnio

    def presionarSalir(self):
        main_win = self.app.top_window()

        # Click en el boton 'Agregar'
        main_win.OficioManual.ThunderRT6UserControlDC10.ThunderRT6PictureBoxDC2.msvb_lib_toolbar.click(coords=(10,10))
        time.sleep(1)

class OMModificar:
    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def modificarOM(self, oficio, anio):
        """Selecciona el boton Modificar e ingresa el oficio a modificar"""
        main_win = self.app.top_window()

        # Click en el boton 'Modificar'
        main_win.OficioManual.ThunderRT6UserControlDC10.ThunderRT6PictureBoxDC2.msvb_lib_toolbar3.click(coords=(10, 10))
        time.sleep(2)

        # Ingresa el oficio a modificar
        newSO = SelectOficio(self.app, self.reporte)
        newSO.seleccionarOficio(oficio, anio)

        self.reporte.agregarLogInformativo("Se selecciona el oficio " + oficio + " del año " + anio + " para modificar")

    def irArchivoAdjunto(self):
        """Accede a la pestaña Archivos Adjuntos de la ventana Oficio Manual [Modificar]"""
        main_win = self.app.top_window()

        # Nos movemos dos veces hacia a la derecha presionando las flechas
        main_win.SSTabCtlWndClass.click()
        main_win.SSTabCtlWndClass.type_keys("{RIGHT}")
        main_win.SSTabCtlWndClass.type_keys("{RIGHT}")

    def buscarArchivoAdjunto(self, file):
        main_win = self.app.top_window()

        # Obtiene la lista de los archivos adjuntos
        list = main_win.SSTabCtlWndClass.child_window(class_name='ListView20WndClass').texts()

        # Si el archivo esta en la lista obtenina
        if file in list:
            print('El archivo: ' + file + ' Es un archivo adjunto del oficio manual')
            self.reporte.agregarLogCondicional("El archivo se encuentra en el oficio", True)
            # Devuelve la posion en la que esta el archivo
            return list.index(file)-1
        else:
            print('No se encontro: ' + file + ' en los archivos adjuntos del oficio manual')
            self.reporte.agregarLogCondicional("El archivo se encuentra en el oficio", False)
            return None

    def eliminarAchivoAdjunto(self, index=0):
        main_win = self.app.top_window()

        try:
            # Selecciona el archivo en la lista de archivos adjuntos
            main_win.SSTabCtlWndClass.child_window(class_name='ListView20WndClass').item(index).click()

            # Obtiene lista de las rutas de los archivos adjuntos
            list = main_win.SSTabCtlWndClass.child_window(class_name='ListView20WndClass').texts()

            # Presiona el boton eliminar
            main_win.SSTabCtlWndClass.Eliminar.click()

            # Acepta dialogo de confirmacion eliminacion
            main_win = self.app.top_window()
            main_win.Button.click()

            # Imprime por pantalla el nombre del archivo eliminado
            print(list[index+1]+" ha sido eliminado de la lista de archivos adjuntos")
            self.reporte.agregarLogInformativo(list[index+1]+" ha sido eliminado de la lista de archivos adjuntos")
        except:
            print("No hay archivo para el indice: " + str(index))

    def verAchivoAdjunto(self, index=0):
        main_win = self.app.top_window()

        try:
            #Selecciona el archivo adjunto
            main_win.SSTabCtlWndClass.child_window(class_name='ListView20WndClass').item(index).click()

            list = main_win.SSTabCtlWndClass.child_window(class_name='ListView20WndClass').texts()
            extension = list[index+1].split('.')
            extension = extension[len(extension)-1]
            main_win.SSTabCtlWndClass.Ver.click()
            time.sleep(5)

            # Dependiendo el tipo de archivo se abren distintas aplicaciones
            # Se abren las aplicaciones y se cierran en cada if

            if extension == 'doc':
                # Se conecta a la aplicación
                wordApp = Application().connect(title_re=".*Word", class_name="OpusApp", timeout=10)
                print(wordApp.OpusApp.texts()[0]+" archivo abierto.")

                time.sleep(3)
                self.reporte.agregarLogInformativoConScreen(wordApp.OpusApp.texts()[0] +
                                                            " archivo abierto.", wordApp.OpusApp)

                # Presiono 'Alt + a' y 's' para salir
                wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('%a')
                wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('s')

            if extension == 'pdf':
                # Se conecta a la aplicación
                pdfApp = Application().connect(title_re=".*pdf*.", class_name="AcrobatSDIWindow", timeout=10)
                print(pdfApp.AcrobatSDIWindow.texts()[0] + " archivo abierto.")

                time.sleep(3)
                self.reporte.agregarLogInformativoConScreen(pdfApp.AcrobatSDIWindow.texts()[0] +
                                                            " archivo abierto.", pdfApp.AcrobatSDIWindow)

                # Presiono 'Alt + a' y 's' para salir
                pdfApp.AcrobatSDIWindow.type_keys('%a')
                pdfApp.AcrobatSDIWindow.type_keys('s')

            if extension == 'jpg':
                # Elegimos como abrir la aplicación
                jpgApp = Application(backend="uia").connect(title='¿Cómo quieres abrir este archivo?')
                jpgApp.Dialog.Aceptar.click()

                # Se conecta a la aplicación
                jpgApp = Application(backend="uia").connect(title_re='.* Visualizador de fotos de Windows',class_name='Photo_Lightweight_Viewer',timeout=10)
                print(jpgApp.Dialog.texts()[0] + " archivo abierto.")

                time.sleep(3)
                self.reporte.agregarLogInformativoConScreen(jpgApp.Dialog.texts()[0] +
                                                            " archivo abierto.", jpgApp.Dialog)
                # Presiono 'Alt + h' y 's' para salir
                jpgApp.Dialog.type_keys('%h')
                jpgApp.Dialog.type_keys('s')

            if extension == 'xls':
                # Se conecta a la aplicación
                xlsApp = Application(backend="uia").connect(title_re='.* .*xls*.', class_name='XLMAIN', timeout=10)
                print(xlsApp.Dialog.texts()[0] + " archivo abierto.")

                time.sleep(3)
                self.reporte.agregarLogInformativoConScreen(xlsApp.Dialog.texts()[0] +
                                                            " archivo abierto.", xlsApp.Dialog)

                # Presiono 'Alt + a' y 's' para salir
                xlsApp.Dialog.type_keys('%a')
                xlsApp.Dialog.type_keys('s')
        except:
            print("No hay archivo para el indice: "+str(index))


    def finalizarModificaciones(self):
        """Presiona aceptar de la ventana Oficio Manual Modificar y acepta los cambios"""
        main_win = self.app.top_window()

        main_win.OficioManual.ThunderRT6UserControlDC10.ThunderRT6PictureBoxDC1.msvb_lib_toolbar2.click()

        try:
          dlg_handles = self.app.window(title="Oficio Manual [Modificar]")
          time.sleep(2)
          dlg_handles.Sí.click()
        except:
          print("Hubo un error al confirmar la operacion Modificar")

        main_win = self.app.top_window()
        main_win.OficioManual.ThunderRT6UserControlDC10.ThunderRT6PictureBoxDC2.msvb_lib_toolbar1.click(coords=(10, 10))
        time.sleep(2)

        self.reporte.agregarLogInformativo("Se finalizan las modificaciones en el oficio")

    def eliminarTodoslosAdjuntos(self):
        main_win = self.app.top_window()

        try:
            # Obtengo la cantidad de la lista
            items=len(main_win.SSTabCtlWndClass.child_window(class_name='ListView20WndClass').texts())-1
            print(items)

            for i in range(items):
                main_win = self.app.top_window()
                # Selecciona el archivo en la lista de archivos adjuntos
                main_win.SSTabCtlWndClass.child_window(class_name='ListView20WndClass').item(0).click()

                # Obtiene lista de las rutas de los archivos adjuntos
                list = main_win.SSTabCtlWndClass.child_window(class_name='ListView20WndClass').texts()

                # Presiona el boton eliminar
                main_win.SSTabCtlWndClass.Eliminar.click()

                # Acepta dialogo de confirmacion eliminacion
                main_win = self.app.top_window()
                main_win.Button.click()

                # Imprime por pantalla el nombre del archivo eliminado
                print(list[1] + " ha sido eliminado de la lista de archivos adjuntos")
                self.reporte.agregarLogInformativo(list[1] + " ha sido eliminado de la lista de archivos adjuntos")
        except:
            print('No hay archivos para eliminar')


def agregarArchivoAdjunto(app,file, reporte):
    try:
        main_win= app.top_window()

        # Selecciona el boton agregar
        main_win.Agregar.click()
        time.sleep(3)

        # Se abre la ventana Seleccionar Archivo
        dialogArchivo = app.window(title="Seleccionar Archivo")

        # Se ingresa el path absoluto del archivo a agregar
        dialogArchivo.Edit.type_keys(file,with_spaces=True)

        # Se selecciona abrir para agregar el archivo
        dialogArchivo.Abrir.click()
        time.sleep(3)

        print('El archivo: '+file+' ha sido agregado a los archivos adjuntos del oficio.')
        reporte.agregarLogInformativo('El archivo: '+file+' ha sido agregado a los archivos adjuntos del oficio.')
    except:
        print('No se pudo adjuntar archivo al oficio.')