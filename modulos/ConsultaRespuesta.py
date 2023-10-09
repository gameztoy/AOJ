import time

import pyautogui
from pywinauto import Application, mouse
from modulos.SeleccionarOficio import SelectOficio


class CR:
    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def seleccionarOficio(self, anio, oficio):
        try:
            main_win = self.app.top_window()
            main_win.ConsultaRespuestasdelOficio.ThunderRT6PictureBoxDC2.Toolbar20WndClass5.msvb_lib_toolbar.click()
            seleccionar = SelectOficio(self.app, self.reporte)
            seleccionar.seleccionarOficioBase(oficio, anio)
        except:
            print("Hubo un error al seleccionar el oficio dentro de Consulta de Respuesta")

    def verOficio(self):
        main_win = self.app.top_window()

        main_win.ThunderRT6UserControlDC.ThunderRT6PictureBoxDC.Toolbar20WndClass2.msvb_lib_toolbar.click(coords=(10, 10))

        time.sleep(1)

        main_win.type_keys('{DOWN}')
        main_win.type_keys('{ENTER}')

    def verCuentaCorriente(self):
        main_win = self.app.top_window()
        try:

            main_win.ThunderRT6UserControlDC.ThunderRT6PictureBoxDC.Toolbar20WndClass2.msvb_lib_toolbar.click(coords=(10, 10))

            time.sleep(1)

            main_win.type_keys('{DOWN}')
            main_win.type_keys('{DOWN}')
            main_win.type_keys('{DOWN}')
            main_win.type_keys('{DOWN}')
            main_win.type_keys('{ENTER}')

            self.reporte.agregarLogInformativo("Se selecciona Consultas -> Cuenta Corriente.")

        except:
            print("Hubo un error al clickear en Cuenta Corriente.")

    def verDatosOficioCreado(self):
        main_win = self.app.top_window()

        #Maximizamos la ventana para sacar la screen
        time.sleep(1)
        mouse.click(coords=(1014, 184))

        self.reporte.agregarLogInformativoConScreen("Detalle del oficio creado", main_win.ShellDocObjectView.InternetExplorer_Server)
        time.sleep(1)
        #Cerramos la ventana maximizada
        mouse.click(coords=(1355, 30))

    def verDatosCuentaCorriente(self):
        main_win = self.app.top_window()
        # Maximizamos la ventana para sacar la screen
        mouse.click(coords=(1014, 184))
        time.sleep(1)
        self.reporte.agregarLogInformativoConScreen("Detalle de Cuenta Corriente", main_win)
        time.sleep(2)
        # Cerramos la ventana maximizada
        mouse.click(coords=(1355, 30))

    def irADetalle(self):
        try:
            main_win = self.app.top_window()
            main_win.TabStrip20WndClass.click()
            main_win.TabStrip20WndClass.type_keys('{RIGHT}')
        except:
            print("Hubo un error al ir a Detalle dentro de Consulta de Respuesta")

    def agregarArchivoAdjunto(self, file='prueba.doc'):
        main_win = self.app.top_window()
        try:
            # Selecciona el boton agregar
            main_win.Agregar.click()
            time.sleep(3)

            # Se abre la ventana Seleccionar Archivo
            dialogArchivo = self.app.window(title="Seleccionar Archivo")

            # Se ingresa el path absoluto del archivo a agregar
            dialogArchivo.Edit.type_keys(file, with_spaces=True)

            # Se selecciona abrir para agregar el archivo
            dialogArchivo.Abrir.click()
            time.sleep(3)

            print('El archivo: ' + file + ' ha sido agregado a los archivos adjuntos del oficio.')
            self.reporte.agregarLogInformativo("El archivo " + file + " ha sido agregado a los archivos adjuntos del oficio")
        except:
            print('No se pudo adjuntar archivo al oficio.')

    def eliminarAchivoAdjunto(self, index=0):
        main_win = self.app.top_window()
        try:
            # Selecciona el archivo en la lista de archivos adjuntos
            main_win.ThunderRT6Frame.child_window(class_name='ListView20WndClass').item(index).click()

            # Obtiene lista de las rutas de los archivos adjuntos
            list = main_win.ThunderRT6Frame.child_window(class_name='ListView20WndClass').texts()

            # Presiona el boton eliminar
            main_win.ThunderRT6Frame.Eliminar.click()

            # Acepta dialogo de confirmacion eliminacion
            main_win = self.app.top_window()
            main_win.Button.click()

            # Imprime por pantalla el nombre del archivo eliminado
            print(list[index + 1] + " ha sido eliminado de la lista de archivos adjuntos")
            self.reporte.agregarLogInformativo(list[index + 1] + " ha sido eliminado de la lista de archivos adjuntos")
        except:
            print("No hay archivo para el indice: " + str(index))

    def verAchivoAdjunto(self, index=0):
        main_win = self.app.top_window()
        try:
            # Selecciona el archivo adjunto
            main_win.ThunderRT6Frame.child_window(class_name='ListView20WndClass').item(index).click()

            list = main_win.ThunderRT6Frame.child_window(class_name='ListView20WndClass').texts()
            extension = list[index + 1].split('.')
            extension = extension[len(extension) - 1]
            main_win.ThunderRT6Frame.Ver.click()
            time.sleep(5)

            # Dependiendo el tipo de archivo se abren distintas aplicaciones
            # Se abren las aplicaciones y se cierran en cada if
            if extension == 'doc':
                # Se conecta a la aplicación
                wordApp = Application().connect(title_re=".*Word", class_name="OpusApp", timeout=10)
                print(wordApp.OpusApp.texts()[0] + " archivo abierto.")

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
                self.reporte.agregarLogInformativoConScreen(pdfApp.AcrobatSDIWindow.texts()[0] + " archivo abierto.", pdfApp.AcrobatSDIWindow)

                # Presiono 'Alt + a' y 's' para salir
                pdfApp.AcrobatSDIWindow.type_keys('%a')
                pdfApp.AcrobatSDIWindow.type_keys('s')

            if extension == 'jpg':
                # Elegimos como abrir la aplicación
                jpgApp = Application(backend="uia").connect(title='¿Cómo quieres abrir este archivo?')
                jpgApp.Dialog.Aceptar.click()

                # Se conecta a la aplicación
                jpgApp = Application(backend="uia").connect(title_re='.* Visualizador de fotos de Windows',
                                                            class_name='Photo_Lightweight_Viewer', timeout=10)
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
            print("No hay archivo para el indice: " + str(index))

    def buscarArchivoAdjunto(self, file='\\\\sfs-1\\testing\\Automatizacion_de_Proyectos\\_AOJ\\Repositorio_archivos\\word.doc'):
        main_win = self.app.top_window()

        try:
            # Obtiene la lista de los archivos adjuntos
            list = main_win.ThunderRT6Frame.child_window(class_name='ListView20WndClass').texts()

            # Si el archivo esta en la lista obtenina
            if file in list:
                print('El archivo: ' + file + ' Es un archivo adjunto del oficio manual')
                self.reporte.agregarLogCondicional("Se encontro el archivo en el oficio", True)
                # Devuelve la posion en la que esta el archivo
                return list.index(file) - 1
            else:
                print('No se encontro: ' + file + ' en los archivos adjuntos del oficio manual')
                self.reporte.agregarLogCondicional("Se encontro el archivo en el oficio", False)
                return None
        except:
            print("Hubo un error al buscar el archivo adjunto: ", file)

    def presionarSalir(self):
        main_win = self.app.top_window()

        # Presiono el boton salir
        main_win.ConsultaRespuestasdelOficio.ThunderRT6PictureBoxDC2.msvb_lib_toolbar.click(coords=(10, 10))
        self.reporte.agregarLogInformativo( "Se presiona Salir de Consulta Respuestas.")

    def presionarBloqDbloq(self):
        try:
            main_win = self.app.top_window()
            main_win.ConsultaRespuestasdelOficio.ThunderRT6PictureBoxDC2.Toolbar20WndClass6.msvb_lib_toolbar.click(coords=(10, 10))
            self.reporte.agregarLogInformativo("Se clickea sobre el boton Bloq/Dbloq Rta")
            time.sleep(2)
        except:
            print("Hubo un error al clickear sobre el boton Bloq/Dbloq Rta")

    def presionarActualizar(self):
        try:
            main_win = self.app.top_window()
            main_win.ConsultaRespuestasdelOficio.ThunderRT6PictureBoxDC2.Toolbar20WndClass4.msvb_lib_toolbar.click(coords=(10, 10))
            self.reporte.agregarLogInformativo("Se clickea sobre el boton Actualizar")
        except:
            print("Hubo un error al clickear sobre el boton Actualizar")
