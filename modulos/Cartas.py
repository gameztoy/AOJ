import os
import time
import cv2
import pyautogui
import pywinauto.mouse as mouse
from pywinauto import Application
from modulos.ReconocimientoTexto import ExtraccionTexto

class EmitirCarta:
    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def presionarSeleccionarOficio(self):
        main_win = self.app.top_window()

        # Presiona seleccionar oficio
        main_win.EmitirCarta.ThunderRT6UserControlDC3.msvb_lib_toolbar.click()

    def seleccionarRecorrido(self):
        main_win = self.app.top_window()

        # Selecciono un recorrido
        main_win.EmitirCarta.ThunderRT6PictureBoxDC.Respuesta.ThunderRT6UserControlDC.ThunderRT6ComboBox.select(2)

        recorrido = main_win.EmitirCarta.ThunderRT6PictureBoxDC.Respuesta.ThunderRT6UserControlDC.ThunderRT6ComboBox.texts()[0]
        self.reporte.agregarLogInformativo("Se selecciono el recorrido: " + recorrido)

    def aceptarEmisionCarta(self):
        main_win = self.app.top_window()

        # Presiona aceptar
        main_win.EmitirCarta.ThunderRT6UserControlDC3.msvb_lib_toolbar2.click(coords=(10, 10))

        # Tomo la ventana "Emitir Cartas [Opciones]"
        main_win = self.app.top_window()

        # Presiono aceptar en el dialog de "Emitir Cartas [Opciones]"
        main_win.ThunderRT6FormDC.ThunderRT6PictureBoxDC3.Aceptar.click()

        time.sleep(15)

        # Cancelamos que guarde la impresión
        guardar = Application(backend="win32").connect(title="Guardar impresión como")
        guardar.Dialog.Cancelar.click()

        time.sleep(5)

        # Aceptamos el dialog de aviso
        dialogEmitir = self.app.window(title="Emitir Carta")
        self.reporte.agregarLogInformativoConScreen("Emisión de carta aceptada", dialogEmitir)
        dialogEmitir.Aceptar.click()

    def presionarSalir(self):
        main_win = self.app.top_window()

        # Presiona salir
        main_win.EmitirCarta.ThunderRT6UserControlDC3.msvb_lib_toolbar6.click(coords=(10, 10))

    def clickearEmitirCarta(self):

        path = os.getcwd()
        path = path.split('automatizacionaoj')[0] + 'automatizacionaoj\\recursos\\emitirCarta.png'

        # Localizo en la pantalla la linea "5- Emitir Carta"
        dim = pyautogui.locateOnScreen(path)

        print(dim)

        # Obtengo las coordenadas
        left = getattr(dim, 'left')
        top = getattr(dim, 'top')

        # Hago click en las coordenadas anteriores + 10 para centrar el click
        mouse.click(coords=(left - 40, top + 10))

    def obtenerEstadoBloqueado(self, estadoEsperado):
        main_win = self.app.top_window()

        # Saco un screenshot de la tabla 'Respuestas' dentro de consultar oficios
        image = main_win.ConsultaRespuestasdelOficio.SPR32X30_SpreadSheet1.capture_as_image()
        # Cambiar el nombre de la ruta
        path = r"\\sfs-1\\Testing\\Tareas en curso\\Adminsitracion de Embargos Judic (OJ)\Pruebas" \
               "\\2022\\Reportes de Prueba\\Reportes Automatizacion" \
               "\\Capturas\\estadoBloqueado.png"
        image.save(path)

        # Recorto la imagen para obtener solo la parte que dice bloqueada
        # img = cv2.imread(path)
        img = cv2.imread(path)
        crop_img = img[1:31, 490:559] #Recorta desde la fila 1 a la 31 y desde la columna 490 hasta la 559
        cv2.imwrite(path, crop_img)

        # Extraigo el texto para obtener el estado bloqueado
        extraccionTexto = ExtraccionTexto(self.app)

        texto = extraccionTexto.extraerTexto(path)

        #Se comenta la anidacion de if, ya que al detectar texto por imagen, tomaba las letras "BL".
        #if ("Si" in texto or "SI" in texto or "sI" in texto or "Bl" in texto or "bl" in texto):
        if ("SI" in texto):
            if(estadoEsperado == "SI"):
                # self.reporte.agregarLogCondicionalConScreen("El estado bloqueado del oficio es SI", True, path)
                self.reporte.agregarLogCondicionalConScreenPrevio("El estado bloqueado del oficio es SI", True, path)
            #else:
                #self.reporte.agregarLogCondicionalConScreen("El estado bloqueado del oficio es Si", False, path)
        else:
            if (estadoEsperado == "NO"):
                # self.reporte.agregarLogCondicionalConScreen("El estado bloqueado del oficio es NO", True, path)
                self.reporte.agregarLogCondicionalConScreenPrevio("El estado bloqueado del oficio es NO", True, path)
            #else:
                #self.reporte.agregarLogCondicionalConScreen("El estado bloqueado del oficio es No", False, path)

    def editarCarta(self):

        wordApp = Application().connect(title_re=".*Word", class_name="OpusApp", timeout=10)
        print(wordApp.OpusApp.texts()[0] + " archivo abierto.")

        print("Edito el archivo")

        # Edito el archivo con un texto para poder guardarlo
        wordApp.OpusApp._WwF._WwB._WwG.type_keys("Se edita el archivo.", with_spaces=True)

        time.sleep(2)
        self.reporte.agregarLogInformativoConScreen("Se edita el archivo", wordApp.OpusApp)

        # Presiono 'Alt + a' y 'g' para guardar
        wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('%a')
        wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('g')

        # Presiono 'Alt + a' y 's' para salir
        wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('%a')
        wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('s')

    def verificarSoloLectura(self, estaActivo):

        wordApp = Application().connect(title_re=".*Word", class_name="OpusApp", timeout=10)
        print(wordApp.OpusApp.texts()[0] + " archivo abierto.")
        time.sleep(5)

        if "Sólo lectura" in wordApp.OpusApp.window_text():
            if estaActivo:
                self.reporte.agregarLogCondicional("El archivo se encuentra en modo solo lectura", True)
            else:
                self.reporte.agregarLogCondicional("El archivo se encuentra en modo solo lectura", False)
        else:
            if estaActivo:
                self.reporte.agregarLogCondicional("El archivo no se encuentra en modo solo lectura", False)
            else:
                self.reporte.agregarLogCondicional("El archivo no se encuentra en modo solo lectura", True)

        # Presiono 'Alt + a' y 's' para salir
        wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('%a')
        wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('s')

class ConfirmarCarta:
    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def presionarSeleccionarOficio(self):
        main_win = self.app.top_window()

        # Presiona seleccionar oficio
        main_win.ThunderRT6FormDC.ThunderRT6UserControlDC.ThunderRT6PictureBoxDC.msvb_lib_toolbar5.click()

    def presionarAceptar(self):
        main_win = self.app.top_window()

        # Presiona seleccionar oficio
        main_win.ThunderRT6FormDC.ThunderRT6UserControlDC.ThunderRT6PictureBoxDC.msvb_lib_toolbar4.click(coords=(10, 10))

        # Aceptamos el dialog de aviso
        dialogFinalizo = self.app.window(class_name="#32770")
        time.sleep(1)
        self.reporte.agregarLogInformativoConScreen("Se confirma la carta.", dialogFinalizo)

        dialogFinalizo.Aceptar.click()

    def presionarSalir(self):
        main_win = self.app.top_window()

        # Presiona seleccionar oficio
        main_win.ThunderRT6FormDC.ThunderRT6UserControlDC.ThunderRT6PictureBoxDC.msvb_lib_toolbar.click(coords=(10, 10))

class RegistrarRecepcionCarta:
    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def presionarAceptar(self):
        main_win = self.app.top_window()

        # Presiona seleccionar oficio
        main_win.ThunderRT6FormDC.ThunderRT6UserControlDC.ThunderRT6PictureBoxDC.msvb_lib_toolbar4.click(coords=(10, 10))

        # Aceptamos el dialog de aviso
        dialogFinalizo = self.app.window(class_name="#32770")
        time.sleep(1)
        self.reporte.agregarLogInformativoConScreen("Se registra la recepcion de carta.", dialogFinalizo)

        dialogFinalizo.Aceptar.click()

class BlanquearCarta:

    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def editarCartaBlanqueada(self):
        wordApp = Application().connect(title_re=".*Word", class_name="OpusApp", timeout=10)
        print(wordApp.OpusApp.texts()[0] + " archivo abierto.")

        print("Edito el archivo blanqueado.")

        # Edito el archivo con un texto para poder guardarlo
        wordApp.OpusApp._WwF._WwB._WwG.type_keys("Se edita el archivo blanqueado. ", with_spaces=True)

        # time.sleep(1)
        # CarpetaScreen.tomarCaptura()
        # self.reporte.agregarLogInformativoConScreen("Se edita el archivo blanqueado", self.reporte.tomarCaptura(nombrecaso,wordApp.OpusApp))

        time.sleep(2)
        self.reporte.agregarLogInformativoConScreen("Se edita el archivo blanqueado.", wordApp.OpusApp)

        # Presiono 'Alt + a' y 'g' para guardar
        wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('%a')
        wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('g')

        # Presiono 'Alt + a' y 's' para salir
        wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('%a')
        wordApp.OpusApp.BarrademenúsMsoCommandBar.type_keys('s')

