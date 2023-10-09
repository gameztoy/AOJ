import csv
import os
import pyautogui
import datetime
from jpype import startJVM, JClass, getDefaultJVMPath, shutdownJVM, isJVMStarted
from datetime import datetime

class Reporte:
    # Modificar la variable aoj_version segun corresponda.
    global aoj_version, common_path
    aoj_version = "Reportes de Prueba"
    common_path = r"\\sfs-1\\Testing\\Tareas en curso\\Adminsitracion de Embargos Judic (OJ)\Pruebas\\2022\\" + \
                  aoj_version + "\\Reportes Automatizacion\\"

    def __init__(self, nombretest, descripciontest):
        path = os.getcwd().split('automatizacionaoj')[0] + "automatizacionaoj\\lib\\"
        path_reporte = common_path + nombretest
        os.makedirs(path_reporte, exist_ok=True)

        classpath = "" + path + "extentreports-2.40.2.jar;" + path + "freemarker-2.3.23.jar"""
        # Se comprueba si hay una instancia de JVM ya iniciada.
        if not isJVMStarted():
            startJVM(getDefaultJVMPath(), "-Djava.class.path=%s" % classpath)

        ExtentReports = JClass('com.relevantcodes.extentreports.ExtentReports')
        ExtentTest = JClass('com.relevantcodes.extentreports.ExtentTest')
        self.LogStatus = JClass('com.relevantcodes.extentreports.LogStatus')
        report_path = os.path.join(path_reporte, str(descripciontest) + ' - ' + self.timenow() + '.html')
        self.extent = ExtentReports(report_path)
        self.test = self.extent.startTest(nombretest, descripciontest)

    def agregarLogInformativo(self, descripcion):
        self.test.log(self.LogStatus.INFO, descripcion)

    def agregarLogCondicional(self, descripcion, condicion):
        if condicion:
            self.test.log(self.LogStatus.PASS, descripcion + " - Exitoso")
        else:
            self.test.log(self.LogStatus.FAIL, descripcion + " - Fallido")

    """Logs para AOJ"""
    def agregarLogInformativoConScreen(self, descripcion, screen):
        self.test.log(self.LogStatus.INFO, self.test.addScreenCapture(self.capturarImagen(screen)) + descripcion)

    def agregarLogCondicionalConScreen(self, descripcion, condicion, screen):
        if condicion:
            self.test.log(self.LogStatus.PASS, self.test.addScreenCapture(self.capturarImagen(screen)) + descripcion + " - Exitoso")
        else:
            self.test.log(self.LogStatus.FAIL, self.test.addScreenCapture(self.capturarImagen(screen)) + descripcion + " - Fallido")

    def agregarLogCondicionalConScreenPrevio(self, descripcion, condicion, screen):
        if condicion:
            self.test.log(self.LogStatus.PASS, self.test.addScreenCapture(screen) + descripcion + " - Exitoso")
        else:
            self.test.log(self.LogStatus.FAIL, self.test.addScreenCapture(screen) + descripcion + " - Fallido")

    """Logs para Crecer"""
    def agregarLogInformativoConScreenCrecer(self, descripcion, screen):
        self.test.log(self.LogStatus.INFO, self.test.addScreenCapture(self.capturarImagenCrecer(screen)) + descripcion)

    def agregarLogCondicionalConScreenCrecer(self, descripcion, condicion, screen):
        if condicion:
            self.test.log(self.LogStatus.PASS, self.test.addScreenCapture(self.capturarImagenCrecer(screen)) + descripcion + " - Exitoso")
        else:
            self.test.log(self.LogStatus.FAIL, self.test.addScreenCapture(self.capturarImagenCrecer(screen)) + descripcion + " - Fallido")

    def terminarReporte(self):
        self.extent.endTest(self.test)
        self.extent.flush()
        # No se apaga la instancia de JVM, ya que no se puede volver a reiniciar.
        #shutdownJVM()

    def timenow(self):
        datenow = datetime.now().strftime("%Y%m%d%H%M%S")[2:]
        return datenow

    """
    Metodo que captura la ventana de AOJ y retorna el path de la imagen para poder agregarla a un log con screen.
    """
    def capturarImagen(self, ventana):
        image = ventana.capture_as_image()
        path_image = common_path + "Capturas\\"
        os.makedirs(path_image, exist_ok=True)
        name_image = "aoj" + self.timenow() + ".png"
        image.save(path_image + name_image)
        return path_image + name_image

    """
    Metodo que captura las ventanas de Crecer y retorna el path de la imagen para poder agregarla a un log con screen.
    """
    def capturarImagenCrecer(self, ventana):
        image = pyautogui.screenshot(region=(0, 0, 1200, 700))
        path_image = common_path + "Capturas\\"
        os.makedirs(path_image, exist_ok=True)
        name_image = "crecer" + self.timenow() + ".png"
        image.save(path_image + name_image)
        return path_image + name_image

    """
    Metodo que generar un csv con cuit, numero y anio de OM.
    """
    def generarCsv(self, denominacion, numero, anio):
        path = os.getcwd()
        path = path.split('automatizacionaoj')[0] + "automatizacionaoj\\reportes\\Oficios Creados\\oficiosCreados.csv"
        f = open(path, "a+")
        writer = csv.writer(f, lineterminator='\n')
        writer.writerow([str(denominacion), str(numero), str(anio)])
