import os
import pyautogui
from datetime import datetime

class CarpetaScreen:

    #nombrecaso = "CMPB33"

    # def nombreCarpeta(nombrecaso):
    #     return nombrecaso

    #ruta = os.getcwd().split('automatizacionaoj')[0] + "automatizacionaoj\\" + "/reportes/capturas/" + nombreCarpeta(nombrecaso)

    def crearCarpeta(nombrecaso):
        if os.path.exists(os.getcwd().split('automatizacionaoj')[0] + "automatizacionaoj\\" + "/reportes/capturas/"+nombrecaso):
            print("Carpeta existente.")
        else:
            os.makedirs(os.getcwd().split('automatizacionaoj')[0] + "automatizacionaoj\\" + "/reportes/capturas/"+nombrecaso)
            print("Se creo el directorio.")

    def tomarCaptura(nombrecaso):
        date_ms = str(datetime.now().date())+(str(datetime.now().time()))[:8]
        simbols = "-:."
        for special in simbols:
            date_ms=date_ms.replace(special, '')
        myscreen = pyautogui.screenshot()
        pathImg = os.getcwd().split('automatizacionaoj')[0] + "automatizacionaoj\\reportes\\capturas\\" + nombrecaso + "\\screen"+date_ms+".png"
        myscreen.save(pathImg)

    #crearCarpeta(nombrecaso)
    #tomarCaptura(nombrecaso)