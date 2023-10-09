import time 
from pywinauto.application import Application


class AOJ:
    def __init__(self):
        self.app = Application(backend="win32").start('C:\AOJ7\AOJ7.exe', timeout=10)
        main_win = self.app.top_window()

        # cerrar dialogo de error
        try:
            dlg_handles = main_win.window(title="ERROR - MAIN")
            if dlg_handles:
                dlg_handles.Button.click()
        except:
            print("No se encontró diálogo al iniciar la app. \n")
        main_win = self.app.top_window()
        main_win.NovedadessinImportar.Salir.click()

    def retornarAOJApp(self):
        return self.app
    
    def closeAOJApp(self):
        time.sleep(1)
        #cerrar aplicacion
        try:
          self.app.top_window().menu_select("Archivo->Salir")
          time.sleep(1)
          dlg_handles = self.app.window(class_name="#32770")
          if dlg_handles:
              time.sleep(2)
              dlg_handles.Button.click()
        except:
          print("Algo falló al querer cerrar la app")