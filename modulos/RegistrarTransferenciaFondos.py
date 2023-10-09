import time
from modulos.AgregarCuentaCorriente import AgregarCuenta
from modulos.SeleccionarOficio import SelectOficio


class RegistrarTransferenciaFondos:

    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def seleccionarOficio(self, oficio, anio):
        time.sleep(2)
        main_win = self.app.top_window().child_window(title="Registrar Transferencias de Fondos", class_name="ThunderRT6FormDC")

        try:
            main_win.ThunderRT6UserControlDC3.msvb_lib_toolbar.click()
            seleccionarOficio = SelectOficio(self.app, self.reporte)
            seleccionarOficio.seleccionarOficioBase(oficio, anio)
        except:
            print("Hubo un error al seleccionar el oficio")

    def agregarCuenta(self, importe):
        main_win = self.app.top_window().child_window(title="Registrar Transferencias de Fondos", class_name="ThunderRT6FormDC")
        main_win.Agregar.click()

        try:
            agregar_win = AgregarCuenta(self.app, self.reporte)
            agregar_win.seleccionarCuentaTransferencia()
            agregar_win.ingresarImporteOrigen(importe)
            agregar_win.aceptar()
        except:
            print("Hubo un error al agregar la cuenta.")

    def presionarAceptarYAvanzar(self):
        main_win = self.app.top_window().child_window(title="Registrar Transferencias de Fondos",
                                                      class_name="ThunderRT6FormDC")
        try:
            main_win.ThunderRT6UserControlDC3.msvb_lib_toolbar3.click(coords=(10, 10))
            time.sleep(3)
            self.reporte.agregarLogCondicional("Se presiona Aceptar y Avanzar", True)
        except:
            print("Hubo un error al presionar Aceptar y Avanzar")

    def presionarSiMovimientosNoConfirmados(self):
        main_win = self.app.window(title='AOJ7',  class_name="#32770")

        try:
            time.sleep(2)
            main_win.Sí.click()
            self.reporte.agregarLogCondicional("Se presiona Sí al detectar movimientos de retencion aun no confirmados en Registrar Transferencias de Fondos", True)
            time.sleep(10)
        except:
            print("Hubo un error al presionar Sí en movimientos no confirmados Registrar Transferencias de Fondos")

    def presionarSalir(self):
        main_win = self.app.top_window().child_window(title="Registrar Transferencias de Fondos",
                                                      class_name="ThunderRT6FormDC")
        try:
            # Click en el boton 'Salir'
            main_win.ThunderRT6UserControlDC3.msvb_lib_toolbar6.click(coords=(10, 10))
            self.reporte.agregarLogCondicional("Se presiona Salir de Registrar Transferencias de Fondos", True)
            time.sleep(1)
        except:
            print("Hubo un error al presionar Salir de Registrar Transferencias de Fondos")
