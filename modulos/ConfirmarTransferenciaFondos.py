import time
from modulos.AgregarCuentaCorriente import AgregarCuenta
from modulos.SeleccionarOficio import SelectOficio


class ConfirmarTransferenciaFondos:

    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def seleccionarOficio(self, oficio, anio):
        time.sleep(2)
        main_win = self.app.top_window().child_window(title="Confirmar Transferencia de Fondos", class_name="ThunderRT6FormDC")

        try:
            main_win.ThunderRT6UserControlDC2.msvb_lib_toolbar5.click()
            seleccionarOficio = SelectOficio(self.app, self.reporte)
            seleccionarOficio.seleccionarOficioBase(oficio, anio)
        except:
            print("Hubo un error al seleccionar el oficio")

    def completarNroMEP(self, numero):
        main_win = self.app.top_window().child_window(title="Confirmar Transferencia de Fondos",
                                                      class_name="ThunderRT6FormDC")
        try:
            main_win.ThunderRT6PictureBoxDC3.ThunderRT6TextBox2.type_keys(numero)
            self.reporte.agregarLogCondicionalConScreen("Se completa el campo Nro MEP.", True, main_win)
        except:
            print("Hubo un error al completar el campo Nro MEP en Confirmar Transferencia de Fondos.")

    def aceptarConfirmarTransferencia(self):
        time.sleep(2)
        main_win = self.app.top_window().child_window(title="Confirmar Transferencia de Fondos",
                                                      class_name="ThunderRT6FormDC")
        try:
            # Click Aceptar en Confirmar Transferencia de Fondos
            main_win.ThunderRT6UserControlDC2.msvb_lib_toolbar4.click(coords=(10, 10))
            self.reporte.agregarLogInformativo("Se hace click en el boton Aceptar en Confirmar Transferencia de Fondos.")
        except:
            print("Hubo un error al hacer click en el boton Aceptar en Confirmar Transferencia de Fondos.")

    def finalizarOperacion(self):
        time.sleep(2)
        # El dialogo con boton aceptar no se encuentra dentro la ventana anterior, por eso se llama directamente a top_window
        main_win = self.app.top_window()
        try:
            # Click Aceptar al finalizar Confirmar Transferencia de Fondos
            self.reporte.agregarLogCondicionalConScreen("Se finaliza Confirmar Transferencia de Fondos.", True, main_win)
            main_win.Aceptar.click()
        except:
            print("Hubo un error al hacer click en el boton Aceptar al finalizar Confirmar Transferencia de Fondos.")

    def presionarSalir(self):
        main_win = self.app.top_window().child_window(title="Confirmar Transferencia de Fondos",
                                                      class_name="ThunderRT6FormDC")
        try:
            time.sleep(2)
            # Click en el boton 'Salir'
            main_win.ThunderRT6UserControlDC2.msvb_lib_toolbar.click(coords=(10, 10))
            self.reporte.agregarLogCondicional("Se presiona Salir de Confirmar Transferencias de Fondos", True)
        except:
            print("Hubo un error al presionar Salir de Confirmar Transferencias de Fondos.")
