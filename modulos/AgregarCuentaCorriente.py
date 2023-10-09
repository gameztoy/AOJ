import csv
import os
import re
import time
import random
from modulos.CuentasCrecer import CuentasEmbargables

class AgregarCuenta:

    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def seleccionarCuenta(self, tipo, index=1):
        try:
            path = os.getcwd()
            pathArchivo = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\cuentasValidas.csv"

            main_win = self.app.top_window().child_window(title="Cuenta Corriente [Agregar]", class_name="ThunderRT6FormDC")
            cuentas = main_win.ComboBox.texts()
            cuentas.pop(0)
            print("Cuentas disponibles en AOJ: ", cuentas)

            cuentasSUC = []
            with open(pathArchivo, newline='') as File:
                reader = csv.reader(File)
                for row in reader:
                    print("Para el cuit: " + str(row[0]), "se encontro la cuenta: ", str(row[1]))
                    self.reporte.agregarLogInformativo("Para el cuit: " + str(row[0]) + " se encontro la cuenta: " + str(row[1]))
                    cuenta = row[1]
                    cuentasSUC.append(cuenta)

            #Flujo para tipo de cuenta Dolar o Pesos
            if tipo == "Dolar":
                try:
                    #Conversion de List a cadena
                    check = '2'
                    tipoCuentas = [idx for idx in cuentasSUC if idx[0].lower() == check.lower()]
                    print("Las cuentas en dolares son : " + str(tipoCuentas))
                    cuentaSeleccionada = random.choice(tipoCuentas)
                    print("Cuenta a embargar: "+cuentaSeleccionada)
                    self.reporte.agregarLogCondicional("Cuenta tipo Dolar a embargar: "+cuentaSeleccionada, True)
                except:
                    print("No se encontraron cuentas de tipo Dolar.")
                    self.reporte.agregarLogCondicional("No se encontraron cuentas de tipo Dolar.", False)
            else:
                try:
                    check0 = '0'
                    check1 = '1'
                    tipoCuentas = [idx for idx in cuentasSUC if idx.lower().startswith(check0.lower())] + [idx for idx in cuentasSUC if idx.lower().startswith(check1.lower())]
                    print("Las cuentas en pesos son : " + str(tipoCuentas))
                    cuentaSeleccionada = random.choice(tipoCuentas)
                    print("Cuenta a embargar: " + cuentaSeleccionada)
                    self.reporte.agregarLogCondicional("Cuenta tipo Pesos a embargar: " + cuentaSeleccionada, True)
                except:
                    print("No se encontraron cuentas de tipo Pesos.")
                    self.reporte.agregarLogCondicional("No se encontraron cuentas de tipo Pesos.", False)

            path = os.getcwd()
            pathArchivoSalida = path.split('automatizacionaoj')[
                                    0] + "automatizacionaoj\\recursos\\datosValidacionCrecer.csv"
            with open(pathArchivo, newline='') as File:
                csv.reader(File)
                f = open(pathArchivoSalida, "w")
                writer = csv.writer(f, lineterminator='\n')
                writer.writerow([cuentaSeleccionada])

            cuentaSeleccionada = [cuenta for cuenta in cuentas if cuentaSeleccionada in cuenta]
            cuentaSeleccionada = " ".join(cuentaSeleccionada)
            main_win.ComboBox.click()
            main_win.ComboBox.select(cuentaSeleccionada)

        except:
            print("Hubo un error al seleccionar la cuenta con indice: ", cuentaSeleccionada, index)

    def seleccionarCuentaTransferencia(self):
        try:
            path = os.getcwd()
            pathArchivoSalida = path.split('automatizacionaoj')[
                                    0] + "automatizacionaoj\\recursos\\datosValidacionCrecer.csv"

            main_win = self.app.top_window().child_window(title="Cuenta Corriente [Agregar]", class_name="ThunderRT6FormDC")
            cuentas = main_win.ComboBox.texts()
            cuentas.pop(0)
            print("Cuentas disponibles: ", cuentas)

            cuentasSUC = []
            with open(pathArchivoSalida, newline='') as File:
                reader = csv.reader(File)
                for row in reader:
                    print("Cuenta para Transferencia: " + str(row[0]))
                    self.reporte.agregarLogInformativo(
                        "Cuenta para Transferencia: " + str(row[0]))
                    cuentaT = row[0]
                    print(cuentaT)
                    #cuentasSUC.append(cuenta)

            #cuentaSeleccionada = [cuenta for cuenta in cuentas if cuentaSeleccionada in cuenta]
            cuentaSeleccionada = [cuenta for cuenta in cuentas if cuentaT in cuenta]
            cuentaSeleccionada = " ".join(cuentaSeleccionada)
            print("cuentaSeleccionada: "+cuentaSeleccionada)
            main_win.ComboBox.click()
            main_win.ComboBox.select(cuentaSeleccionada)
        except:
            print("Hubo un error al seleccionar la cuenta para Transferencia: ", cuentaSeleccionada)

    def ingresarImporteOrigen(self, importe):
        try:
            main_win = self.app.top_window().child_window(title="Cuenta Corriente [Agregar]", class_name="ThunderRT6FormDC")
            main_win.Edit2.type_keys(importe)

            self.reporte.agregarLogInformativo("Se ingresa el importe de origen")

            # Captura de datos en Cuenta Corriente [Agregar].
            self.reporte.agregarLogInformativoConScreen("Datos en Cuenta Corriente [Agregar].", main_win)
        except:
            print("Hubo un error al ingresar el importe de origen: "+importe)

    def aceptar(self):
        try:
            main_win = self.app.top_window().child_window(title="Cuenta Corriente [Agregar]", class_name="ThunderRT6FormDC")
            main_win.child_window(class_name='ThunderRT6UserControlDC', ctrl_index=10).msvb_lib_toolbar2.click()
        except:
            print("Hubo un error al presionar aceptar en la ventana Cuenta Corriente [Agregar]")

    def obtenerCuentasEmbargables(self, cuit):
        try:
            cuentasCrecer = CuentasEmbargables()
            cuentasCrecer.buscarCuentasPorCUIT(cuit)
        except:
            print("Hubo un error al querer obtener las cuentas embargables")
