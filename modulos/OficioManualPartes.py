import time


class OficioManualPartes:
    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def ingresarDenominacion(self, cuit):
        app_ventana_main = self.app.top_window()

        # abre ventana 'Oficio Manual Parte [Agregar]',boton 'Buscar' Denominacion
        app_ventana_main.ThunderRT6FormDC.Button2.click()
        time.sleep(1)

        app_ventana_main = self.app.top_window()
        # abre ventana 'Seleccionar Cliente',completa Campo 'CUIT'
        app_ventana_main.SeleccionarCliente.Edit.type_keys(cuit)

        # Boton Aceptar
        app_ventana_main.SeleccionarCliente.msvb_lib_toolbar.click()
        print("Persona seleccionada: ",cuit)
        self.reporte.agregarLogInformativo("Se selecciona la persona con cuit: " + str(cuit))

        # cerrar dialogo de error ODBC, a configurar
        try:
            dlg_handles = self.app.window(class_name="#32770")
            if dlg_handles:
                time.sleep(5)
                dlg_handles.Button.click()
        except:
            print("")

        # Esperar a que el cliente se cargue en pantalla
        time.sleep(10)

    def presionarSinPersona(self):
        app_ventana_main = self.app.top_window()
        # tilda el checkbox 'Sin Persona'
        app_ventana_main.ThunderRT6FormDC.SinPersona.click()
        time.sleep(1)

        print("Sin persona seleccionado")
        self.reporte.agregarLogInformativo("Sin persona seleccionado")

    def realizarCruceIndividual(self):
        app_ventana_main = self.app.top_window()

        # Vuelve a 'Oficio Manual Parte [Agregar]', boton 'Cruce Individual'
        app_ventana_main.ThunderRT6FormDCwindow.ThunderRT6UserControlDC17.msvb_lib_toolbar.click()
        time.sleep(1)

        app_ventana_main = self.app.top_window()
        # abre ventana 'Cruce Individual', Boton Buscar
        app_ventana_main.CruceIndividual.ThunderRT6UserControlDC3.ReBar.Toolbar20WndClass3.msvb_lib_toolbar.click()
        time.sleep(1)

        app_ventana_main = self.app.top_window()
        # abre ventana 'Seleccionar Cliente', Donde aparecen los datos del cliente seleccionado, Boton 'Aceptar'
        app_ventana_main.SeleccionarCliente.ThunderRT6UserControlDC.msvb_lib_toolbar.click()
        time.sleep(1)

        app_ventana_main = self.app.top_window()
        # vuelve a la ventana 'Cruce Individual', Boton Aceptar
        app_ventana_main.CruceIndividual.ThunderRT6UserControlDC.ReBar.ThunderRT6PictureBoxDC2.msvb_lib_toolbar2.click()
        time.sleep(1)

        print("Cruce Individual realizado con exito")
        self.reporte.agregarLogCondicional("Cruce individual realizado", True)

    def seleccionarTipoOficio(self, tipoOficio):
        app_ventana_main = self.app.top_window()

        # Tipo Oficio
        app_ventana_main.ThunderRT6FormDC.ThunderRT6UserControlDC2.ThunderRT6ComboBox.select(tipoOficio)
        self.reporte.agregarLogInformativo("Se selecciona el tipo de oficio: " + tipoOficio)
        time.sleep(1)

    def seleccionarVigencia(self, vigencia):
        app_ventana_main = self.app.top_window()
        # Vigencia
        app_ventana_main.ThunderRT6FormDC.ThunderRT6UserControlDC3.ThunderRT6ComboBox.select(vigencia)
        self.reporte.agregarLogInformativo("Se selecciona la vigencia: " + vigencia)
        time.sleep(1)

    def seleccionarTransferencia(self, transferencia):
        app_ventana_main = self.app.top_window()

        # Transferencia
        app_ventana_main.ThunderRT6FormDC.ThunderRT6UserControlDC5.ThunderRT6ComboBox.select(transferencia)
        self.reporte.agregarLogInformativo("Se selecciona la transferencia: " + transferencia)
        time.sleep(1)

    def seleccionarMoneda(self, moneda):
        app_ventana_main = self.app.top_window()

        # Moneda: Pesos
        app_ventana_main.ThunderRT6FormDC.ThunderRT6UserControlDC6.ThunderRT6ComboBox.select(moneda)
        self.reporte.agregarLogInformativo("Se selecciona la moneda: " + moneda)
        time.sleep(1)

    def seleccionarSinMontos(self):
        app_ventana_main = self.app.top_window()

        # Tilda el check box 'Sin montos'
        app_ventana_main.ThunderRT6FormDC.SinMontos.click()
        self.reporte.agregarLogInformativo("Sin montos seleccionado")
        time.sleep(1)

    def ingresarCapitalYIntereses(self, capital, intereses):
        app_ventana_main = self.app.top_window()

        # ingresa Campo 'capital'
        app_ventana_main.ThunderRT6FormDC.ThunderRT6UserControlDC8.Edit.type_keys(capital)
        time.sleep(1)
        # ingresa Campo 'intereses'
        app_ventana_main.ThunderRT6FormDC.ThunderRT6UserControlDC9.Edit.type_keys(intereses)
        time.sleep(1)

        self.reporte.agregarLogInformativo("Se ingresan capital: " + str(capital) + " y intereses: " + str(intereses))

    def enlazarOficio(self, numeroOM, anioOM):
        main_win = self.app.top_window()

        # Clickeamos el boton para enlazar un oficio
        main_win.ThunderRT6PictureBoxDC.Button3.click()

        main_win = self.app.top_window()

        # Seleccionamos Oficios Manuales de la lista
        main_win.SeleccionarOficio.ComboBox.select("Oficios Manuales")

        # Ingresamos el numero de oficio
        main_win.SeleccionarOficio.ThunderRT6PictureBoxDC1.Edit1.type_keys(numeroOM)

        # Ingresamos el año
        main_win.SeleccionarOficio.ThunderRT6PictureBoxDC1.Edit2.type_keys(anioOM)

        # Presionamos aceptar
        main_win.SeleccionarOficio.ThunderRT6PictureBoxDC5.msvb_lib_toolbar.click(coords=(10, 10))

        time.sleep(1)
        self.reporte.agregarLogInformativo("Se enlaza el oficio: " + numeroOM + "- año: " + anioOM)
        print("Oficio enlazado al oficio:", numeroOM, "- año: ", anioOM)

    def presionarAceptar(self):
        app_ventana_main = self.app.top_window()

        #Boton Aceptar en 'Oficio Manual Parte [Agregar]'
        app_ventana_main.ThunderRT6FormDCwindow.ThunderRT6UserControlDC17.ThunderRT6PictureBoxDC2.Toolbar20WndClass2.msvb_lib_toolbar.click()
        time.sleep(1)

        # Confirma dialogo de:
        # No se encontró ninguna programación para las caracteristicas del oficio especificado. Desea continuar.
        try:
            dlg_handles = self.app.window(title="ATENCION - Oficio Manual Parte [Agregar]")
            time.sleep(1)
            dlg_handles.Sí.click()
        except:
            print("No se encontro el dialogo de 'No se encontró ninguna programación para las caracteristicas del oficio especificado. Desea continuar'")