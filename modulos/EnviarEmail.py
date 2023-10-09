import time
from pywinauto import Application


class EnviarEmail:

    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def aceptarEnvio(self):
        """ Acepta el envío del email y los avisos que esté genera
            :param app: aplicación
        """
        main_win = self.app.top_window()

        # Presionamos aceptar
        main_win.EnviarEmail.ThunderRT6UserControlDC2.msvb_lib_toolbar2.click(coords=(10, 10))

        # Aceptamos el mensaje
        self.app.EnviarEmail.Sí.click()
        time.sleep(1)
        self.reporte.agregarLogInformativoConScreen("Se acepta el envio del email",
                                                    self.app.EnviarEmail)

        # Si se muestra el aviso de saltear paso continuamos
        try:
            # Capturamos la ventana del mensaje de advertencia de outlook
            outlookApp = Application().connect(title="Microsoft Outlook", ctrl_index=0, visible_only=True, timeout=5)

            if outlookApp.Dialog.is_visible():
                time.sleep(5)
                outlookApp.MicrosoftOutlook.Sí.click()
                outlookApp.MicrosoftOutlook.type_keys('{ENTER}')
        except:
            print("No se encontro el dialogo de advertencia de outlook")

        # Aceptamos el mensaje de que se envio correctamente
        self.app.EnviarEmail.Aceptar.click()

    def presionarSeleccionarOficio(self):
        main_win = self.app.top_window()

        # Presiono el boton Seleccionar oficio
        main_win.EnviarEmail.ThunderRT6UserControlDC2.Toolbar20WndClass.msvb_lib_toolbar.click()

    def presionarSalir(self):
        main_win = self.app.top_window().child_window(title="Enviar Email",
                                                      class_name="ThunderRT6FormDC")
        try:
            # Click en el boton 'Salir'
            main_win.ThunderRT6UserControlDC2.msvb_lib_toolbar8.click(coords=(10, 10))
            self.reporte.agregarLogCondicional("Se presiona Salir de Enviar Email", True)
            time.sleep(1)
        except:
            print("Hubo un error al presionar Salir de Enviar Email")
            self.reporte.agregarLogCondicional("Hubo un error al presionar Salir de Enviar Email", False)

class EnviarEmailsPorLote:

    def __init__(self,app, reporte):
        self.app = app
        self.reporte = reporte

    def seleccionarOficioParaEnviarMails(self, tipoOficio):
        ''' Se seleccionan los oficios en base al filtro indicado
            En este caso solo se filtra por tipo de oficio
            De requerirse más filtros se pueden agregar
            Esta pantalla es diferente a la de seleccionar oficio de Consulta de Respuesta

        :param tipoOficio: Tipo de oficio a filtrar
        '''
        main_win = self.app.top_window()

        try:
            # Presionamos el boton Seleccionar Oficio
            main_win.EnviarEmails.ThunderRT6UserControlDC.msvb_lib_toolbar.click()
            time.sleep(3)

            main_win = self.app.top_window()

            # Seleccionamos el tipo de oficio
            main_win.SeleccionarOficio.ThunderRT6PictureBoxDC2.ComboBox14.select(tipoOficio)

            self.reporte.agregarLogInformativo("Se filtran por los oficios de tipo: " + tipoOficio)

            # Presionamos aceptar
            main_win.SeleccionarOficio.ThunderRT6PictureBoxDC.Aceptar.click()
        except:
            print("Hubo un error al seleccionar el oficio")

    def aceptarEnvioDeEmails(self):
        """ Acepta el envío de los emails y los avisos que esté genera
            Tambien cierra las ventanas de Resultados del proceso y Enviar Emails
        """
        main_win = self.app.top_window()

        try:
            # Presionamos el boton Aceptar
            main_win.EnviarEmails.ThunderRT6UserControlDC.msvb_lib_toolbar2.click(coords=(10,10))

            # Presionamos Sí
            time.sleep(1)
            self.app.EnviarEmail.Sí.click()

            # Aceptamos el mensaje de que se envio correctamente
            self.app.EnviarEmails.Aceptar.click()
            self.reporte.agregarLogInformativoConScreen("Mails enviados correctamente", self.app.EnviarEmails)
            time.sleep(1)

            # Presionamos Salir en la ventana de 'Resumen del Proceso'
            main_win = self.app.top_window()
            main_win.ThunderRT6FormDC.ThunderRT6PictureBoxDC.Salir.click()
            time.sleep(1)

            # Presionamos Salir en la ventana de 'Enviar Emails'
            main_win = self.app.top_window()
            main_win.EnviarEmails.ThunderRT6UserControlDC.msvb_lib_toolbar6.click(coords=(10,10))
            time.sleep(1)

            # Aceptamos el diálogo de confimación
            self.app.EnviarEmail.Sí.click()
        except:
            print("Hubo un error al aceptar el envío de mails por lote")