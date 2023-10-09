import time


class SelectOficio:
    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def seleccionarOficio(self, nroOficio, anio):
        main_win = self.app.top_window()

        try:
            main_win.SeleccionarOficio.ThunderRT6PictureBoxDC.editar.type_keys(nroOficio)
            main_win.SeleccionarOficio.ThunderRT6PictureBoxDC.editar2.type_keys(anio)
            main_win = main_win.child_window(class_name='ThunderRT6UserControlDC', ctrl_index=4)
            main_win.Toolbar20WndClass.msvb_lib_toolbar.click()
            time.sleep(1)

            self.reporte.agregarLogInformativo("Se selecciona el oficio " + nroOficio + " para el año " + anio)
        except:
            print("Hubo un error al seleccionar el oficio: ", nroOficio, " del año: ", anio)

    def seleccionarOficioBase(self, nroOficio, anio, base='Oficios Manuales'):
        """
        :param app:
        :param nroOficio:
        :param anio:
        :param base: Por defecto la base de busque es 'Oficios Manuales'
        :return:
        """
        main_win = self.app.top_window()

        main_win.SeleccionarOficio.ReBar.ComboBox.select(base)
        main_win.SSTabCtlWndClass.Edit2.type_keys(anio)
        main_win.SSTabCtlWndClass.Edit4.type_keys(nroOficio)
        main_win.AceptarButton.click()
        try:
            if self.app.SeleccionarPaso.is_visible():
                self.app.SeleccionarPaso.Sí.click()
        except:
            print('No se encontro la ventana para saltear el paso')

        self.reporte.agregarLogInformativo("Se selecciona el oficio " + str(nroOficio) + " para el año " + str(anio))