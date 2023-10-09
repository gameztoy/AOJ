class Acciones:

    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def desplegarMenu(self, menu):
        """ Despliega el menú indicado por parametro
            :param menu: nombre del menú a desplegar
        """
        listaMenus = {"Emails": 1,
                      "Cartas": 2,
                      "Respuestas": 4,
                      "Transferencias": 8}
        main_win = self.app.top_window()

        try:
            cantidad = listaMenus.get(menu)

            # Presiono la tecla hacia abajo la cantidad indicada de veces para moverme entre las opciones
            while cantidad > 0:
                main_win.Acciones.type_keys('{DOWN}')
                cantidad = cantidad - 1

            # Presiono la tecla hacia la derecha para desplegar el menú
            main_win.Acciones.type_keys('{RIGHT}')

            self.reporte.agregarLogInformativo("Se despliega el menú: " + menu)
        except:
            print("No se encontro el menú: ", menu)

    def abrirAccion(self, accion):
        """ Abre la acción indicado por parametro dentro del menú desplegado
            :param accion: nombre de la acción a abrir
        """
        listaAcciones = {"Emitir Carta Individual": 1,
                         "Confirmar Carta": 2,
                         "Emitir Carta por Lote": 3,
                         "Registrar Recepción de Carta": 4,
                         "Enviar Email Individual": 1,
                         "Enviar Email por Lote": 2,
                         "Registrar Respuestas de Email": 3,
                         "Asignar Sucursal/Oficial Directa": 4,
                         "Registrar Respuestas de Embargos": 1,
                         "Registrar Transferencia de Fondos": 1,
                         "Confirmar Transferencia de Fondos": 3}

        main_win = self.app.top_window()

        try:
            cantidad = listaAcciones.get(accion)

            # Presiono la tecla hacia abajo la cantidad indicada de veces para moverme entre las acciones
            while cantidad > 0:
                main_win.Acciones.type_keys('{DOWN}')
                cantidad = cantidad - 1

            # Presiono la tecla Enter para abrir la acción
            main_win.Acciones.type_keys('{ENTER}')

            self.reporte.agregarLogInformativo("Se abre la acción: " + accion)
        except:
            print("No se encontro la acción: ", accion)