import time


def irA(app, submenu, reporte):
    try:
        main_win = app.top_window()
        main_win.menu_select(submenu)
        reporte.agregarLogInformativo("Ingresamos a " + submenu)
        time.sleep(1)
    except:
        print("Hubo un error al querer ir al submenu: ", submenu)
        reporte.agregarLogInformativo("Hubo un error al querer ir al submenu: " + submenu)

