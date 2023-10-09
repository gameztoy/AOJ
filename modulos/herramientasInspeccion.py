import os


def save_control_identifiers( ventana):
        #print('Capturando Control Identifiers de: '+ventana.element_info.name)
        path = os.getcwd()
        print(path.split('automatizacionaoj'))
        #path = path.split('automatizacionaoj')[0] + 'automatizacionaoj\\captura\\'+ventana.element_info.name+'.txt'
        path = path.split('automatizacionaoj')[0] + 'automatizacionaoj\\captura\\a.txt'
        filename = os.path.abspath(path)
        print(filename)
        ventana.print_control_identifiers(filename=filename)


def marcar_ventana(ventana):
        ventana.draw_outline()