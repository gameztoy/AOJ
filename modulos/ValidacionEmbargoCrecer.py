import csv
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from modulos.AgregarCuentaCorriente import AgregarCuenta

class CuentasPorCuentas:

    def __init__(self, app, reporte):
        self.app = app
        self.reporte = reporte

    def buscarCuentasPorCuenta(self):
        driver = webdriver.Chrome(executable_path=r"\\sfs-1\Testing\Automatizacion_de_Proyectos\Herramientas\chromedriver 75\chromedriver.exe")

        driver.get("http://sapa2101lx-vip.bancocredicoop.coop/t24ts702/servlet/BrowserServlet")
        driver.maximize_window()
        driver.find_element(By.CSS_SELECTOR, '#signOnName').send_keys('F00402')
        driver.find_element(By.CSS_SELECTOR, '#sign-in').click()
        self.reporte.agregarLogCondicional("Se inicia sesion en Crecer", True)
        time.sleep(2)

        frame = driver.find_element(By.CSS_SELECTOR, "frame[id^='banner']")
        driver.switch_to.frame(frame)
        driver.maximize_window()
        time.sleep(2)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@id='commandValue']"))).send_keys('?328')
        driver.find_element(By.CSS_SELECTOR, '#cmdline_img').click()
        time.sleep(2)

        #Posicionamiento sobre la ventana T24
        t24 = driver.window_handles[1]
        driver.switch_to.window(t24)
        driver.maximize_window()
        driver.find_element(By.CSS_SELECTOR, '#pane_>ul:nth-child(4)>li>span').click()
        driver.find_element(By.CSS_SELECTOR, '#pane_>ul:nth-child(4)>li>ul>li:nth-child(2)>span').click()
        driver.find_element(By.CSS_SELECTOR, "a[onclick*='Consulta de Cuentas por Cuenta']").click()

        #Posicionamiento sobre la ventana Consulta de Cuentas por Cuenta
        cxc = driver.window_handles[2]
        driver.switch_to.window(cxc)
        driver.maximize_window()
        driver.find_element(By.CSS_SELECTOR, "[name='value:1:1:1']").clear()

        datoSeleccionado = []
        try:
            path = os.getcwd()
            pathArchivo = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\datosValidacionCrecer.csv"
            with open(pathArchivo, newline='') as File:
                reader = csv.reader(File)
                for row in reader:
                    cuenta = row[0]
                    datoSeleccionado.append(cuenta)
            print("Cuenta para validar en Crecer: " + str(row[0]))
        except:
            print("Hubo un error al leer el archivo.")

        driver.find_element(By.CSS_SELECTOR, "[name='value:1:1:1']").send_keys(datoSeleccionado)
        driver.find_element(By.CSS_SELECTOR, "a[alt='Ejecutar']").click()
        time.sleep(2)
        driver.find_element(By.CSS_SELECTOR, "img[alt='Select Drilldown']").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//span[normalize-space()='Datos Impositivos y Estado Cuenta']").click()
        try:
            driver.find_element(By.XPATH, "//a[@id='fieldCaption:L.AC.TIP.BLO']").is_displayed()
        except:
            print('Tipo de bloqueo no encontrado')

        # Validacion de tipo de bloqueo 2 y bloqueo parcial
        # Flujo para distinguir posicion del monto del embargo, segun tipo de cuenta.
        primerdigito = " ".join(datoSeleccionado)
        primerdigito = primerdigito[0]
        if primerdigito == '0':
            try:
                bloqueoparcial = driver.find_element(By.CSS_SELECTOR, '#enri_L\.AC\.TIP\.BLO').get_attribute(
                    "textContent")
                assert bloqueoparcial == "BLOQUEO PARCIAL"
                print('El bloqueo es: ' + bloqueoparcial)
                self.reporte.agregarLogCondicionalConScreenCrecer("Tipo de bloqueo es: " + bloqueoparcial, True, cxc)
            except:
                print('No se ha encontrado Tipo de bloqueo 2 BLOQUEO PARCIAL')
                self.reporte.agregarLogCondicional("Tipo de bloqueo es: " + bloqueoparcial, False)
        else:
            try:
                bloqueoparcial = driver.find_element(By.CSS_SELECTOR, '#enri_L\.AC\.TIP\.BLO').get_attribute(
                    "textContent")
                assert bloqueoparcial == "BLOQUEO PARCIAL"
                print('El bloqueo es: ' + bloqueoparcial)
                self.reporte.agregarLogCondicionalConScreenCrecer("Tipo de bloqueo es: " + bloqueoparcial, True, cxc)
            except:
                print('No se ha encontrado Tipo de bloqueo 2 BLOQUEO PARCIAL.')
                self.reporte.agregarLogCondicional("Tipo de bloqueo es: " + bloqueoparcial, False)

        driver.switch_to.window(cxc)
        driver.close()
        driver.switch_to.window(t24)
        driver.close()
        driver.quit()

    def consultaMovPorFecha(self, importe):
        driver = webdriver.Chrome(executable_path=r"\\sfs-1\Testing\Automatizacion_de_Proyectos\Herramientas\chromedriver 75\chromedriver.exe")
        driver.maximize_window()
        driver.get("http://sapa2101lx-vip.bancocredicoop.coop/t24ts702/servlet/BrowserServlet")
        driver.find_element(By.CSS_SELECTOR, '#signOnName').send_keys('F00402')
        driver.find_element(By.CSS_SELECTOR, '#sign-in').click()
        self.reporte.agregarLogCondicional("Se inicia sesion en Crecer", True)
        time.sleep(2)

        #Menu principal Crecer
        frame = driver.find_element(By.CSS_SELECTOR, "frame[id^='banner']")
        driver.switch_to.frame(frame)
        time.sleep(2)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@id='commandValue']"))).send_keys('?328')
        driver.find_element(By.CSS_SELECTOR, '#cmdline_img').click()
        time.sleep(2)

        #Posicionamiento sobre la ventana T24
        t24 = driver.window_handles[1]
        driver.switch_to.window(t24)
        driver.maximize_window()
        driver.find_element(By.CSS_SELECTOR, '#pane_>ul:nth-child(4)>li>span').click()
        driver.find_element(By.CSS_SELECTOR, '#pane_>ul:nth-child(4)>li>ul>li:nth-child(2)>span').click()
        driver.find_element(By.CSS_SELECTOR, "a[onclick*='Consulta de Mov. por Fecha Proceso']").click()

        #Posicionamiento sobre la ventana Movimiento por fecha de Cuentas
        cxc = driver.window_handles[2]
        driver.switch_to.window(cxc)
        driver.maximize_window()
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR, "[name='value:1:1:1']").clear()

        datoSeleccionado = []
        try:
            path = os.getcwd()
            pathArchivo = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\datosValidacionCrecer.csv"
            with open(pathArchivo, newline='') as File:
                reader = csv.reader(File)
                for row in reader:
                    cuenta = row[0]
                    datoSeleccionado.append(cuenta)
        except:
            print("Hubo un error al leer el archivo.")

        driver.find_element(By.CSS_SELECTOR, "[name='value:1:1:1']").send_keys(datoSeleccionado)
        driver.find_element(By.CSS_SELECTOR, "a[alt='Ejecutar']").click()
        time.sleep(2)

        #Flujo para distinguir posicion del monto del embargo, segun tipo de cuenta.
        primerdigito = " ".join(datoSeleccionado)
        primerdigito = primerdigito[0]
        if primerdigito == '0':
            try:
                montodebito = driver.find_element(By.CSS_SELECTOR, '#r2>td:nth-child(5)').get_attribute("textContent")
                assert montodebito == importe
                print('El importe indicado: '+importe+' es igual al montoDebito: '+montodebito)
                self.reporte.agregarLogCondicionalConScreenCrecer("El importe indicado: " + importe + " es igual al montoDebito: " + montodebito, True, cxc)
            except:
                print('No se ha encontrado el montoDebito correspondiente')
                self.reporte.agregarLogCondicional("El importe indicado: " + importe + " NO es igual al montoDebito: " + montodebito, False)
        else:
            try:
                montodebito = driver.find_element(By.CSS_SELECTOR, '#r1>td:nth-child(5)').get_attribute("textContent")
                assert montodebito == importe
                print('El importe indicado: '+importe+' es igual al montoDebito: '+montodebito)
                self.reporte.agregarLogCondicionalConScreenCrecer("El importe indicado: " + importe + " es igual al montoDebito: " + montodebito, True, cxc)
            except:
                print('No se ha encontrado el montoDebito correspondiente')
                self.reporte.agregarLogCondicional("El importe indicado: " + importe + " NO es igual al montoDebito: " + montodebito, False)

        driver.switch_to.window(cxc)
        driver.close()
        driver.switch_to.window(t24)
        driver.close()
        driver.quit()

    def buscarCuentasPorCuentaSBloqueo(self):
        driver = webdriver.Chrome(executable_path=r"\\sfs-1\Testing\Automatizacion_de_Proyectos\Herramientas\chromedriver 75\chromedriver.exe")

        driver.get("http://sapa2101lx-vip.bancocredicoop.coop/t24ts702/servlet/BrowserServlet")
        driver.maximize_window()
        driver.find_element(By.CSS_SELECTOR, '#signOnName').send_keys('F00402')
        driver.find_element(By.CSS_SELECTOR, '#sign-in').click()
        self.reporte.agregarLogCondicional("Se inicia sesion en Crecer", True)
        time.sleep(2)

        frame = driver.find_element(By.CSS_SELECTOR, "frame[id^='banner']")
        driver.switch_to.frame(frame)
        driver.maximize_window()
        time.sleep(2)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@id='commandValue']"))).send_keys('?328')
        driver.find_element(By.CSS_SELECTOR, '#cmdline_img').click()
        time.sleep(2)

        #Posicionamiento sobre la ventana T24
        t24 = driver.window_handles[1]
        driver.switch_to.window(t24)
        driver.maximize_window()
        driver.find_element(By.CSS_SELECTOR, '#pane_>ul:nth-child(4)>li>span').click()
        driver.find_element(By.CSS_SELECTOR, '#pane_>ul:nth-child(4)>li>ul>li:nth-child(2)>span').click()
        driver.find_element(By.CSS_SELECTOR, "a[onclick*='Consulta de Cuentas por Cuenta']").click()

        #Posicionamiento sobre la ventana Consulta de Cuentas por Cuenta
        cxc = driver.window_handles[2]
        driver.switch_to.window(cxc)
        driver.maximize_window()
        driver.find_element(By.CSS_SELECTOR, "[name='value:1:1:1']").clear()

        datoSeleccionado = []
        try:
            path = os.getcwd()
            pathArchivo = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\datosValidacionCrecer.csv"
            with open(pathArchivo, newline='') as File:
                reader = csv.reader(File)
                for row in reader:
                    cuenta = row[0]
                    datoSeleccionado.append(cuenta)
            print("Cuenta para validar en Crecer: " + str(row[0]))
        except:
            print("Hubo un error al leer el archivo.")

        driver.find_element(By.CSS_SELECTOR, "[name='value:1:1:1']").send_keys(datoSeleccionado)
        driver.find_element(By.CSS_SELECTOR, "a[alt='Ejecutar']").click()
        time.sleep(2)
        driver.find_element(By.CSS_SELECTOR, "img[alt='Select Drilldown']").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//span[normalize-space()='Datos Impositivos y Estado Cuenta']").click()

        try:
            driver.find_element(By.XPATH, "//a[@id='fieldCaption:L.AC.TIP.BLO']").is_displayed()
        except:
            print('Tipo de bloqueo no encontrado.')

        # Validacion sin bloqueo
        try:
            bloqueoparcial = driver.find_element(By.CSS_SELECTOR, '#enri_L\.AC\.TIP\.BLO').get_attribute(
                "textContent")
            assert bloqueoparcial == ""
            print('No hay tipo de bloqueo.')
            self.reporte.agregarLogCondicionalConScreenCrecer("No hay numero ni tipo de bloqueo.", True, cxc)
        except:
            print('Se ha encontrado un tipo de bloqueo.')
            self.reporte.agregarLogCondicional("Tipo de bloqueo es: " + bloqueoparcial, False)

        driver.switch_to.window(cxc)
        driver.close()
        driver.switch_to.window(t24)
        driver.close()
        driver.quit()
