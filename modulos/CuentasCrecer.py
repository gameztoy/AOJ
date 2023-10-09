import csv
import os
import sys
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

class CuentasEmbargables:

    def __init__(self):
        self.listaCuentasNOEMB = ['6050.1070', '6001.6035', '6001.6037','1040.2100','1015.1015'
                                  '6001.6200', '6001.6201', '6001.6202','6001.6203','1005.1003'
                                  '001.6204', '1040.2110', '1040.2111', '1040.2112','1019.1065'
                                  '1040.2113', '1040.2114', '6001.6055', '6001.6066','8015.8015'
                                  '6001.6051', '6001.6056', '6001.6057', '6001.6070','6001.6050']

    def buscarCuentasPorArchivo(self, pathArchivo):
        chrome_options = Options()
        chrome_options.add_argument("--headless")

        driver = webdriver.Chrome(options=chrome_options,
                                  executable_path=r"\\sfs-1\Testing\Automatizacion_de_Proyectos\Herramientas\chromedriver 75\chromedriver.exe")
        driver.get("http://sapa2101lx-vip.bancocredicoop.coop/t24ts702/servlet/BrowserServlet")
        driver.find_element(By.CSS_SELECTOR, '#signOnName').send_keys('ANOVELLO')
        driver.find_element(By.CSS_SELECTOR, '#sign-in').click()

        frame = driver.find_element(By.CSS_SELECTOR, "frame[id^='menu']")
        driver.switch_to.frame(frame)
        driver.find_element(By.CSS_SELECTOR, '#pane_>ul:nth-child(2)>li>span').click()
        driver.find_element(By.CSS_SELECTOR, '#pane_>ul:nth-child(2)>li>ul>li:nth-child(1)>span').click()
        window_crecer_principal = driver.window_handles[0]

        path = os.getcwd()
        pathArchivoSalida = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\cuentasValidas.csv"

        with open(pathArchivo, newline='') as File:
            reader = csv.reader(File)
            f = open(pathArchivoSalida, "w")
            writer = csv.writer(f, lineterminator='\n')
            writer.writerow(['CUIL', 'NroCuenta'])
            for row in reader:
                print(row)
                cuit = row[0]
                dniFirmante = row[0][2:10]

                driver.find_element(By.CSS_SELECTOR, "a[onclick*='Consulta de Cuentas por Firmante']").click()
                window_consulta_Cuentas_Firmante = driver.window_handles[1]
                driver.switch_to.window(window_consulta_Cuentas_Firmante)
                driver.find_element(By.CSS_SELECTOR, "[name='value:1:1:1']").clear()
                driver.find_element(By.CSS_SELECTOR, "[name='value:1:1:1']").send_keys(dniFirmante)
                driver.find_element(By.CSS_SELECTOR, "a[alt='Ejecutar']").click()

                listaDeCuentas = driver.find_elements(By.CSS_SELECTOR, '#datadisplay>tbody>tr')
                cuentasValida = []
                for elementoLista in listaDeCuentas:
                    campos = elementoLista.find_elements(By.CSS_SELECTOR, 'td')

                    # Si la cuenta se encuentra en la lista de cuentas no embargables o esta cerrada
                    if (campos[2].text in self.listaCuentasNOEMB or campos[5].text == "Cerrada"):
                        if (campos[5].text == "Cerrada"):
                            print(campos[0].text, ' - ', campos[2].text, '--CUENTA CERRADA--')
                        else:
                            print(campos[0].text, ' - ', campos[2].text, '--CUENTA NO EMBARGABLE--')
                    else:
                        cuentasValida.append(campos[0].text)

                if len(cuentasValida) > 0:
                    print('Cuentas Validas para el cuit: ', cuit)
                else:
                    print(cuit, ' no posee cuentas Embargables')

                for cuenta in cuentasValida:
                    print(cuenta)
                    writer.writerow([cuit, cuenta])

                print("\n")

                driver.close()
                driver.switch_to.window(window_crecer_principal)
                frame = driver.find_element(By.CSS_SELECTOR, "frame[id^='menu']")
                driver.switch_to.frame(frame)

        driver.close()

    def buscarCuentasPorCUIT(self, cuit):
        chrome_options = Options()
        chrome_options.add_argument("--headless")

        driver = webdriver.Chrome(options=chrome_options,
                                  executable_path=r"\\sfs-1\Testing\Automatizacion_de_Proyectos\Herramientas\chromedriver 75\chromedriver.exe")
        print("Inicia busqueda de Cuentas por CUIT en Crecer...")
        driver.get("http://sapa2101lx-vip.bancocredicoop.coop/t24ts702/servlet/BrowserServlet")
        driver.maximize_window()
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR, '#signOnName').send_keys('ANOVELLO')
        driver.find_element(By.CSS_SELECTOR, '#sign-in').click()

        frame = driver.find_element(By.CSS_SELECTOR, "frame[id^='menu']")
        driver.switch_to.frame(frame)
        driver.find_element(By.CSS_SELECTOR, '#pane_>ul:nth-child(2)>li>span').click()
        driver.find_element(By.CSS_SELECTOR, '#pane_>ul:nth-child(2)>li>ul>li:nth-child(1)>span').click()
        window_crecer_principal = driver.window_handles[0]

        path = os.getcwd()
        pathArchivoSalida = path.split('automatizacionaoj')[0] + "automatizacionaoj\\recursos\\cuentasValidas.csv"

        f = open(pathArchivoSalida, "w")
        writer = csv.writer(f, lineterminator='\n')

        # Se busca a la persona por CUIT/CUIL
        driver.find_element(By.CSS_SELECTOR, "a[onclick*='Consulta de Cuentas por Firmante']").click()
        window_consulta_Cuentas_Firmante = driver.window_handles[1]
        driver.switch_to.window(window_consulta_Cuentas_Firmante)
        driver.maximize_window()
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR, "[name='value:1:1:1']").clear()
        time.sleep(2)
        driver.find_element(By.CSS_SELECTOR, "[name='value:1:1:1']").send_keys(cuit)
        time.sleep(2)
        driver.find_element(By.CSS_SELECTOR, "a[alt='Ejecutar']").click()

        listaDeCuentas = driver.find_elements(By.CSS_SELECTOR, '#datadisplay>tbody>tr')
        cuentasValida = []
        for elementoLista in listaDeCuentas:
            campos = elementoLista.find_elements(By.CSS_SELECTOR, 'td')

            # Si la cuenta se encuentra en la lista de cuentas no embargables o esta cerrada
            if (campos[2].text in self.listaCuentasNOEMB or campos[5].text == "Cerrada"):
                if (campos[5].text == "Cerrada"):
                    print(campos[0].text, ' - ', campos[2].text, '--CUENTA CERRADA--')
                else:
                    print(campos[0].text, ' - ', campos[2].text, '--CUENTA NO EMBARGABLE--')
            else:
                cuentasValida.append(campos[0].text)

        if len(cuentasValida) > 0:
            for cuenta in cuentasValida:
                print('Para el CUIT: '+cuit, 'se encontro la cuenta: '+cuenta)
                writer.writerow([cuit, cuenta])
        else:
            print(cuit, ' no posee cuentas Embargables')
            sys.exit(1)

        f.close()
        driver.close()
        driver.switch_to.window(window_crecer_principal)
        frame = driver.find_element(By.CSS_SELECTOR, "frame[id^='menu']")
        driver.switch_to.frame(frame)
        driver.close()
