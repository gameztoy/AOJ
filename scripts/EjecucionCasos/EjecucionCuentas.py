import os

from modulos.CuentasCrecer import CuentasEmbargables

#Obtenemos el path del archivo a leer
path = os.getcwd()
pathArchivo = path.split('automatizacionaoj')[0]+"automatizacionaoj\\recursos\\cuitSuc2.csv"

cuentasCrecer = CuentasEmbargables()

# Buscamos las cuentas no embargables en todos los cuits del archivo
cuentasCrecer.buscarCuentasPorArchivo(pathArchivo)

# Buscamos las cuentas no embargables para una sola persona
#cuentasCrecer.buscarCuentasPorCUIT("27203123995")

