

#Detect if the package is installed
import pkg_resources
import subprocess

def check_and_install_packages(package_list):
    installed_packages = [pkg.key for pkg in pkg_resources.working_set]
    for package_name in package_list:
        if package_name in installed_packages:
            print(f"{package_name} está instalado.")
        else:
            print(f"{package_name} no está instalado. Instalando...")
            !pip install {package_name}
            print(f"{package_name} se ha instalado correctamente.")

# Ejemplo de uso
packages = ['selenium', 'requests', 'webdriver-manager', 'win10toast'] #Añadir aquí los paquetes necesarios
check_and_install_packages(packages)

#!pip install selenium
#!pip install requests
#!pip install webdriver-manager
#!pip install win10toast

#import libraries
from selenium import webdriver
import locale

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
import time

#get notifications
from win10toast import ToastNotifier
toaster = ToastNotifier()

#P
from datetime import datetime
from pandas.tseries.holiday import AbstractHolidayCalendar, Holiday
import pandas as pd
import datetime
from datetime import date
from datetime import datetime

from pandas.tseries.holiday import *
from pandas.tseries.offsets import CustomBusinessDay
from pandas.tseries.offsets import MonthBegin

#Función para avanzar al siguiente lunes
def strict_next_monday(dt: datetime) -> datetime:
    if dt.weekday() > 0:
        return dt + timedelta(7-dt.weekday())
    return dt

#Colombian hollydais
class ColombianBusinessCalendar(AbstractHolidayCalendar):
    rules = [
        # festivos fijos
        Holiday('Año nuevo', month=1, day=1),
        Holiday('Día del trabajo', month=5, day=1),
        Holiday('Día de la independencia', month=7, day=20),
        Holiday('Batalla de Boyacá', month=8, day=7),
        Holiday('Inmaculada Concepción', month=12, day=8),
        Holiday('Navidad', month=12, day=25),
        # festivos relativos a la pascua
        Holiday('Jueves santo', month=1, day=1, offset=[Easter(), Day(-3)]),
        Holiday('Viernes santo', month=1, day=1, offset=[Easter(), Day(-2)]),
        Holiday('Ascención de Jesús', month=1, day=1, offset=[Easter(), Day(43)]),
        Holiday('Corpus Christi', month=1, day=1, offset=[Easter(), Day(64)]),
        Holiday('Sagrado Corazón de Jesús', month=1, day=1, offset=[Easter(), Day(71)]),
        # festivos desplazables (Emiliani)
        Holiday('Epifanía del señor', month=1, day=6, observance=strict_next_monday),
        Holiday('Día de San José', month=3, day=19, observance=strict_next_monday),
        Holiday('San Pedro y San Pablo', month=6, day=29, observance=strict_next_monday),
        Holiday('Asunción de la Virgen', month=8, day=15, observance=strict_next_monday),
        Holiday('Día de la raza', month=10, day=12, observance=strict_next_monday),
        Holiday('Todos los santos', month=11, day=1, observance=strict_next_monday),
        Holiday('Independencia de Cartagena', month=11, day=11, observance=strict_next_monday)
    ]
    
    def isHoliday(self, dt: datetime) -> bool:
        """
        Check if the given date is a holiday.
        """
        return self.holidays(start=dt, end=dt, return_name=True).size > 0

calendar_col = ColombianBusinessCalendar()

holidays = calendar_col.holidays(datetime(2023, 1, 1), end=datetime(2023, 12, 31)) #Cambiar cuando se acabe el año 
festivos_col = [holiday.date() for holiday in holidays]

import calendar
import numpy as np
day = datetime.now() #traigo el día
day.strftime("%d")
day_format = day.strftime("%d") #numero de día
day_format_01=day_format
day_format=int(day_format) #formato entero

month_format = day.strftime("%m") #numero de día
month_format_01=month_format
month_format=int(month_format)
year_format = day.strftime("%Y") #numero de día
year_format=int(year_format)





cal= calendar.Calendar() #traigo calendario
cal.setfirstweekday(calendar.SUNDAY) #para que empiece en domingo (0,0) en python (1,1 ) en Adminfo
A=(cal.monthdayscalendar(year_format, month_format)) # traigo semana en formato lista



A=np.array(A) #convierto a matriz con numpy para poder trabajar



W=np.where(A == day_format) #busco la posición de el día actual en la matrix anterior
                                                                                                                                                         
fila_dia=str(int(W[0])+1) #acomodo idex python con index adminfo
columna_dia=str(int(W[1])+1)

W_1=np.where(A == 1) #busca la posición del primer día
fila_primer_dia=str(int(W_1[0])+1) #acomodo idex python con index adminfo
columna_primer_dia=str(int(W_1[1])+1)

locale.setlocale(locale.LC_ALL, 'es-ES') 

import os
try:
    os.remove("//server_rute\\operation\\"+str(year_format)+"\\"+str(month_format_01)+"."+str(day.strftime("%B").capitalize())+" "+str(year_format)+"\\Creación listas de marcación\\exclusiones_seguimiento\\adminfo_seguimiento_"+str(year_format)+str(month_format_01)+str(day_format_01)+".zip")
    os.remove("//server_rute\\operation\\"+str(year_format)+"\\"+str(month_format_01)+"."+str(day.strftime("%B").capitalize())+" "+str(year_format)+"\\Creación listas de marcación\\exclusiones_seguimiento\\informe_compromiso_"+str(year_format)+str(month_format_01)+str(day_format_01)+".zip")

except: 
    print('continuemos')

from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
chrome_options = webdriver.ChromeOptions()

dt_obj = datetime.now()
dt_obj.strftime("%Y%m%d")
date_format = dt_obj.strftime("%Y%m%d")


locale.setlocale(locale.LC_ALL, '')
prefs = {'download.default_directory':
                                        "//server_rute\\operation\\\\"+str(year_format)+"\\"+str(month_format_01)+"."+str(day.strftime("%B").capitalize())+" "+str(year_format)+"\\Creación listas de marcación\\exclusiones_seguimiento"}
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_experimental_option('prefs', prefs)


chrome_options.page_load_strategy = 'none'
driver = webdriver.Chrome(#desired_capabilities=capa,
                          service=Service(ChromeDriverManager().install()),
                          options=chrome_options)
#driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

#abrir pagina

driver.get('https://adminfo.apps.bancolombia.com/')

try: 
    #ingresar usuario 
    WebDriverWait(driver, 500).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div/form/div[3]/div[2]/input"))).send_keys('ingridri')

    #ingresar contraseña
    WebDriverWait(driver, 500).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div/form/div[3]/div[4]/input"))).send_keys('$Ingrid#_Lyam0728++')


    #dar click en el boton. entrar
    WebDriverWait(driver, 500).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div/form/div[3]/div[5]/button"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de ingresar al aplicativo' : {str(e)}", duration=10, threaded=True)






#Hacer click en botón Gestión
try:
    WebDriverWait(driver, 500).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[3]/div/a[2]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en el botón 'Gestión' : {str(e)}", duration=10, threaded=True)

#dar click en Informes de agentes externos
try:
    WebDriverWait(driver, 500).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div[3]/table/tbody/tr[2]/td[2]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Informes de agentes externos' : {str(e)}", duration=10, threaded=True)

#Click en pestaña de informes
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/ul/li[1]/a/span/p/span"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Informes' : {str(e)}", duration=10, threaded=True)

#Click en resumen de seguimiento
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/ul/li[1]/ul/li[2]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Resumen de seguimiento' : {str(e)}", duration=10, threaded=True)

#Click en agrupador
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[1]/div[1]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Agrupador' : {str(e)}", duration=10, threaded=True)

#Click en detalle en pantalla
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[1]/div[1]/select/option[15]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Detalle en pantalla' : {str(e)}", duration=10, threaded=True)

#Click en Columna
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[2]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Columna' : {str(e)}", duration=10, threaded=True)

#Click en grabador
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[2]/select/option[14]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Grabador' : {str(e)}", duration=10, threaded=True)

#Click en Operador
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[4]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Operador' : {str(e)}", duration=10, threaded=True)

#Click en Mayor o Igual Que
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[4]/select[1]/option[8]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Mayor o igual que' : {str(e)}", duration=10, threaded=True)

#Escribir en valor 1
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[6]/input"))).send_keys(1)
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de escribir '1' : {str(e)}", duration=10, threaded=True)

#Click en + para agregar criterio
try:
    WebDriverWait(driver, 500).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[7]/a[1]/span"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en '+ para agregar criterio' : {str(e)}", duration=10, threaded=True)


#Click en Generar Informe
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[4]/a[3]/i"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Generar Informe' : {str(e)}", duration=10, threaded=True)

#Click en Agregar todos
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[2]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Agregar todos' : {str(e)}", duration=10, threaded=True)

#Fase 1: Adminfo Seguimiento, vamos a seleccionar los que no quiero traer

#Click en Nro Documento                                              
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[5]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Número Documento' : {str(e)}", duration=10, threaded=True)

#Click en Días Mora                                                    
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Días Mora' : {str(e)}", duration=10, threaded=True)

#Click en Cuotas Vencidas     
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Cuotas Vencidas' : {str(e)}", duration=10, threaded=True)

#Click en Asesor                                                        
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Asesor' : {str(e)}", duration=10, threaded=True)

#Click en Regional                
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Regional' : {str(e)}", duration=10, threaded=True)

#Click en Tipo 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Tipo' : {str(e)}", duration=10, threaded=True)

#Click en Teléfono 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Teléfono' : {str(e)}", duration=10, threaded=True)

#Click en Nombre Abogado
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[9]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Nombre Abogado' : {str(e)}", duration=10, threaded=True)

#Click en Último Abogado 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[9]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Último Abogado' : {str(e)}", duration=10, threaded=True)

#Click en Fecha Mejor Código De Gestión 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Fecha Mejor Código De Gestión' : {str(e)}", duration=10, threaded=True)

#Click en Mejor Código De Gestión 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Mejor Código De Gestión' : {str(e)}", duration=10, threaded=True)

#Click en Cuenta 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Cuenta' : {str(e)}", duration=10, threaded=True)

#Click en Descripción Mejor Código de Gestión 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Descripción Mejor Código de Gestión' : {str(e)}", duration=10, threaded=True)

#Click en Consecutivo Demanda
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Consecutivo Demanda' : {str(e)}", duration=10, threaded=True)

#Click en Estado Producto 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Estado Producto' : {str(e)}", duration=10, threaded=True)

#Click en Estado Producto     
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Estado Producto' : {str(e)}", duration=10, threaded=True)

#Click en Desc.Causal Rechazo  
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Desc.Causal Rechazo' : {str(e)}", duration=10, threaded=True)

#Click en Fecha de Promesa 2    
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Fecha de Promesa 2' : {str(e)}", duration=10, threaded=True)

#Click en Fecha Venci    
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Fecha Venci' : {str(e)}", duration=10, threaded=True)

#Click en Nit Deudor   
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Nit Deudor' : {str(e)}", duration=10, threaded=True)

#Click en Nit Compañia  
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Nit Compañia' : {str(e)}", duration=10, threaded=True)

#Click en Nit Tercero      
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Nit Tercero' : {str(e)}", duration=10, threaded=True)

#Click en Ocupación
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Ocupación' : {str(e)}", duration=10, threaded=True)

#Click en Profesión 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Profesión' : {str(e)}", duration=10, threaded=True)

#Click en T_compromi   
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'T_compromi' : {str(e)}", duration=10, threaded=True)

#Click en T_entrada      
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'T_entrada' : {str(e)}", duration=10, threaded=True)

#Click en Hora Grabación    
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Hora Grabación' : {str(e)}", duration=10, threaded=True)

#Click en T_igraba       
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'T_igraba' : {str(e)}", duration=10, threaded=True)

#Click en Sitio 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Sitio' : {str(e)}", duration=10, threaded=True)

#Click en Lote         
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Lote' : {str(e)}", duration=10, threaded=True)

#Click en Hora De Compromiso       
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Hora De Compromiso' : {str(e)}", duration=10, threaded=True)

#Click en Cobrador     
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    # Si ocurre una excepción, muestra una notificación de error
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Cobrador' : {str(e)}", duration=10, threaded=True)

#Click en Código De Cobro Anterior 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Código De Cobro Anterior' : {str(e)}", duration=10, threaded=True)

#Click en Codcob_mec_no     
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Codcob_mec_no' : {str(e)}", duration=10, threaded=True)

#Click en Codcob_mec_nor_tra       
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Codcob_mec_nor_tra' : {str(e)}", duration=10, threaded=True)

#Click en Codcob_mec_nor_util       
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Codcob_mec_nor_util' : {str(e)}", duration=10, threaded=True)

#Click en Código De Cobro Anterior
#WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[2]/ul/li[16]/a"))).click()

#Click en Codigo Causal      
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Codigo Causal' : {str(e)}", duration=10, threaded=True)

#Click en Descripción Causal   
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[13]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Descripción Causal' : {str(e)}", duration=10, threaded=True)

#Click en Código De Contacto    
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[13]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Código De Contacto' : {str(e)}", duration=10, threaded=True)

#Click en Control        
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[13]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Agregar todos' : {str(e)}", duration=10, threaded=True)

#Click en Consdirec        
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[13]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Consdirec' : {str(e)}", duration=10, threaded=True)

#Click en T_entrada
#WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[2]/ul/li[9]/a"))).click()

#Click en Cuadrante         
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[14]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'cuadrante' : {str(e)}", duration=10, threaded=True)

#Click en Fecha De Actuación      
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[14]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Fecha de Actuación' : {str(e)}", duration=10, threaded=True)

#Click en Fecha Corte        
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[14]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Fecha Corte' : {str(e)}", duration=10, threaded=True)

#Click en Codigo Causal Rechazo       
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[16]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Codigo Causal Rechazo' : {str(e)}", duration=10, threaded=True)

#Click en Generar Informe
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/div/a[3]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Generar Informe' : {str(e)}", duration=10, threaded=True)

#Click en Excel
try:
    WebDriverWait(driver, 1000).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/div/span[1]/a[1]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Excel' : {str(e)}", duration=10, threaded=True)

#Borra campos de Nombre Archivo
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/div/div[3]/form/div[2]/div[1]/input"))).clear()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de borrar el nombre por defecto : {str(e)}", duration=10, threaded=True)

#Nombre de Adminfo_Seguimiento+Fecha
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/div/div[3]/form/div[2]/div[1]/input"))).send_keys(["adminfo_seguimiento_" +str(date_format)])
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de agregar el nuevo nombre : {str(e)}", duration=10, threaded=True)



#Click en Exportar
try:
    WebDriverWait(driver, 1000).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/div/div[3]/form/div[4]/input[1]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de exportar el archivo : {str(e)}", duration=10, threaded=True)

toaster = ToastNotifier()

# Muestra la notificación
toaster.show_toast("Bot de Adminfo", "Ya se está descargando el archivo de Adminfo_Seguimiento ", duration=10)



###### Fase 2: Informe compromisos 

#Click en pestaña Informes
try:
    WebDriverWait(driver, 0).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/ul/li[1]/a/span/p/span"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en la pestaña 'Informes' : {str(e)}", duration=10, threaded=True)

#Click en Informes de Compromisos
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/ul/li[1]/ul/li[1]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en 'Informes de Compromisos' : {str(e)}", duration=10, threaded=True)

#Click en Agrupar Por
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[1]/div[1]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en 'Agrupar Por' : {str(e)}", duration=10, threaded=True)

#Click en Detalle en Pantalla
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[1]/div[1]/select/option[10]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en 'Detalle en pantalla' : {str(e)}", duration=10, threaded=True)

#Click en Columna
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[2]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en 'Columna' : {str(e)}", duration=10, threaded=True)

#Click en Fecha Generacion
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[2]/select/option[7]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en 'Fecha Generacion' : {str(e)}", duration=10, threaded=True)

#Click en icono de calendario 
A0=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[2]")))

try:
    A0.click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en el Calendario : {str(e)}", duration=10, threaded=True)

#Selecciona la fecha
#A=WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div[2]/button[1]')))
hour=day.hour 
if hour<9:
    if datetime.now().strftime('%A') == 'lunes':
        A=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/table/tbody/tr['+str(int(fila_dia)-1)+']/td['+str(int( columna_dia)+5)+']')))
    elif datetime.now().date() - timedelta(days=1) in festivos_col:
        if datetime.now().strftime('%A') == 'martes':
            A=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/table/tbody/tr['+str(int(fila_dia)-1)+']/td['+str(int( columna_dia)+4)+']')))
        else: 
            A=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/table/tbody/tr['+str(fila_dia)+']/td['+str(int( columna_dia)-2)+']')))
    else:    
        A=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/table/tbody/tr['+str(fila_dia)+']/td['+str(int( columna_dia)-1)+']')))
else: 
    A=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/table/tbody/tr['+str(fila_dia)+']/td['+str(columna_dia)+']')))

A.click()   
    #driver.find_element("xpath", '//*[@id="ui-datepicker-div"]/table/tbody/tr['+str(fila_dia)+']/td['+str(columna_dia)+']').click() #/html/body/div[6]/div[2]/button[1] A.click()

B=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[1]"))) 
A1= WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div[2]/button[2]')))
   
try:
   
   B.click()
except:
        A1.click()
        B.click()   

#Click en Generar Informe
C=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/div/a[3]/i")))
try:
   
   C.click()
except:
        A1.click()
        C.click()   

#Click en Generar Informe
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/a[3]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en 'Generar Informe' : {str(e)}", duration=10, threaded=True)

#Click en CSV
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/span[1]/a[3]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en 'CSV' : {str(e)}", duration=10, threaded=True)

#Borra los campos de Nombre Archivo
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/div[2]/form/div[2]/div[1]/input"))).clear()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de borrar el nombre por defecto : {str(e)}", duration=10, threaded=True)

#Coloca Informe_Compromiso_ + Fecha
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/div[2]/form/div[2]/div[1]/input"))).send_keys(["informe_compromiso_" +str(date_format)])
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de poner el nuevo nombre al archivo : {str(e)}", duration=10, threaded=True)

#Click en Exportar
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/div[2]/form/div[3]/input[1]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en 'Exportar' : {str(e)}", duration=10, threaded=True)

#Cerrar el navegador Google Chrome
#driver.quit()

#Fin de la descarga
toaster.show_toast("Bot de Adminfo", "Ya se está descargando el archivo de Informe_Compromiso", duration=10)




###### FASE 3: Adminfo Seguimiento Mensual

try:
    os.remove("//RFNTRFS02/Analitica$/Servicing\\Bancolombia\\"+str(year_format)+"\\"+str(month_format_01)+"."+str(day.strftime("%B").capitalize())+" "+str(year_format)+"\\Creación listas de marcación\\exclusiones_seguimiento\\adminfo_seguimiento_"+str(day.strftime("%B").capitalize())+".zip")
except: 
    print('continuemos')

#Click en pestaña de informes
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/ul/li[1]/a/span/p/span"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Informes' : {str(e)}", duration=10, threaded=True)

#Click en resumen de seguimiento
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/ul/li[1]/ul/li[2]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Resumen de seguimiento' : {str(e)}", duration=10, threaded=True)

#Click en agrupador
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[1]/div[1]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Agrupador' : {str(e)}", duration=10, threaded=True)

#Click en detalle en pantalla
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[1]/div[1]/select/option[15]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Detalle en pantalla' : {str(e)}", duration=10, threaded=True)

#Click en Columna
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[2]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Columna' : {str(e)}", duration=10, threaded=True)

#Click en grabador
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[2]/select/option[14]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Grabador' : {str(e)}", duration=10, threaded=True)

#Click en Operador
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[4]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Operador' : {str(e)}", duration=10, threaded=True)

#Click en Mayor o Igual Que
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[4]/select[1]/option[8]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Mayor o igual que' : {str(e)}", duration=10, threaded=True)

#Escribir en valor 1
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[6]/input"))).send_keys(1)
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de escribir '1' : {str(e)}", duration=10, threaded=True)

#Click en + para agregar criterio
try:
    WebDriverWait(driver, 500).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[7]/a[1]/span"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en '+ para agregar criterio' : {str(e)}", duration=10, threaded=True)

#Click en Columna 
try: 
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[2]/div/a"))).click()    
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'columna' : {str(e)}", duration=10, threaded=True)    

#Click en Fecha Gestion 
try: 
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[11]/ul/li[6]/div"))).click()    
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Fecha Gestión' : {str(e)}", duration=10, threaded=True) 

#Click en Operador                                                      
try: 
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[4]/div/a"))).click()    
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Operador' : {str(e)}", duration=10, threaded=True)    

#Click en Mayor o Igual que 
try: 
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[11]/ul/li[8]/div"))).click()    
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'mayor o igual que' : {str(e)}", duration=10, threaded=True) 

#Click en Calendario                                                                                                                              
A0=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[7]/a[2]"))) 
try: 
    A0.click()    
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Calendario' : {str(e)}", duration=10, threaded=True) 

#Click en primer dia del mes
A=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/table/tbody/tr['+str(int(fila_primer_dia))+']/td['+str(int( columna_primer_dia)-1)+']/a')))
try: 
    A.click()   
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'El primer día del mes' : {str(e)}", duration=10, threaded=True)

#Click en Cerrar Calendario /html/body/div[6]/div[2]/button[2]
#WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[6]/div[2]/button[2]"))).click()
#WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[6]/div[2]/button[2]"))).click()

#Click en + Agregar criterios
#WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[7]/a[1]/span"))).click()    

#Cierra el calendario y da click en +
B=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[7]/a[1]/span"))) 
A1= WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div[2]/button[2]')))
   
try:
   
   B.click()
except:
        A1.click()
        B.click()

#Click en Operador               
try: 
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[4]/div/a"))).click()    
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Operador' : {str(e)}", duration=10, threaded=True)    

#Click en Menor o Igual que 
D=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[11]/ul/li[9]/div")))   
try: 
    D.click()   
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Menor o Igual que ' : {str(e)}", duration=10, threaded=True)

#Click en Calendario                                                                                                                              
A0=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[7]/a[2]")))
try: 
    A0.click()    
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Calendario' : {str(e)}", duration=10, threaded=True)

#Click en dia actual 
A=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/table/tbody/tr['+str(int(fila_dia))+']/td['+str(int( columna_dia)-1)+']/a')))
try: 
    A.click()    
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Día actual' : {str(e)}", duration=10, threaded=True)

#Click en Cerrar Calendario /html/body/div[6]/div[2]/button[2]
#WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[6]/div[2]/button[2]"))).click()
#WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[6]/div[2]/button[2]"))).click()

#Click en + Agregar criterios
#WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[7]/a[1]/span"))).click() 

#Cierra el calendario y da click en +
B=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div/div[1]/div/div[1]/span/table/tbody/tr/td[7]/a[1]/span"))) 
A1= WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div[2]/button[2]')))
   
try:
   
   B.click()
except:
        A1.click()
        B.click()

#Click en Generar Informe 
try: 
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[1]/div[1]/a[1]"))).click()      
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Generar Informe' : {str(e)}", duration=10, threaded=True)

#Click en Agregar todos
try: 
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[2]/a"))).click()    
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Agregar todos' : {str(e)}", duration=10, threaded=True)

#Click en Nro Documento                                              
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[5]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Número Documento' : {str(e)}", duration=10, threaded=True)


#Click en Días Mora                                                    
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Días Mora' : {str(e)}", duration=10, threaded=True)

#Click en Cuotas Vencidas     
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Cuotas Vencidas' : {str(e)}", duration=10, threaded=True)

#Click en Asesor                                                        
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Asesor' : {str(e)}", duration=10, threaded=True)

#Click en Regional                
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Regional' : {str(e)}", duration=10, threaded=True)

#Click en Tipo 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Tipo' : {str(e)}", duration=10, threaded=True)

#Click en Teléfono 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[7]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Teléfono' : {str(e)}", duration=10, threaded=True)

#Click en Nombre Abogado
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[9]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Nombre Abogado' : {str(e)}", duration=10, threaded=True)

#Click en Último Abogado 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[9]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Último Abogado' : {str(e)}", duration=10, threaded=True)

#Click en Fecha Mejor Código De Gestión 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Fecha Mejor Código De Gestión' : {str(e)}", duration=10, threaded=True)

#Click en Mejor Código De Gestión 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Mejor Código De Gestión' : {str(e)}", duration=10, threaded=True)

#Click en Cuenta 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Cuenta' : {str(e)}", duration=10, threaded=True)

#Click en Descripción Mejor Código de Gestión 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Descripción Mejor Código de Gestión' : {str(e)}", duration=10, threaded=True)

#Click en Consecutivo Demanda
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Consecutivo Demanda' : {str(e)}", duration=10, threaded=True)

#Click en Estado Producto 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Estado Producto' : {str(e)}", duration=10, threaded=True)

#Click en Estado Producto     
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Estado Producto' : {str(e)}", duration=10, threaded=True)

#Click en Desc.Causal Rechazo  
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Desc.Causal Rechazo' : {str(e)}", duration=10, threaded=True)

#Click en Fecha de Promesa 2    
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Fecha de Promesa 2' : {str(e)}", duration=10, threaded=True)

#Click en Fecha Venci    
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Fecha Venci' : {str(e)}", duration=10, threaded=True)

#Click en Nit Deudor   
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Nit Deudor' : {str(e)}", duration=10, threaded=True)

#Click en Nit Compañia  
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Nit Compañia' : {str(e)}", duration=10, threaded=True)

#Click en Nit Tercero      
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Nit Tercero' : {str(e)}", duration=10, threaded=True)

#Click en Ocupación
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Ocupación' : {str(e)}", duration=10, threaded=True)

#Click en Profesión 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Profesión' : {str(e)}", duration=10, threaded=True)

#Click en T_compromi   
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'T_compromi' : {str(e)}", duration=10, threaded=True)

#Click en T_entrada      
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'T_entrada' : {str(e)}", duration=10, threaded=True)

#Click en Hora Grabación    
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Hora Grabación' : {str(e)}", duration=10, threaded=True)

#Click en T_igraba       
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'T_igraba' : {str(e)}", duration=10, threaded=True)

#Click en Sitio 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Sitio' : {str(e)}", duration=10, threaded=True)

#Click en Lote         
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Lote' : {str(e)}", duration=10, threaded=True)

#Click en Hora De Compromiso       
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Hora De Compromiso' : {str(e)}", duration=10, threaded=True)

#Click en Cobrador     
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    # Si ocurre una excepción, muestra una notificación de error
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento", f"Se ha producido un error al momento de dar click en 'Cobrador' : {str(e)}", duration=10, threaded=True)

#Click en Código De Cobro Anterior 
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Código De Cobro Anterior' : {str(e)}", duration=10, threaded=True)

#Click en Codcob_mec_no     
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Codcob_mec_no' : {str(e)}", duration=10, threaded=True)

#Click en Codcob_mec_nor_tra       
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Codcob_mec_nor_tra' : {str(e)}", duration=10, threaded=True)

#Click en Codcob_mec_nor_util       
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Codcob_mec_nor_util' : {str(e)}", duration=10, threaded=True)

#Click en Codigo Causal      
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[12]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Codigo Causal' : {str(e)}", duration=10, threaded=True)

#Click en Descripción Causal   
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[13]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Descripción Causal' : {str(e)}", duration=10, threaded=True)

#Click en Código De Contacto    
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[13]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Código De Contacto' : {str(e)}", duration=10, threaded=True)

#Click en Control        
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[13]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Agregar todos' : {str(e)}", duration=10, threaded=True)

#Click en Consdirec        
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[13]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Consdirec' : {str(e)}", duration=10, threaded=True)

#Click en Cuadrante         
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[14]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'cuadrante' : {str(e)}", duration=10, threaded=True)

#Click en Fecha De Actuación      
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[14]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Fecha de Actuación' : {str(e)}", duration=10, threaded=True)

#Click en Fecha Corte        
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[14]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Fecha Corte' : {str(e)}", duration=10, threaded=True)

#Click en Codigo Causal Rechazo       
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/fieldset/div[3]/ul/li[16]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Codigo Causal Rechazo' : {str(e)}", duration=10, threaded=True)

#Click en Generar Informe
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/div/a[3]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Generar Informe' : {str(e)}", duration=10, threaded=True)

#Click en Excel
try:
    WebDriverWait(driver, 1000).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/div/span[1]/a[1]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de dar click en 'Excel' : {str(e)}", duration=10, threaded=True)

#Borra campos de Nombre Archivo
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/div/div[3]/form/div[2]/div[1]/input"))).clear()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de borrar el nombre por defecto : {str(e)}", duration=10, threaded=True) 

#Nombre de Adminfo_Seguimiento+Mes
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/div/div[3]/form/div[2]/div[1]/input"))).send_keys(["adminfo_seguimiento_" +str(day.strftime("%B").capitalize())]) 
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de agregar el nuevo nombre : {str(e)}", duration=10, threaded=True)

#Click en Exportar
try:
    WebDriverWait(driver, 1000).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div[2]/div[2]/div/div/div[3]/form/div[4]/input[1]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Adminfo Seguimiento Mes", f"Se ha producido un error al momento de exportar el archivo : {str(e)}", duration=10, threaded=True)

toaster = ToastNotifier()

# Muestra la notificación
toaster.show_toast("Bot de Adminfo", "Ya se está descargando el archivo de Adminfo_Seguimiento Mes", duration=10)



#### FASE 4: Informe Compromiso Mes

try:
    os.remove("//RFNTRFS02/Analitica$/Servicing\\Bancolombia\\"+str(year_format)+"\\"+str(month_format_01)+"."+str(day.strftime("%B").capitalize())+" "+str(year_format)+"\\Creación listas de marcación\\exclusiones_seguimiento\\informe_compromiso_"+str(day.strftime("%B").capitalize())+".zip")
except: 
    print('continuemos')

#Click en pestaña Informes
try:
    WebDriverWait(driver, 0).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/ul/li[1]/a/span/p/span"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en la pestaña 'Informes' : {str(e)}", duration=10, threaded=True)

#Click en Informes de Compromisos
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/ul/li[1]/ul/li[1]/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Informes de Compromisos' : {str(e)}", duration=10, threaded=True)

#Click en Agrupar Por
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[1]/div[1]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Agrupar Por' : {str(e)}", duration=10, threaded=True)

#Click en Detalle en Pantalla
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[1]/div[1]/select/option[10]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Detalle en pantalla' : {str(e)}", duration=10, threaded=True)

#Click en Columna
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[2]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en 'Columna' : {str(e)}", duration=10, threaded=True)

#Click en Fecha Generacion
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[2]/select/option[7]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Fecha Generacion' : {str(e)}", duration=10, threaded=True)

#Click en Operador
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[4]/div/a"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Fecha Generacion' : {str(e)}", duration=10, threaded=True)

#Click en Mayor o Igual que
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[11]/ul/li[8]/div"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Fecha Generacion' : {str(e)}", duration=10, threaded=True)

#Click en icono de calendario 
A0=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[2]")))
try:
    A0.click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso", f"Se ha producido un error al momento de dar click en el Calendario : {str(e)}", duration=10, threaded=True)

#Selecciona la fecha del primer día
A=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/table/tbody/tr['+str(int(fila_primer_dia))+']/td['+str(int( columna_primer_dia))+']/a')))
try:
    A.click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Fecha Generacion' : {str(e)}", duration=10, threaded=True)

#Click en + 
#WebDriverWait(driver, 500).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[1]"))).click()

#Cierra el calendario y da click en +
B=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[1]"))) 
A1= WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div[2]/button[2]')))
   
try:
   
   B.click()
except:
        A1.click()
        B.click()

#Click en Operador
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[4]/div/a"))).click()   
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Fecha Generacion' : {str(e)}", duration=10, threaded=True)

#Click en Menor o Igual que
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[11]/ul/li[9]/div"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Fecha Generacion' : {str(e)}", duration=10, threaded=True)

#Click en icono de calendario 
A0=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[2]")))
try:
    A0.click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en el Calendario : {str(e)}", duration=10, threaded=True)

#Selecciona la fecha de hoy
A=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/table/tbody/tr['+str(fila_dia)+']/td['+str(columna_dia)+']/a')))
try:
    A.click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Fecha Generacion' : {str(e)}", duration=10, threaded=True)

#Click en + 
#WebDriverWait(driver, 500).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[1]"))).click()

#Cierra el calendario y da click en +
B=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[1]"))) 
A1= WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div[2]/button[2]')))
   
try:
   
   B.click()
except:
        A1.click()
        B.click()

#Click en Generar Informe
C=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/div/a[3]/i")))
try:
   
   C.click()
except:
        A1.click()
        C.click()   

#Click en Generar Informe
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/a[3]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Generar Informe' : {str(e)}", duration=10, threaded=True)

#Click en CSV
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/span[1]/a[3]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'CSV' : {str(e)}", duration=10, threaded=True)

#Borra los campos de Nombre Archivo
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/div[2]/form/div[2]/div[1]/input"))).clear()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de borrar el nombre por defecto : {str(e)}", duration=10, threaded=True)

#Coloca Informe_Compromiso_ + Mes
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/div[2]/form/div[2]/div[1]/input"))).send_keys(["informe_compromiso_" +str(day.strftime("%B").capitalize())])
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de poner el nuevo nombre al archivo : {str(e)}", duration=10, threaded=True)

#Click en Exportar
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/div[2]/form/div[3]/input[1]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Compromiso Mes", f"Se ha producido un error al momento de dar click en 'Exportar' : {str(e)}", duration=10, threaded=True)

#Cerrar el navegador Google Chrome
#driver.quit()

#Fin de la descarga
toaster.show_toast("Bot de Adminfo", "Ya se está descargando el archivo de Informe_Compromiso Mes", duration=10)



### FASE 5: Informe Pagos Mes

try:
    os.remove("//RFNTRFS02/Analitica$/Servicing\\Bancolombia\\"+str(year_format)+"\\"+str(month_format_01)+"."+str(day.strftime("%B").capitalize())+" "+str(year_format)+"\\Creación listas de marcación\\exclusiones_seguimiento\\informe_pagos_"+str(day.strftime("%B").capitalize())+".zip")
except: 
    print('continuemos')

#Click en pestaña de informes
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/ul/li[1]/a/span/p/span"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Pestaña de informes' : {str(e)}", duration=10, threaded=True)

#Click en Informes de Pagos
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/ul/li[1]/ul/li[3]/a"))).click()  
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Informes de Pago' : {str(e)}", duration=10, threaded=True)

#Click en Agrupar
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[1]/div[1]/div/a/span[2]/b"))).click()  
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Agrupar' : {str(e)}", duration=10, threaded=True)

#Click en Detalle en Pantalla
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[10]/ul/li[10]/div"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Detalle en Pantalla' : {str(e)}", duration=10, threaded=True)

#Click en Columna
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[2]/div/a/span[2]/b"))).click()   
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Columna' : {str(e)}", duration=10, threaded=True)

#Click en Fecha Pago
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[10]/ul/li[5]/div"))).click()  
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Fecha Pago' : {str(e)}", duration=10, threaded=True)

#Click en Operador
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[4]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Operador' : {str(e)}", duration=10, threaded=True)

#Click en Mayor o Igual que
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[11]/ul/li[8]/div"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Mayor o Igual que' : {str(e)}", duration=10, threaded=True)

#Click en Calendario
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[2]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Calendario' : {str(e)}", duration=10, threaded=True)

#Click en Primer dia del mes 
A=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/table/tbody/tr['+str(int(fila_primer_dia))+']/td['+str(int( columna_primer_dia))+']/a')))
try:
    A.click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en el primer día del mes : {str(e)}", duration=10, threaded=True)

#Click en Cerrar Calendario y click en +
B=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[1]"))) 
A1= WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div[2]/button[2]')))
   
try:
   
   B.click()
except:
        A1.click()
        B.click()  

#Click en Operador
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[4]/div/a/span[2]/b"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Operador' : {str(e)}", duration=10, threaded=True)

#Click en Menor o Igual que
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[11]/ul/li[9]/div"))).click() 
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Menor o igual que' : {str(e)}", duration=10, threaded=True)

#Click en Calendario
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[2]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Calendario' : {str(e)}", duration=10, threaded=True)

#Click en dia actual
A=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/table/tbody/tr['+str(fila_dia)+']/td['+str(columna_dia)+']/a')))
try:
    A.click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en el día actual : {str(e)}", duration=10, threaded=True)

#Click en Cerrar Calendario y click en +
B=WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/fieldset/span[2]/div/div[1]/span/table/tbody/tr/td[7]/a[1]"))) 
A1= WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div[2]/button[2]')))
   
try:
   
   B.click()
except:
        A1.click()
        B.click() 

#Click en Generar Informe
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div/div/a[3]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Generar Informe' : {str(e)}", duration=10, threaded=True)

#Click en Generar Informe
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/a[3]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Generar Informe' : {str(e)}", duration=10, threaded=True)

#Click en CSV
try:
    WebDriverWait(driver, 500).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/span[1]/a[3]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'CSV' : {str(e)}", duration=10, threaded=True)

#Borrar Nombre Archivo
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/div[2]/form/div[2]/div[1]/input"))).clear()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click al borrar el nombre del archivo : {str(e)}", duration=10, threaded=True)

#Coloca informe_pagos_ + Mes
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/div[2]/form/div[2]/div[1]/input"))).send_keys(["informe_pagos_" +str(day.strftime("%B").capitalize())])
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de poner el nuevo nombre del archivo' : {str(e)}", duration=10, threaded=True)

#Click en Exportar
try:
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/div[3]/div/div[2]/div[2]/div/div/div[2]/form/div[3]/input[1]"))).click()
except Exception as e:
    toaster.show_toast("Error en el proceso - Informe Pagos Mes", f"Se ha producido un error al momento de dar click en 'Exportar' : {str(e)}", duration=10, threaded=True)

toaster = ToastNotifier()

# Muestra la notificación
toaster.show_toast("Bot de Adminfo", "Ya se está descargando el archivo de Informe Pagos Mes", duration=10)

