#Librerias
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import pandas as pd

#Obtengo los datos del excel
datos_excel = pd.read_excel('Datos.xlsx')
#Asigno valor a las variables
username=datos_excel.loc[0, 'Username']
password=datos_excel.loc[0, 'Pass']
contribuyente=datos_excel.loc[0, 'Nombre completo']
fechadesde=datos_excel.loc[0, 'desde']
fechahasta=datos_excel.loc[0, 'hasta']

# Opciones de navegacion
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')
options.add_argument('--disable-extensions')

#Ingresar la ubicacion de chromedriver
driver_path = 'chromedriver_win32/chromedriver.exe'

driver = webdriver.Chrome(options=options)

# Inicializar navegador
driver.get('https://auth.afip.gob.ar')

# Comenzas a operar la pagina
WebDriverWait(driver, 5) \
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#F1\:username'))) \
    .send_keys(str(username))   
WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'input#F1\:btnSiguiente')))\
    .click()
WebDriverWait(driver, 5) \
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input#F1\:password'))) \
    .send_keys(str(password))   
WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'input#F1\:btnIngresar')))\
    .click()
WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'#serviciosMasUtilizados > div > div > div > div:nth-child(5) > div > a')))\
    .click()
#Nos desplazamos un poco para visualizar el servicio buscado

driver.execute_script("window.scrollBy(0, 400);")
time.sleep(1)
driver.execute_script("window.scrollBy(0, 400);")
time.sleep(1)

#Busco el servicio comprobantes en linea
WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.XPATH, "//h3[contains(text(), 'COMPROBANTES EN LÍNEA')]")))\
    .click()
time.sleep(1)
# Obtener las ventanas abiertas
window_handles = driver.window_handles

# Cambiar al control de la última pestaña
driver.switch_to.window(window_handles[-1])
WebDriverWait(driver, 5)\
    .until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.btn_empresa[value='{}']".format(contribuyente))))\
    .click()

WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR,'#btn_consultas')))\
    .click()

# Cambiar Fecha Desde
fechadesde = fechadesde.strftime("%d/%m/%Y")  # Convertir a cadena de formato "dd/mm/yyyy"
fechahasta = fechahasta.strftime("%d/%m/%Y")  # Convertir a cadena de formato "dd/mm/yyyy"
fecha_emision_desde = driver.find_element(By.ID, "fed")

# Borrar el valor existente
fecha_emision_desde.clear()

# Ingresar un nuevo valor
nuevo_valor = fechadesde
fecha_emision_desde.send_keys(nuevo_valor)

# Cambiar fecha Hasta
fecha_emision_hasta = driver.find_element(By.ID, "feh")

# Borrar el valor existente
fecha_emision_hasta.clear()

# Ingresar un nuevo valor
nuevo_valor = fechahasta
fecha_emision_hasta.send_keys(nuevo_valor)

#Buscar
WebDriverWait(driver, 5) \
    .until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='button'][value='Buscar'][style='width:100px;']"))) \
    .click()
# Encontrar todos los elementos de botón dentro de la tabla
botones = driver.find_elements(By.CSS_SELECTOR, "tr input[type='button'][value='Ver']")

if len(botones) == 0:
    print("No hay facturas para descargar. Finalizando...")
    
    break
# Hacer clic en cada botón uno a uno para descargar facturas
for boton in botones:
    boton.click()
    time.sleep(1)  # Pausa de 1 segundo antes de hacer clic en el siguiente botón