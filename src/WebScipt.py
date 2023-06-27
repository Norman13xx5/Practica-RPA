import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Ruta del ejecutable webdriver, asegúrate de tener la versión correcta
webdriver_path = 'D:/brayan.diaz/12. python/01. Script01/env/chromedriver.exe'

############# Esto no se que sea, pero necesita esto por la version de selenium y la version de webdriver #################

# Inicializar el objeto Options y configurar las opciones
driver_options = Options()
driver_options.add_argument("--ignore-local-proxy")

# Inicializar el servicio de ChromeDriver
service = Service(webdriver_path)

# Inicializar el driver de Chrome
driver = webdriver.Chrome(service=service, options=driver_options)

############# Esto no se que sea, pero necesita esto por la version de selenium y la version de webdriver #################

# Aquí hacemos un arreglo con todas las opciones que existen en el campo sexo
sexo_dict = {'Male':  '/html/body/div/div[2]/form/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div/span/div/div[1]/label/div/div[1]/div/div[3]/div',
             'Female': '/html/body/div/div[2]/form/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div/span/div/div[2]/label/div/div[1]/div/div[3]/div',
             "I'd rather not answer": '//*[@id="i23"]/div[3]/div'
}

# Arreglo del SelectButton
rubros_dict = {'Mining': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[2]/div[3]/span',
               'Health Care': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[2]/div[4]/span',  
               'Energy': '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[2]/div[5]/span',
               'Manufacturing':'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[2]/div[6]/span'
}

# Leer el archivo Excel
df = pd.read_csv('D:/brayan.diaz/12. python/01. Script01/env/Datos.csv', encoding='latin-1')

# Recorrer las filas del DataFrame
for row, datos in df.iterrows():
    id = datos['ID']    
    nombre = datos['NOMBRES']
    pateterno = datos['APEPATERNO']
    materno = datos['APEMATERNO']
    genero = datos['SEXO']
    email = datos['EMAIL']
    rubro = datos['RUBRO']

    # Ruta de navegación
    driver.get('https://docs.google.com/forms/d/e/1FAIpQLSdqbcMwmFgGV7vm-lEhxC-tB0cby9ErjnjmmqN2JJ4GPrSRJA/viewform')
    time.sleep(3)

    # Input ID
    driver.find_element("xpath", '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(id)

    # Input Nombre
    driver.find_element("xpath", '//*//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(nombre)

    # Input Apellido
    driver.find_element("xpath", '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(pateterno + ' ' + materno)

    # SelectButton
    sexo = genero
    driver.find_element("xpath", sexo_dict[sexo]).click()

    # Input Correo
    driver.find_element("xpath", '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(email)

    # RadioButton
    driver.find_element("xpath", '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div[1]/div[1]/span').click()
    time.sleep(2)

    # Opción a escoger
    rubro = rubro
    # Opción escogida
    driver.find_element("xpath", rubros_dict[rubro]).click()
    time.sleep(2)

    # Btn Enviar
    driver.find_element("xpath", '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span').click()
    time.sleep(2)

# Cerrar el navegador
driver.close()
