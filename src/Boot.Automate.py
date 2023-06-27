import pyautogui
import time
import sys
import subprocess
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# ------------------------//-------------------------------//------------------------------//----------------------- #
# Menú de opciones
OPSI = "RPA Transformar Excel to PDF"
OPNO = "RPA Realizar Evolución"
OPCLOSE = "RPA Web Scraping"

# Crear el mensaje.
opt = pyautogui.confirm(
    "¿Que deseas realizar?",
    "CMISL Automatizacion",
    [OPSI, OPNO, OPCLOSE]
)

ruta_icosalud = "C:/Users/brayan.diaz/Downloads/image/Opcion2/01icosalud.png"
ruta_historia_clinica = "C:/Users/brayan.diaz/Downloads/image/Opcion2/02historia_clinica.png"
ruta_corazon = "C:/Users/brayan.diaz/Downloads/image/Opcion2/03corazon.png"
ruta_evolucion = "C:/Users/brayan.diaz/Downloads/image/Opcion2/04icono_evolucion.png"
ruta_insertar_evolucion = "C:/Users/brayan.diaz/Downloads/image/Opcion2/05icono_evolucion.png"
ruta_grabar_evolucion = "C:/Users/brayan.diaz/Downloads/image/Opcion2/06grabar_evolucion.png"
ruta_salir_evolucion = "C:/Users/brayan.diaz/Downloads/image/Opcion2/07salir.png"


tipo_evolucion = "NOTA MEDICAS"
tipo_evolucion = tipo_evolucion.upper()
subjetivo = "El paciente se presenta con una serie de manifestaciones en el sistema tegumentario, lo que genera preocupación y malestar. El individuo ha notado la presencia de lesiones cutáneas que causan picazón intensa y enrojecimiento en varias áreas del cuerpo, especialmente en el tronco y las extremidades. Además, ha experimentado sequedad en la piel y descamación en algunas zonas específicas. Estas condiciones han afectado su calidad de vida y le han causado incomodidad física y emocional. El paciente busca una evaluación médica exhaustiva para comprender la causa subyacente y recibir el tratamiento adecuado."
objetivo = "En el examen físico, se observa la presencia de múltiples lesiones en la piel, caracterizadas por eritema y descamación. Las áreas afectadas muestran una apariencia inflamada y muestran signos de excoriaciones debido al rascado intenso. Las lesiones se distribuyen de manera simétrica en el tronco y las extremidades superiores e inferiores. Además, se observa una sequedad generalizada en la piel, especialmente en los codos y las rodillas."
sistema_tegumentario = "El sistema tegumentario del paciente presenta diversas alteraciones cutáneas. Las lesiones inflamatorias en forma de placas eritematosas indican la presencia de una posible dermatitis. Los síntomas asociados, como picazón intensa y descamación, sugieren una reacción alérgica o una enfermedad inflamatoria crónica de la piel. La sequedad cutánea generalizada podría ser indicativa de una disfunción en la barrera protectora de la piel, lo que aumenta la susceptibilidad a irritantes externos y la pérdida de agua transepidérmica."
analisis = "Dada la presentación clínica del paciente, se plantea como diagnóstico diferencial la posibilidad de una dermatitis atópica o un brote de psoriasis. Ambas condiciones son enfermedades crónicas de la piel que pueden manifestarse con lesiones cutáneas inflamatorias, prurito y descamación. Se requiere realizar pruebas complementarias, como un examen microscópico de una muestra de piel y pruebas alérgicas, para confirmar el diagnóstico."


# Opción 1
if opt == OPSI:
    # Abrir explorador de archivos
    pyautogui.hotkey("win", "e")
    # Esperar 2 segundos
    time.sleep(2)
    # Ruta Imagen
    ruta_imagen = "C:/Users/brayan.diaz/Downloads/image/01icono_descarga.png"
    # Localizar imagen
    posicion_icono = pyautogui.locateCenterOnScreen(ruta_imagen)
    # Haz clic en el centro del icono
    pyautogui.click(posicion_icono)
    # Esperar 2 segundos
    time.sleep(2)
    # Ruta Imagen
    ruta_imagen2 = "C:/Users/brayan.diaz/Downloads/image/02excel.png"
    # Localizar imagen
    posicion_icono2 = pyautogui.locateCenterOnScreen(ruta_imagen2)
    if posicion_icono2 is not None:
      pyautogui.doubleClick(posicion_icono2)
    else:
      print("No se encontró el icono en el escritorio")
      sys.exit()
    # Esperar 2 segundos
    time.sleep(2)
    pyautogui.click(71, 267)
    pyautogui.click(310, 204)
    pyautogui.hotkey("ctrl", "e")
    time.sleep(2)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(1)
    ruta_archivo = "C:/Users/brayan.diaz/Downloads/image/03archivo.png"
    posicion_archivo = pyautogui.locateCenterOnScreen(ruta_archivo)
    pyautogui.click(posicion_archivo)
    time.sleep(1)
    guardar_como = "C:/Users/brayan.diaz/Downloads/image/04guardar_como.png"
    posicion_guardar_como = pyautogui.locateCenterOnScreen(guardar_como)
    pyautogui.click(posicion_guardar_como)
    time.sleep(1)
    examinar = "C:/Users/brayan.diaz/Downloads/image/05examinar.png"
    posicion_examinar = pyautogui.locateCenterOnScreen(examinar)
    pyautogui.click(posicion_examinar)
    time.sleep(1)
    text = "C:/Users/brayan.diaz/Downloads/image/06text.png"
    posicion_text = pyautogui.locateCenterOnScreen(text)
    pyautogui.doubleClick(posicion_text)
    time.sleep(1)
    pyautogui.click(563, 471)
    time.sleep(1)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)
    pyautogui.hotkey("enter")
    time.sleep(1)
    pyautogui.hotkey("alt", "f4")
    time.sleep(1)

    print("Se ha generado 1 proceso")

    sys.exit()

# Opción 2
elif opt == OPNO:

  user = pyautogui.prompt("Ingrese el Usuario:")
  password = "Northman4321"
  registro = pyautogui.prompt("Ingrese el Registro:")
  # Ruta Imagen
  time.sleep(1)
  # Localizar imagen
  posicion_icosalud = pyautogui.locateCenterOnScreen(ruta_icosalud)
  time.sleep(1)
  print(type(posicion_icosalud))
  # Valicación Imagen Icono Icosalud
  if posicion_icosalud is None:
      print("No se encontró el icono de icosalud")
      sys.exit()
 
  pyautogui.click(posicion_icosalud)
  time.sleep(1)
  pyautogui.typewrite(user)
  time.sleep(1)
  pyautogui.press('tab')
  time.sleep(1)
  pyautogui.typewrite(password)
  pyautogui.press('enter')
  time.sleep(1)
  posicion_historia_clinica = pyautogui.locateCenterOnScreen(ruta_historia_clinica)
  time.sleep(1)

  # Valicación Imagen Historia Clinica
  if posicion_historia_clinica is None:
    print("No se encontró el modulo historia clinica")
    sys.exit()
  
  pyautogui.click(posicion_historia_clinica)
  time.sleep(1)
  posicion_corazon = pyautogui.locateCenterOnScreen(ruta_corazon)
  time.sleep(1)

    # Valicación Imagen Icono de corazon
  if posicion_corazon is None:
    print("No se encontró el icono en forma de corazón")
    sys.exit()
    
  pyautogui.click(posicion_corazon)
  time.sleep(1)
  pyautogui.typewrite(registro)
  time.sleep(1)
  pyautogui.press('enter')
  time.sleep(1)
  pyautogui.press('enter')
  time.sleep(1)
  posicion_evolucion = pyautogui.locateCenterOnScreen(ruta_evolucion)

  # Valicación Imagen Icono de evolución
  if posicion_evolucion is None:
      print("No se encontró el icono de evolucion")
      sys.exit()
      
  pyautogui.click(posicion_evolucion)
  time.sleep(1)
  posicion_insertar_evolucion = pyautogui.locateCenterOnScreen(ruta_insertar_evolucion)

  # Validación Imagen botón de insertar evolución
  if posicion_insertar_evolucion is None:
      print("No se encontró el botón de insertar evolucion")
      sys.exit()
  pyautogui.click(posicion_insertar_evolucion)
  time.sleep(1)
  pyautogui.typewrite(tipo_evolucion)
  time.sleep(1)
  pyautogui.press('enter')
  time.sleep(1)
  pyautogui.press('esc')
  time.sleep(1)
  pyautogui.typewrite(subjetivo)
  time.sleep(1)
  pyautogui.press('tab')
  pyautogui.typewrite(objetivo)
  time.sleep(1)
  pyautogui.press('tab')
  pyautogui.typewrite(sistema_tegumentario)
  time.sleep(1)
  pyautogui.press('tab')
  pyautogui.typewrite(analisis)
  posicion_grabar_evolucion = pyautogui.locateCenterOnScreen(ruta_grabar_evolucion)
  
  # Validación Imagen del botón que se llama Grabar_Evolución
  if posicion_grabar_evolucion is None:
      print("No se encontró el botón que se llama Grabar Evolución")
      sys.exit()
  pyautogui.click(posicion_grabar_evolucion)
  time.sleep(1)
  posicion_salir_evolucion = pyautogui.locateCenterOnScreen(ruta_salir_evolucion)

  # Validación Imagen del botón que se llama Salir
  if posicion_salir_evolucion is None:
      print("No se encontró el botón que se llama Salir")
      sys.exit()
          
  pyautogui.click(posicion_salir_evolucion)
  time.sleep(1)
  pyautogui.press('enter')
  print("Se ha realizado la evolción de tipo Nota Medicas al registro", "#", registro)
  sys.exit()

#  Opción 3
elif opt == OPCLOSE:


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


      



