import pyautogui
import time
import sys

# ------------------------//-------------------------------//------------------------------//----------------------- #
OPSI = "Iniciar Si"
OPNO = "Iniciar No"
OPCLOSE = "Salir"
# Crear el mensaje.
opt = pyautogui.confirm(
    "¿Que deseas realizar?",
    "Repollo Automatizacion",
    [OPSI, OPNO, OPCLOSE]
)
# Opción 1
if opt == OPSI:
    # Abrir explorador de archivos
    pyautogui.hotkey("win", "e")
    time.sleep(2)

# Opción 2
elif opt == OPNO:
    # Abrir explorador de archivos
    pyautogui.hotkey("win", "e")
    time.sleep(2)
    # Cerrar explorador de archivos
    pyautogui.hotkey("alt", "f4")
    time.sleep(2) 

# Opción 3
elif opt == OPCLOSE:
        sys.exit()

