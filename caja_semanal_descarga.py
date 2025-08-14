import subprocess
import time
import pyautogui
import ctypes
import calendar
import datetime
import os

# ============================================================
# CONFIGURACIÓN DE FECHAS
# ============================================================

hoy = datetime.date.today()
inicio_mes = hoy.replace(day=1)
ultimo_dia = calendar.monthrange(hoy.year, hoy.month)[1]
fin_mes = hoy.replace(day=ultimo_dia)

fecha_inicio = inicio_mes.strftime('%d/%m/%Y')
fecha_fin = fin_mes.strftime('%d/%m/%Y')
fecha_swap = '30/06/2025'

# ============================================================
# UTILIDADES DEL SISTEMA
# ============================================================

def is_capslock_on():
    """Verifica si la tecla Caps Lock está activada."""
    return ctypes.WinDLL("User32.dll").GetKeyState(0x14) & 0x0001

def ensure_capslock_off():
    """Desactiva Caps Lock si está activado."""
    if is_capslock_on():
        pyautogui.press('capslock')
        print("Caps Lock estaba activado. Se ha desactivado.")
    else:
        print("Caps Lock ya estaba desactivado.")

def ya_existe_detalle_swap(ruta_descargas, prefijo="Detalle_Swap", extension=".xlsx"):
    """Verifica si ya existe un archivo 'Detalle_Swap' creado este mes."""
    hoy = datetime.date.today()
    inicio_mes = datetime.datetime(hoy.year, hoy.month, 1)

    for archivo in os.listdir(ruta_descargas):
        if archivo.startswith(prefijo) and archivo.endswith(extension):
            ruta_archivo = os.path.join(ruta_descargas, archivo)
            fecha_creacion = datetime.datetime.fromtimestamp(os.path.getctime(ruta_archivo))
            if fecha_creacion >= inicio_mes:
                return True
    return False

# ============================================================
# INICIO DE APLICACIÓN
# ============================================================
print("Ejecutando el archivo CMD...")

# En esta versión se hace uso del nuevo OPEFIN, esperemos que mejore la 
# Velocidad de descarga de los informes
subprocess.Popen(r"C:\OPEFIN NUEVO\OPEFIN 2.cmd", shell=True)

ensure_capslock_off()

print("Esperando 40 segundos para que la aplicación se cargue...")
time.sleep(30)

# En esta versión se omite la autenticación de la aplicación

# ============================================================
# NAVEGACIÓN A INFORMES
# ============================================================

def abrir_modulo_informes():
    """Navega al módulo de informes y abre la sección de Excel."""
    pyautogui.click((28, 305))  # Módulo INFORMES
    pyautogui.doubleClick((142, 340))  # Petición de Informes Excel
    time.sleep(1)

abrir_modulo_informes()

# ============================================================
# DESCARGA DE INFORMES
# ============================================================

def descargar_informe(nombre, fecha_ini=None, fecha_fin=None, solo_fecha=False, tiempo_espera=30):
    """
    Busca y descarga un informe por nombre y fechas.
    """
    time.sleep(3)
    pyautogui.click((67, 115))  # Barra de búsqueda del nuevo Infisa
    print(f"Buscando informe: {nombre}")
    pyautogui.write(nombre, interval=0.1)
    pyautogui.press('enter')
    pyautogui.press('enter')

    if solo_fecha and fecha_ini:
        pyautogui.write(fecha_ini, interval=0.1)
    elif fecha_ini and fecha_fin:
        pyautogui.write(fecha_ini, interval=0.1)
        pyautogui.press('tab')
        pyautogui.write(fecha_fin, interval=0.1)

    pyautogui.click((688, 424))  # Botón de descarga/ Alias rayito/ O lanzar informe
    print(f"Informe '{nombre}' solicitado.")

    time.sleep(tiempo_espera) # Es muy importante este número ya que es el que 
    #van a esperar todos los informes. Por defecto son 30s si no se especifica
    #es mejor especificarlo ya que esto puede variar a veces entre meses.

    pyautogui.hotkey('alt', 'tab') # Esto es para devolverme a Infisa y siga el RPA con normalidad
    time.sleep(2)
    print("Confirmación posterior a descarga realizada.")

# ------------------------------------------------------------
# Descargar informes
# ------------------------------------------------------------

pyautogui.click((1252, 309))  # Botón de flecha abajo para consultar todos los Informes de Excel
descargar_informe(nombre="Liquidaciones NDF", fecha_ini=fecha_inicio, fecha_fin=fecha_fin, tiempo_espera=180)

pyautogui.click((1252, 309))  # Botón de flecha abajo para consultar todos los Informes de Excel
descargar_informe(nombre="Intereses de Deuda", fecha_ini=fecha_inicio, fecha_fin=fecha_fin, tiempo_espera=15)

# Verificar si ya existe el archivo de Detalle Swap
ruta_descargas = r"C:\Users\ntorreslo\Downloads"

if not ya_existe_detalle_swap(ruta_descargas):
    pyautogui.click((1252, 309))  # Botón de flecha abajo para consultar todos los Informes de Excel
    descargar_informe(nombre="Detalle Swap", fecha_ini=fecha_swap, solo_fecha=True)
else:
    print("Ya existe un archivo 'Detalle_Swap' este mes. Se omite la descarga.")

# ============================================================
# CIERRE DE APLICACIÓN
# ============================================================

pyautogui.click((663, 297))  # Botón salir/ Es una puerta roja en el costado izquierdo superior
time.sleep(2)
pyautogui.hotkey('alt', 'f4')
pyautogui.hotkey('alt', 'tab') # En este caso, me quiero devolver al IDE a continuar programando o devolverme a hacer lo que esté haciendo
print("Aplicación cerrada.")
