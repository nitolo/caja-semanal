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
fecha_swap = '29/05/2025'

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
subprocess.Popen(r"C:\Users\ntorreslo\OneDrive - Telefonica\Escritorio\OPEFIN 1.cmd", shell=True)

ensure_capslock_off()

print("Esperando 40 segundos para que la aplicación se cargue...")
time.sleep(40)

# ============================================================
# AUTENTICACIÓN EN LA APLICACIÓN
# ============================================================

pyautogui.press('enter')
time.sleep(3)
pyautogui.press('delete')

pyautogui.press('capslock')
pyautogui.write('OPFN', interval=0.1)
pyautogui.press('capslock')

pyautogui.press('tab')
pyautogui.write('opfn', interval=0.1)
pyautogui.press('tab')
pyautogui.write('afinprod', interval=0.1)
pyautogui.press('tab')
pyautogui.press('enter')

time.sleep(8)

# ============================================================
# NAVEGACIÓN A INFORMES
# ============================================================

def abrir_modulo_informes():
    """Navega al módulo de informes y abre la sección de Excel."""
    pyautogui.click((32, 342))  # Módulo informes
    pyautogui.doubleClick((151, 389))  # Informes Excel
    time.sleep(1)

abrir_modulo_informes()

# ============================================================
# DESCARGA DE INFORMES
# ============================================================

def descargar_informe(nombre, fecha_ini=None, fecha_fin=None, solo_fecha=False):
    """
    Busca y descarga un informe por nombre y fechas.
    """
    time.sleep(3)
    pyautogui.click((92, 136))  # Barra de búsqueda
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

    pyautogui.click((687, 440))  # Botón de descarga
    print(f"Informe '{nombre}' solicitado.")
    time.sleep(30) # Es muy importante este número ya que es el que van a esperar todos los informes

    pyautogui.hotkey('alt', 'tab')
    time.sleep(2)
    print("Confirmación posterior a descarga realizada.")

# ------------------------------------------------------------
# Descargar informes
# ------------------------------------------------------------
pyautogui.click((1249, 325))  # Confirmación
descargar_informe(nombre="Liquidaciones NDF", fecha_ini=fecha_inicio, fecha_fin=fecha_fin)
pyautogui.click((1249, 325))  # Confirmación
descargar_informe(nombre="Intereses de Deuda", fecha_ini=fecha_inicio, fecha_fin=fecha_fin)

# Verificar si ya existe el archivo de Detalle Swap
ruta_descargas = r"C:\Users\ntorreslo\Downloads"

if not ya_existe_detalle_swap(ruta_descargas):
    pyautogui.click((1249, 325))  # Confirmación
    descargar_informe(nombre="Detalle Swap", fecha_ini=fecha_swap, solo_fecha=True)
else:
    print("Ya existe un archivo 'Detalle_Swap' este mes. Se omite la descarga.")

# ============================================================
# CIERRE DE APLICACIÓN
# ============================================================

pyautogui.click((661, 307))  # Botón salir
time.sleep(1)
pyautogui.hotkey('alt', 'f4')
print("Aplicación cerrada.")

