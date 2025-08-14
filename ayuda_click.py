import pyautogui
import time

# En este punto, yo lo que quiero saber es dónde debo hacer el click
# Realmente es muy jodido hacer a ojo y conociendo los puntos cardinales

print("Mueva el mouse rápido. Solo hay cinco minutos")
time.sleep(5)

# Captura la posición actual del mouse
x, y = pyautogui.position()
print(f"La posición del botón es: ({x}, {y})")
