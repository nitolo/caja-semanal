import subprocess
import sys
import time
import logging
from pathlib import Path

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('script_execution.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def ejecutar_script(nombre_script):
    """
    Ejecuta un script Python y maneja errores si ocurren.
    
    Args:
        nombre_script (str): Nombre del archivo de script a ejecutar
        
    Returns:
        bool: True si se ejecutó correctamente, False en caso contrario
    """
    try:
        # Verificar que el archivo existe
        script_path = Path(nombre_script)
        if not script_path.exists():
            logger.error(f"El archivo {nombre_script} no existe")
            return False
            
        logger.info(f"Ejecutando {nombre_script}...")
        
        # Ejecutar el script con captura de salida para mejor debugging
        result = subprocess.run(
            [sys.executable, nombre_script], 
            check=True,
            capture_output=True,
            text=True,
            timeout=300  # Timeout de 5 minutos
        )
        
        logger.info(f"{nombre_script} ejecutado correctamente")
        
        # Mostrar salida si hay algo importante
        if result.stdout.strip():
            logger.info(f"Salida de {nombre_script}:\n{result.stdout}")
            
        return True
        
    except subprocess.CalledProcessError as e:
        logger.error(f"Error al ejecutar {nombre_script}")
        logger.error(f"Código de salida: {e.returncode}")
        if e.stdout:
            logger.error(f"Salida estándar: {e.stdout}")
        if e.stderr:
            logger.error(f"Error estándar: {e.stderr}")
        return False
        
    except subprocess.TimeoutExpired:
        logger.error(f"Timeout: {nombre_script} tardó más de 5 minutos en ejecutarse")
        return False
        
    except FileNotFoundError:
        logger.error(f"No se pudo encontrar el intérprete de Python o el archivo {nombre_script}")
        return False
        
    except Exception as e:
        logger.error(f"Error inesperado al ejecutar {nombre_script}: {e}")
        return False

def formatear_tiempo(duracion):
    """
    Formatea la duración en un formato legible.
    
    Args:
        duracion (float): Duración en segundos
        
    Returns:
        str: Duración formateada
    """
    if duracion < 60:
        return f"{duracion:.2f} segundos"
    elif duracion < 3600:
        minutos = int(duracion // 60)
        segundos = int(duracion % 60)
        return f"{minutos} minutos y {segundos} segundos"
    else:
        horas = int(duracion // 3600)
        minutos = int((duracion % 3600) // 60)
        segundos = int(duracion % 60)
        return f"{horas} horas, {minutos} minutos y {segundos} segundos"

def main():
    """
    Función principal que ejecuta los scripts en secuencia.
    """
    scripts = [
        "caja_semanal_descarga.py",
        "caja_semanal_limpieza.py", 
        "caja_semanal_tipodecambio.py"
    ]
    
    logger.info("=== INICIANDO EJECUCIÓN DE SCRIPTS ===")
    inicio = time.time()
    
    scripts_exitosos = 0
    scripts_fallidos = 0
    
    # Ejecutar los scripts en orden
    for script in scripts:
        if ejecutar_script(script):
            scripts_exitosos += 1
        else:
            scripts_fallidos += 1
            logger.error(f"Falló la ejecución de {script}. Deteniendo proceso.")
            break  # Detener si falla algún script
    
    fin = time.time()
    duracion = fin - inicio
    
    # Resumen de ejecución
    logger.info("=== RESUMEN DE EJECUCIÓN ===")
    logger.info(f"Scripts ejecutados exitosamente: {scripts_exitosos}")
    logger.info(f"Scripts fallidos: {scripts_fallidos}")
    logger.info(f"Tiempo total de ejecución: {formatear_tiempo(duracion)}")
    
    if scripts_fallidos > 0:
        logger.error("Proceso completado con errores")
        return False
    else:
        logger.info("Todos los scripts ejecutados correctamente")
        return True

if __name__ == "__main__":
    try:
        success = main()
        if not success:
            sys.exit(1)  # Usar sys.exit() en lugar de exit()
    except KeyboardInterrupt:
        logger.info("Proceso interrumpido por el usuario")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Error crítico: {e}")
        sys.exit(1)
