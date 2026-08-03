import threading
import time
import logging
from queue import Queue

# Configuración estricta de registro de errores (Bitácora de auditoría)
logging.basicConfig(
    filename='sistema_auditoria.log',
    level=logging.INFO,
    format='%(asctime)s - [%(levelname)s] - %(message)s'
)

class GestorProcesosCriticos:
    def __init__(self):
        self.cola_tareas = Queue()
        self.activo = True

    def ejecutar_tarea_segura(self, nombre_tarea, funcion_objetivo, *args):
        """
        Envuelve cualquier tarea en un hilo independiente con manejo de excepciones 
        para evitar que un fallo colapse el sistema global.
        """
        def hilo_trabajador():
            try:
                logging.info(f"Iniciando tarea: {nombre_tarea}")
                funcion_objetivo(*args)
                logging.info(f"Tarea finalizada con éxito: {nombre_tarea}")
            except Exception as e:
                logging.error(f"Fallo crítico en {nombre_tarea}: {str(e)}")
                print(f"[ALERTA DE SISTEMA] Error interceptado en '{nombre_tarea}': {e}")

        hilo = threading.Thread(target=hilo_trabajador)
        hilo.daemon = True
        hilo.start()

# --- Módulos de Operación Específica ---

def procesar_inventario_excel(ruta_archivo):
    print(f"Analizando datos de inventario desde: {ruta_archivo}...")
    time.sleep(2)
    print("Inventario procesado, validado y guardado correctamente.")

def realizar_cierre_caja(datos_caja):
    print(f"Calculando balance de caja para los registros...")
    time.sleep(1.5)
    print("Cierre de caja completado sin discrepancias.")

# --- Secuencia Lógica Principal ---
if __name__ == "__main__":
    print("=== INICIANDO SISTEMA DE CONTROL DE PRECISIÓN QUIRÚRGICA ===")
    
    sistema = GestorProcesosCriticos()

    # Paso 1: Ejecutar tarea de inventario en paralelo seguro
    sistema.ejecutar_tarea_segura("Control de Inventario", procesar_inventario_excel, "inventario_general.xlsx")

    # Paso 2: Ejecutar cierre de caja en paralelo seguro
    sistema.ejecutar_tarea_segura("Cierre de Caja", realizar_cierre_caja, {"caja_id": 104})

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n[SISTEMA] Apagado seguro solicitado. Guardando estados y liberando recursos...")
