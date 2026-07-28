import io
import pandas as pd

class ProcesadorExcel:
    """
    Clase de nivel profesional para la carga, validación, procesamiento 
    y exportación de datos de inventario y caja mediante Pandas.
    """
    def __init__(self, archivo_subido):
        self.archivo_subido = archivo_subido
        self.df = None
        self.errores = []

    def cargar_datos(self):
        """Carga el archivo Excel de manera segura manejando excepciones de formato y estructura vacía."""
        try:
            # Lee el archivo subido mediante un buffer de bytes con openpyxl
            self.df = pd.read_excel(self.archivo_subido, engine='openpyxl')
            
            if self.df.empty:
                self.errores.append("El archivo Excel cargado está vacío.")
                return False
                
            return True
        except ValueError as ve:
            self.errores.append(f"Error de formato en el archivo Excel: {str(ve)}")
            return False
        except Exception as e:
            self.errores.append(f"Error crítico al leer el archivo Excel: {str(e)}")
            return False

    def validar_estructura(self, columnas_requeridas):
        """Verifica de forma robusta que las columnas obligatorias existan, normalizando espacios y mayúsculas."""
        if self.df is None:
            self.errores.append("No hay datos cargados para validar la estructura.")
            return False

        # Normalizamos las columnas actuales del DataFrame (eliminando espacios y llevando a minúsculas)
        columnas_actuales = {str(col).strip().lower(): col for col in self.df.columns}
        faltantes = []

        for req in columnas_requeridas:
            req_normalizado = req.strip().lower()
            if req_normalizado not in columnas_actuales:
                faltantes.append(req)

        if faltantes:
            self.errores.append(f"Faltan las siguientes columnas obligatorias: {', '.join(faltantes)}")
            return False
            
        return True

    def procesar_inventario_y_caja(self):
        """Ejecuta los cálculos lógicos con saneamiento estricto de tipos de datos y prevención de NaNs."""
        if self.df is None:
            self.errores.append("No hay datos disponibles para procesar.")
            return None

        try:
            # Limpieza básica de filas totalmente vacías
            df_limpio = self.df.dropna(how='all').copy()

            # Normalizar los nombres de las columnas en el DataFrame de trabajo para asegurar coincidencia interna
            df_limpio.columns = [str(col).strip().lower() for col in df_limpio.columns]

            # Procesamiento de inventario (Cantidad * Precio)
            if 'cantidad' in df_limpio.columns and 'precio' in df_limpio.columns:
                df_limpio['cantidad'] = pd.to_numeric(df_limpio['cantidad'], errors='coerce').fillna(0.0)
                df_limpio['precio'] = pd.to_numeric(df_limpio['precio'], errors='coerce').fillna(0.0)
                df_limpio['total_inventario'] = df_limpio['cantidad'] * df_limpio['precio']

            # Procesamiento adicional de caja o totales si aplican columnas secundarias
            if 'pagado' in df_limpio.columns and 'total_inventario' in df_limpio.columns:
                df_limpio['pagado'] = pd.to_numeric(df_limpio['pagado'], errors='coerce').fillna(0.0)
                df_limpio['diferencia_caja'] = df_limpio['total_inventario'] - df_limpio['pagado']

            return df_limpio
        except Exception as e:
            self.errores.append(f"Error durante el procesamiento matemático de los datos: {str(e)}")
            return None

    def exportar_excel(self, df_resultado):
        """Exporta los datos procesados de vuelta a un archivo Excel optimizado y listo para descarga."""
        try:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_resultado.to_excel(writer, index=False, sheet_name='Resultado_BotExcel')
            output.seek(0)
            return output
        except Exception as e:
            self.errores.append(f"Error al generar el archivo Excel de salida: {str(e)}")
            return None
