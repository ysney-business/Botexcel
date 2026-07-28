import io
import logging
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Procesador Remoto Resiliente",
    layout="centered"
)

# Configuración de logs optimizada para Streamlit Cloud (salida por consola)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class ProcesadorRemotoResiliente:
    def __init__(self, archivo_entrada: Path, dataframe_en_memoria=None) -> None:
        self.archivo_entrada = archivo_entrada
        self.df_procesado = dataframe_en_memoria

    def _detectar_y_leer(self, uploaded_file) -> pd.DataFrame:
        extension = self.archivo_entrada.suffix.lower()
        
        # Creamos un stream seguro en memoria para evitar problemas de punteros en Streamlit
        bytes_data = uploaded_file.getvalue()
        
        try:
            if extension in ['.xls', '.xlsx']:
                return pd.read_excel(io.BytesIO(bytes_data), dtype=str, engine='openpyxl')
            
            elif extension == '.csv':
                try:
                    return pd.read_csv(io.BytesIO(bytes_data), dtype=str, encoding='utf-8')
                except UnicodeDecodeError:
                    logging.warning("Fallo UTF-8 en CSV. Reintentando con codificación latin-1.")
                    return pd.read_csv(io.BytesIO(bytes_data), dtype=str, encoding='latin-1')
            
            elif extension in ['.txt', '.log']:
                try:
                    return pd.read_csv(io.BytesIO(bytes_data), sep=None, engine='python', dtype=str)
                except Exception as e:
                    raise ValueError(f"No se pudo estructurar el archivo de texto plano: {e}") from e
            
            else:
                raise ValueError(f"Extensión no soportada: {extension}")
                
        except Exception as e:
            logging.error(f"Error crítico al leer el archivo {self.archivo_entrada.name}: {e}")
            raise

    def sanitizar_datos(self, uploaded_file) -> None:
        df = self._detectar_y_leer(uploaded_file)
        
        df = df.fillna("")
        df.columns = df.columns.astype(str).str.strip()
        
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()
            
        self.df_procesado = df
        logging.info(f"Archivo {self.archivo_entrada.name} sanitizado exitosamente.")

    def exportar_a_bytes(self) -> bytes:
        if self.df_procesado is None or self.df_procesado.empty:
            raise ValueError("El DataFrame está vacío o no ha sido procesado.")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            self.df_procesado.to_excel(writer, index=False)
        output.seek(0)
        return output.getvalue()

st.title("Procesador Remoto Multifacético")
st.markdown("Sube cualquier archivo de trabajo recibido (CSV, Excel, TXT, Log), el sistema lo depurará con precisión y entregará un Excel limpio.")

archivo_subido = st.file_uploader("Selecciona o arrastra tu archivo de entrada", type=['csv', 'xlsx', 'xls', 'txt', 'log'])

if archivo_subido is not None:
    ruta_archivo = Path(archivo_subido.name)
    st.info(f"Archivo detectado: {ruta_archivo.name}")
    
    if st.button("Procesar y Sanitizar Datos", type="primary"):
        try:
            with st.spinner("Ejecutando análisis quirúrgico y normalización de datos..."):
                procesador = ProcesadorRemotoResiliente(ruta_archivo)
                procesador.sanitizar_datos(archivo_subido)
                excel_bytes = procesador.exportar_a_bytes()
            
            st.success("Archivo procesado con éxito.")
            
            # Muestra opcional de vista previa de los datos en pantalla
            st.dataframe(procesador.df_procesado.head(10))
            
            nombre_salida = f"limpio_{ruta_archivo.stem}.xlsx"
            
            st.download_button(
                label="Descargar Excel Normalizado",
                data=excel_bytes,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as error:
            st.error(f"Ocurrió una incidencia controlada: {error}")
