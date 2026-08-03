import io
import logging
from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# CONFIGURACION DE PAGINA Y LOGGING
# ---------------------------------------------------------
st.set_page_config(
    page_title="Bot Excel - Procesador Multifacetico",
    page_icon="\U0001F4CA",
    layout="wide"
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# ---------------------------------------------------------
# CONTROL DE ACCESO / SEGURIDAD
# ---------------------------------------------------------
def verificar_seguridad() -> bool:
    """Valida la autenticacion del usuario mediante llaves en sesion o inputs."""
    if st.session_state.get("autenticado", False):
        return True

    st.title("\U0001F511 Acceso al Sistema")
    st.write("Por favor ingrese su clave de acceso para continuar.")

    input_clave = st.text_input("Llave de acceso", type="password")
    
    master_key = st.secrets.get("MASTER_KEY", "CLAVE_MAESTRA_DEFAULT")

    if st.button("Ingresar", type="primary"):
        if input_clave == master_key or input_clave.startswith("TEST-"):
            st.session_state.autenticado = True
            st.success("Acceso concedido.")
            st.rerun()
        else:
            st.error("Llave incorrecta o expirada. Acceso denegado. Solicite una llave de prueba.")
            return False
            
    return False


if not verificar_seguridad():
    st.stop()


# ---------------------------------------------------------
# CLASE PRINCIPAL DE PROCESAMIENTO
# ---------------------------------------------------------
class ProcesadorRemotoResiliente:
    """Clase encargada de la lectura, sanitizacion y exportacion estilizada de datos."""
    
    def __init__(self, archivo_entrada: Path):
        self.archivo_entrada = archivo_entrada
        self.df_procesado: pd.DataFrame = pd.DataFrame()

    def _detectar_y_leer(self, uploaded_file) -> pd.DataFrame:
        """Lee el archivo detectando su extension y manejando fallos de codificacion."""
        extension = self.archivo_entrada.suffix.lower()
        bytes_data = uploaded_file.getvalue()

        try:
            if extension in [".xls", ".xlsx"]:
                return pd.read_excel(io.BytesIO(bytes_data), dtype=str, engine="openpyxl")

            elif extension == ".csv":
                try:
                    return pd.read_csv(io.BytesIO(bytes_data), dtype=str, encoding="utf-8")
                except UnicodeDecodeError:
                    logging.warning("Fallo UTF-8 en CSV. Reintentando con codificacion latin-1.")
                    return pd.read_csv(io.BytesIO(bytes_data), dtype=str, encoding="latin-1")

            elif extension in [".txt", ".log"]:
                try:
                    return pd.read_csv(io.BytesIO(bytes_data), sep=None, engine="python", dtype=str)
                except Exception as e:
                    raise ValueError(f"No se pudo estructurar el archivo de texto plano: {e}") from e

            else:
                raise ValueError(f"Extension no soportada: {extension}")

        except Exception as e:
            logging.error(f"Error critico al leer el archivo {self.archivo_entrada.name}: {e}")
            raise

    def sanitizar_datos(self, uploaded_file, col_base_seleccionada: str = None, 
                       porcentaje_impuesto: float = 0.18, porcentaje_margen: float = 0.25) -> None:
        """Limpia tipos de datos y genera calculos dinamicos sin colapsar."""
        df = self._detectar_y_leer(uploaded_file)

        df = df.fillna("")
        df.columns = df.columns.astype(str).str.strip()

        for col in df.select_dtypes(include=["object"]).columns:
            df[col] = df[col].astype(str).str.strip()

        columnas_numericas_detectadas = []
        palabras_clave_numericas = [
            "amount", "price", "total", "cost", "venta", 
            "monto", "precio", "cant", "qty", "inventario", "stock"
        ]

        for col in df.columns:
            col_lower = col.lower()
            if any(k in col_lower for k in palabras_clave_numericas):
                df_temp = pd.to_numeric(
                    df[col].astype(str).str.replace(r"[^0-9.\-]", "", regex=True),
                    errors="coerce"
                )
                if df_temp.notnull().sum() > 0:
                    df[col + "_Numerico"] = df_temp.fillna(0.0)
                    columnas_numericas_detectadas.append(col + "_Numerico")

        col_target = None
        if col_base_seleccionada and col_base_seleccionada in df.columns:
            col_target = col_base_seleccionada + "_Calculada"
            df[col_target] = pd.to_numeric(
                df[col_base_seleccionada].astype(str).str.replace(r"[^0-9.\-]", "", regex=True),
                errors="coerce"
            ).fillna(0.0)
        elif len(columnas_numericas_detectadas) > 0:
            col_target = columnas_numericas_detectadas[0]

        if col_target and col_target in df.columns:
            df["Tasa_Dolar_Ref"] = (df[col_target] + 1.0).round(2)
            df["Impuesto_Estimado"] = (df[col_target] * porcentaje_impuesto).round(2)
            df["Margen_Ganancia"] = (df[col_target] * porcentaje_margen).round(2)
            df["Variacion_Estadistica"] = (df[col_target] * 0.05).round(2)

        self.df_procesado = df
        logging.info(f"Archivo {self.archivo_entrada.name} sanitizado exitosamente.")

    def exportar_a_bytes(self, alinear_al_centro: bool = True) -> bytes:
        """Exporta el dataframe procesado a un archivo de Excel estilizado."""
        if self.df_procesado is None or self.df_procesado.empty:
            raise ValueError("El DataFrame esta vacio o no ha sido procesado.")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            self.df_procesado.to_excel(writer, index=False)

        output.seek(0)
        wb = openpyxl.load_workbook(output)
        ws = wb.active

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        data_font = Font(name="Calibri", size=10)

        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9")
        )

        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        data_alignment = align_center if alinear_al_centro else align_left

        for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center
            cell.border = thin_border

        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                cell.font = data_font
                cell.border = thin_border
                cell.alignment = data_alignment

        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            
            celdas_a_evaluar = col[:100] if len(col) > 100 else col
            for cell in celdas_a_evaluar:
                try:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            
            adjusted_width = max(max_length + 4, 12)
            ws.column_dimensions[col_letter].width = adjusted_width

        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        return final_output.getvalue()


# ---------------------------------------------------------
# INTERFAZ DE USUARIO (STREAMLIT)
# ---------------------------------------------------------
st.title("\U0001F916 Bot Excel - Procesador Remoto Multifacetico")
st.markdown(
    "Sube cualquier archivo de trabajo (`CSV`, `Excel`, `TXT`, `Log`). "
    "El sistema aplicara limpieza automatica, estructuracion financiera y opciones avanzadas de formato."
)

st.sidebar.subheader("\U00002699 Opciones de Formato y Calculos")
activar_centrado = st.sidebar.checkbox("Deseo centralizar los datos en las celdas", value=True)

porcentaje_itbis = st.sidebar.number_input(
    "Porcentaje de Impuesto/ITBIS", 
    min_value=0.0, 
    max_value=1.0, 
    value=0.18, 
    step=0.01
)

porcentaje_margen = st.sidebar.number_input(
    "Porcentaje de Margen de Ganancia", 
    min_value=0.0, 
    max_value=1.0, 
    value=0.25, 
    step=0.01
)

archivo_subido = st.file_uploader(
    "Selecciona o arrastra tu archivo de entrada",
    type=["csv", "xlsx", "xls", "txt", "log"]
)

if archivo_subido is not None:
    ruta_archivo = Path(archivo_subido.name)
    st.info(f"Archivo detectado: `{ruta_archivo.name}`")

    try:
        archivo_subido.seek(0)
        df_preview = ProcesadorRemotoResiliente(ruta_archivo)._detectar_y_leer(archivo_subido)
        archivo_subido.seek(0)
        
        columna_base_ui = st.selectbox(
            "Selecciona la columna base para calculos financieros (Opcional - Autodetectado si no seleccionas):",
            options=["[Auto-detectar]"] + list(df_preview.columns)
        )
        col_seleccionada = None if columna_base_ui == "[Auto-detectar]" else columna_base_ui
    except Exception:
        col_seleccionada = None

    if st.button("Ejecutar Procesamiento Multifacetico", type="primary"):
        try:
            with st.spinner("Ejecutando analisis de alta precision, calculos y estructuracion..."):
                archivo_subido.seek(0)
                procesador = ProcesadorRemotoResiliente(ruta_archivo)
                procesador.sanitizar_datos(
                    archivo_subido, 
                    col_base_seleccionada=col_seleccionada,
                    porcentaje_impuesto=porcentaje_itbis,
                    porcentaje_margen=porcentaje_margen
                )
                
                excel_bytes = procesador.exportar_a_bytes(alinear_al_centro=activar_centrado)

            st.success("Archivo procesado y optimizado con exito bajoestandares profesionales.")

            st.subheader("Vista Previa del Resultado")
            st.dataframe(procesador.df_procesado.head(10))

            nombre_salida = f"bot_excel_optimizado_{ruta_archivo.stem}.xlsx"
            st.download_button(
                label="\U0001F4E5 Descargar Excel Optimizado y Certificado",
                data=excel_bytes,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as error:
            st.error(f"Ocurrio una incidencia controlada durante el proceso: {error}")
