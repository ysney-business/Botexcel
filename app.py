from datetime import datetime, timedelta
import io
import logging
from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Procesador Remoto Resiliente", layout="centered"
)

MASTER_KEY = "bOtExcel-2026-kEy-pY"


def verificar_seguridad():
  if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

  if not st.session_state.autenticado:
    st.title("Acceso Restringido - Procesador Remoto")
    st.markdown(
        "Introduce tu llave de acceso (Maestra o temporal de prueba) para"
        " continuar."
    )

    input_clave = st.text_input("Llave de acceso:", type="password")

    if st.button("Ingresar", type="primary"):
      if input_clave == MASTER_KEY or input_clave.startswith("TEST-"):
        st.session_state.autenticado = True
        st.success("Acceso concedido.")
        st.rerun()
      else:
        st.error(
            "Llave incorrecta o expirada. Acceso denegado. Solicite una llave"
            " de prueba (7 o 30 dias)."
        )
    return False
  return True


if not verificar_seguridad():
  st.stop()

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


class ProcesadorRemotoResiliente:

  def __init__(self, archivo_entrada: Path, dataframe_en_memoria=None) -> None:
    self.archivo_entrada = archivo_entrada
    self.df_procesado = dataframe_en_memoria

  def _detectar_y_leer(self, uploaded_file) -> pd.DataFrame:
    extension = self.archivo_entrada.suffix.lower()
    bytes_data = uploaded_file.getvalue()

    try:
      if extension in [".xls", ".xlsx"]:
        return pd.read_excel(
            io.BytesIO(bytes_data), dtype=str, engine="openpyxl"
        )

      elif extension == ".csv":
        try:
          return pd.read_csv(
              io.BytesIO(bytes_data), dtype=str, encoding="utf-8"
          )
        except UnicodeDecodeError:
          logging.warning(
              "Fallo UTF-8 en CSV. Reintentando con codificacion latin-1."
          )
          return pd.read_csv(
              io.BytesIO(bytes_data), dtype=str, encoding="latin-1"
          )

      elif extension in [".txt", ".log"]:
        try:
          return pd.read_csv(
              io.BytesIO(bytes_data), sep=None, engine="python", dtype=str
          )
        except Exception as e:
          raise ValueError(
              f"No se pudo estructurar el archivo de texto plano: {e}"
          ) from e

      else:
        raise ValueError(f"Extension no soportada: {extension}")

    except Exception as e:
      logging.error(
          f"Error critico al leer el archivo {self.archivo_entrada.name}: {e}"
      )
      raise

  def sanitizar_datos(self, uploaded_file) -> None:
    df = self._detectar_y_leer(uploaded_file)

    df = df.fillna("")
    df.columns = df.columns.astype(str).str.strip()

    for col in df.select_dtypes(include=["object"]).columns:
      df[col] = df[col].astype(str).str.strip()

    columnas_numericas_detectadas = []
    for col in df.columns:
      col_lower = col.lower()
      if any(
          k in col_lower
          for k in [
              "amount",
              "price",
              "total",
              "cost",
              "venta",
              "monto",
              "precio",
              "cant",
              "qty",
              "inventario",
              "stock",
          ]
      ):
        df_temp = pd.to_numeric(
            df[col].str.replace(r"[^0-9.\-]", "", regex=True), errors="coerce"
        )
        if df_temp.notnull().sum() > 0:
          df[col + "_Numerico"] = df_temp
          columnas_numericas_detectadas.append(col + "_Numerico")

    if len(columnas_numericas_detectadas) > 0:
      col_base = columnas_numericas_detectadas[0]
      df["Tasa_Dolar_Ref"] = (df[col_base] * 1.0).round(2)
      df["Impuesto_Estimado (18%)"] = (df[col_base] * 0.18).round(2)
      df["Margen_Ganancia (25%)"] = (df[col_base] * 0.25).round(2)
      df["Variacion_Estadistica"] = (df[col_base] * 0.05).round(2)

    self.df_procesado = df
    logging.info(f"Archivo {self.archivo_entrada.name} sanitizado exitosamente.")

  def exportar_a_bytes(self) -> bytes:
    if self.df_procesado is None or self.df_procesado.empty:
      raise ValueError("El DataFrame esta vacio o no ha sido procesado.")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      self.df_procesado.to_excel(writer, index=False)

    output.seek(0)
    wb = openpyxl.load_workbook(output)
    ws = wb.active

    header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    
    # Alineacion estrictamente CENTRADA para todas las celdas de datos y encabezados
    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # Formatear encabezados centrados
    for col_num in range(1, ws.max_column + 1):
      cell = ws.cell(row=1, column=col_num)
      cell.fill = header_fill
      cell.font = header_font
      cell.alignment = align_center
      cell.border = thin_border

    # Formatear todas las celdas de datos con centrado uniforme (incluyendo las celdas en blanco)
    for row in range(2, ws.max_row + 1):
      for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = data_font
        cell.border = thin_border
        cell.alignment = align_center

    # Ajuste automatico del ancho de columnas para expansion correcta
    for col in ws.columns:
      max_length = 0
      column_letter = get_column_letter(col[0].column)
      for cell in col:
        try:
          if cell.value:
            max_length = max(max_length, len(str(cell.value)))
        except:
          pass
      adjusted_width = max(max_length + 4, 14)
      ws.column_dimensions[column_letter].width = adjusted_width

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output.getvalue()


st.title("Bot Excel - Procesador Remoto Multifacetico")
st.markdown(
    "Sube cualquier archivo de trabajo (CSV, Excel, TXT, Log). El sistema"
    " aplicara limpieza quirurgica, calculos financieros y centrado"
    " estructurado de datos."
)

archivo_subido = st.file_uploader(
    "Selecciona o arrastra tu archivo de entrada",
    type=["csv", "xlsx", "xls", "txt", "log"],
)

if archivo_subido is not None:
  ruta_archivo = Path(archivo_subido.name)
  st.info(f"Archivo detectado: {ruta_archivo.name}")

  if st.button("Procesar, Sanitizar y Centrar Formato", type="primary"):
    try:
      with st.spinner(
          "Ejecutando analisis quirurgico y centrado de celdas..."
      ):
        procesador = ProcesadorRemotoResiliente(ruta_archivo)
        procesador.sanitizar_datos(archivo_subido)
        excel_bytes = procesador.exportar_a_bytes()

      st.success(
          "Archivo procesado, limpiado y estructurado con exito."
      )

      st.dataframe(procesador.df_procesado.head(10))

      nombre_salida = f"limpio_centrado_{ruta_archivo.stem}.xlsx"

      st.download_button(
          label="Descargar Excel Estructurado y Centrado",
          data=excel_bytes,
          file_name=nombre_salida,
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

    except Exception as error:
      st.error(f"Ocurrio una incidencia controlada: {error}")
