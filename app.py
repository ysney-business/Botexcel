from datetime import datetime
from io import BytesIO
import pandas as pd
from openpyxl import load_workbook
import streamlit as st

# ==========================================
# CONFIGURACIÓN DE SEGURIDAD Y LICENCIAMIENTO
# ==========================================
# Definición de parámetros de control de acceso (Llave y Tiempo Limitado)
LICENCIA_ACTIVA = True
CLAVE_ACCESSO_MAESTRA = "BOTEXCEL-2026-KEY"
FECHA_EXPIRACION = datetime.strptime("2026-12-31", "%Y-%m-%d")

# ==========================================
# INTERFAZ VISUAL (STREAMLIT)
# ==========================================
st.set_page_config(
    page_title="BOTEXCEL - Automatización Segura",
    page_icon="\U0001F916",
    layout="centered",
)

st.title("\U0001F916 BOTEXCEL - Sistema de Automatización")
st.write(
    "Plataforma protegida para la gestión, limpieza y cálculo automatizado de"
    " hojas de cálculo."
)

# Panel lateral para control de licencias y llaves
st.sidebar.header("\U0001F510 Control de Licenciamiento")
modo_prueba = st.sidebar.checkbox("Activar validación por llave", value=True)

acceso_concedido = True

if modo_prueba:
  clave_ingresada = st.sidebar.text_input(
      "Ingrese su llave de acceso:", type="password"
  )
  if clave_ingresada != CLAVE_ACCESSO_MAESTRA:
    acceso_concedido = False
    if clave_ingresada != "":
      st.sidebar.error("Llave de acceso incorrecta.")
    else:
      st.sidebar.warning("Por favor, ingrese su llave para operar el bot.")

# Verificación de tiempo limitado (Fecha de expiración)
if datetime.now() > FECHA_EXPIRACION:
  acceso_concedido = False
  st.error(
      "\u26A0\uFE0F La licencia temporal de este software ha expirado."
      " Contacte al administrador."
  )

# ==========================================
# NÚCLEO DE LA APLICACIÓN
# ==========================================
if acceso_concedido and LICENCIA_ACTIVA:
  st.success(
      "\u2705 Sistema desbloqueado y operativo. Autoría y propiedad"
      " intelectual protegidas."
  )

  archivo_subido = st.file_uploader(
      "Seleccione o arrastre su archivo de Excel", type=["xlsx", "xls"]
  )

  if archivo_subido is not None:
    try:
      # Lectura segura del archivo en memoria
      df = pd.read_excel(archivo_subido)

      st.write("### \U0001F4CA Vista previa de los datos originales:")
      st.dataframe(df.head())

      if st.button("Ejecutar Automatización Completa"):
        # Procesamiento Lógico (Relleno rápido, autosuma, filtros)

        # 1. Limpieza y normalización de textos en columnas de texto si existen
        for col in df.select_dtypes(include=["object"]).columns:
          df[col] = df[col].astype(str).str.strip().str.title()

        # 2. Extracción condicional del primer nombre si existe la columna Cliente
        if "Cliente" in df.columns:
          df["Nombre_Procesado"] = (
              df["Cliente"].astype(str).str.strip().str.split().str[0]
          )

        # 3. Cálculo de métricas y autosuma si existe la columna Venta o similares
        columnas_numericas = [
            c for c in df.select_dtypes(include=["number"]).columns
        ]

        if "Venta" in df.columns:
          total_ventas = df["Venta"].sum()
          st.info(
              "\U0001F4C8 La autosuma total calculada por el bot (Venta):"
              f" {total_ventas:,.2f}"
          )
        elif len(columnas_numericas) > 0:
          col_principal = columnas_numericas[0]
          total_columna = df[col_principal].sum()
          st.info(
              f"\U0001F4C8 La autosuma total calculada por el bot"
              f" ({col_principal}): {total_columna:,.2f}"
          )

        # 4. Filtrado estricto para evitar filas completamente nulas o vacías
        df_filtrado = df.dropna(how="all")

        # 5. Generación de archivo de salida en memoria con openpyxl
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df_filtrado.to_excel(
              writer, sheet_name="Reporte_Automatizado", index=False
          )
        processed_data = output.getvalue()

        st.success(
            "\U0001F389 ¡Proceso completado con éxito y de forma privada!"
        )

        # Botón de descarga del archivo procesado
        st.download_button(
            label="\U0001F4E5 Descargar Excel Automatizado",
            data=processed_data,
            file_name="reporte_final_automatizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
      st.error(f"\u274C Error crítico en el procesamiento del archivo: {e}")
else:
  st.warning(
      "\U0001F512 Acceso restringido. Introduzca la llave de seguridad válida"
      " en el panel lateral para continuar."
  )