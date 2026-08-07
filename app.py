import sys
from pathlib import Path
from io import BytesIO

BASE_DIR = Path(__file__).resolve().parent
SRC_DIR = BASE_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import pandas as pd
import streamlit as st

from config import RAW_SHEET_NAME, SHEET_CONFIG_LINEA
from extractor_maestro import procesar_dataframe_dinamico, CargarMaestro

import re

ILLEGAL_CHARS_RE = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')

def limpiar_para_excel(df):
    df_limpio = df.copy()

    # 1. Desinfectar nombres de columnas
    df_limpio.columns = [
        ILLEGAL_CHARS_RE.sub('', str(col)).replace('_x0000_', '').replace('_x000D_', '').strip()
        for col in df_limpio.columns
    ]

    # 2. Quitar zonas horarias en fechas
    for col in df_limpio.columns:
        if pd.api.types.is_datetime64_any_dtype(df_limpio[col]):
            if hasattr(df_limpio[col].dt, 'tz') and df_limpio[col].dt.tz is not None:
                df_limpio[col] = df_limpio[col].dt.tz_localize(None)

    # 3. Limpiar celdas de texto
    for col in df_limpio.columns:
        if df_limpio[col].dtype == 'object' or df_limpio[col].dtype == 'string':
            df_limpio[col] = df_limpio[col].apply(
                lambda x: (
                    ILLEGAL_CHARS_RE.sub('', str(x))
                    .replace('_x0000_', '')
                    .replace('_x000D_', '\n')
                    .lstrip('=')
                    if pd.notna(x) else x
                )
            )
    return df_limpio


st.set_page_config(page_title="Clasificador de Importaciones — Veritrade", page_icon="🗂️", layout="wide")

st.title("🗂️ Clasificador de Importaciones — Veritrade")
st.caption(
    "Sube el archivo de datos crudos y el maestro de reglas de la línea de producto "
    "que quieras procesar (UPS, Tableros, o cualquier otra ya armada con el mismo formato)."
)

col1, col2 = st.columns(2)
with col1:
    archivo_raw = st.file_uploader("1. Archivo de datos crudos (.xlsx)", type=["xlsx"])
with col2:
    archivo_maestro = st.file_uploader("2. Maestro de reglas (.xlsx)", type=["xlsx"])

hoja_raw = st.text_input("Nombre de la hoja del archivo crudo", value=RAW_SHEET_NAME)


def leer_info_maestro(archivo):
    archivo.seek(0)
    maestro_info = CargarMaestro(ruta_excel=archivo)
    archivo.seek(0)
    return maestro_info


maestro_info = None
if archivo_maestro is not None:
    try:
        maestro_info = leer_info_maestro(archivo_maestro)
        linea_detectada = maestro_info.config_linea.get("LINEA_PRODUCTO", "Producto")
        st.caption(f"📄 Maestro detectado: línea de producto **{linea_detectada}**")
    except Exception as e:
        st.warning(
            f"No pude leer la configuración del maestro (hoja '{SHEET_CONFIG_LINEA}'). "
            f"Puedo seguir procesando, pero los nombres/métricas quedarán genéricos. Detalle: {e}"
        )

procesar = st.button("▶️ Procesar", type="primary", disabled=not (archivo_raw and archivo_maestro))


def procesar_datos_optimizados(raw_file, maestro_file, sheet):
    raw_file.seek(0)
    maestro_file.seek(0)
    df_raw = pd.read_excel(raw_file, sheet_name=sheet, engine="calamine")
    df_resultado = procesar_dataframe_dinamico(df_raw, ruta_maestro=maestro_file)
    return df_resultado


if procesar:
    linea = "Producto"
    valores_marca_default = {"Marca Generica", "Marca Componentes"}
    if maestro_info is not None:
        linea = maestro_info.config_linea.get("LINEA_PRODUCTO", "Producto")
        valores_marca_default = set(maestro_info.dict_defaults.values()) or valores_marca_default

    try:
        with st.spinner("Procesando clasificación de forma optimizada..."):
            df_resultado = procesar_datos_optimizados(archivo_raw, archivo_maestro, hoja_raw)

        st.success(f"✅ Proceso completado: {len(df_resultado)} filas clasificadas ({linea}).")

        buffer = BytesIO()
        df_para_excel = limpiar_para_excel(df_resultado)
        
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_para_excel.to_excel(writer, index=False)
        
        excel_data = buffer.getvalue()

        st.download_button(
            label="⬇️ Descargar resultado (.xlsx)",
            data=excel_data,
            file_name=f"Resultado_Clasificacion_{linea}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("📊 Resumen rápido"):
            c1, c2, c3 = st.columns(3)
            c1.metric("Total filas", len(df_resultado))

            if "Es_Producto_Principal" in df_resultado.columns:
                c2.metric(
                    f"{linea} identificados",
                    int(df_resultado["Es_Producto_Principal"].sum()),
                )

            if "Marca_Extraida" in df_resultado.columns:
                sin_marca = df_resultado["Marca_Extraida"].isin(valores_marca_default).sum()
                c3.metric("Sin marca identificada", int(sin_marca))

    except Exception as e:
        st.error(f"❌ Ocurrió un error durante el procesamiento: {e}")
        st.exception(e)
else:
    st.info("Sube ambos archivos para habilitar el botón de procesar.")