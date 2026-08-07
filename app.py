import sys
from pathlib import Path
from io import BytesIO
import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
SRC_DIR = BASE_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from config import RAW_SHEET_NAME, SHEET_CONFIG_LINEA
from extractor_maestro import procesar_dataframe_dinamico, CargarMaestro


st.set_page_config(page_title="Clasificador de Importaciones — Veritrade", page_icon="🗂️", layout="wide")

st.title("🗂️ Clasificador de Importaciones — Veritrade")
st.caption(
    "Sube el archivo de datos crudos y el maestro de reglas de la línea de producto "
    "que quieras procesar (UPS, Tableros, o cualquier otra con la misma estructura)."
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
            f"No se pudo leer la configuración del maestro (hoja '{SHEET_CONFIG_LINEA}'). "
            f"Se usará configuración genérica. Detalle: {e}"
        )

procesar = st.button("▶️ Procesar", type="primary", disabled=not (archivo_raw and archivo_maestro))


def procesar_datos_identico_al_main(raw_file, maestro_file, sheet):
    # Resetear punteros de archivos subidos
    raw_file.seek(0)
    maestro_file.seek(0)

    # 1. Leer exactamente igual que en main.py (usando openpyxl)
    df_raw = pd.read_excel(raw_file, sheet_name=sheet, engine="openpyxl")

    # 2. Procesar clasificación
    df_resultado = procesar_dataframe_dinamico(df_raw, ruta_maestro=maestro_file)
    return df_resultado


if procesar:
    linea = "Producto"
    valores_marca_default = {"Marca Generica", "Marca Componentes"}
    if maestro_info is not None:
        linea = maestro_info.config_linea.get("LINEA_PRODUCTO", "Producto")
        valores_marca_default = set(maestro_info.dict_defaults.values()) or valores_marca_default

    try:
        with st.spinner("Procesando clasificación..."):
            df_resultado = procesar_datos_identico_al_main(archivo_raw, archivo_maestro, hoja_raw)

        st.success(f"✅ Proceso completado: {len(df_resultado)} filas clasificadas ({linea}).")

        # 3. Guardar en memoria exactamente igual que en main.py (sin alterar tipos ni datos)
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, index=False)
        
        excel_data = output_buffer.getvalue()

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