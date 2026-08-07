import sys
from pathlib import Path
from io import BytesIO
import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
SRC_DIR = BASE_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from extractor_maestro import (
    procesar_dataframe_dinamico,
    CargarMaestro,
    sanitizar_dataframe_para_excel,
)

st.set_page_config(page_title="Clasificador de Importaciones — Veritrade", page_icon="🗂️", layout="wide")

st.title("🗂️ Clasificador de Importaciones — Veritrade")
st.caption(
    "Sube el archivo de datos crudos y el maestro de reglas de cualquier línea de producto "
    "(UPS, Interruptores, Tableros, etc.)."
)

col1, col2 = st.columns(2)
with col1:
    archivo_raw = st.file_uploader("1. Archivo de datos crudos (.xlsx)", type=["xlsx"])
with col2:
    archivo_maestro = st.file_uploader("2. Maestro de reglas (.xlsx)", type=["xlsx"])

hoja_raw = st.text_input("Nombre de la hoja del archivo crudo", value="Datos")


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
        vars_cat = ", ".join(maestro_info.variables_categoricas)
        vars_num = ", ".join(maestro_info.variables_potencia)
        st.success(f"📄 **Maestro detectado:** Línea **{linea_detectada}** | **Variables:** {vars_cat}, {vars_num}")
    except Exception as e:
        st.warning(f"No se pudo leer la configuración del maestro. Detalle: {e}")

procesar = st.button("▶️ Procesar", type="primary", disabled=not (archivo_raw and archivo_maestro))

if procesar:
    linea = maestro_info.config_linea.get("LINEA_PRODUCTO", "Producto") if maestro_info else "Producto"

    try:
        with st.spinner("Procesando clasificación dinámica..."):
            archivo_raw.seek(0)
            archivo_maestro.seek(0)

            df_raw = pd.read_excel(archivo_raw, sheet_name=hoja_raw, engine="openpyxl")
            df_resultado = procesar_dataframe_dinamico(df_raw, ruta_maestro=archivo_maestro)

        st.success(f"✅ Proceso completado: {len(df_resultado)} filas clasificadas ({linea}).")

        # Sanitizar para evitar corrupción XML en Streamlit Cloud / Linux
        df_export = sanitizar_dataframe_para_excel(df_resultado)

        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            df_export.to_excel(writer, index=False, sheet_name="Clasificacion")
        
        output_buffer.seek(0)
        excel_data = output_buffer.getvalue()

        st.download_button(
            label=f"⬇️ Descargar resultado ({linea}) (.xlsx)",
            data=excel_data,
            file_name=f"Resultado_Clasificacion_{linea}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.dataframe(df_resultado.head(10), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Ocurrió un error durante el procesamiento: {e}")
        st.exception(e)
else:
    st.info("Sube ambos archivos para habilitar el botón de procesar.")