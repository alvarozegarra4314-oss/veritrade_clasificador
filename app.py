import sys
import warnings
from pathlib import Path
from io import BytesIO
import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ---------------------------------------------------------------------
st.set_page_config(page_title="Clasificador de Importaciones — Veritrade", page_icon="🗂️", layout="wide")

# Ignorar la advertencia de obsolescencia de la librería de Gemini
warnings.filterwarnings("ignore", category=FutureWarning, module="google.generativeai")

# CSS personalizado para emular el diseño web (Botón principal grande y métricas con fondo)
st.markdown("""
<style>
    /* Estilizar el botón principal de procesar */
    .stButton>button[kind="primary"] {
        height: 3.5rem;
        font-size: 1.2rem;
        font-weight: bold;
        border-radius: 0.5rem;
    }
    /* Darle un fondo suave a las métricas para que parezcan tarjetas */
    div[data-testid="metric-container"] {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 1rem;
        border-radius: 0.5rem;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------
# IMPORTACIONES LOCALES
# ---------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
SRC_DIR = BASE_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from src.pipeline import procesar_dataframe_dinamico
from src.maestro.loader import CargarMaestro
from src.excel_io import sanitizar_dataframe_para_excel
from src.ia_rescate import RescatadorIA, GENAI_DISPONIBLE
from src.maestro_optimizer import guardar_maestro_optimizado

# ---------------------------------------------------------------------
# 0. INICIALIZACIÓN DE VARIABLES DE SESIÓN (SESSION STATE)
# ---------------------------------------------------------------------
if "df_resultado" not in st.session_state:
    st.session_state.df_resultado = None
if "df_export_data" not in st.session_state:
    st.session_state.df_export_data = None
if "maestro_opt_data" not in st.session_state:
    st.session_state.maestro_opt_data = None
if "resumen_opt" not in st.session_state:
    st.session_state.resumen_opt = None
if "linea_producto" not in st.session_state:
    st.session_state.linea_producto = "Producto"
if "proceso_completado" not in st.session_state:
    st.session_state.proceso_completado = False
if "kpis" not in st.session_state:
    st.session_state.kpis = {}

def _obtener_api_key_de_secrets() -> str:
    try:
        return st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        return ""

@st.cache_data(show_spinner=False)
def ejecutar_pipeline_reglas_cached(bytes_raw: bytes, bytes_maestro: bytes, hoja: str):
    df_raw = pd.read_excel(BytesIO(bytes_raw), sheet_name=hoja, engine="openpyxl")
    maestro = CargarMaestro(BytesIO(bytes_maestro))
    return df_raw, maestro

# =====================================================================
# ENCABEZADO
# =====================================================================
st.title("🗂️ Clasificador de Importaciones — Veritrade")
st.markdown("Sube el archivo de datos crudos y el maestro de reglas para procesar e identificar variables técnicas y marcas comerciales.")
st.write("") # Espaciador

# =====================================================================
# SECCIÓN 1: CARGA DE ARCHIVOS
# =====================================================================
c_raw, c_maestro = st.columns(2)

hoja_raw_valida = False
archivo_raw = None
archivo_maestro = None
maestro_info = None

with c_raw:
    with st.container(border=True):
        st.subheader("1. Archivo de Datos Crudos")
        st.caption("Sube el archivo Excel con las descripciones a analizar.")
        archivo_raw = st.file_uploader("Arrastra tu archivo .xlsx aquí", type=["xlsx"], label_visibility="collapsed")
        
        hoja_raw = st.text_input("Nombre de la hoja a procesar", value="Datos")
        
        if archivo_raw is not None:
            try:
                archivo_raw.seek(0)
                hojas_disponibles_raw = pd.ExcelFile(archivo_raw, engine="openpyxl").sheet_names
                if hoja_raw in hojas_disponibles_raw:
                    hoja_raw_valida = True
                else:
                    st.error(f"❌ La hoja **'{hoja_raw}'** no existe. Disponibles: {', '.join(hojas_disponibles_raw)}")
            except Exception as e:
                st.error(f"No se pudo leer el archivo: {e}")

with c_maestro:
    with st.container(border=True):
        st.subheader("2. Maestro de Reglas")
        st.caption("Catálogo de marcas, condiciones y variables paramétricas.")
        archivo_maestro = st.file_uploader("Arrastra tu archivo maestro .xlsx aquí", type=["xlsx"], label_visibility="collapsed", key="up_maestro")
        
        if archivo_maestro is not None:
            try:
                archivo_maestro.seek(0)
                maestro_info = CargarMaestro(ruta_excel=archivo_maestro)
                linea_detectada = maestro_info.config_linea.get("LINEA_PRODUCTO", "Producto")
                n_cond = len(getattr(maestro_info, "condicionales", []))
                
                st.success(f"✅ **Línea:** {linea_detectada} | **Reglas activas:** {n_cond}")
            except Exception as e:
                st.warning(f"Error al leer el maestro: {e}")

# =====================================================================
# SECCIÓN 2: CONFIGURACIÓN DE IA
# =====================================================================
st.write("") # Espaciador
with st.container(border=True):
    st.markdown("### 🤖 Rescate por IA Generativa")
    st.caption("Delega a Gemini el análisis de las descripciones que el motor de reglas no logre resolver. *(Opcional pero recomendado)*.")
    
    if not GENAI_DISPONIBLE:
        st.error("⚠️ El paquete 'google-generativeai' no está instalado.")
        usar_ia = False
        api_key = ""
    else:
        usar_ia = st.toggle("Activar motor de rescate (Gemini 3.1 Flash Lite)", value=False)
        
        api_key = None
        if usar_ia:
            c_key, c_rpm = st.columns(2)
            with c_key:
                api_key = st.text_input("API Key", type="password", value=_obtener_api_key_de_secrets())
                if not api_key:
                    st.warning("Se requiere API Key de Gemini.")
            with c_rpm:
                rpm_limite = st.slider("Límite de Peticiones (RPM)", min_value=1, max_value=60, value=12)

# =====================================================================
# SECCIÓN 3: ACCIÓN PRINCIPAL (PROCESAMIENTO)
# =====================================================================
st.write("") # Espaciador
listo_para_procesar = (archivo_raw and archivo_maestro and hoja_raw_valida and (not usar_ia or api_key))

procesar = st.button(
    "▶️ PROCESAR CLASIFICACIÓN", 
    type="primary", 
    use_container_width=True, 
    disabled=not listo_para_procesar
)

if procesar:
    linea = maestro_info.config_linea.get("LINEA_PRODUCTO", "Producto") if maestro_info else "Producto"
    
    try:
        archivo_raw.seek(0)
        archivo_maestro.seek(0)

        with st.spinner("Aplicando reglas deterministas..."):
            df_raw, maestro = ejecutar_pipeline_reglas_cached(
                archivo_raw.getvalue(), archivo_maestro.getvalue(), hoja_raw
            )

        rescatador = None
        if usar_ia and api_key:
            rescatador = RescatadorIA(api_key=api_key, maestro=maestro, rpm_limite=rpm_limite)
            bar_ia = st.progress(0.0, text="Preparando rescate por IA...")

            def _actualizar_progreso(fase, i, total):
                bar_ia.progress(min(i / total, 1.0), text=f"Rescatando con IA: {i}/{total} descripciones únicas")

            df_resultado = procesar_dataframe_dinamico(
                df_raw, maestro, rescatador_ia=rescatador, progreso_callback=_actualizar_progreso
            )
            bar_ia.empty()
        else:
            with st.spinner("Procesando clasificación dinámica..."):
                df_resultado = procesar_dataframe_dinamico(df_raw, maestro)

        # Guardar en sesión
        st.session_state.df_resultado = df_resultado
        st.session_state.linea_producto = linea
        st.session_state.proceso_completado = True
        
        # Calcular y guardar KPIs
        total_filas = len(df_resultado)
        kpis = {"total": total_filas, "rescatados": 0, "cache": 0, "nuevas": 0, "errores": 0}
        
        if rescatador is not None:
            kpis["rescatados"] = int(df_resultado["Rescatado_Por_IA"].sum()) if "Rescatado_Por_IA" in df_resultado else 0
            kpis["cache"] = rescatador.llamadas_desde_cache
            kpis["errores"] = rescatador.errores
            
            propuestas = getattr(rescatador, "propuestas_aprendizaje", None)
            if propuestas and (propuestas.get("nuevas_marcas") or propuestas.get("nuevas_caracteristicas")):
                archivo_maestro.seek(0)
                buffer_maestro_opt = BytesIO()
                resumen_opt = guardar_maestro_optimizado(
                    ruta_maestro_original=archivo_maestro, propuestas=propuestas, ruta_salida=buffer_maestro_opt
                )
                st.session_state.maestro_opt_data = buffer_maestro_opt.getvalue()
                st.session_state.resumen_opt = resumen_opt
                kpis["nuevas"] = resumen_opt.get("marcas_agregadas", 0) + resumen_opt.get("caracteristicas_agregadas", 0)
            else:
                st.session_state.maestro_opt_data = None
                st.session_state.resumen_opt = None
        else:
            st.session_state.maestro_opt_data = None
            st.session_state.resumen_opt = None
            
        st.session_state.kpis = kpis

        # Buffer del resultado final
        df_export = sanitizar_dataframe_para_excel(df_resultado)
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            df_export.to_excel(writer, index=False, sheet_name="Clasificacion")
        st.session_state.df_export_data = output_buffer.getvalue()

    except Exception as e:
        st.error(f"❌ Ocurrió un error: {e}")
        st.exception(e)

# =====================================================================
# SECCIÓN 4: ÁREA DE RESULTADOS (PERSISTENTE)
# =====================================================================
if st.session_state.proceso_completado and st.session_state.df_resultado is not None:
    st.write("") # Espaciador
    st.divider()
    
    col_titulo, col_tag = st.columns([4, 1])
    with col_titulo:
        st.markdown("### 📥 Resultados y Descargas")
    with col_tag:
        st.success("✅ Proceso Finalizado")

    # KPIs al estilo Dashboard
    kpis = st.session_state.kpis
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Filas", f"{kpis.get('total', 0):,}")
    m2.metric("Rescatados IA", f"{kpis.get('rescatados', 0):,}")
    m3.metric("Ahorro Caché", f"{kpis.get('cache', 0):,}")
    m4.metric("Nuevas Reglas", f"+{kpis.get('nuevas', 0)}")
    
    if kpis.get("errores", 0) > 0:
        st.warning(f"⚠️ {kpis['errores']} descripciones tuvieron errores de conexión con Gemini.")

    st.write("") # Espaciador
    
    # Botones de Descarga Amplios
    d1, d2 = st.columns(2)
    with d1:
        if st.session_state.maestro_opt_data is not None:
            st.download_button(
                label=f"🧠 Descargar Maestro Optimizado",
                data=st.session_state.maestro_opt_data,
                file_name=f"Maestro_Optimizado_{st.session_state.linea_producto}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="btn_descarga_maestro"
            )
        else:
            st.button("🧠 Maestro Optimizado (Sin aprendizajes nuevos)", disabled=True, use_container_width=True)
            
    with d2:
        if st.session_state.df_export_data is not None:
            st.download_button(
                label=f"📊 Descargar Resultado (Excel)",
                data=st.session_state.df_export_data,
                file_name=f"Resultado_{st.session_state.linea_producto}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="btn_descarga_resultado"
            )

    # Vista previa en tabla limpia
    st.write("") # Espaciador
    st.markdown("#### Vista Previa de los Datos (Primeras 15 filas)")
    st.dataframe(st.session_state.df_resultado.head(15), use_container_width=True, hide_index=True)