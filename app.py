import sys
import warnings
from pathlib import Path
from io import BytesIO
import pandas as pd
import streamlit as st

# Ignorar la advertencia de obsolescencia de la librería de Gemini para mantener limpia la consola
warnings.filterwarnings("ignore", category=FutureWarning, module="google.generativeai")

BASE_DIR = Path(__file__).resolve().parent
SRC_DIR = BASE_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from src.pipeline import procesar_dataframe_dinamico
from src.maestro.loader import CargarMaestro
from src.excel_io import sanitizar_dataframe_para_excel
from src.ia_rescate import RescatadorIA, GENAI_DISPONIBLE
from src.maestro_optimizer import guardar_maestro_optimizado

st.set_page_config(page_title="Clasificador de Importaciones — Veritrade", page_icon="🗂️", layout="wide")

# =====================================================================
# 0. INICIALIZACIÓN DE VARIABLES DE SESIÓN (SESSION STATE)
# =====================================================================
# Esto evita que los datos y botones de descarga desaparezcan al hacer clic
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

# =====================================================================
# 1. INTERFAZ PRINCIPAL Y CARGA DE ARCHIVOS
# =====================================================================
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

hojas_disponibles_raw = []
hoja_raw_valida = False
if archivo_raw is not None:
    try:
        archivo_raw.seek(0)
        hojas_disponibles_raw = pd.ExcelFile(archivo_raw, engine="openpyxl").sheet_names
        archivo_raw.seek(0)
    except Exception as e:
        st.error(f"No se pudo leer el archivo crudo: {e}")

    if hojas_disponibles_raw:
        if hoja_raw in hojas_disponibles_raw:
            hoja_raw_valida = True
        else:
            st.error(
                f"❌ La hoja **'{hoja_raw}'** no existe en el archivo que subiste. "
                f"Hojas disponibles en este archivo: {', '.join(f'`{h}`' for h in hojas_disponibles_raw)}"
            )

# =====================================================================
# 2. CONFIGURACIÓN DE IA
# =====================================================================
st.divider()
st.subheader("🤖 Rescate por IA Generativa (opcional)")

usar_ia = st.checkbox(
    "Activar rescate por IA (Gemini) para registros que las reglas no lograron clasificar",
    value=False,
    disabled=not GENAI_DISPONIBLE,
    help=(
        "La IA solo actúa sobre registros donde marca o el producto principal "
        "quedaron sin resolver por el motor de reglas. Nunca sobreescribe un "
        "resultado ya confirmado por reglas."
    ),
)
if not GENAI_DISPONIBLE:
    st.caption(
        "⚠️ El paquete 'google-generativeai' no está instalado en este entorno. "
        "Agrégalo a requirements.txt para habilitar esta opción."
    )

api_key = None
LIMITE_ADVERTENCIA_UNICAS = 2000

def _obtener_api_key_de_secrets() -> str:
    try:
        return st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        return ""

if usar_ia:
    api_key = st.text_input(
        "API Key de Gemini",
        type="password",
        value=_obtener_api_key_de_secrets(),
        help="Preferible configurarla en .streamlit/secrets.toml como GEMINI_API_KEY en producción.",
    )
    if not api_key:
        st.warning("Ingresa tu API Key de Gemini para poder usar el rescate por IA.")

    rpm_limite = st.slider(
        "Límite de peticiones por minuto (RPM) a Gemini",
        min_value=1, max_value=60, value=12,
        help=(
            "El tier gratuito de Gemini 1.5 Flash suele permitir ~15 RPM. "
            "Bájalo si sigues viendo errores de cuota; súbelo si tienes un plan de pago."
        ),
    )

@st.cache_data(show_spinner=False)
def ejecutar_pipeline_reglas_cached(bytes_raw: bytes, bytes_maestro: bytes, hoja: str):
    df_raw = pd.read_excel(BytesIO(bytes_raw), sheet_name=hoja, engine="openpyxl")
    maestro = CargarMaestro(BytesIO(bytes_maestro))
    return df_raw, maestro

maestro_info = None
if archivo_maestro is not None:
    try:
        archivo_maestro.seek(0)
        maestro_info = CargarMaestro(ruta_excel=archivo_maestro)
        linea_detectada = maestro_info.config_linea.get("LINEA_PRODUCTO", "Producto")
        vars_cat = ", ".join(maestro_info.variables_categoricas)
        vars_num = ", ".join(maestro_info.variables_potencia)
        st.success(f"📄 **Maestro detectado:** Línea **{linea_detectada}** | **Variables:** {vars_cat}, {vars_num}")

        n_cond = len(getattr(maestro_info, "condicionales", []))
        n_cond_descartadas = len(getattr(maestro_info, "condicionales_descartadas", []))
        if n_cond or n_cond_descartadas:
            st.caption(f"🔀 Condicionales activas para esta línea: **{n_cond}**")
            if n_cond_descartadas:
                with st.expander(f"⚠️ {n_cond_descartadas} condicional(es) descartada(s) por no aplicar a esta línea"):
                    st.dataframe(pd.DataFrame(maestro_info.condicionales_descartadas), use_container_width=True)
    except Exception as e:
        st.warning(f"No se pudo leer la configuración del maestro. Detalle: {e}")


# =====================================================================
# 3. LÓGICA DE PROCESAMIENTO
# =====================================================================
procesar = st.button(
    "▶️ Procesar",
    type="primary",
    disabled=not (archivo_raw and archivo_maestro and hoja_raw_valida and (not usar_ia or api_key)),
)

if procesar:
    linea = maestro_info.config_linea.get("LINEA_PRODUCTO", "Producto") if maestro_info else "Producto"

    try:
        archivo_raw.seek(0)
        archivo_maestro.seek(0)

        with st.spinner("Aplicando reglas deterministas..."):
            df_raw, maestro = ejecutar_pipeline_reglas_cached(
                archivo_raw.getvalue(),
                archivo_maestro.getvalue(),
                hoja_raw,
            )

        rescatador = None
        if usar_ia and api_key:
            rescatador = RescatadorIA(api_key=api_key, maestro=maestro, rpm_limite=rpm_limite)
            progress_bar = st.progress(0.0, text="Preparando rescate por IA...")

            def _actualizar_progreso(fase, i, total):
                if total >= LIMITE_ADVERTENCIA_UNICAS and i == 1:
                    st.warning(
                        f"⚠️ Se detectaron {total} descripciones únicas pendientes de rescate. "
                        "Esto puede tardar varios minutos y consumir cuota de la API."
                    )
                progress_bar.progress(
                    min(i / total, 1.0),
                    text=f"Rescatando con IA: {i}/{total} descripciones únicas",
                )

            df_resultado = procesar_dataframe_dinamico(
                df_raw, maestro,
                rescatador_ia=rescatador,
                progreso_callback=_actualizar_progreso,
            )
            progress_bar.empty()
        else:
            with st.spinner("Procesando clasificación dinámica..."):
                df_resultado = procesar_dataframe_dinamico(df_raw, maestro)

        # ------------------------------------------------------------------
        # GUARDAR RESULTADOS EN LA SESIÓN
        # ------------------------------------------------------------------
        st.session_state.df_resultado = df_resultado
        st.session_state.linea_producto = linea
        st.session_state.proceso_completado = True

        # Renderizar auditoría de IA de inmediato
        if rescatador is not None:
            st.divider()
            st.subheader("📊 Auditoría de Rescate IA")
            n_rescatados = int(df_resultado["Rescatado_Por_IA"].sum()) if "Rescatado_Por_IA" in df_resultado else 0
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Registros rescatados por IA", n_rescatados,
                       f"{n_rescatados / len(df_resultado):.1%} del total" if len(df_resultado) else None)
            c2.metric("Llamadas reales a la API", rescatador.llamadas_realizadas)
            c3.metric("Resueltas desde caché", rescatador.llamadas_desde_cache)
            c4.metric("Errores / cuota excedida", rescatador.errores)

            if rescatador.errores > 0:
                st.warning(
                    f"⚠️ {rescatador.errores} descripciones no pudieron ser rescatadas. "
                    "Esas filas conservan el resultado del motor de reglas."
                )
                with st.expander("Detalle de errores de IA"):
                    st.write(
                        f"- Cuota excedida (429) tras reintentos: **{rescatador.errores_cuota}**\n"
                        f"- Errores de API (no cuota): **{rescatador.errores_api}**\n"
                        f"- Respuesta JSON inválida: **{rescatador.errores_formato}**\n"
                        f"- Otros (red/timeout): **{rescatador.errores_otros}**"
                    )

            # --- Maestro Optimizado en Memoria ---
            propuestas = getattr(rescatador, "propuestas_aprendizaje", None)
            hay_aprendizaje = propuestas and (
                propuestas.get("nuevas_marcas") or propuestas.get("nuevas_caracteristicas") or propuestas.get("revisar_manual")
            )

            if hay_aprendizaje:
                archivo_maestro.seek(0)
                buffer_maestro_opt = BytesIO()
                resumen_opt = guardar_maestro_optimizado(
                    ruta_maestro_original=archivo_maestro,
                    propuestas=propuestas,
                    ruta_salida=buffer_maestro_opt,
                )
                st.session_state.maestro_opt_data = buffer_maestro_opt.getvalue()
                st.session_state.resumen_opt = resumen_opt
            else:
                st.session_state.maestro_opt_data = None
                st.session_state.resumen_opt = None
        else:
            st.session_state.maestro_opt_data = None
            st.session_state.resumen_opt = None

        # --- DataFrame Exportable en Memoria ---
        df_export = sanitizar_dataframe_para_excel(df_resultado)
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            df_export.to_excel(writer, index=False, sheet_name="Clasificacion")
        st.session_state.df_export_data = output_buffer.getvalue()

    except Exception as e:
        st.error(f"❌ Ocurrió un error durante el procesamiento: {e}")
        st.exception(e)

# Validaciones visuales antes de procesar
elif not procesar:
    if usar_ia and not api_key:
        st.info("Ingresa tu API Key de Gemini para habilitar el botón de procesar.")
    elif archivo_raw and not hoja_raw_valida:
        st.info("Corrige el nombre de la hoja del archivo crudo para habilitar el botón de procesar.")
    elif not archivo_raw or not archivo_maestro:
        st.info("Sube ambos archivos para habilitar el botón de procesar.")


# =====================================================================
# 4. ÁREA DE RESULTADOS Y DESCARGAS (PERSISTENTE)
# =====================================================================
# Esta sección siempre se renderizará si hay datos guardados en session_state
if st.session_state.proceso_completado and st.session_state.df_resultado is not None:
    st.divider()
    st.subheader("📥 Descargas y Resultados")
    
    col_d1, col_d2 = st.columns(2)
    
    # 4.1 Botón de Maestro Optimizado
    if st.session_state.maestro_opt_data is not None:
        with col_d1:
            resumen = st.session_state.resumen_opt
            texto_ayuda = ""
            if resumen:
                texto_ayuda = f"{resumen.get('marcas_agregadas', 0)} marca(s) y {resumen.get('caracteristicas_agregadas', 0)} característica(s) nueva(s)."
            
            st.download_button(
                label=f"⬇️ Descargar Maestro Optimizado ({st.session_state.linea_producto})",
                data=st.session_state.maestro_opt_data,
                file_name=f"Maestro_Optimizado_{st.session_state.linea_producto}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help=texto_ayuda,
                key="btn_descarga_maestro"  # Key única para evitar conflicto
            )
    else:
        with col_d1:
            st.caption("No se generó Maestro Optimizado (no se activó IA o no hubo aprendizaje).")

    # 4.2 Botón de Resultado de Clasificación
    if st.session_state.df_export_data is not None:
        with col_d2:
            st.download_button(
                label=f"⬇️ Descargar resultado ({st.session_state.linea_producto})",
                data=st.session_state.df_export_data,
                file_name=f"Resultado_Clasificacion_{st.session_state.linea_producto}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_descarga_resultado" # Key única para evitar conflicto
            )

    # 4.3 Previsualización de los datos
    st.markdown("### Vista previa de los datos")
    st.dataframe(st.session_state.df_resultado.head(15), use_container_width=True)