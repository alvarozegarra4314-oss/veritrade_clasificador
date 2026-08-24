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
from src.texto_utils import identificar_columnas_descripcion

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
if "df_pendientes" not in st.session_state:
    st.session_state.df_pendientes = None
if "linea_detectada" not in st.session_state:
    st.session_state.linea_detectada = "Producto"

def _obtener_api_key_de_secrets() -> str:
    try:
        return st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        return ""

# Valores que el motor considera "marca sin resolver" (defaults del maestro)
VALORES_MARCA_SIN_RESOLVER = {
    "MARCA GENERICA", "MARCA PRINCIPAL", "MARCA COMPONENTES",
    "S/M", "SIN MARCA", "GENERICO", "NO APLICA",
}


@st.cache_data(show_spinner=False)
def _listar_hojas_y_filas(bytes_raw: bytes):
    """Lista las hojas de un Excel, su cantidad de filas y columnas con nombre (rápido)."""
    import openpyxl
    wb = openpyxl.load_workbook(BytesIO(bytes_raw), read_only=True, data_only=True)
    info = []
    for ws in wb.worksheets:
        # Leer solo la primera fila para contar columnas con nombre real
        cols_nombradas = 0
        for fila in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            cols_nombradas = sum(1 for v in fila if v is not None and not str(v).startswith("Unnamed"))
        info.append({"nombre": ws.title, "filas": int(ws.max_row or 0), "cols_nombradas": cols_nombradas})
    wb.close()
    return info


def _hoja_recomendada(info_hojas):
    """
    Preselecciona la hoja de datos real:
    prioriza nombres tipo 'Veritrade'/'data'/'2025'/'2024' CON columnas
    nombradas (descarta hojas auxiliares de gráficos que salen 'Unnamed').
    """
    if not info_hojas:
        return None

    def peso(item):
        nombre = item["nombre"].upper()
        es_datos = any(k in nombre for k in ("VERITRADE", "DATA", "2025", "2024"))
        tiene_cols = item.get("cols_nombradas", 0) >= 3
        bonus = 10_000_000 if (es_datos and tiene_cols) else (1_000_000 if es_datos else 0)
        return bonus + item["filas"]

    return max(info_hojas, key=peso)["nombre"]


def _cargar_maestro_incluido():
    """Devuelve (bytes, nombre) del maestro incluido en el proyecto, o (None, None)."""
    ruta = BASE_DIR / "data" / "maestro" / "Maestro_UPS_v2.xlsx"
    if ruta.exists():
        return ruta.read_bytes(), ruta.name
    return None, None


@st.cache_data(show_spinner=False)
def ejecutar_pipeline_reglas_cached(bytes_raw: bytes, bytes_maestro: bytes, hoja: str):
    df_raw = pd.read_excel(BytesIO(bytes_raw), sheet_name=hoja, engine="openpyxl")
    maestro = CargarMaestro(BytesIO(bytes_maestro))
    return df_raw, maestro


@st.cache_data(show_spinner=False)
def _vista_previa_cruda(bytes_raw: bytes, hoja: str) -> pd.DataFrame:
    """Encabezados + 5 filas de una hoja. Cacheado para no re-parsear el
    Excel completo en cada rerun (la validación de columnas corre aquí)."""
    return pd.read_excel(BytesIO(bytes_raw), sheet_name=hoja, nrows=5)


def _mensaje_columnas_no_reconocidas(columnas) -> str:
    """Mensaje accionable cuando la hoja no tiene columnas de descripción."""
    lista_cols = ", ".join(map(str, list(columnas)[:15])) + (" ..." if len(columnas) > 15 else "")
    return (
        "⚠️ La hoja seleccionada no tiene ninguna columna de descripción reconocible, "
        "así que no se puede clasificar.\n\n"
        "**Qué buscamos:** columnas cuyo nombre contenga *DESCRIPCION*, *DETALLE*, "
        "*MERCADERIA* o *COMMODITY*. Las administrativas (*PARTIDA*, *ARANCEL*, "
        "*NANDINA*, *SUBPARTIDA*) se ignoran a propósito.\n\n"
        f"**Columnas encontradas:** {lista_cols}\n\n"
        "Selecciona otra hoja aquí arriba o verifica que el archivo sea el export "
        "Veritrade correcto."
    )

# =====================================================================
# ENCABEZADO
# =====================================================================
st.title("🗂️ Clasificador de Importaciones — Veritrade")
st.markdown("Sube tu archivo de importaciones y obtén la clasificación por producto y marca. **No necesitas saber de reglas:** la herramienta aplica el maestro de la línea automáticamente y, si activas la IA, rescata lo que no logra resolver.")
st.write("") # Espaciador

# =====================================================================
# SECCIÓN 1: CARGA DE ARCHIVOS
# =====================================================================
c_raw, c_maestro = st.columns(2)

hoja_raw_valida = False
archivo_raw = None
maestro_bytes = None
maestro_nombre = None
maestro_info = None

with c_raw:
    with st.container(border=True):
        st.subheader("1. Archivo de Datos Crudos")
        st.caption("Sube el archivo Excel con las descripciones a analizar.")
        archivo_raw = st.file_uploader("Arrastra tu archivo .xlsx aquí", type=["xlsx"], label_visibility="collapsed")

        if archivo_raw is not None:
            try:
                archivo_raw.seek(0)
                info_hojas = _listar_hojas_y_filas(archivo_raw.getvalue())
                nombres_hojas = [i["nombre"] for i in info_hojas]

                if nombres_hojas:
                    recomendada = _hoja_recomendada(info_hojas)
                    idx_default = nombres_hojas.index(recomendada) if recomendada in nombres_hojas else 0
                    hoja_raw = st.selectbox(
                        "Hoja a procesar",
                        nombres_hojas,
                        index=idx_default,
                        help="Se preseleccionó automáticamente la hoja con más datos.",
                    )
                    filas_estimadas = info_hojas[idx_default]["filas"]
                    cols_nombradas = info_hojas[idx_default].get("cols_nombradas", 0)
                    st.caption(f"📄 Hoja **{hoja_raw}** — ~{filas_estimadas:,} filas · {cols_nombradas} columnas")

                    # Validación temprana: la hoja debe tener columnas de
                    # descripción reconocibles ANTES de permitir procesar.
                    # Así el usuario descubre el problema aquí y no tras un error.
                    df_prev = None
                    try:
                        df_prev = _vista_previa_cruda(archivo_raw.getvalue(), hoja_raw)
                    except Exception as e:
                        st.error(f"No se pudo leer la hoja '{hoja_raw}': {e}")

                    if df_prev is not None:
                        cols_desc_detectadas = identificar_columnas_descripcion(df_prev.columns)
                        if cols_desc_detectadas:
                            hoja_raw_valida = True
                            st.caption(f"🔎 Columnas de descripción detectadas: **{', '.join(cols_desc_detectadas)}**")
                            with st.expander("👀 Vista previa del archivo crudo"):
                                st.caption(f"Columnas detectadas: {len(df_prev.columns)}")
                                st.dataframe(df_prev, use_container_width=True, hide_index=True)
                        else:
                            hoja_raw_valida = False
                            st.error(_mensaje_columnas_no_reconocidas(df_prev.columns))
                else:
                    st.error("El archivo no contiene hojas.")
                    hoja_raw = None
            except Exception as e:
                st.error(f"No se pudo leer el archivo: {e}")
                hoja_raw = None
        else:
            hoja_raw = None

with c_maestro:
    with st.container(border=True):
        st.subheader("2. Maestro de Reglas")
        st.caption("Reglas de marcas, condiciones y variables paramétricas.")

        fuente_maestro = st.radio(
            "Fuente del maestro",
            ["Usar maestro incluido (UPS)", "Subir mi propio maestro"],
            horizontal=True,
        )

        if fuente_maestro == "Subir mi propio maestro":
            archivo_maestro_up = st.file_uploader(
                "Arrastra tu archivo maestro .xlsx aquí",
                type=["xlsx"], label_visibility="collapsed", key="up_maestro",
            )
            if archivo_maestro_up is not None:
                archivo_maestro_up.seek(0)
                maestro_bytes = archivo_maestro_up.getvalue()
                maestro_nombre = archivo_maestro_up.name
        else:
            maestro_bytes, maestro_nombre = _cargar_maestro_incluido()
            if maestro_bytes is None:
                st.warning("⚠️ No se encontró el maestro incluido. Cambia a 'Subir mi propio maestro'.")

        if maestro_bytes:
            try:
                maestro_info = CargarMaestro(ruta_excel=BytesIO(maestro_bytes))
                linea_detectada = maestro_info.config_linea.get("LINEA_PRODUCTO", "Producto")
                n_cond = len(getattr(maestro_info, "condicionales", []))
                st.session_state.linea_detectada = linea_detectada
                st.success(f"✅ **Maestro:** {maestro_nombre} | **Línea:** {linea_detectada} | **Reglas activas:** {n_cond}")
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
listo_para_procesar = (archivo_raw and maestro_bytes and hoja_raw_valida and (not usar_ia or api_key))

procesar = st.button(
    "▶️ PROCESAR CLASIFICACIÓN", 
    type="primary", 
    use_container_width=True, 
    disabled=not listo_para_procesar
)

if procesar:
    linea = st.session_state.get("linea_detectada", "Producto")

    # Barra única para TODO el proceso (reglas + rescate IA). Se repinta solo
    # cuando el avance justifica un refresco (~cada 0.5%) para no penalizar
    # la velocidad del pipeline con miles de actualizaciones de UI.
    barra = st.progress(0.0, text="Preparando procesamiento...")
    _estado_progreso = {"ultimo_pct": -1.0}

    def _actualizar_progreso(fase, i, total):
        if not total or total <= 0:
            return
        pct = min(i / total, 1.0)
        if pct < 1.0 and pct - _estado_progreso["ultimo_pct"] < 0.005:
            return
        _estado_progreso["ultimo_pct"] = pct
        if fase == "reglas":
            texto = f"Fase 1/2 · Reglas deterministas: {i:,} de {total:,} filas"
        else:
            texto = f"Fase 2/2 · Rescate con IA: {i:,} de {total:,} descripciones únicas"
        barra.progress(pct, text=texto)

    try:
        archivo_raw.seek(0)

        df_raw, maestro = ejecutar_pipeline_reglas_cached(
            archivo_raw.getvalue(), maestro_bytes, hoja_raw
        )

        rescatador = None
        if usar_ia and api_key:
            rescatador = RescatadorIA(api_key=api_key, maestro=maestro, rpm_limite=rpm_limite)

        df_resultado = procesar_dataframe_dinamico(
            df_raw, maestro, rescatador_ia=rescatador, progreso_callback=_actualizar_progreso
        )
        barra.progress(1.0, text="✅ Clasificación completada")

        # Guardar en sesión
        st.session_state.df_resultado = df_resultado
        st.session_state.linea_producto = linea
        st.session_state.proceso_completado = True

        # ---- KPIs + Cobertura del análisis ----
        total_filas = len(df_resultado)
        kpis = {
            "total": total_filas, "rescatados": 0, "cache": 0, "nuevas": 0, "errores": 0,
            "con_producto": 0, "sin_producto": 0, "sin_marca": 0, "pendientes": 0,
        }

        if "Producto_Declarado" in df_resultado:
            kpis["con_producto"] = int(df_resultado["Producto_Declarado"].notna().sum())
            kpis["sin_producto"] = int(df_resultado["Producto_Declarado"].isna().sum())

        if "Marca_Extraida" in df_resultado:
            kpis["sin_marca"] = int(
                df_resultado["Marca_Extraida"].astype(str).str.upper().isin(VALORES_MARCA_SIN_RESOLVER).sum()
            )

        # Pendientes = sin producto o sin marca
        pendiente_mask = pd.Series(False, index=df_resultado.index)
        if "Producto_Declarado" in df_resultado:
            pendiente_mask |= df_resultado["Producto_Declarado"].isna()
        if "Marca_Extraida" in df_resultado:
            pendiente_mask |= df_resultado["Marca_Extraida"].astype(str).str.upper().isin(VALORES_MARCA_SIN_RESOLVER)
        kpis["pendientes"] = int(pendiente_mask.sum())
        st.session_state.df_pendientes = df_resultado[pendiente_mask].copy()

        if rescatador is not None:
            kpis["rescatados"] = int(df_resultado["Rescatado_Por_IA"].sum()) if "Rescatado_Por_IA" in df_resultado else 0
            kpis["cache"] = rescatador.llamadas_desde_cache
            kpis["errores"] = rescatador.errores

            propuestas = getattr(rescatador, "propuestas_aprendizaje", None)
            if propuestas and (propuestas.get("nuevas_marcas") or propuestas.get("nuevas_caracteristicas")):
                buffer_maestro_opt = BytesIO()
                resumen_opt = guardar_maestro_optimizado(
                    ruta_maestro_original=BytesIO(maestro_bytes), propuestas=propuestas, ruta_salida=buffer_maestro_opt
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

    except ValueError as e:
        # Errores de validación (ej. columnas de descripción no reconocidas):
        # el mensaje ya viene redactado para el usuario final.
        barra.empty()
        st.error(f"❌ No se pudo procesar el archivo:\n\n{e}")
    except Exception as e:
        barra.empty()
        st.error(f"❌ No se pudo completar la clasificación. Detalle: {e}")
        with st.expander("🛠️ Detalles técnicos (compartir con soporte)"):
            st.exception(e)
    finally:
        barra.empty()

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

    # ---- Cobertura del análisis (transparencia para el cliente) ----
    total = max(kpis.get("total", 1), 1)
    con_producto = kpis.get("con_producto", 0)
    sin_producto = kpis.get("sin_producto", 0)
    sin_marca = kpis.get("sin_marca", 0)
    pendientes = kpis.get("pendientes", 0)

    st.markdown("#### 🎯 Cobertura del análisis")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("✅ Con producto", f"{con_producto:,}", help="Filas donde se identificó el tipo de producto (ej. UPS).")
    c2.metric("❌ Sin producto", f"{sin_producto:,}", help="Filas donde no se pudo identificar el producto.")
    c3.metric("🏷️ Sin marca", f"{sin_marca:,}", help="Filas con marca genérica o sin marca.")
    c4.metric("⚠️ Pendientes", f"{pendientes:,}", help="Filas incompletas (sin producto o sin marca).")
    st.progress(min(con_producto / total, 1.0), text=f"Identificación de producto: {con_producto / total:.1%} del total")

    # ---- Registros pendientes de revisión ----
    df_pendientes = st.session_state.get("df_pendientes")
    if pendientes > 0 and df_pendientes is not None and len(df_pendientes) > 0:
        with st.expander(f"🔍 Ver los {pendientes:,} registros pendientes de revisión"):
            st.caption("Registros donde no se identificó el producto o la marca quedó genérica. Descárgalos para depurar el maestro.")
            st.dataframe(df_pendientes.head(100), use_container_width=True, hide_index=True)
            if len(df_pendientes) > 100:
                st.caption(f"Mostrando 100 de {len(df_pendientes):,} filas. Descarga el Excel para ver todas.")
            buffer_pend = BytesIO()
            df_pend_export = sanitizar_dataframe_para_excel(df_pendientes)
            with pd.ExcelWriter(buffer_pend, engine="openpyxl") as writer:
                df_pend_export.to_excel(writer, index=False, sheet_name="Pendientes")
            st.download_button(
                label="📥 Descargar pendientes (Excel)",
                data=buffer_pend.getvalue(),
                file_name=f"Pendientes_{st.session_state.linea_producto}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_descarga_pendientes",
                use_container_width=True,
            )

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

    # Vista previa con filtro de texto
    st.write("") # Espaciador
    st.markdown("#### 📋 Vista Previa de los Datos")
    df_resultado = st.session_state.df_resultado
    filtro = st.text_input("🔎 Filtrar por marca o producto (texto libre)", value="")
    df_vista = df_resultado
    if filtro.strip():
        mask = df_resultado.astype(str).apply(
            lambda col: col.str.contains(filtro.strip(), case=False, na=False)
        ).any(axis=1)
        df_vista = df_resultado[mask]
    st.caption(f"Mostrando {len(df_vista):,} de {len(df_resultado):,} filas")
    st.dataframe(df_vista.head(50), use_container_width=True, hide_index=True)