"""Worker del pipeline ejecutado en un proceso independiente.

Por qué un proceso y no un hilo: el motor de reglas es CPU-bound (regex sobre
cada fila). En un hilo comparte el GIL con el servidor de Streamlit, así que en
instancias de un solo nucleo -como las del plan gratuito de Streamlit Cloud- el
hilo acapara la CPU y la interfaz se queda congelada. En un proceso aparte la
memoria no se comparte y el hilo de la interfaz sigue respondiendo.
"""
import sys
from io import BytesIO
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

# Valores que el motor considera "marca sin resolver".
VALORES_MARCA_SIN_RESOLVER = {
    "MARCA GENERICA", "MARCA PRINCIPAL", "MARCA COMPONENTES",
    "S/M", "SIN MARCA", "GENERICO", "NO APLICA",
}


def ejecutar_pipeline(cola, bytes_raw, bytes_maestro, hoja):
    """Clasifica y publica el resultado por la cola.

    La cola recibe mensajes con la forma ("tipo", ...):
      ("progreso", fase, i, total)
      ("fin", DataFrame, kpis, filas_pendientes)
      ("error", mensaje, None, None)
    """
    import pandas as pd

    from src.excel_io import sanitizar_dataframe_para_excel  # noqa: F401
    from src.maestro.loader import CargarMaestro
    from src.pipeline import procesar_dataframe_dinamico

    try:
        df_raw = pd.read_excel(BytesIO(bytes_raw), sheet_name=hoja, engine="openpyxl")
        maestro = CargarMaestro(BytesIO(bytes_maestro))

        def _cb(fase, i, total):
            cola.put(("progreso", fase, i, total))

        resultado = procesar_dataframe_dinamico(
            df_raw, maestro, rescatador_ia=None, progreso_callback=_cb
        )

        total_filas = len(resultado)
        kpis = {
            "total": total_filas, "rescatados": 0, "cache": 0, "nuevas": 0, "errores": 0,
            "con_producto": 0, "sin_producto": 0, "sin_marca": 0, "pendientes": 0,
        }

        var_principal = maestro.variable_producto_principal
        if var_principal in resultado:
            kpis["con_producto"] = int(resultado[var_principal].notna().sum())
            kpis["sin_producto"] = int(resultado[var_principal].isna().sum())

        if "Marca_Extraida" in resultado:
            kpis["sin_marca"] = int(
                resultado["Marca_Extraida"].astype(str).str.upper()
                .isin(VALORES_MARCA_SIN_RESOLVER).sum()
            )

        pend_mask = pd.Series(False, index=resultado.index)
        if var_principal in resultado and "Marca_Extraida" in resultado:
            pend_mask = (
                resultado[var_principal].isna()
                & resultado["Marca_Extraida"].astype(str).str.upper()
                .isin(VALORES_MARCA_SIN_RESOLVER)
            )
        kpis["pendientes"] = int(pend_mask.sum())

        cola.put(("fin", resultado, kpis, resultado[pend_mask].copy()))

    except Exception as exc:  # noqa: BLE001
        cola.put(("error", f"{exc}", None, None))