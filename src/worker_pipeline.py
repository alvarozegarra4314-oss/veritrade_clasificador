"""Standalone worker for the CPU-bound classification pipeline."""
import argparse
import json
import os
import pickle
import sys
import time
import traceback
from io import BytesIO
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

VALORES_MARCA_SIN_RESOLVER = {
    "MARCA GENERICA", "MARCA PRINCIPAL", "MARCA COMPONENTES",
    "S/M", "SIN MARCA", "GENERICO", "NO APLICA",
}


def _escribir_estado(ruta: Path, **valores) -> None:
    ruta.write_text(json.dumps(valores, ensure_ascii=False), encoding="utf-8")


def _clasificar(bytes_raw: bytes, bytes_maestro: bytes, hoja: str, progreso_callback):
    import pandas as pd

    from src.maestro.loader import CargarMaestro
    from src.pipeline import procesar_dataframe_dinamico

    df_raw = pd.read_excel(BytesIO(bytes_raw), sheet_name=hoja, engine="openpyxl")
    maestro = CargarMaestro(BytesIO(bytes_maestro))
    resultado = procesar_dataframe_dinamico(
        df_raw, maestro, rescatador_ia=None, progreso_callback=progreso_callback
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
        marcas = resultado["Marca_Extraida"].astype(str).str.upper()
        kpis["sin_marca"] = int(marcas.isin(VALORES_MARCA_SIN_RESOLVER).sum())
    else:
        marcas = None

    pend_mask = pd.Series(False, index=resultado.index)
    if var_principal in resultado and marcas is not None:
        pend_mask = resultado[var_principal].isna() & marcas.isin(VALORES_MARCA_SIN_RESOLVER)
    kpis["pendientes"] = int(pend_mask.sum())

    return {
        "df_resultado": resultado,
        "df_pendientes": resultado[pend_mask].copy(),
        "kpis": kpis,
        "maestro_opt_data": None,
        "resumen_opt": None,
    }


def ejecutar_desde_archivos(
    raw_path: Path,
    maestro_path: Path,
    hoja: str,
    resultado_path: Path,
    estado_path: Path,
) -> int:
    ultima_emision = {"tiempo": 0.0, "pct": -1.0}
    _escribir_estado(
        estado_path,
        state="running",
        phase="read",
        progress_pct=0.0,
        message="Leyendo archivo y maestro...",
        updated_at=time.time(),
    )

    def _progreso(fase, actual, total):
        if not total or total <= 0:
            return
        pct = min(actual / total, 1.0)
        ahora = time.monotonic()
        if pct < 1.0 and pct - ultima_emision["pct"] < 0.005 and ahora - ultima_emision["tiempo"] < 1.0:
            return
        if fase == "reglas":
            mensaje = f"Reglas: {actual:,} de {total:,} filas ({pct:.1%})"
        else:
            mensaje = f"Fase 2/2 · IA: {actual:,} de {total:,} descripciones ({pct:.1%})"
        _escribir_estado(
            estado_path,
            state="running",
            phase=fase,
            progress_pct=pct,
            message=mensaje,
            updated_at=time.time(),
        )
        ultima_emision.update(tiempo=ahora, pct=pct)

    try:
        resultado = _clasificar(
            raw_path.read_bytes(), maestro_path.read_bytes(), hoja, _progreso
        )
        temporal_resultado = resultado_path.with_suffix(resultado_path.suffix + ".tmp")
        with temporal_resultado.open("wb") as archivo:
            pickle.dump(resultado, archivo, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(temporal_resultado, resultado_path)
        _escribir_estado(
            estado_path,
            state="completed",
            phase="done",
            progress_pct=1.0,
            message="Clasificación completada.",
            updated_at=time.time(),
        )
        return 0
    except Exception as exc:  # noqa: BLE001
        traceback.print_exc(file=sys.stderr)
        _escribir_estado(
            estado_path,
            state="error",
            phase="error",
            progress_pct=0.0,
            message=f"{type(exc).__name__}: {exc}",
            updated_at=time.time(),
        )
        return 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--raw", type=Path, required=True)
    parser.add_argument("--maestro", type=Path, required=True)
    parser.add_argument("--hoja", required=True)
    parser.add_argument("--resultado", type=Path, required=True)
    parser.add_argument("--estado", type=Path, required=True)
    args = parser.parse_args()
    return ejecutar_desde_archivos(
        args.raw, args.maestro, args.hoja, args.resultado, args.estado
    )


if __name__ == "__main__":
    raise SystemExit(main())