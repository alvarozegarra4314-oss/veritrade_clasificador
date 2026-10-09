"""Mide el tiempo real de cada fase del pipeline con el archivo de 13k filas.

Script auxiliar de diagnóstico; se elimina al terminar.
"""
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE / "src"))

import pandas as pd

from src.maestro.loader import CargarMaestro
from src.pipeline import procesar_dataframe_dinamico

RAW = BASE / "data" / "raw" / "Veritrade_ups_22_626_crudo.xlsx"
HOJA = "Veritrade"
MAESTRO = BASE / "data" / "maestro" / "Maestro_UPS_v2_14Setiembre.xlsx"

t0 = time.perf_counter()
df = pd.read_excel(RAW, sheet_name=HOJA, engine="openpyxl")
t1 = time.perf_counter()
print(f"lectura Excel      : {t1 - t0:6.2f} s  ({len(df):,} filas x {df.shape[1]} columnas)")

maestro = CargarMaestro(MAESTRO)
t2 = time.perf_counter()
print(f"carga maestro      : {t2 - t1:6.2f} s")

# Muestreo del avance para ver cada cuánto reporta el pipeline.
marcas = []
inicio = time.perf_counter()


def progreso(fase, i, total):
    ahora = time.perf_counter() - inicio
    if not marcas or ahora - marcas[-1][0] >= 15:
        marcas.append((ahora, fase, i, total))
        print(f"  t+{ahora:6.1f}s  {fase:<8} {i:>7,}/{total:,}")


res = procesar_dataframe_dinamico(df, maestro, rescatador_ia=None, progreso_callback=progreso)
t3 = time.perf_counter()

print(f"clasificacion      : {t3 - t2:6.2f} s")
print(f"TOTAL              : {t3 - t0:6.2f} s")
print(f"filas/s            : {len(df) / max(t3 - t2, 0.001):,.0f}")
print(f"columnas resultado : {len(res.columns)}")
print(f"marcas detectadas  : {res['Marca_Extraida'].notna().sum():,}")