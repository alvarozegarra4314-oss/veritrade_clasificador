# Limpieza, normalización y helpers de texto/regex

import re
import unicodedata
import pandas as pd
import numpy as np


def limpiar_texto(texto: str) -> str:
    """Normaliza y remueve tildes/caracteres especiales, devolviendo texto en mayúsculas."""
    if not texto or pd.isna(texto):
        return ""
    texto_str = str(texto)
    nfkd_form = unicodedata.normalize('NFKD', texto_str)
    sin_tildes = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return sin_tildes.upper().strip()


def parse_bool(val) -> bool:
    """Convierte texto o números a booleano de forma segura."""
    if pd.isna(val):
        return False
    if isinstance(val, bool):
        return val
    val_str = str(val).strip().upper()
    return val_str in ['TRUE', 'VERDADERO', '1', 'T', 'SI', 'YES']


def parse_float(val, default: float = 1.0) -> float:
    """Convierte un multiplicador a float de forma segura sin retornar NaN."""
    if pd.isna(val):
        return default
    try:
        f = float(val)
        return f if not np.isnan(f) else default
    except (ValueError, TypeError):
        return default


def construir_patron_desde_palabras(lista_palabras: list) -> str:
    """Compila una lista de palabras clave en un patrón OR regex con escape de caracteres."""
    partes = []
    for palabra in lista_palabras:
        palabra_limpia = limpiar_texto(str(palabra))
        if not palabra_limpia:
            continue
        escapada = re.escape(palabra_limpia).replace(r'\ ', r'\s+')
        partes.append(escapada)
    return '|'.join(partes)


def identificar_columnas_descripcion(df_columns) -> list:
    """Detecta de forma dinámica las columnas que contienen descripciones comerciales."""
    cols_desc = []
    for col in df_columns:
        c_upper = str(col).strip().upper()
        if any(k in c_upper for k in ['DESCRIPCION', 'DESC_', 'DESC ', 'DETALLE', 'MERCADERIA', 'COMMODITY']):
            cols_desc.append(col)
    
    if not cols_desc:
        cols_desc = list(df_columns)
        
    return cols_desc