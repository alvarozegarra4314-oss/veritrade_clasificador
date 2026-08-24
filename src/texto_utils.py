# Limpieza, normalización y helpers de texto/regex

import re
import unicodedata
import pandas as pd
import numpy as np


def limpiar_texto(texto: str) -> str:
    """Normaliza y remueve tildes/caracteres especiales, devolviendo texto en mayúsculas."""
    # Ojo con el orden: `not pd.NA` lanza TypeError ("boolean value of NA is
    # ambiguous"), así que los faltantes se detectan ANTES de evaluar verdad.
    if texto is None:
        return ""
    if not isinstance(texto, str):
        try:
            if pd.isna(texto):
                return ""
        except (TypeError, ValueError):
            return ""  # no-escalar (lista, etc.) se trata como vacío
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


def identificar_columnas_descripcion(columnas) -> list:
    """
    Identifica dinámicamente las columnas con descripciones comerciales de producto
    y excluye las columnas de descripción arancelaria o de aduanas.
    """
    # Palabras clave que identifican descripciones de producto
    PALABRAS_INCLUSION = [
        'DESCRIPCION', 'DESC_', 'DESC ', 'DESCP', 
        'DETALLE', 'MERCADERIA', 'COMMODITY'
    ]
    
    # Palabras clave que DESCARTAN la columna (descripciones administrativas/arancelarias)
    PALABRAS_EXCLUSION = [
        'PARTIDA', 'ARANCEL', 'NANDINA', 'SUBPARTIDA', 
        'CAPITULO', 'POSICION', 'DECLARACION'
    ]

    cols_identificadas = []

    for col in columnas:
        col_str = str(col).strip()
        col_upper = col_str.upper()

        # 1. Si contiene alguna palabra de exclusión (ej. "PARTIDA"), se ignora de inmediato
        if any(excl in col_upper for excl in PALABRAS_EXCLUSION):
            continue

        # 2. Si contiene alguna palabra de inclusión, se agrega a la lista de descripciones
        if any(incl in col_upper for incl in PALABRAS_INCLUSION):
            cols_identificadas.append(col_str)

    return cols_identificadas