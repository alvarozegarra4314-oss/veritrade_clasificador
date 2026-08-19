# -*- coding: utf-8 -*-
"""Generacion de reportes Excel con resultados y resumen ejecutivo."""

from __future__ import annotations

import os
import re
import tempfile
from io import BytesIO
from pathlib import Path
from typing import IO

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

COLOR_HEADER = "1F4E78"
COLOR_SUBHEADER = "D9EAF7"
COLOR_TEXT = "FFFFFF"
COLOR_BODY = "FFFFFF"
COLOR_BORDER = "D9D9D9"
BORDER = Border(*(Side(style="thin", color=COLOR_BORDER) for _ in range(4)))


def _nombre_columna(columnas, patrones: tuple[str, ...]) -> str | None:
    for columna in columnas:
        normalizada = re.sub(r"[^A-Z0-9]", "", str(columna).upper())
        if any(patron in normalizada for patron in patrones):
            return columna
    return None


def _numero(serie: pd.Series) -> pd.Series:
    """Convierte valores numericos con formatos comunes de Excel/Veritrade."""
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce").fillna(0)
    texto = serie.astype(str).str.strip()
    tiene_coma = texto.str.contains(",", regex=False)
    texto = texto.where(
        ~tiene_coma,
        texto.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
    )
    return pd.to_numeric(texto, errors="coerce").fillna(0)


def _normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Garantiza encabezados no vacios y unicos, requisito de Excel Table."""
    resultado = df.copy()
    nombres = []
    vistos: dict[str, int] = {}
    for indice, columna in enumerate(resultado.columns, start=1):
        base = re.sub(r"[\x00-\x1f]", "", str(columna)).strip() or f"Columna_{indice}"
        clave = base.casefold()
        vistos[clave] = vistos.get(clave, 0) + 1
        nombres.append(base if vistos[clave] == 1 else f"{base}_{vistos[clave]}")
    resultado.columns = nombres
    return resultado


def _es_caracteristica(columna: str, marca_columna: str | None) -> bool:
    nombre = str(columna).upper()
    excluidas = {"DESCRIPCION", "PRODUCTO_TEXTO", "MODELO_SERIE", "ORIGEN", "RESCATADO", "ES_PRODUCTO"}
    if columna == marca_columna or any(palabra in nombre for palabra in excluidas):
        return False
    return any(palabra in nombre for palabra in ("PRODUCTO", "TIPO", "SALIDA", "TECNOLOGIA", "FASES", "POTENCIA", "KVA", "TENSION", "VOLTAJE"))


def _resumen_marcas(df: pd.DataFrame, marca_columna: str | None, qty_columna: str | None, fob_columna: str | None) -> pd.DataFrame:
    if not marca_columna:
        return pd.DataFrame(columns=["Marca", "Registros", "Qty2_Total", "FOB_Total", "Participacion"])
    trabajo = pd.DataFrame({"Marca": df[marca_columna].fillna("Sin marca").replace("", "Sin marca")})
    trabajo["Registros"] = 1
    if qty_columna:
        trabajo["Qty2_Total"] = _numero(df[qty_columna])
    if fob_columna:
        trabajo["FOB_Total"] = _numero(df[fob_columna])
    resumen = trabajo.groupby("Marca", as_index=False).sum(numeric_only=True).sort_values("Registros", ascending=False)
    resumen["Participacion"] = resumen["Registros"] / max(len(df), 1)
    resumen.insert(0, "Ranking", range(1, len(resumen) + 1))
    return resumen


def _resumen_caracteristicas(df: pd.DataFrame, marca_columna: str | None, qty_columna: str | None, fob_columna: str | None) -> pd.DataFrame:
    columnas = [columna for columna in df.columns if _es_caracteristica(columna, marca_columna)]
    filas = []
    for columna in columnas:
        trabajo = pd.DataFrame({"Caracteristica": columna, "Valor": df[columna].fillna("Sin valor").replace("", "Sin valor")})
        trabajo["Registros"] = 1
        if qty_columna:
            trabajo["Qty2_Total"] = _numero(df[qty_columna])
        if fob_columna:
            trabajo["FOB_Total"] = _numero(df[fob_columna])
        filas.append(trabajo.groupby(["Caracteristica", "Valor"], as_index=False).sum(numeric_only=True))
    if not filas:
        return pd.DataFrame(columns=["Caracteristica", "Valor", "Registros", "Qty2_Total", "FOB_Total"])
    return pd.concat(filas, ignore_index=True).sort_values("Registros", ascending=False)


def _generar_libro(df: pd.DataFrame, destino) -> None:
    """Escribe el libro completo con xlsxwriter, sin reabrirlo con openpyxl."""
    df = _normalizar_columnas(df)
    marca_columna = _nombre_columna(df.columns, ("MARCAEXTRAIDA", "MARCADECLARADA", "MARCAESTANDARIZADA", "MARCA"))
    qty_columna = _nombre_columna(df.columns, ("QTY2", "CANTIDAD", "QUANTITY"))
    fob_columna = _nombre_columna(df.columns, ("FOBTOTAL", "FOB"))
    resumen_marcas = _resumen_marcas(df, marca_columna, qty_columna, fob_columna)
    resumen_caracteristicas = _resumen_caracteristicas(df, marca_columna, qty_columna, fob_columna)

    with pd.ExcelWriter(destino, engine="xlsxwriter") as writer:
        workbook = writer.book
        header = workbook.add_format({"bold": True, "font_color": COLOR_TEXT, "bg_color": COLOR_HEADER, "border": 1, "border_color": COLOR_BORDER})
        title = workbook.add_format({"bold": True, "font_size": 18, "font_color": COLOR_HEADER})
        section = workbook.add_format({"bold": True, "font_size": 13, "font_color": COLOR_HEADER})
        subtitle = workbook.add_format({"italic": True, "font_color": "666666"})
        percent = workbook.add_format({"num_format": "0.0%", "border": 1, "border_color": COLOR_BORDER})
        body = workbook.add_format({"border": 1, "border_color": COLOR_BORDER, "bg_color": COLOR_BODY})
        number = workbook.add_format({"num_format": "#,##0.00", "border": 1, "border_color": COLOR_BORDER})

        df.to_excel(writer, index=False, sheet_name="Clasificacion")
        resultados = writer.sheets["Clasificacion"]
        resultados.hide_gridlines(2)
        resultados.freeze_panes(1, 0)
        resultados.add_table(0, 0, len(df), len(df.columns) - 1, {
            "name": "Tabla_Clasificacion",
            "style": "Table Style Medium 2",
            "columns": [{"header": str(columna)} for columna in df.columns],
        })
        resultados.set_column(0, len(df.columns) - 1, 15, body)

        resumen = workbook.add_worksheet("Resumen Ejecutivo")
        writer.sheets["Resumen Ejecutivo"] = resumen
        resumen.hide_gridlines(2)
        resumen.freeze_panes(3, 0)
        resumen.write("A1", "Resumen Ejecutivo de Clasificacion", title)
        resumen.write("A2", f"Registros analizados: {len(df):,} | Marca: {marca_columna or 'no disponible'} | Qty2: {qty_columna or 'no disponible'} | FOB: {fob_columna or 'no disponible'}", subtitle)

        def escribir_seccion(fila, titulo, explicacion, datos, nombre_tabla):
            resumen.write(fila, 0, titulo, section)
            resumen.write(fila + 1, 0, explicacion, subtitle)
            inicio = fila + 3
            datos.to_excel(writer, sheet_name="Resumen Ejecutivo", startrow=inicio, startcol=0, index=False)
            fin = inicio + len(datos)
            if len(datos):
                resumen.add_table(inicio, 0, fin, len(datos.columns) - 1, {
                    "name": nombre_tabla,
                    "style": "Table Style Medium 2",
                    "columns": [{"header": str(columna)} for columna in datos.columns],
                })
            return fin + 3

        fila = escribir_seccion(4, "Ranking y participacion por marca", "Cantidad de registros, volumen, FOB total y participacion sobre el total.", resumen_marcas, "Tabla_Resumen_Marcas")
        escribir_seccion(fila, "Resumen por caracteristica", "Distribucion de valores clasificados con conteo, Qty2 y FOB cuando existen en el resultado.", resumen_caracteristicas, "Tabla_Resumen_Caracteristicas")
        resumen.set_column(0, 8, 18)


def generar_reporte_excel(df: pd.DataFrame, destino) -> None:
    """Genera el Excel final en una ruta o BytesIO, sin alterar el DataFrame."""
    if df is None or df.empty:
        raise ValueError("No hay datos para generar el reporte Excel.")
    if isinstance(destino, BytesIO):
        _generar_libro(df, destino)
        return
    ruta = Path(destino)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    buffer = BytesIO()
    try:
        _generar_libro(df, buffer)
        fd, temporal = tempfile.mkstemp(suffix=".xlsx", dir=ruta.parent)
        with os.fdopen(fd, "wb") as archivo:
            archivo.write(buffer.getvalue())
        os.replace(temporal, ruta)
    finally:
        if "temporal" in locals() and os.path.exists(temporal):
            os.remove(temporal)
