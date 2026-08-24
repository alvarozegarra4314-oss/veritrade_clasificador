# -*- coding: utf-8 -*-
"""
Tests unitarios del motor de reglas y utilidades de texto.

Cubren las funciones con más casos borde del pipeline:
  - normalizar_numero_extraido (rangos, miles, decimales con coma)
  - indicadores de "sin marca" y validación de candidatos a marca
  - negaciones previas (NO ONLINE, SIN BATERIA)
  - extracción posicional desde Descripcion 1
  - operadores de condicionales (>=, BETWEEN, ==)
  - limpieza de texto y detección de columnas de descripción

Ejecutar:  python -m pytest tests/ -v
"""

import pandas as pd
import pytest

from src.maestro.reglas import (
    normalizar_numero_extraido,
    es_indicador_sin_marca,
    es_candidato_marca_valido,
    tiene_negacion_previa,
    extraer_producto_y_modelo_desc1,
    evaluar_potencia_numerica_condicion,
    evaluar_categorica_condicion,
)
from src.texto_utils import (
    limpiar_texto,
    parse_bool,
    identificar_columnas_descripcion,
    construir_patron_desde_palabras,
)


# ----------------------------------------------------------------------
# normalizar_numero_extraido
# ----------------------------------------------------------------------
@pytest.mark.parametrize(
    "entrada,esperado",
    [
        ("1,5", 1.5),            # coma decimal
        ("110-220", 220.0),      # rango -> límite superior
        ("10/15", 15.0),         # rango con barra
        ("1,000", 1000.0),       # separador de miles con coma
        ("1.000", 1000.0),       # separador de miles con punto
        ("220V", 220.0),         # texto pegado
        ("", None),
        (None, None),
        ("ABC", None),
    ],
)
def test_normalizar_numero_extraido(entrada, esperado):
    assert normalizar_numero_extraido(entrada) == esperado


# ----------------------------------------------------------------------
# Indicadores "sin marca"
# ----------------------------------------------------------------------
@pytest.mark.parametrize(
    "texto,esperado",
    [
        ("S/M", True),
        ("SIN MARCA", True),
        ("GENERICO", True),
        ("N/A", True),
        ("NO APLICA", True),
        ("APC", False),
        ("SCHNEIDER ELECTRIC", False),
        ("", False),
        (None, False),
    ],
)
def test_es_indicador_sin_marca(texto, esperado):
    assert es_indicador_sin_marca(texto) is esperado


# ----------------------------------------------------------------------
# Validación de candidatos a marca
# ----------------------------------------------------------------------
def test_candidato_marca_valida():
    assert es_candidato_marca_valido("APC", set()) is True
    assert es_candidato_marca_valido("Schneider", set()) is True


@pytest.mark.parametrize(
    "candidato",
    [
        "SUA3000",              # contiene números
        "MARCA DEMASIADO LARGA PARA SER VALIDA",  # >25 caracteres
        "UNO DOS TRES CUATRO",  # más de 3 palabras
        "TOMA",                 # palabra de descarte
        "KVA",                  # unidad técnica
        "SIN MARCA",            # indicador explícito sin marca
        "APC/SCHNEIDER",        # símbolo prohibido
        "",                     # vacío
    ],
)
def test_candidato_marca_invalida(candidato):
    assert es_candidato_marca_valido(candidato, set()) is False


def test_candidato_marca_con_stopwords():
    assert es_candidato_marca_valido("MODELO", {"MODELO"}) is False


# ----------------------------------------------------------------------
# Negaciones previas (NO ONLINE, SIN BATERIA...)
# ----------------------------------------------------------------------
def test_negacion_previa_detectada():
    texto = "EQUIPO NO ONLINE 10KVA"
    pos = texto.index("ONLINE")
    assert tiene_negacion_previa(texto, pos) is True


def test_negacion_previa_ausente():
    texto = "UPS ONLINE 10KVA"
    pos = texto.index("ONLINE")
    assert tiene_negacion_previa(texto, pos) is False


# ----------------------------------------------------------------------
# Extracción posicional desde Descripcion 1
# ----------------------------------------------------------------------
def test_producto_y_modelo_desc1_completo():
    producto, modelo = extraer_producto_y_modelo_desc1(
        "ALLSAIW 10K PRO 3/3 - UPS, ALLSAI, W KPRO"
    )
    assert producto == "ALLSAIW 10K PRO 3/3 - UPS"
    assert modelo == "W KPRO"


def test_producto_y_modelo_desc1_sin_tercera_posicion():
    producto, modelo = extraer_producto_y_modelo_desc1("PRODUCTO, MARCA")
    assert producto == "PRODUCTO"
    assert modelo is None


def test_producto_y_modelo_desc1_vacio():
    assert extraer_producto_y_modelo_desc1("") == (None, None)


# ----------------------------------------------------------------------
# Operadores de condicionales (hoja 5_Condicionales)
# ----------------------------------------------------------------------
@pytest.mark.parametrize(
    "valor,operador,v1,v2,esperado",
    [
        (10, ">=", 10, None, True),
        (9.99, ">=", 10, None, False),
        (15, "BETWEEN", 10, 20, True),
        (50, "BETWEEN", 10, 20, False),
        (15, "BETWEEN", 20, 10, True),   # extremos invertidos se normalizan
        (None, ">", 1, None, False),
        ("abc", ">", 1, None, False),    # no numérico -> False seguro
        (5, "==", 5, None, True),
        (5, "!=", 6, None, True),
    ],
)
def test_evaluar_potencia_numerica_condicion(valor, operador, v1, v2, esperado):
    assert evaluar_potencia_numerica_condicion(valor, operador, v1, v2) is esperado


@pytest.mark.parametrize(
    "actual,operador,esperado_val,resultado",
    [
        ("Trifasico", "==", "TRIFASICO", True),   # comparación insensible a mayúsculas
        ("Online", "==", "Offline", False),
        ("Online", "!=", "Offline", True),
        (None, "==", "X", False),
    ],
)
def test_evaluar_categorica_condicion(actual, operador, esperado_val, resultado):
    assert (
        evaluar_categorica_condicion(actual, operador, esperado_val) is resultado
    )


# ----------------------------------------------------------------------
# Utilidades de texto
# ----------------------------------------------------------------------
def test_limpiar_texto_quita_tildes_y_mayusculiza():
    assert limpiar_texto("Descripción Comercial") == "DESCRIPCION COMERCIAL"


@pytest.mark.parametrize("entrada", ["", None, pd.NA])
def test_limpiar_texto_vacio(entrada):
    assert limpiar_texto(entrada) == ""


@pytest.mark.parametrize(
    "valor,esperado",
    [("SI", True), ("true", True), (1, True), (True, True),
     ("NO", False), ("0", False), (pd.NA, False)],
)
def test_parse_bool(valor, esperado):
    assert parse_bool(valor) is esperado


def test_identificar_columnas_descripcion_incluye_y_excluye():
    columnas = [
        "Partida Aduanera",
        "Descripcion Comercial",
        "Descripcion de la Partida Aduanera",  # administrativa -> excluida
        "Marca",
    ]
    assert identificar_columnas_descripcion(columnas) == ["Descripcion Comercial"]


def test_identificar_columnas_descripcion_sin_coincidencias():
    assert identificar_columnas_descripcion(["Marca", "Qty 2"]) == []


def test_construir_patron_desde_palabras_espacios_flexibles():
    patron = construir_patron_desde_palabras(["UPS Online"])
    assert patron == r"UPS\s+ONLINE"
