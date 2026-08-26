# -*- coding: utf-8 -*-
"""
Generador de Maestros de Clasificación a partir de Veritrade Crudo
-----------------------------------------------------------------
Automatiza la creación de un archivo maestro (con la misma estructura
que Maestro_Plantilla.xlsx) a partir de:

1. Un archivo Veritrade crudo del PM (miles de filas).
2. El maestro plantilla como referencia de formato.
3. Contexto de dominio que el PM aporta (características, marcas, patrones).

Flujo:
  Paso 1: Muestreo estratificado determinístico (pandas puro, sin LLM).
  Paso 2: Generación del maestro via una única llamada a Gemini.
  Paso 3: Escritura del Excel con el mismo esquema de hojas.
  Paso 4 (opcional): Validación contra ground truth del PM.
"""

from __future__ import annotations

import json
import re
import hashlib
import logging
from io import BytesIO
from pathlib import Path
from datetime import datetime
from typing import Optional

import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from src.texto_utils import limpiar_texto, identificar_columnas_descripcion
from src.excel_estilos import (
    aplicar_estilo_hoja_excel,
    COLOR_HEADER_BG,
    COLOR_HEADER_TEXT,
    COLOR_BORDER,
)

logger = logging.getLogger("creador_maestro")

# Mínimo y máximo de filas en la muestra enviada a la IA
MUESTRA_MINIMA = 80
MUESTRA_MAXIMA = 200
MUESTRA_DEFAULT = 150


# =====================================================================
# PASO 1: MUESTREO ESTRATIFICADO (sin LLM)
# =====================================================================

def _normalizar_para_dedup(texto: str) -> str:
    """Normaliza un texto para detectar duplicados casi-idénticos."""
    c = limpiar_texto(texto)
    # Eliminar números de modelo/SKU que varían entre filas del mismo producto
    c = re.sub(r'\b[A-Z]{0,3}\d{3,}\b', '', c)
    c = re.sub(r'\s+', ' ', c).strip()
    return c


def _firmas_texto(texto: str) -> dict:
    """
    Extrae 'firmas' de un texto para estratificación:
    - longitud (corta/media/larga)
    - palabras clave técnicas detectadas
    - presencia de números técnicos (voltaje, potencia, etc.)
    """
    clean = limpiar_texto(texto)
    palabras = clean.split()
    n_palabras = len(palabras)

    # Bucket de longitud
    if n_palabras <= 5:
        bucket_len = "corta"
    elif n_palabras <= 15:
        bucket_len = "media"
    else:
        bucket_len = "larga"

    # Detección de patrones numéricos técnicos
    tiene_voltaje = bool(re.search(r'\b(?:1[12]0|220|230|240|110|208|380|400|480)\s*[Vv]', clean))
    tiene_potencia = bool(re.search(r'\b\d+(?:\.\d+)?\s*(?:K?VA|KW|W|AMP)', clean))
    tiene_fases = bool(re.search(r'\b(?:1[FfPp]|2[FfPp]|3[FfPp]|monof[aá]sico|trif[aá]sico|bif[aá]sico)', clean))

    return {
        "n_palabras": n_palabras,
        "bucket_len": bucket_len,
        "tiene_voltaje": tiene_voltaje,
        "tiene_potencia": tiene_potencia,
        "tiene_fases": tiene_fases,
    }


def _concatenar_descripciones(fila: pd.Series, columnas_desc: list[str]) -> str:
    """Concatena todas las columnas de descripción de una fila en un solo texto."""
    partes = []
    for col in columnas_desc:
        val = fila.get(col)
        if pd.notna(val) and str(val).strip():
            partes.append(str(val).strip())
    return " | ".join(partes)


def muestrear_veritrade(
    df_raw: pd.DataFrame,
    n_muestra: int = MUESTRA_DEFAULT,
    columna_filtro_valor: Optional[str] = None,
    columna_filtro_nombre: Optional[str] = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Muestreo estratificado determinístico de un DataFrame Veritrade.

    Estrategia:
    1. Detectar columnas de descripción automáticamente.
    2. Concatenar descripciones para dedup y estratificación.
    3. Deduplicar por texto normalizado (mismo producto → 1 entrada).
    4. Muestrear equilibrando: diversidad de redacción × diversidad de proveedores.
    5. Devolver la muestra + el dataset completo deduplicado.

    Retorna: (df_muestra, df_completo_dedup)
    """
    df = df_raw.copy()

    # Detectar columnas de descripción
    cols_desc = identificar_columnas_descripcion(df.columns)
    if not cols_desc:
        raise ValueError(
            "No se detectaron columnas de descripción en el archivo. "
            f"Columnas disponibles: {', '.join(map(str, df.columns[:15]))}"
        )

    # Columna de texto concatenado para análisis
    df["_texto_concat"] = df.apply(lambda row: _concatenar_descripciones(row, cols_desc), axis=1)

    # Filtrar filas vacías
    df = df[df["_texto_concat"].str.strip().astype(bool)].copy()

    if len(df) == 0:
        raise ValueError("El archivo no contiene filas con datos de descripción.")

    # --- Filtrado opcional por partida arancelaria ---
    if columna_filtro_nombre and columna_filtro_valor:
        col_filtro = [c for c in df.columns if columna_filtro_nombre.upper() in str(c).upper()]
        if col_filtro:
            df = df[df[col_filtro[0]].astype(str).str.contains(columna_filtro_valor, case=False, na=False)]

    # --- Deduplicación por texto normalizado ---
    df["_dedup_key"] = df["_texto_concat"].apply(_normalizar_para_dedup)
    df_completo = df.drop_duplicates(subset=["_dedup_key"], keep="first").copy()
    df_completo = df_completo.drop(columns=["_dedup_key"], errors="ignore")

    # --- Estratificación y muestreo ---
    n_objetivo = min(max(n_muestra, MUESTRA_MINIMA), MUESTRA_MAXIMA, len(df_completo))

    if len(df_completo) <= n_objetivo:
        # Si hay menos filas que el objetivo, devolver todas
        df_muestra = df_completo.copy()
    else:
        # Firmas para estratificación
        firmas = df_completo["_texto_concat"].apply(_firmas_texto)
        df_completo["_bucket_len"] = firmas.apply(lambda f: f["bucket_len"])
        df_completo["_tiene_tech"] = firmas.apply(
            lambda f: f["tiene_voltaje"] or f["tiene_potencia"] or f["tiene_fases"]
        )

        # Distribuir cuotas por bucket de longitud (proporcional pero con mínimo)
        buckets = df_completo["_bucket_len"].value_counts()
        cuotas = {}
        for bucket, count in buckets.items():
            cuotas[bucket] = max(int(n_objetivo * count / len(df_completo)), 10)

        # Ajustar cuotas para que sumen n_objetivo
        total_cuotas = sum(cuotas.values())
        if total_cuotas > n_objetivo:
            # Reducir proporcionalmente
            factor = n_objetivo / total_cuotas
            cuotas = {k: max(int(v * factor), 5) for k, v in cuotas.items()}

        muestras_bucket = []
        for bucket, cuota in cuotas.items():
            sub = df_completo[df_completo["_bucket_len"] == bucket]
            if len(sub) <= cuota:
                muestras_bucket.append(sub)
            else:
                # Dentro de cada bucket, priorizar diversidad técnica
                sub_tech = sub[sub["_tiene_tech"]]
                sub_no_tech = sub[~sub["_tiene_tech"]]
                n_tech = min(len(sub_tech), cuota // 2 + 1)
                n_no_tech = min(len(sub_no_tech), cuota - n_tech)
                elegidos = pd.concat([
                    sub_tech.sample(n=n_tech, random_state=42) if len(sub_tech) > 0 else sub_tech,
                    sub_no_tech.sample(n=n_no_tech, random_state=42) if len(sub_no_tech) > 0 else sub_no_tech,
                ])
                muestras_bucket.append(elegidos)

        df_muestra = pd.concat(muestras_bucket)

        # Si quedan cortas, rellenar con muestreo aleatorio determinista
        if len(df_muestra) < n_objetivo:
            restantes = df_completo[~df_completo.index.isin(df_muestra.index)]
            faltan = n_objetivo - len(df_muestra)
            if len(restantes) > 0:
                extras = restantes.sample(n=min(faltan, len(restantes)), random_state=42)
                df_muestra = pd.concat([df_muestra, extras])

    # Limpiar columnas auxiliares
    for col in ["_texto_concat", "_bucket_len", "_tiene_tech", "_dedup_key"]:
        df_muestra = df_muestra.drop(columns=[col], errors="ignore")
        df_completo = df_completo.drop(columns=[col], errors="ignore")

    df_muestra = df_muestra.head(n_objetivo)
    return df_muestra.reset_index(drop=True), df_completo.reset_index(drop=True)


# =====================================================================
# PASO 2: CONSTRUCCIÓN DEL PROMPT Y LLAMADA A LA IA
# =====================================================================

def _leer_estructura_template(ruta_template: Path | BytesIO) -> dict:
    """
    Lee el maestro plantilla y extrae la estructura de cada hoja
    (nombres de columnas + 1-2 filas de ejemplo) para incluir
    en el prompt como referencia de formato.
    """
    estructura = {}
    try:
        if isinstance(ruta_template, Path):
            xls = pd.ExcelFile(ruta_template)
        else:
            ruta_template.seek(0)
            xls = pd.ExcelFile(ruta_template)

        for sheet_name in xls.sheet_names:
            if "INSTRUCCION" in sheet_name.upper():
                continue  # Saltar hoja de instrucciones
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, nrows=5)
                cols = [str(c).strip() for c in df.columns]
                ejemplos = []
                for _, row in df.head(2).iterrows():
                    fila_ej = {}
                    for c in df.columns:
                        val = row[c]
                        if pd.notna(val):
                            fila_ej[str(c).strip()] = str(val)[:80]
                    if fila_ej:
                        ejemplos.append(fila_ej)
                estructura[sheet_name] = {
                    "columnas": cols,
                    "ejemplos": ejemplos,
                    "n_filas_reales": len(df),
                }
            except Exception:
                continue
        xls.close()
    except Exception as e:
        logger.warning(f"No se pudo leer el template: {e}")

    return estructura


def _construir_prompt_creador(
    estructura_template: dict,
    df_muestra: pd.DataFrame,
    dominio: dict,
    producto: str,
) -> tuple[str, str]:
    """
    Construye el system prompt y el user prompt para la generación del maestro.

    Retorna: (system_prompt, user_prompt)
    """
    # Preparar las descripciones de la muestra
    cols_desc = identificar_columnas_descripcion(df_muestra.columns)
    descripciones = []
    for i, row in df_muestra.iterrows():
        texto = _concatenar_descripciones(row, cols_desc)
        if texto.strip():
            descripciones.append(f"[{i+1}] {texto[:500]}")

    texto_descripciones = "\n".join(descripciones)

    # Preparar contexto de dominio
    texto_caracteristicas = dominio.get("caracteristicas", "No especificadas")
    texto_marcas = dominio.get("marcas_conocidas", "No especificadas")
    texto_patrones = dominio.get("patrones_tecnicos", "No especificados")

    # Serializar estructura del template como referencia
    ref_template = json.dumps(estructura_template, indent=1, ensure_ascii=False, default=str)
    if len(ref_template) > 4000:
        ref_template = ref_template[:4000] + "\n... ( truncado por longitud)"

    system_prompt = f"""Eres un experto en clasificación de productos de comercio exterior (importaciones/exportaciones).
Tu tarea es crear un ARCHIVO MAESTRO de clasificación para la línea de productos "{producto}" basándote en una muestra de descripciones de importaciones reales.

El maestro es un archivo Excel con hojas específicas que un motor de clasificación automática consume. Cada hoja tiene un propósito y formato exacto.

## ESTRUCTURA DEL MAESTRO (formato que DEBES respetar exactamente)

{ref_template}

## INSTRUCCIONES POR HOJA

### Hoja "0b_Config_Linea"
Parámetros clave-valor. Siempre incluir:
- PARAMETRO="LINEA_PRODUCTO", VALOR="{producto}"
- PARAMETRO="COL_DESCRIPCION", VALOR="Descripcion Comercial"

### Hoja "1_Marcas"
Diccionario de marcas: Patrón detectado en texto → Marca estandarizada → Prioridad.
- Prioridad: 1=alta (regla humana), 2=media, 3=baja (aprendizaje IA).
- El "Patrón" es una palabra o frase tal como aparece en las descripciones (para matching directo).
- Incluye TODAS las marcas que identifiques en la muestra. Sé exhaustivo.
- La IA debe NORMALIZAR marcas (ej: "Schneider", "Schneider Electric", "SE" → todos apuntan a "SCHNEIDER ELECTRIC").

### Hoja "1b_Palabras_Ignorar" (Stopwords)
Columna única "Palabra". Incluir palabras que NUNCA deben considerarse marca:
- Términos genéricos: "GENérico", "SIN MARCA", "NO APLICA", etc.
- Palabras técnicas comunes de la categoría: "MODULO", "CABLE", "DISPOSITIVO", etc.
- Reutilizar stopwords genéricos Y agregar específicos de "{producto}".

### Hoja "1c_Marca_Por_Defeito"
Marca a asignar cuando nada coincide:
- Si producto principal=True → nombre del producto (ej: "{producto}")
- Si producto principal=False → "MARCA COMPONENTES"

### Hoja "2_Caracteristicas"
Palabras clave → valor de cada característica categórica.
Cada fila: Variable | Palabra_Clave | Valor_Resultado | Prioridad
- "Variable" es el nombre de la característica (ej: "Tecnologia", "Fases", "Formato", "Gama")
- "Palabra_Clave" son palabras/fragmente que, al aparecer en la descripción, indican ese valor
- "Valor_Resultado" es el valor normalizado de la característica
- Usar regex simple: alternaciones con | (ej: "TRIF|3F|3~|TRIFASICO|3 FASES")
- Prioridad: 1=alta, 2=media, 3=baja

### Hoja "3_Tecnico_Potencia_NOEDIT"
Extracción de valores numéricos técnicos con multiplicadores de unidad.
Cada fila: Variable | Patron | Multiplicador | Valor_Min | Valor_Max | Unidad
- El Patrón es un regex que captura el número (con grupo de captura)
- El Multiplicador convierte a la unidad base (ej: KVA→VA es *1000)
- Solo incluir si hay patrones numéricos relevantes para "{producto}".

### Hoja "4_Tecnico_RegexMarca_NOEDIT"
Patrones regex avanzados para marcas difíciles.
Solo incluir si hay marcas que requieren matching complejo (ej: abreviaturas ambiguas).

### Hoja "5_Condicionales"
Reglas tipo SI-ENTONCES para resolver ambigüedades.
Cada fila: Regla_ID | Variable_Resultado | Valor_Resultado | Prioridad | Variable_Condicion | Operador | Valor_1 | Valor_2 | Es_Numerica
- Ejemplo: SI potencia >= 10000 Y potencia <= 100000 ENTONCES Gama="Alta"

## REGLAS CRÍTICAS
1. Las descripciones de importación son textos de FACTURA: marcas, modelos, especificaciones, usos mezclados.
2. Para "1_Marcas": ser EXHAUSTIVO con marcas reales. No inventar marcas que no aparezcan en las descripciones.
3. Para "2_Caracteristicas": las PALABRAS CLAVE deben aparecer LITERALMENTE en las descripciones de la muestra.
4. Para "3_Tecnico_Potencia": los patrones regex deben capturar correctamente los números de las descripciones.
5. NUNCA inventar datos. Si algo no se puede inferir de la muestra, dejar la hoja vacía o con un comentario.
6. Devolver SOLO el JSON válido, sin texto adicional."""

    user_prompt = f"""## CONTEXTO DEL PM

**Producto a clasificar:** {producto}

**Características diferenciadoras que usa el PM:**
{texto_caracteristicas}

**Marcas conocidas en esta categoría (el PM las menciona):**
{texto_marcas}

**Patrones técnicos numéricos relevantes:**
{texto_patrones}

## MUESTRA DE DESCRIPCIONES REALES ({len(descripciones)} filas únicas)

{texto_descripciones}

---

Genera el JSON del maestro siguiendo EXACTAMENTE el formato del template. Devuelve ÚNICAMENTE el JSON, sin explicaciones."""

    return system_prompt, user_prompt


def _schema_json_maestro() -> dict:
    """Schema JSON que Gemini debe devolver (una key por hoja del maestro)."""
    return {
        "type": "OBJECT",
        "properties": {
            "0b_Config_Linea": {
                "type": "ARRAY",
                "description": "Filas de configuración: cada una con PARAMETRO y VALOR.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "PARAMETRO": {"type": "STRING"},
                        "VALOR": {"type": "STRING"},
                    },
                },
            },
            "1_Marcas": {
                "type": "ARRAY",
                "description": "Diccionario de marcas: patrón de búsqueda → nombre estándar + prioridad.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Patrón detectado en texto": {"type": "STRING"},
                        "Marca estandarizada": {"type": "STRING"},
                        "Prioridad": {"type": "INTEGER"},
                    },
                },
            },
            "1b_Palabras_Ignorar": {
                "type": "ARRAY",
                "description": "Stopwords: palabras que nunca deben ser marca.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Palabra": {"type": "STRING"},
                    },
                },
            },
            "1c_Marca_Por_Defeito": {
                "type": "ARRAY",
                "description": "Marca por defecto cuando nada coincide.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Producto Principal": {"type": "BOOLEAN"},
                        "Marca_Por_Defecto": {"type": "STRING"},
                    },
                },
            },
            "2_Caracteristicas": {
                "type": "ARRAY",
                "description": "Reglas de características categóricas.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Variable": {"type": "STRING"},
                        "Palabra_Clave": {"type": "STRING"},
                        "Valor_Resultado": {"type": "STRING"},
                        "Prioridad": {"type": "INTEGER"},
                    },
                },
            },
            "3_Tecnico_Potencia_NOEDIT": {
                "type": "ARRAY",
                "description": "Patrones de extracción numérica técnica.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Variable": {"type": "STRING"},
                        "Patron": {"type": "STRING"},
                        "Multiplicador": {"type": "NUMBER"},
                        "Valor_Min": {"type": "NUMBER"},
                        "Valor_Max": {"type": "NUMBER"},
                        "Unidad": {"type": "STRING"},
                        "Orden_Prioridad": {"type": "INTEGER"},
                    },
                },
            },
            "4_Tecnico_RegexMarca_NOEDIT": {
                "type": "ARRAY",
                "description": "Regex avanzados de marca (casos difíciles).",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Patron_Regex": {"type": "STRING"},
                        "Marca_Estandar": {"type": "STRING"},
                        "Orden_Prioridad": {"type": "INTEGER"},
                    },
                },
            },
            "5_Condicionales": {
                "type": "ARRAY",
                "description": "Reglas condicionales SI-ENTONCES.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Regla_ID": {"type": "STRING"},
                        "Variable_Resultado": {"type": "STRING"},
                        "Valor_Resultado": {"type": "STRING"},
                        "Prioridad": {"type": "INTEGER"},
                        "Variable_Condicion": {"type": "STRING"},
                        "Operador": {"type": "STRING"},
                        "Valor_1": {"type": "STRING"},
                        "Valor_2": {"type": "STRING"},
                        "Es_Numerica": {"type": "BOOLEAN"},
                    },
                },
            },
        },
    }


def generar_maestro_con_ia(
    api_key: str,
    modelo: str,
    df_muestra: pd.DataFrame,
    ruta_template: Path | BytesIO,
    producto: str,
    dominio: dict,
    progreso_callback=None,
) -> dict:
    """
    Genera el maestro completo via una única llamada a Gemini.

    Args:
        api_key: API key de Google Gemini.
        modelo: Nombre del modelo (ej: "gemini-3.5-flash-lite").
        df_muestra: DataFrame con la muestra estratificada (100-200 filas).
        ruta_template: Ruta al maestro plantilla (Path o BytesIO).
        producto: Nombre del producto a clasificar (ej: "Estabilizadores").
        dominio: dict con "caracteristicas", "marcas_conocidas", "patrones_tecnicos".
        progreso_callback: fn(fase: str, msg: str) opcional para UI.

    Returns:
        dict con claves = nombres de hoja, valores = DataFrames listos para escribir.
    """
    try:
        from google import genai
        from google.genai import types as genai_types
    except ImportError:
        raise RuntimeError(
            "La librería 'google-genai' no está instalada. "
            "Instálala con: pip install google-genai"
        )

    if progreso_callback:
        progreso_callback("template", "Leyendo estructura del maestro plantilla...")

    # 1. Leer estructura del template
    estructura = _leer_estructura_template(ruta_template)

    # 2. Construir prompts
    if progreso_callback:
        progreso_callback("prompt", "Construyendo prompt para la IA...")

    system_prompt, user_prompt = _construir_prompt_creador(
        estructura, df_muestra, dominio, producto
    )

    # 3. Llamada a Gemini con JSON mode
    if progreso_callback:
        progreso_callback("ia", f"Enviando a {modelo}... (~{len(df_muestra)} descripciones)")

    cliente = genai.Client(api_key=api_key)

    response_schema = _schema_json_maestro()

    respuesta = cliente.models.generate_content(
        model=modelo,
        contents=user_prompt,
        config=genai_types.GenerateContentConfig(
            system_instruction=system_prompt,
            response_mime_type="application/json",
            response_schema=response_schema,
            temperature=0.2,
        ),
    )

    texto_respuesta = respuesta.text or ""
    if not texto_respuesta.strip():
        raise RuntimeError("La IA devolvió una respuesta vacía.")

    # 4. Parsear JSON
    if progreso_callback:
        progreso_callback("parseo", "Parseando respuesta de la IA...")

    try:
        datos = json.loads(texto_respuesta)
    except json.JSONDecodeError as e:
        # Intentar extraer JSON de un bloque de código markdown
        match = re.search(r'```(?:json)?\s*\n(.*?)\n```', texto_respuesta, re.DOTALL)
        if match:
            datos = json.loads(match.group(1))
        else:
            raise RuntimeError(f"La IA devolvió JSON inválido: {e}\n\nRespuesta (primeros 500 chars): {texto_respuesta[:500]}")

    # 5. Convertir a DataFrames
    if progreso_callback:
        progreso_callback("dataframe", "Convirtiendo a DataFrames...")

    hojas = _parsear_respuesta_a_dataframes(datos, producto)

    return hojas


def _parsear_respuesta_a_dataframes(datos: dict, producto: str) -> dict[str, pd.DataFrame]:
    """
    Convierte el JSON de la IA en un dict de DataFrames, uno por hoja.
    """
    hojas = {}

    # 0b_Config_Linea
    cfg_rows = datos.get("0b_Config_Linea", [])
    if not cfg_rows:
        cfg_rows = [
            {"PARAMETRO": "LINEA_PRODUCTO", "VALOR": producto},
            {"PARAMETRO": "COL_DESCRIPCION", "VALOR": "Descripcion Comercial"},
        ]
    hojas["0b_Config_Linea"] = pd.DataFrame(cfg_rows)

    # 1_Marcas
    marcas_rows = datos.get("1_Marcas", [])
    if marcas_rows:
        hojas["1_Marcas"] = pd.DataFrame(marcas_rows)
    else:
        hojas["1_Marcas"] = pd.DataFrame(columns=[
            "Patrón detectado en texto", "Marca estandarizada", "Prioridad"
        ])

    # 1b_Palabras_Ignorar
    sw_rows = datos.get("1b_Palabras_Ignorar", [])
    if sw_rows:
        hojas["1b_Palabras_Ignorar"] = pd.DataFrame(sw_rows)
    else:
        hojas["1b_Palabras_Ignorar"] = pd.DataFrame(columns=["Palabra"])

    # 1c_Marca_Por_Defeito
    default_rows = datos.get("1c_Marca_Por_Defeito", [])
    if not default_rows:
        default_rows = [
            {"Producto Principal": True, "Marca_Por_Defecto": producto.upper()},
            {"Producto Principal": False, "Marca_Por_Defecto": "MARCA COMPONENTES"},
        ]
    hojas["1c_Marca_Por_Defeito"] = pd.DataFrame(default_rows)

    # 2_Caracteristicas
    carac_rows = datos.get("2_Caracteristicas", [])
    if carac_rows:
        hojas["2_Caracteristicas"] = pd.DataFrame(carac_rows)
    else:
        hojas["2_Caracteristicas"] = pd.DataFrame(columns=[
            "Variable", "Palabra_Clave", "Valor_Resultado", "Prioridad"
        ])

    # 3_Tecnico_Potencia_NOEDIT
    pot_rows = datos.get("3_Tecnico_Potencia_NOEDIT", [])
    if pot_rows:
        hojas["3_Tecnico_Potencia_NOEDIT"] = pd.DataFrame(pot_rows)
    else:
        hojas["3_Tecnico_Potencia_NOEDIT"] = pd.DataFrame(columns=[
            "Variable", "Patron", "Multiplicador", "Valor_Min", "Valor_Max", "Unidad", "Orden_Prioridad"
        ])

    # 4_Tecnico_RegexMarca_NOEDIT
    regex_rows = datos.get("4_Tecnico_RegexMarca_NOEDIT", [])
    if regex_rows:
        hojas["4_Tecnico_RegexMarca_NOEDIT"] = pd.DataFrame(regex_rows)
    else:
        hojas["4_Tecnico_RegexMarca_NOEDIT"] = pd.DataFrame(columns=[
            "Patron_Regex", "Marca_Estandar", "Orden_Prioridad"
        ])

    # 5_Condicionales
    cond_rows = datos.get("5_Condicionales", [])
    if cond_rows:
        hojas["5_Condicionales"] = pd.DataFrame(cond_rows)
    else:
        hojas["5_Condicionales"] = pd.DataFrame(columns=[
            "Regla_ID", "Variable_Resultado", "Valor_Resultado", "Prioridad",
            "Variable_Condicion", "Operador", "Valor_1", "Valor_2", "Es_Numerica"
        ])

    return hojas


# =====================================================================
# PASO 3: ESCRITURA DEL EXCEL
# =====================================================================

def guardar_maestro_nuevo(
    hojas: dict[str, pd.DataFrame],
    producto: str,
    ruta_salida: Path | BytesIO | None = None,
    resumen_ia: dict | None = None,
) -> bytes:
    """
    Escribe el maestro generado como un archivo Excel con el mismo
    esquema de hojas que Maestro_Plantilla.xlsx.

    Returns:
        bytes del archivo Excel (para download en Streamlit).
    """
    buf = BytesIO()

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Escribir hojas en orden
        orden_hojas = [
            "0b_Config_Linea",
            "1_Marcas",
            "1b_Palabras_Ignorar",
            "1c_Marca_Por_Defeito",
            "2_Caracteristicas",
            "3_Tecnico_Potencia_NOEDIT",
            "4_Tecnico_RegexMarca_NOEDIT",
            "5_Condicionales",
        ]

        for nombre_hoja in orden_hojas:
            df = hojas.get(nombre_hoja)
            if df is not None and len(df) > 0:
                df.to_excel(writer, index=False, sheet_name=nombre_hoja)

        # Hoja de Instrucciones (al final, como referencia)
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
        df_instrucciones = pd.DataFrame([
            ("Producto", producto),
            ("Fecha de generación", ahora),
            ("Generador", "App Clasificador Veritrade — Generador Automático v1"),
            ("", ""),
            ("INSTRUCCIONES DE USO", ""),
            ("", ""),
            ("Este maestro fue generado automáticamente por IA a partir de una muestra de descripciones reales."),
            ("Revisa CADA hoja antes de usarlo en producción:", ""),
            ("  1. 1_Marcas: Verifica que las marcas sean correctas y completas."),
            ("  2. 2_Caracteristicas: Ajusta las palabras clave y valores según tu conocimiento del producto."),
            ("  3. 3_Tecnico_Potencia: Valida los patrones numéricos y multiplicadores."),
            ("  4. 5_Condicionales: Revisa las reglas SI-ENTONCES."),
            ("", ""),
            ("NOTAS TÉCNICAS", ""),
            ("", ""),
            ("Las hojas marcadas _NOEDIT son generadas por IA y pueden requerir ajuste fino."),
            ("La hoja 1b_Palabras_Ignorar (Stopwords) es genérica y puede compartirse entre líneas."),
            ("", ""),
            ("ESTRUCTURA DE HOJAS", ""),
            ("  0b_Config_Linea — Parámetros de configuración de la línea"),
            ("  1_Marcas — Diccionario de marcas (patrón → nombre estándar)"),
            ("  1b_Palabras_Ignorar — Stopwords (palabras que nunca son marca)"),
            ("  1c_Marca_Por_Defeito — Marca por defecto según si es producto principal"),
            ("  2_Caracteristicas — Reglas de características categóricas"),
            ("  3_Tecnico_Potencia_NOEDIT — Extracción de valores numéricos técnicos"),
            ("  4_Tecnico_RegexMarca_NOEDIT — Regex avanzados de marca"),
            ("  5_Condicionales — Reglas condicionales SI-ENTONCES"),
        ], columns=["Parámetro", "Valor"])

        df_instrucciones.to_excel(writer, index=False, sheet_name="Instrucciones")

        # Aplicar estilos a cada hoja
        for nombre_hoja in writer.sheets:
            if nombre_hoja == "Instrucciones":
                continue
            df_hoja = hojas.get(nombre_hoja)
            if df_hoja is not None and len(df_hoja) > 0:
                try:
                    aplicar_estilo_hoja_excel(writer.sheets[nombre_hoja], df_hoja)
                except Exception:
                    pass  # Estilo es cosmético, no bloqueante

    return buf.getvalue()


# =====================================================================
# PASO 4: VALIDACIÓN (opcional)
# =====================================================================

def validar_maestro(
    df_validacion: pd.DataFrame,
    maestro_bytes: bytes,
    columnas_a_validar: list[str] | None = None,
) -> dict:
    """
    Valida un maestro contra clasificaciones manuales del PM.

    Args:
        df_validacion: DataFrame con las filas clasificadas manualmente por el PM.
                       Debe tener las mismas columnas que el output del clasificador
                       más columnas "_manual" (ej: "Marca_Extraida_manual").
        maestro_bytes: Bytes del maestro generado.
        columnas_a_validar: Lista de columnas a comparar (si None, auto-detectar).

    Returns:
        dict con {columna: {coinciden: int, total: int, pct: float}}.
    """
    from src.maestro.loader import CargarMaestro
    from src.pipeline import procesar_dataframe_dinamico

    # Cargar maestro y procesar
    maestro = CargarMaestro(BytesIO(maestro_bytes))
    df_clasificado = procesar_dataframe_dinamico(df_validacion, maestro)

    # Auto-detectar columnas a validar
    if columnas_a_validar is None:
        col_manuales = [c for c in df_validacion.columns if c.endswith("_manual")]
        columnas_a_validar = [c.replace("_manual", "") for c in col_manuales]

    resultados = {}
    for col in columnas_a_validar:
        col_manual = f"{col}_manual"
        if col_manual not in df_validacion.columns or col not in df_clasificado.columns:
            continue

        manual = df_validacion[col_manual].astype(str).str.strip().str.upper()
        auto = df_clasificado[col].astype(str).str.strip().str.upper()

        coinciden = (manual == auto).sum()
        total = len(manual)
        resultados[col] = {
            "coinciden": int(coinciden),
            "total": int(total),
            "pct": float(coinciden / total) if total > 0 else 0.0,
        }

    return resultados
