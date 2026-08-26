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
MUESTRA_MAXIMA = 500
MUESTRA_DEFAULT = 500


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
    El prompt es AUTOCONTENIDO: extrae toda la información directamente de las
    descripciones reales, sin depender de que el PM responda el formulario.

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

    # Contexto adicional del PM (solo si lo proporcionó)
    contexto_extra = ""
    if dominio.get("caracteristicas"):
        contexto_extra += f"\n**El PM indica que estas son las características diferenciadoras:**\n{dominio['caracteristicas']}\n"
    if dominio.get("marcas_conocidas"):
        contexto_extra += f"\n**Marcas conocidas por el PM:**\n{dominio['marcas_conocidas']}\n"
    if dominio.get("patrones_tecnicos"):
        contexto_extra += f"\n**Patrones técnicos que el PM menciona:**\n{dominio['patrones_tecnicos']}\n"

    # Serializar estructura del template como referencia
    ref_template = json.dumps(estructura_template, indent=1, ensure_ascii=False, default=str)
    if len(ref_template) > 4000:
        ref_template = ref_template[:4000] + "\n... ( truncado por longitud)"

    system_prompt = f"""Eres un experto MUNDIAL en clasificación de productos de comercio exterior (importaciones/exportaciones).
Tu tarea es crear un ARCHIVO MAESTRO de clasificación COMPLETO y DETALLADO para la línea de productos "{producto}".

El maestro es un archivo Excel con hojas específicas que un motor de clasificación automática consume.
Debes generar REGLAS EXTENSAS, EXHAUSTIVAS y PRECISAS basándote en las descripciones de importación reales.

## IMPORTANTE: CANTIDADES MÍNIMAS OBLIGATORIAS
El maestro de referencia para esta categoría tiene:
- 72 marcas → TÚ debes generar MÍNIMO 50 marcas
- 284 reglas de características → TÚ debes generar MÍNIMO 150 reglas
- 16 patrones técnicos → TÜ debes generar MÍNIMO 10 patrones
- 4 reglas condicionales → TÚ debes generar MÍNIMO 8 reglas
- 21 stopwords → TÚ debes generar MÍNIMO 30 stopwords
Si generas MENOS que estos mínimos, tu respuesta es INACEPTABLE. Vuelve a analizar.

## ESTRUCTURA DEL MAESTRO (formato que DEBES respetar exactamente)

{ref_template}

## INSTRUCCIONES DETALLADAS POR HOJA — USA LOS NOMBRES DE COLUMNAS DEL TEMPLATE EXACTAMENTE

### Hoja "0b_Config_Linea"
Columnas EXACTAS: "Parametro" | "Valor"
Siempre incluir estas 3 filas:
- Parametro="LINEA_PRODUCTO", Valor="{producto}"
- Parametro="VARIABLE_PRODUCTO_PRINCIPAL", Valor="Tipo_Producto_Detallado"
- Parametro="VALOR_PRODUCTO_PRINCIPAL", Valor="{producto} Sistema Completo"

### Hoja "1_Marcas" — MÍNIMO 50 marcas (el template de referencia tiene 72)
Columnas EXACTAS: "Patron_Busqueda" | "Marca_Estandar" | "Prioridad"
- "Patron_Busqueda": palabra o frase TAL COMO aparece en las descripciones (para matching directo).
- "Marca_Estandar": nombre normalizado de la marca.
- Prioridad: 1=alta, 2=media, 3=baja.
- NORMALIZAR: variantes de la misma marca → mismo nombre estándar.
- Ejemplo de normalización:
  "APC", "A.P.C", "APC BY SCHNEIDER", "AMERICAN POWER CONVERSION" → "APC"
  "SCHNEIDER", "SCHNEIDER ELECTRIC", "SE" → "SCHNEIDER ELECTRIC"
  "EATON", "EATON POWER QUALITY" → "EATON"
- Incluye TAMBIÉN: siglas de marcas, nombres parciales, variantes con/espacios/puntos/guiones.
- Sé EXHAUSTIVO: si ves una marca en las descripciones, DEBE aparecer aquí.

### Hoja "1b_Palabras_Ignorar" — MÍNIMO 30 stopwords (el template tiene 21)
Columnas EXACTAS: "Palabra_Ignorar" | "Categoria"
- "Palabra_Ignorar": palabra que NUNCA debe considerarse marca.
- "Categoria": clasificación de la palabra ("Término Técnico", "Origen / Comercial", "Entidad Legal", "Genérico").
- Incluir AMBAS columnas para cada fila.
- Obligatorio incluir: FUENTE, ALIMENTACION, PODER, CARGADOR, POWER, ADAPTADOR, ESTABILIZADOR, ONLINE, CHINA, TAIWAN, USA, S.A., SAC, LTD, UNIDAD, PIEZA, LOTE, CAJA, KIT, MODULO, CABLE, DISPOSITIVO, GENERICO, SIN MARCA, NO APLICA.
- Agregar más términos técnicos específicos de "{producto}" que aparezcan en las descripciones.

### Hoja "1c_Marca_Por_Defecto"
Columnas EXACTAS: "Es_Producto_Principal" | "Marca_Default"
- Fila 1: Es_Producto_Principal=True, Marca_Default="Marca Generica"
- Fila 2: Es_Producto_Principal=False, Marca_Default="Marca Componentes"

### Hoja "2_Caracteristicas" — MÍNIMO 150 reglas (el template tiene 284)
Columnas EXACTAS: "Variable" | "Valor_Resultado" | "Palabra_Clave" | "Prioridad" | "Comentario"
- "Variable": nombre de la característica categórica.
- "Valor_Resultado": valor normalizado de la característica.
- "Palabra_Clave": palabra o frase que aparece en las descripciones (o regex simple con |).
- "Prioridad": 1=alta, 2=media, 3=baja.
- "Comentario": explicación breve de por qué esta regla existe.

**VARIABLES OBLIGATORIAS (debes crear múltiples filas para CADA una):**
- "Formato_Montaje": Rack, Torre, Rack/Torre (RT), 1U, 2U, 3U, Modular, Piso, Compacto, Portatil, etc.
- "Tipo_Tecnologia": On-Line / Doble Conversion, Off-Line / Standby, Line-Interactive, Online, Offline, Standby, etc.
- "Capacidad_Bateria": Alta, Media, Baja, Extended, Interna, Externa, etc.
- "Salida_Fases": Monofasico, Bifasico, Trifasico, 1F, 2F, 3F, etc.
- "Voltaje_Entrada": 110V, 220V, 110/220V, Bivoltaje, Universal, 12VDC, 24VDC, 48VDC, etc.
- "Uso_Aplicacion": Domiciliario, Oficina, Industrial, Servidor, Data Center, Hospitalario, Telecomunicaciones, etc.
- "Gama": Basica, Media, Alta, Premium, Enterprise, etc.
- "Tipo_Producto": UPS Sistema Completo, UPS Solo Bateria, Modulo de Bateria, Kit de Bateria, PDU, Regulador, etc.

**Para CADA variable, crear MÍNIMO 10-15 filas** con TODAS las variaciones de escritura:
- Ejemplo para Tipo_Tecnologia: ONLINE, ON-LINE, ON LINE, ONLINE DOBLE CONVERSION, DOBLE CONVERSION, DOUBLE CONVERSION, OFFLINE, OFF-LINE, OFF LINE, STANDBY, STAND-BY, STAND BY, LINE INTERACTIVE, LINE-INTERACTIVE, LINEAR, AUXILIAR, etc.
- Ejemplo para Formato_Montaje: RACK, RACK/TOWER, RACK TOWER, RT, TORRE, TOWER, 1U, 2U, 3U, 4U, 6U, 8U, MODULAR, PISO, FLOOR, COMPACT, PORTABLE, etc.

### Hoja "3_Tecnico_Potencia_NOEDIT" — MÍNIMO 10 patrones (el template tiene 16)
Columnas EXACTAS: "Variable" | "Pattern_Regex" | "Multiplicador_kVA" | "Orden_Prioridad" | "Comentario"
- "Variable": nombre de la variable técnica (Amperaje, Capacidad_Bateria, Frecuencia, Potencia_Watts, Potencia_kVA, Voltaje_Entrada, etc.)
- "Pattern_Regex": regex con grupo de captura que extrae el número.
- "Multiplicador_kVA": factor de conversión (1.0 para unidades base, 1000 para K→base, 0.001 para base→K).
- "Orden_Prioridad": 1=primero en revisarse.
- "Comentario": explicación del patrón.

**Patrones OBLIGATORIOS (adapta según lo que veas en las descripciones):**
- Amperaje: \b(\d+(?:[\.,]\d+)?)\s*AMPS?\b → Multiplicador=1.0
- Amperaje genérico: \b(\d+(?:[\.,]\d+)?)\s+A\b (con word boundary) → Multiplicador=1.0
- Capacidad_Bateria: \b(\d+(?:[\.,]\d+)?)\s*AH\b → Multiplicador=1.0
- Frecuencia: \b(\d+(?:[\.,]\d+)?)\s*HZ\b → Multiplicador=1.0
- Potencia_Watts: \b(\d+(?:[\.,]\d+)?)\s*WATTS?\b → Multiplicador=1.0
- Potencia_Watts (abreviado): \b(\d+(?:[\.,]\d+)?)\s+W\b (con word boundary) → Multiplicador=1.0
- Potencia_kVA directo: \b(\d+(?:[\.,]\d+)?)\s*KVA\b → Multiplicador=1.0
- Potencia_kVA desde VA: \b(\d+(?:[\.,]\d+)?)\s*VA\b → Multiplicador=0.001
- Potencia_kVA desde KW: \b(\d+(?:[\.,]\d+)?)\s*KW\b → Multiplicador=1.0
- Voltaje DC: \b(\d+(?:[\.,]\d+)?)\s*VDC\b → Multiplicador=1.0
- Voltaje AC: \b(\d+(?:[\.,]\d+)?)\s*VAC\b → Multiplicador=1.0
- Voltaje genérico: \b(1[0-9]|[2-9]\d|[1-5]\d{{2}}|600)\s*V\b → Multiplicador=1.0

### Hoja "4_Tecnico_RegexMarca_NOEDIT" — MÍNIMO 5 patrones
Columnas EXACTAS: "Orden_Prioridad" | "Etiqueta_Patron" | "Pattern_Regex"
- "Orden_Prioridad": 1=primero.
- "Etiqueta_Patron": nombre descriptivo del patrón.
- "Pattern_Regex": regex que captura la marca.

**Patrones a incluir:**
- \bMARCA[:\s]+([A-Z0-9\.\-]+) → Detección explícita de "MARCA: X"
- \bBRAND[:\s]+([A-Z0-9\.\-]+) → Detección explícita de "BRAND: X"
- \bM/([A-Z0-9\.\-]+) → Formato abreviado "M/X"
- Patrones específicos para marcas ambiguas que identifiques en las descripciones.

### Hoja "5_Condicionales" — MÍNIMO 8 reglas
Columnas EXACTAS: "Regla_ID" | "Variable_Resultado" | "Valor_Resultado" | "Prioridad" | "Variable_Condicion" | "Operador" | "Valor_1" | "Valor_2" | "Comentario"
- "Regla_ID": número secuencial (1, 2, 3...).
- "Variable_Resultado": variable que se asigna si se cumple la condición.
- "Valor_Resultado": valor a asignar.
- "Prioridad": 1=primera en evaluarse.
- "Variable_Condicion": variable que se evalúa.
- "Operador": ==, >, <, >=, <=, BETWEEN, CONTAINS, NOT_CONTAINS.
- "Valor_1" y "Valor_2": valores de comparación (Valor_2 solo para BETWEEN).
- "Comentario": explicación de la regla.

**Ejemplos de reglas OBLIGATORIAS:**
- SI Potencia_kVA BETWEEN 0 Y 3 → Salida_Fases="Monofásico" (UPS pequeños son monofásicos)
- SI Potencia_kVA BETWEEN 3 Y 10 → Salida_Fases="Bifásico" (rango intermedio)
- SI Potencia_kVA > 10 → Salida_Fases="Trifásico" (UPS grandes)
- SI Tipo_Tecnologia=="On-Line Doble Conversion" Y Salida_Fases=="Trifásico" → confirmar
- SI Voltaje_Entrada CONTAINS "110/220" → Voltaje_Entrada="Bivoltaje"
- SI descripcion CONTAINS "INDUSTRIAL" → Uso_Aplicacion="Industrial"
- SI descripcion CONTAINS "DATA CENTER" → Uso_Aplicacion="Data Center"
- SI Capacidad_Bateria=="Extended" → Gama="Alta"

## ANÁLISIS OBLIGATORIO DE LAS DESCRIPCIONES

Antes de generar el JSON, haz un conteo mental:
1. Lista TODAS las marcas que encuentras (deben ser 50+).
2. Lista TODAS las variantes de escritura para cada característica (deben ser 150+ filas en total).
3. Lista TODOS los patrones numéricos (deben ser 10+).
4. Lista TODAS las palabras que NO son marca (deben ser 30+).
5. Crea reglas condicionales que resuelvan ambigüedades (deben ser 8+).

Si no llegas a los mínimos, REVISA de nuevo las descripciones. Siempre hay más de lo que parece a primera vista.

## REGLAS CRÍTICAS
1. Las descripciones de importación son textos de FACTURA: marcas, modelos, especificaciones, usos mezclados.
2. Sé EXHAUSTIVO: mejor sobregenerar que subgenerar. El PM puede borrar, pero no sabe qué falta.
3. Los nombres de columnas DEBEN ser EXACTAMENTE los del template (ver arriba).
4. NUNCA inventar datos que no estén en las descripciones. Si un patrón no existe, no lo inventes.
5. Los patrones regex deben ser FUNCIONALES (probados mentalmente contra las descripciones).
6. Devolver SOLO el JSON válido, sin texto adicional."""

    user_prompt = f"""## CONTEXTO DEL PRODUCTO
**Producto a clasificar:** {producto}
{contexto_extra}
## MUESTRA DE DESCRIPCIONES REALES ({len(descripciones)} filas únicas extraídas de Veritrade)

{texto_descripciones}

---

## INSTRUCCIONES DE GENERACIÓN — CANTIDADES MÍNIMAS OBLIGATORIAS

El maestro de referencia para UPS tiene 72 marcas y 284 características. TÚ debes generar algo similar para "{producto}".

### REGLAS DE CANTIDAD (NO NEGOCIABLES):
1. "1_Marcas": MÍNIMO 50 filas. Si no tienes 50 marcas, buscas mal. Revisa abbreviaturas, siglas, variantes ortográficas.
2. "2_Caracteristicas": MÍNIMO 150 filas. Para CADA Variable (Formato_Montaje, Tipo_Tecnologia, etc.) crea 10-20 filas con TODAS las variaciones de escritura.
3. "3_Tecnico_Potencia_NOEDIT": MÍNIMO 10 filas. Cada tipo de unidad (VA, KVA, KW, W, AH, HZ, V, VAC, VDC, AMPS) necesita su propio patrón.
4. "5_Condicionales": MÍNIMO 8 filas. Reglas que resuelvan ambigüedades reales.
5. "1b_Palabras_Ignorar": MÍNIMO 30 filas. Términos técnicos, orígenes, entidades legales, genéricos.

### EJEMPLO DE RIQUEZA ESPERADA para "2_Caracteristicas":
Si la Variable es "Tipo_Tecnologia", debes crear filas para:
ONLINE, ON-LINE, ON LINE, ONLINE DOBLE CONVERSION, DOBLE CONVERSION, DOUBLE CONVERSION,
OFFLINE, OFF-LINE, OFF LINE, STANDBY, STAND-BY, STAND BY,
LINE INTERACTIVE, LINE-INTERACTIVE, LINEAR, AUXILIAR, etc.
CADA variación es una fila separada con su Palabra_Clave y Prioridad.

### EJEMPLO DE RIQUEZA ESPERADA para "1_Marcas":
Incluir TODAS las marcas que aparezcan en las descripciones, incluyendo variantes:
APC, A.P.C, AMERICAN POWER CONVERSION → normalizar a "APC"
SCHNEIDER, SCHNEIDER ELECTRIC, SE → normalizar a "SCHNEIDER ELECTRIC"
EATON, EATON POWER QUALITY → normalizar a "EATON"
Cada variación de escritura de la MISMA marca también va como fila aparte.

## DEVUELVE ÚNICAMENTE EL JSON, SIN EXPLICACIONES."""

    return system_prompt, user_prompt


def _schema_json_maestro() -> dict:
    """Schema JSON que Gemini debe devolver (una key por hoja del maestro).
    Los nombres de columnas coinciden EXACTAMENTE con Maestro_Plantilla.xlsx."""
    return {
        "type": "OBJECT",
        "properties": {
            "0b_Config_Linea": {
                "type": "ARRAY",
                "description": "Filas de configuración: cada una con Parametro y Valor.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Parametro": {"type": "STRING"},
                        "Valor": {"type": "STRING"},
                    },
                },
            },
            "1_Marcas": {
                "type": "ARRAY",
                "description": "Diccionario de marcas: patron de busqueda → nombre estándar + prioridad. MÍNIMO 50 filas.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Patron_Busqueda": {"type": "STRING"},
                        "Marca_Estandar": {"type": "STRING"},
                        "Prioridad": {"type": "INTEGER"},
                    },
                },
            },
            "1b_Palabras_Ignorar": {
                "type": "ARRAY",
                "description": "Stopwords: palabras que nunca deben ser marca. MÍNIMO 30 filas.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Palabra_Ignorar": {"type": "STRING"},
                        "Categoria": {"type": "STRING"},
                    },
                },
            },
            "1c_Marca_Por_Defecto": {
                "type": "ARRAY",
                "description": "Marca por defecto cuando nada coincide.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Es_Producto_Principal": {"type": "BOOLEAN"},
                        "Marca_Default": {"type": "STRING"},
                    },
                },
            },
            "2_Caracteristicas": {
                "type": "ARRAY",
                "description": "Reglas de características categóricas. MÍNIMO 150 filas. Cada Variable debe tener 10-20 variaciones de escritura.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Variable": {"type": "STRING"},
                        "Valor_Resultado": {"type": "STRING"},
                        "Palabra_Clave": {"type": "STRING"},
                        "Prioridad": {"type": "INTEGER"},
                        "Comentario": {"type": "STRING"},
                    },
                },
            },
            "3_Tecnico_Potencia_NOEDIT": {
                "type": "ARRAY",
                "description": "Patrones de extracción numérica técnica. MÍNIMO 10 filas.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Variable": {"type": "STRING"},
                        "Pattern_Regex": {"type": "STRING"},
                        "Multiplicador_kVA": {"type": "NUMBER"},
                        "Orden_Prioridad": {"type": "INTEGER"},
                        "Comentario": {"type": "STRING"},
                    },
                },
            },
            "4_Tecnico_RegexMarca_NOEDIT": {
                "type": "ARRAY",
                "description": "Regex avanzados de marca (casos difíciles). MÍNIMO 5 filas.",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "Orden_Prioridad": {"type": "INTEGER"},
                        "Etiqueta_Patron": {"type": "STRING"},
                        "Pattern_Regex": {"type": "STRING"},
                    },
                },
            },
            "5_Condicionales": {
                "type": "ARRAY",
                "description": "Reglas condicionales SI-ENTONCES. MÍNIMO 8 filas.",
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
                        "Comentario": {"type": "STRING"},
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
    Los nombres de columnas coinciden con Maestro_Plantilla.xlsx.
    """
    hojas = {}

    # 0b_Config_Linea
    cfg_rows = datos.get("0b_Config_Linea", [])
    if not cfg_rows:
        cfg_rows = [
            {"Parametro": "LINEA_PRODUCTO", "Valor": producto},
            {"Parametro": "VARIABLE_PRODUCTO_PRINCIPAL", "Valor": "Tipo_Producto_Detallado"},
            {"Parametro": "VALOR_PRODUCTO_PRINCIPAL", "Valor": f"{producto} Sistema Completo"},
        ]
    hojas["0b_Config_Linea"] = pd.DataFrame(cfg_rows)

    # 1_Marcas
    marcas_rows = datos.get("1_Marcas", [])
    if marcas_rows:
        hojas["1_Marcas"] = pd.DataFrame(marcas_rows)
    else:
        hojas["1_Marcas"] = pd.DataFrame(columns=[
            "Patron_Busqueda", "Marca_Estandar", "Prioridad"
        ])

    # 1b_Palabras_Ignorar
    sw_rows = datos.get("1b_Palabras_Ignorar", [])
    if sw_rows:
        hojas["1b_Palabras_Ignorar"] = pd.DataFrame(sw_rows)
    else:
        hojas["1b_Palabras_Ignorar"] = pd.DataFrame(columns=["Palabra_Ignorar", "Categoria"])

    # 1c_Marca_Por_Defecto
    default_rows = datos.get("1c_Marca_Por_Defecto", [])
    if not default_rows:
        default_rows = [
            {"Es_Producto_Principal": True, "Marca_Default": "Marca Generica"},
            {"Es_Producto_Principal": False, "Marca_Default": "Marca Componentes"},
        ]
    hojas["1c_Marca_Por_Defecto"] = pd.DataFrame(default_rows)

    # 2_Caracteristicas
    carac_rows = datos.get("2_Caracteristicas", [])
    if carac_rows:
        hojas["2_Caracteristicas"] = pd.DataFrame(carac_rows)
    else:
        hojas["2_Caracteristicas"] = pd.DataFrame(columns=[
            "Variable", "Valor_Resultado", "Palabra_Clave", "Prioridad", "Comentario"
        ])

    # 3_Tecnico_Potencia_NOEDIT
    pot_rows = datos.get("3_Tecnico_Potencia_NOEDIT", [])
    if pot_rows:
        hojas["3_Tecnico_Potencia_NOEDIT"] = pd.DataFrame(pot_rows)
    else:
        hojas["3_Tecnico_Potencia_NOEDIT"] = pd.DataFrame(columns=[
            "Variable", "Pattern_Regex", "Multiplicador_kVA", "Orden_Prioridad", "Comentario"
        ])

    # 4_Tecnico_RegexMarca_NOEDIT
    regex_rows = datos.get("4_Tecnico_RegexMarca_NOEDIT", [])
    if regex_rows:
        hojas["4_Tecnico_RegexMarca_NOEDIT"] = pd.DataFrame(regex_rows)
    else:
        hojas["4_Tecnico_RegexMarca_NOEDIT"] = pd.DataFrame(columns=[
            "Orden_Prioridad", "Etiqueta_Patron", "Pattern_Regex"
        ])

    # 5_Condicionales
    cond_rows = datos.get("5_Condicionales", [])
    if cond_rows:
        hojas["5_Condicionales"] = pd.DataFrame(cond_rows)
    else:
        hojas["5_Condicionales"] = pd.DataFrame(columns=[
            "Regla_ID", "Variable_Resultado", "Valor_Resultado", "Prioridad",
            "Variable_Condicion", "Operador", "Valor_1", "Valor_2", "Comentario"
        ])

    return hojas


# =====================================================================
# PASO 3: ESCRITURA DEL EXCEL
# =====================================================================

def _normalizar_hoja(df: pd.DataFrame, columnas_esperadas: list[str]) -> pd.DataFrame:
    """Filtra un DataFrame a solo las columnas esperadas, en el orden correcto.
    Columnas extra de la IA se descartan; columnas faltantes se crean vacías."""
    for col in columnas_esperadas:
        if col not in df.columns:
            df[col] = None
    return df[columnas_esperadas]


# Columnas esperadas por hoja (referencia centralizada — coinciden con Maestro_Plantilla.xlsx)
COLUMNAS_HOJAS = {
    "0b_Config_Linea": ["Parametro", "Valor"],
    "1_Marcas": ["Patron_Busqueda", "Marca_Estandar", "Prioridad"],
    "1b_Palabras_Ignorar": ["Palabra_Ignorar", "Categoria"],
    "1c_Marca_Por_Defecto": ["Es_Producto_Principal", "Marca_Default"],
    "2_Caracteristicas": ["Variable", "Valor_Resultado", "Palabra_Clave", "Prioridad", "Comentario"],
    "3_Tecnico_Potencia_NOEDIT": ["Variable", "Pattern_Regex", "Multiplicador_kVA", "Orden_Prioridad", "Comentario"],
    "4_Tecnico_RegexMarca_NOEDIT": ["Orden_Prioridad", "Etiqueta_Patron", "Pattern_Regex"],
    "5_Condicionales": ["Regla_ID", "Variable_Resultado", "Valor_Resultado", "Prioridad", "Variable_Condicion", "Operador", "Valor_1", "Valor_2", "Comentario"],
}


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
        # Escribir hojas en orden, normalizando columnas
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
                cols_ok = COLUMNAS_HOJAS.get(nombre_hoja)
                if cols_ok:
                    df = _normalizar_hoja(df, cols_ok)
                df.to_excel(writer, index=False, sheet_name=nombre_hoja)

        # Hoja de Instrucciones (al final, como referencia)
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
        df_instrucciones = pd.DataFrame({
            "Parámetro": [
                "Producto",
                "Fecha de generación",
                "Generador",
                "",
                "INSTRUCCIONES DE USO",
                "",
                "Este maestro fue generado automáticamente por IA a partir de una muestra de descripciones reales.",
                "Revisa CADA hoja antes de usarlo en producción:",
                "  1. 1_Marcas: Verifica que las marcas sean correctas y completas.",
                "  2. 2_Caracteristicas: Ajusta las palabras clave y valores según tu conocimiento del producto.",
                "  3. 3_Tecnico_Potencia: Valida los patrones numéricos y multiplicadores.",
                "  4. 5_Condicionales: Revisa las reglas SI-ENTONCES.",
                "",
                "NOTAS TÉCNICAS",
                "",
                "Las hojas marcadas _NOEDIT son generadas por IA y pueden requerir ajuste fino.",
                "La hoja 1b_Palabras_Ignorar (Stopwords) es genérica y puede compartirse entre líneas.",
                "",
                "ESTRUCTURA DE HOJAS",
                "  0b_Config_Linea — Parámetros de configuración de la línea",
                "  1_Marcas — Diccionario de marcas (patrón → nombre estándar)",
                "  1b_Palabras_Ignorar — Stopwords (palabras que nunca son marca)",
                "  1c_Marca_Por_Defeito — Marca por defecto según si es producto principal",
                "  2_Caracteristicas — Reglas de características categóricas",
                "  3_Tecnico_Potencia_NOEDIT — Extracción de valores numéricos técnicos",
                "  4_Tecnico_RegexMarca_NOEDIT — Regex avanzados de marca",
                "  5_Condicionales — Reglas condicionales SI-ENTONCES",
            ],
            "Valor": [
                producto,
                ahora,
                "App Clasificador Veritrade — Generador Automático v1",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ],
        })

        df_instrucciones.to_excel(writer, index=False, sheet_name="Instrucciones")

        # Aplicar estilos a cada hoja
        for nombre_hoja in writer.sheets:
            if nombre_hoja == "Instrucciones":
                continue
            df_hoja = hojas.get(nombre_hoja)
            if df_hoja is not None and len(df_hoja) > 0:
                try:
                    cols_ok = COLUMNAS_HOJAS.get(nombre_hoja)
                    df_estilo = _normalizar_hoja(df_hoja.copy(), cols_ok) if cols_ok else df_hoja
                    aplicar_estilo_hoja_excel(writer.sheets[nombre_hoja], df_estilo)
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
