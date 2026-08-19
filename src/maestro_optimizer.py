# -*- coding: utf-8 -*-
"""
Aprendizaje incremental del Maestro a partir de los rescates de IA
---------------------------------------------------------------------
Cada vez que Gemini resuelve una marca o una característica que las
reglas NO pudieron resolver, ese conocimiento se pierde en la próxima
corrida (se vuelve a gastar cuota en la misma descripción si aparece
en otro archivo). Este módulo cierra ese ciclo:

    reglas (rápido, gratis) -> IA (rescate) -> "Maestro_Optimizado.xlsx"
    -> próxima corrida usa el maestro optimizado -> menos rescates de IA

Dos niveles de confianza, tratados distinto a propósito:

1. MARCAS: si la IA declara una marca, se confía en el valor completo
   (viene de un LLM leyendo la descripción completa, criterio conservador
   por prompt) y se agrega como nueva fila en la hoja de Marcas del
   maestro, con baja prioridad (se evalúa DESPUÉS de todas las reglas
   humanas ya existentes, nunca las pisa).

2. CARACTERÍSTICAS CATEGÓRICAS: aquí NO basta con confiar en el valor;
   una regla categórica necesita una PALABRA/FRASE CLAVE verificable
   en el texto para poder reutilizarse después con regex. Por eso:
     - Si el valor que dio la IA (o una palabra significativa de él)
       aparece LITERALMENTE en la descripción -> se propone esa palabra
       como nuevo patrón, con baja prioridad.
     - Si no aparece literalmente (la IA infirió el valor sin una
       palabra ancla clara) -> NO se auto-agrega ninguna regla; el caso
       queda en la hoja de auditoría "Log_Aprendizaje_IA" para revisión
       humana. Auto-generar un regex "adivinado" sería peligroso: podría
       sobre-generalizar y mal-clasificar descripciones futuras.

Todo lo aprendido (auto-agregado o no) queda trazado en la hoja
"Log_Aprendizaje_IA" del archivo de salida, con la descripción de
origen, para poder auditar/depurar el maestro con el tiempo.
"""

from __future__ import annotations

import re
import os
import tempfile
from datetime import datetime
from typing import Optional

import pandas as pd

from src.texto_utils import limpiar_texto
from src.maestro.reglas import es_candidato_marca_valido
from src.excel_estilos import aplicar_estilo_hoja_excel, restaurar_hoja_instrucciones


# ----------------------------------------------------------------------
# Detección de columnas (mismo criterio que loader.py, reutilizado aquí
# para poder escribir en las columnas reales del archivo, sean cuales sean)
# ----------------------------------------------------------------------
def _col(df: pd.DataFrame, *keywords: str) -> Optional[str]:
    cols = [c for c in df.columns if any(kw.upper() in str(c).upper() for kw in keywords)]
    return cols[0] if cols else None


def _palabra_ancla_en_texto(valor: str, desc_clean: str) -> Optional[str]:
    """
    Busca una "palabra ancla" verificable para un valor categórico dentro
    de la descripción original. Devuelve el texto exacto tal como aparece
    (para poder usarlo como patrón regex confiable), o None si no hay
    ninguna coincidencia literal segura.

    Esto es un FALLBACK heurístico para cuando la IA no entregó el campo
    "<variable>__evidencia" (ej. respuestas de una corrida anterior, o un
    modelo que ignoró la instrucción). El camino principal y preferido es
    usar directamente la evidencia literal que la IA devuelve — ver
    `_ancla_desde_evidencia_ia`.
    """
    if not valor or not desc_clean:
        return None

    valor_clean = limpiar_texto(str(valor))

    # 1. Coincidencia exacta del valor completo dentro del texto
    if valor_clean and valor_clean in desc_clean:
        return valor_clean

    # 2. Si el valor tiene varias palabras, probamos con la palabra más
    #    larga (normalmente la más específica / menos ambigua)
    palabras = [p for p in valor_clean.split() if len(p) >= 4]
    palabras.sort(key=len, reverse=True)
    for palabra in palabras:
        if re.search(fr'(?:^|(?<=\W)){re.escape(palabra)}(?:$|(?=\W))', desc_clean):
            return palabra

    return None


def _ancla_desde_evidencia_ia(evidencia: str, desc_clean: str) -> Optional[str]:
    """
    Valida y normaliza la evidencia literal que la propia IA devolvió
    (campo "<variable>__evidencia"). Solo se acepta si, tras limpiar
    tildes/mayúsculas igual que el resto del pipeline, aparece REALMENTE
    como subcadena de la descripción — nunca se confía a ciegas en lo
    que dice el modelo, se verifica contra el texto fuente.
    """
    if not evidencia:
        return None
    ev_clean = limpiar_texto(str(evidencia)).strip()
    if not ev_clean or len(ev_clean) < 1:
        return None
    if ev_clean in desc_clean:
        return ev_clean
    return None


# ----------------------------------------------------------------------
# Paso 1: construir las propuestas de aprendizaje a partir de los
# resultados exitosos de la fase de rescate IA
# ----------------------------------------------------------------------
def construir_propuestas_aprendizaje(
    resultados_ia: dict,          # {desc_clean: ResultadoRescateIA}
    maestro,
    variables_cat: list[str],
) -> dict:
    """
    Analiza los resultados exitosos de IA y separa lo aprendido en:
      - nuevas_marcas: [{"patron", "estandar", "desc_origen"}]
      - nuevas_caracteristicas: [{"variable", "valor", "palabra_clave", "desc_origen"}]
      - revisar_manual: [{"tipo", "variable", "valor", "desc_origen", "motivo"}]
    No escribe nada a disco todavía; eso lo hace `guardar_maestro_optimizado`.
    """
    nuevas_marcas = []
    nuevas_marcas_vistas = set()

    nuevas_caracteristicas = []
    nuevas_caract_vistas = set()

    revisar_manual = []

    for desc_clean, resultado in resultados_ia.items():
        if not resultado or not resultado.exito:
            continue
        valores = resultado.valores or {}

        # --- Marca ---
        marca_ia = valores.get("marca")
        if marca_ia:
            patron = limpiar_texto(str(marca_ia))
            ya_en_maestro = any(
                re.search(fr'(?:^|(?<=\W)){re.escape(p)}(?:$|(?=\W))', desc_clean)
                for p, _ in maestro.lista_marcas
            )
            if patron and not ya_en_maestro and patron not in nuevas_marcas_vistas:
                if es_candidato_marca_valido(marca_ia, maestro.stopwords):
                    nuevas_marcas_vistas.add(patron)
                    nuevas_marcas.append({
                        "patron": patron,
                        "estandar": str(marca_ia).strip(),
                        "desc_origen": desc_clean,
                    })
                else:
                    revisar_manual.append({
                        "tipo": "Marca",
                        "variable": "marca",
                        "valor": marca_ia,
                        "desc_origen": desc_clean,
                        "motivo": "La IA propuso una marca que no pasa la validación de formato "
                                  "(números, símbolos, muy larga, palabra descartada, etc.)",
                    })

        # --- Características categóricas ---
        for var in variables_cat:
            valor_ia = valores.get(var)
            if not valor_ia:
                continue

            clave_dedup = (var, str(valor_ia).strip().upper())

            # Camino principal: evidencia literal que la propia IA entregó
            # (campo "<var>__evidencia"), verificada contra el texto fuente.
            evidencia_ia = valores.get(f"{var}__evidencia")
            palabra_ancla = _ancla_desde_evidencia_ia(evidencia_ia, desc_clean)

            # Fallback: heurística vieja (buscar el propio valor normalizado
            # en el texto), por si la IA no devolvió el campo de evidencia
            # (ej. resultados cacheados de una corrida anterior a este fix).
            if not palabra_ancla:
                palabra_ancla = _palabra_ancla_en_texto(valor_ia, desc_clean)

            if palabra_ancla:
                if clave_dedup not in nuevas_caract_vistas:
                    nuevas_caract_vistas.add(clave_dedup)
                    nuevas_caracteristicas.append({
                        "variable": var,
                        "valor": str(valor_ia).strip(),
                        "palabra_clave": palabra_ancla,
                        "desc_origen": desc_clean,
                    })
            else:
                revisar_manual.append({
                    "tipo": "Caracteristica",
                    "variable": var,
                    "valor": valor_ia,
                    "desc_origen": desc_clean,
                    "motivo": "La IA infirió el valor pero no se encontró una palabra clave "
                              "literal en el texto para generar una regla regex confiable "
                              "(ni en el campo de evidencia ni por heurística). Requiere "
                              "revisión humana antes de agregarse al maestro.",
                })

    return {
        "nuevas_marcas": nuevas_marcas,
        "nuevas_caracteristicas": nuevas_caracteristicas,
        "revisar_manual": revisar_manual,
    }


# ----------------------------------------------------------------------
# Paso 2: volcar las propuestas sobre una copia del maestro original
# ----------------------------------------------------------------------
def guardar_maestro_optimizado(ruta_maestro_original, propuestas: dict, ruta_salida) -> dict:
    """
    Lee TODAS las hojas del maestro original, agrega (append, nunca
    sobreescribe) las filas aprendidas en las hojas de Marcas y
    Características con la prioridad más baja posible (se evalúan de
    últimas, así que jamás compiten con una regla humana existente), y
    escribe el resultado en `ruta_salida` junto con una hoja de
    auditoría "Log_Aprendizaje_IA".

    Devuelve un resumen {marcas_agregadas, caracteristicas_agregadas,
    pendientes_revision} para mostrar en la UI de Streamlit.
    """
    hojas = pd.read_excel(ruta_maestro_original, sheet_name=None)

    resumen = {"marcas_agregadas": 0, "caracteristicas_agregadas": 0, "pendientes_revision": 0}
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    filas_log = []

    # --- Hoja de Marcas ---
    nombre_hoja_marcas = next((n for n in hojas if "MARCA" in n.upper() and "DEFECTO" not in n.upper()), None)
    if nombre_hoja_marcas and propuestas.get("nuevas_marcas"):
        df_marcas = hojas[nombre_hoja_marcas]
        col_pat = _col(df_marcas, "PATRON", "BUSQUEDA")
        col_est = _col(df_marcas, "MARCA", "ESTANDAR")
        col_prio = _col(df_marcas, "PRIORIDAD")

        if col_pat and col_est:
            # Patrones ya existentes en el maestro real (no solo lo visto en
            # esta corrida) -> evita re-agregar si este mismo maestro
            # optimizado se retroalimenta como base la próxima vez.
            patrones_existentes = {
                limpiar_texto(str(p)).strip().upper()
                for p in df_marcas[col_pat].dropna()
            }

            prio_base = (df_marcas[col_prio].max() + 1) if col_prio and df_marcas[col_prio].notna().any() else 999
            filas_nuevas = []
            omitidas = 0
            for item in propuestas["nuevas_marcas"]:
                patron_norm = limpiar_texto(item["patron"]).strip().upper()
                if patron_norm in patrones_existentes:
                    omitidas += 1
                    filas_log.append({
                        "Fecha": ahora, "Tipo": "Marca", "Variable": "marca",
                        "Valor_Aprendido": item["estandar"], "Patron_O_Palabra_Clave": item["patron"],
                        "Descripcion_Origen": item["desc_origen"],
                        "Estado": "Omitido — el patrón ya existe en el maestro",
                    })
                    continue
                patrones_existentes.add(patron_norm)  # también dedupe dentro de esta misma corrida
                fila = {c: None for c in df_marcas.columns}
                fila[col_pat] = item["patron"]
                fila[col_est] = item["estandar"]
                if col_prio:
                    fila[col_prio] = prio_base
                filas_nuevas.append(fila)
                filas_log.append({
                    "Fecha": ahora, "Tipo": "Marca", "Variable": "marca",
                    "Valor_Aprendido": item["estandar"], "Patron_O_Palabra_Clave": item["patron"],
                    "Descripcion_Origen": item["desc_origen"], "Estado": "Agregado automáticamente",
                })
            if filas_nuevas:
                hojas[nombre_hoja_marcas] = pd.concat(
                    [df_marcas, pd.DataFrame(filas_nuevas)], ignore_index=True
                )
            resumen["marcas_agregadas"] = len(filas_nuevas)

    # --- Hoja de Características ---
    nombre_hoja_carac = next((n for n in hojas if "CARACTERISTICA" in n.upper()), None)
    if nombre_hoja_carac and propuestas.get("nuevas_caracteristicas"):
        df_carac = hojas[nombre_hoja_carac]
        col_kw = _col(df_carac, "PALABRA", "CLAVE")
        col_prio = _col(df_carac, "PRIORIDAD", "ORDEN_PRIORIDAD")

        if col_kw and "Variable" in df_carac.columns and "Valor_Resultado" in df_carac.columns:
            # Dedupe real contra el maestro base, en dos niveles:
            #  1) misma (variable, palabra_clave) -> ya existe exactamente esa regla
            #  2) misma (variable, valor) -> ya hay ALGUNA regla que produce ese
            #     valor para esa variable, aunque la palabra clave sea distinta;
            #     agregar otra sería redundante y solo infla el maestro.
            claves_existentes = set()
            valores_existentes = set()
            for _, fila_existente in df_carac.iterrows():
                var_e = str(fila_existente.get("Variable", "")).strip().upper()
                val_e = str(fila_existente.get("Valor_Resultado", "")).strip().upper()
                kw_e = limpiar_texto(str(fila_existente.get(col_kw, ""))).strip().upper()
                if var_e:
                    if kw_e:
                        claves_existentes.add((var_e, kw_e))
                    if val_e:
                        valores_existentes.add((var_e, val_e))

            prio_base = (df_carac[col_prio].max() + 1) if col_prio and df_carac[col_prio].notna().any() else 999
            filas_nuevas = []
            for item in propuestas["nuevas_caracteristicas"]:
                var_norm = item["variable"].strip().upper()
                val_norm = item["valor"].strip().upper()
                kw_norm = limpiar_texto(item["palabra_clave"]).strip().upper()

                if (var_norm, kw_norm) in claves_existentes:
                    filas_log.append({
                        "Fecha": ahora, "Tipo": "Caracteristica", "Variable": item["variable"],
                        "Valor_Aprendido": item["valor"], "Patron_O_Palabra_Clave": item["palabra_clave"],
                        "Descripcion_Origen": item["desc_origen"],
                        "Estado": "Omitido — la palabra clave ya existe en el maestro para esa variable",
                    })
                    continue
                if (var_norm, val_norm) in valores_existentes:
                    filas_log.append({
                        "Fecha": ahora, "Tipo": "Caracteristica", "Variable": item["variable"],
                        "Valor_Aprendido": item["valor"], "Patron_O_Palabra_Clave": item["palabra_clave"],
                        "Descripcion_Origen": item["desc_origen"],
                        "Estado": "Omitido — ya existe una regla que produce ese valor para esa variable",
                    })
                    continue

                claves_existentes.add((var_norm, kw_norm))
                valores_existentes.add((var_norm, val_norm))

                fila = {c: None for c in df_carac.columns}
                fila["Variable"] = item["variable"]
                fila["Valor_Resultado"] = item["valor"]
                fila[col_kw] = item["palabra_clave"]
                if col_prio:
                    fila[col_prio] = prio_base
                filas_nuevas.append(fila)
                filas_log.append({
                    "Fecha": ahora, "Tipo": "Caracteristica", "Variable": item["variable"],
                    "Valor_Aprendido": item["valor"], "Patron_O_Palabra_Clave": item["palabra_clave"],
                    "Descripcion_Origen": item["desc_origen"], "Estado": "Agregado automáticamente",
                })
            if filas_nuevas:
                hojas[nombre_hoja_carac] = pd.concat(
                    [df_carac, pd.DataFrame(filas_nuevas)], ignore_index=True
                )
            resumen["caracteristicas_agregadas"] = len(filas_nuevas)

    # --- Pendientes de revisión manual (no se tocan reglas, solo auditoría) ---
    for item in propuestas.get("revisar_manual", []):
        filas_log.append({
            "Fecha": ahora, "Tipo": item["tipo"], "Variable": item["variable"],
            "Valor_Aprendido": item["valor"], "Patron_O_Palabra_Clave": None,
            "Descripcion_Origen": item["desc_origen"], "Estado": f"PENDIENTE REVISIÓN — {item['motivo']}",
        })
    resumen["pendientes_revision"] = len(propuestas.get("revisar_manual", []))

    if filas_log:
        df_log_nuevo = pd.DataFrame(filas_log)
        if "Log_Aprendizaje_IA" in hojas:
            hojas["Log_Aprendizaje_IA"] = pd.concat([hojas["Log_Aprendizaje_IA"], df_log_nuevo], ignore_index=True)
        else:
            hojas["Log_Aprendizaje_IA"] = df_log_nuevo

    ruta_salida = os.fspath(ruta_salida)
    carpeta_salida = os.path.dirname(os.path.abspath(ruta_salida))
    fd_temporal, ruta_temporal = tempfile.mkstemp(suffix=".xlsx", dir=carpeta_salida)
    os.close(fd_temporal)
    try:
        with pd.ExcelWriter(ruta_temporal, engine="openpyxl") as writer:
            for nombre_hoja, df in hojas.items():
                hoja_final = nombre_hoja[:31]
                df.to_excel(writer, sheet_name=hoja_final, index=False)
                ws = writer.sheets[hoja_final]
                es_log = hoja_final == "Log_Aprendizaje_IA"
                aplicar_estilo_hoja_excel(ws, df, es_log=es_log)

        # La primera hoja es documentación, no datos tabulares. Se restaura
        # desde el original después de cerrar el writer para conservar su diseño.
        restaurar_hoja_instrucciones(ruta_temporal, ruta_maestro_original)
        os.replace(ruta_temporal, ruta_salida)
    finally:
        if os.path.exists(ruta_temporal):
            os.remove(ruta_temporal)

    return resumen
