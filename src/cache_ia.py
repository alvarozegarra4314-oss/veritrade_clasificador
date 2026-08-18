# -*- coding: utf-8 -*-
"""
Caché persistente (SQLite) para resultados de rescate por IA.
----------------------------------------------------------------
Objetivo: que una descripción ya rescatada en una corrida anterior
(incluso en otro día, otro archivo, otra sesión de Streamlit) jamás
vuelva a consumir cuota de la API. Se guarda por hash de
`desc_clean` + nombre de modelo, ya que dos modelos podrían dar
esquemas distintos.

Uso típico:
    cache = CacheIA(ruta_db="data/cache_ia.sqlite")
    resultado = cache.obtener(desc_clean, modelo)
    if resultado is None:
        resultado = llamar_api(...)
        cache.guardar(desc_clean, modelo, resultado)
"""

from __future__ import annotations

import json
import sqlite3
import hashlib
import threading
from pathlib import Path
from typing import Optional


class CacheIA:
    def __init__(self, ruta_db: str | Path = "data/cache_ia.sqlite"):
        self.ruta_db = Path(ruta_db)
        self.ruta_db.parent.mkdir(parents=True, exist_ok=True)
        # check_same_thread=False porque Streamlit puede reusar el objeto
        # entre reruns/hilos; protegemos con un Lock propio.
        self._conn = sqlite3.connect(str(self.ruta_db), check_same_thread=False)
        self._lock = threading.Lock()
        self._crear_tabla()

    def _crear_tabla(self):
        with self._lock:
            self._conn.execute(
                """
                CREATE TABLE IF NOT EXISTS cache_ia (
                    clave TEXT PRIMARY KEY,
                    desc_clean TEXT NOT NULL,
                    modelo TEXT NOT NULL,
                    valores_json TEXT NOT NULL,
                    creado_ts REAL DEFAULT (strftime('%s','now'))
                )
                """
            )
            self._conn.commit()

    @staticmethod
    def _clave(desc_clean: str, modelo: str) -> str:
        # `modelo` puede incluir un sufijo de versión de prompt (ej.
        # "gemini-3.1-flash-lite::promptv2"), no solo el nombre del modelo
        # de Gemini. Esto es intencional: así, cuando se cambia el
        # PROMPT_SISTEMA (p.ej. para priorizar características técnicas
        # sobre la marca), las descripciones YA cacheadas con el prompt
        # viejo se tratan como "no vistas" y se vuelven a mandar a la IA,
        # en vez de servir para siempre resultados obtenidos con lógica
        # de extracción desactualizada.
        base = f"{modelo}::{desc_clean}"
        return hashlib.sha256(base.encode("utf-8")).hexdigest()

    def obtener(self, desc_clean: str, modelo: str) -> Optional[dict]:
        clave = self._clave(desc_clean, modelo)
        with self._lock:
            cur = self._conn.execute(
                "SELECT valores_json FROM cache_ia WHERE clave = ?", (clave,)
            )
            fila = cur.fetchone()
        if fila is None:
            return None
        try:
            return json.loads(fila[0])
        except (json.JSONDecodeError, TypeError):
            return None

    def guardar(self, desc_clean: str, modelo: str, valores: dict):
        clave = self._clave(desc_clean, modelo)
        with self._lock:
            self._conn.execute(
                """
                INSERT INTO cache_ia (clave, desc_clean, modelo, valores_json)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(clave) DO UPDATE SET
                    valores_json = excluded.valores_json,
                    creado_ts = strftime('%s','now')
                """,
                (clave, desc_clean, modelo, json.dumps(valores, ensure_ascii=False)),
            )
            self._conn.commit()

    def obtener_muchos(self, descripciones: list[str], modelo: str) -> dict[str, dict]:
        """Devuelve {desc_clean: valores} solo para las que YA están en caché."""
        if not descripciones:
            return {}
        claves = {self._clave(d, modelo): d for d in descripciones}
        placeholders = ",".join("?" * len(claves))
        with self._lock:
            cur = self._conn.execute(
                f"SELECT clave, valores_json FROM cache_ia WHERE clave IN ({placeholders})",
                list(claves.keys()),
            )
            filas = cur.fetchall()
        resultado = {}
        for clave, valores_json in filas:
            desc = claves.get(clave)
            if desc is None:
                continue
            try:
                resultado[desc] = json.loads(valores_json)
            except (json.JSONDecodeError, TypeError):
                continue
        return resultado

    def cerrar(self):
        with self._lock:
            self._conn.close()
