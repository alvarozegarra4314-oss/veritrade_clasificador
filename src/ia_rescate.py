# -*- coding: utf-8 -*-
"""
Módulo de Rescate por IA Generativa (Google Gemini) — v2
-----------------------------------------------------------
Misma responsabilidad que la v1 (rescatar SOLO lo que las reglas no
resolvieron), pero optimizado para archivos masivos de Veritrade:

1. BATCHING REAL: en vez de 1 petición por descripción, se agrupan
   `batch_size` (20-50) descripciones en un único prompt y se le pide
   a Gemini un ARRAY JSON de igual longitud (response_schema tipo ARRAY).
   Esto reduce, p.ej., 1000 rescates a ~20-50 peticiones.

2. CACHÉ PERSISTENTE (SQLite, ver cache_ia.py): antes de armar los
   lotes, se descartan las descripciones que ya fueron rescatadas en
   una corrida anterior (cualquier día, cualquier archivo). Además se
   mantiene una caché en memoria (por sesión) como capa 0, más rápida
   que ir a SQLite.

3. BACKOFF EXPONENCIAL ante 429 / cuota agotada, con reintentos y
   jitter, aplicado a nivel de LOTE (si un lote falla por cuota, se
   reintenta el lote completo, no fila por fila). Si tras los
   reintentos el lote sigue fallando, esas descripciones simplemente
   quedan "no rescatadas" y el pipeline continúa con el resultado de
   las reglas (degradación segura, nunca se rompe la corrida).

Compatibilidad: la API pública (`rescatar`, `rescatar_lote`) se
mantiene igual que en v1, así que `pipeline.py` NO necesita cambios.
"""

from __future__ import annotations

import json
import time
import random
import logging
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

from src.cache_ia import CacheIA
from src.config import MODELO_IA_DEFAULT

logger = logging.getLogger("ia_rescate")

try:
    # SDK actual de Google (el paquete google-generativeai está deprecado
    # y sin soporte desde 2025 — ver README del repo google-gemini).
    from google import genai
    from google.genai import types as genai_types
    from google.genai import errors as genai_errors
    GENAI_DISPONIBLE = True
except ImportError:  # La librería es opcional: el pipeline debe seguir funcionando sin IA
    GENAI_DISPONIBLE = False
    genai = None
    genai_types = None
    genai_errors = None


# Modelo por defecto centralizado en src/config.py (editable ahí o desde la UI).
MODELO_IA = MODELO_IA_DEFAULT
# Se incrementa cada vez que cambia PROMPT_SISTEMA (o el orden del prompt
# de usuario) de forma que afecte los VALORES extraídos. Se concatena al
# nombre del modelo para formar la clave de caché (ver CacheIA._clave),
# así un cambio de prompt invalida automáticamente los resultados viejos
# en vez de servirlos para siempre desde SQLite. Súbela también si tocas
# `_bloque_catalogo` o `_prompt_lote` de forma que cambie el resultado.
PROMPT_VERSION = "v2-prioriza-caracteristicas"
# Bajado de 30 -> 15: lotes más chicos = más "atención relativa" por
# descripción = mejor recall de características técnicas. Sacrifica algo
# de throughput (más peticiones), pero con rpm_limite alto el impacto en
# tiempo total es bajo. Ajustable vía parámetro batch_size del constructor.
BATCH_SIZE_DEFAULT = 15


# ----------------------------------------------------------------------
# Construcción de schemas
# ----------------------------------------------------------------------
def _construir_schema_item(variables_cat: list[str], variables_pot: list[str]) -> dict:
    """
    Schema de UN registro rescatado (incluye 'id' para reasociar con la fila).

    IMPORTANTE: por cada variable categórica se pide también un campo
    gemelo "<var>__evidencia" con el fragmento EXACTO y literal copiado
    de la descripción que justifica el valor elegido. Sin esto, el
    optimizador de maestro (maestro_optimizer.py) no puede generar una
    regla regex confiable, porque el VALOR normalizado (ej. "Trifasico")
    casi nunca aparece tal cual en el texto crudo (que suele decir "3F",
    "TRIF.", "3~", etc.) — se necesita la evidencia, no el valor, para
    poder aprender un patrón de búsqueda reutilizable.
    """
    propiedades = {
        "id": {
            "type": "INTEGER",
            "description": "Índice del ítem dentro del lote enviado (posición 0-based en la lista de entrada).",
        },
        "marca": {
            "type": "STRING",
            "nullable": True,
            "description": "Marca comercial detectada, o null si el texto no la declara o dice explícitamente 'sin marca'.",
        },
    }
    for var in variables_cat:
        propiedades[var] = {
            "type": "STRING",
            "nullable": True,
            "description": f"Valor categórico para '{var}', o null si no se puede inferir con certeza.",
        }
        propiedades[f"{var}__evidencia"] = {
            "type": "STRING",
            "nullable": True,
            "description": (
                f"Fragmento EXACTO y literal (subcadena tal cual, sin corregir ni traducir) "
                f"copiado de la descripción original que evidencia el valor de '{var}'. "
                f"Si '{var}' quedó en null, este campo también debe ser null."
            ),
        }
    for var in variables_pot:
        propiedades[var] = {
            "type": "NUMBER",
            "nullable": True,
            "description": f"Valor numérico técnico para '{var}' (sin unidades), o null si no aparece.",
        }
    return {
        "type": "OBJECT",
        "properties": propiedades,
        "required": list(propiedades.keys()),
    }


def _construir_schema_lote(variables_cat: list[str], variables_pot: list[str]) -> dict:
    """Schema del LOTE completo: un ARRAY de items."""
    return {
        "type": "ARRAY",
        "items": _construir_schema_item(variables_cat, variables_pot),
    }


PROMPT_SISTEMA = """\
Eres un experto clasificador TÉCNICO de descripciones aduaneras de importación (Veritrade) \
de equipos eléctricos e industriales (UPS, tableros, interruptores, tomacorrientes, etc.).

TU PRIORIDAD #1 son las CARACTERÍSTICAS TÉCNICAS (variables categóricas y numéricas), \
NO la marca. La marca es un dato secundario, fácil de detectar, que casi siempre aparece \
al inicio del texto — por eso es la tentación natural rellenarla y descuidar el resto. \
Vas a ser evaluado principalmente por cuántas características técnicas correctas extraes, \
no por si detectas la marca. Dedica tu esfuerzo de lectura a comparar el texto contra el \
catálogo de valores permitidos de CADA variable categórica, y a buscar los números técnicos \
(potencia, voltaje, corriente, etc.), aunque estén abreviados, en otro idioma o en medio \
del texto.

Vas a recibir una LISTA NUMERADA de descripciones comerciales crudas. Debes devolver \
un ARRAY JSON con EXACTAMENTE un objeto por cada descripción recibida, en el mismo \
orden, incluyendo el campo "id" con el índice (0-based) que corresponde a esa entrada \
en la lista recibida. Nunca omitas, fusiones ni reordenes ítems: si una descripción \
no tiene información suficiente, igual debes incluir su objeto con los campos en null.

PROCEDIMIENTO OBLIGATORIO para cada descripción (síguelo internamente, sin escribirlo \
en la salida — la salida es solo el JSON):
1. Lee la descripción completa, sin detenerte apenas encuentres una marca.
2. Recorre, UNA POR UNA, todas las variables categóricas de la lista de "VARIABLES \
   CATEGÓRICAS Y SUS VALORES PERMITIDOS". Para cada una, busca en el texto cualquier \
   abreviatura, sigla, número o palabra (en español, inglés u otro idioma) que corresponda \
   a alguno de los valores permitidos de esa variable. No te detengas en la primera \
   variable que resuelvas: revisa TODAS antes de continuar.
3. Recorre, UNA POR UNA, todas las variables numéricas de la lista de "VARIABLES \
   NUMÉRICAS A EXTRAER" y busca su valor técnico en el texto (con o sin unidades).
4. Solo al final, identifica la marca comercial (si la hay).
5. Arma el objeto de salida con TODOS los campos: características primero en tu análisis, \
   luego marca.

Reglas de extracción, con criterio conservador pero exhaustivo (no te quedes corto por \
priorizar velocidad; revisa todo el texto, no solo el inicio):

- Para variables categóricas, usa EXCLUSIVAMENTE uno de los valores permitidos que se \
te indiquen en la lista de opciones válidas para esa variable; si ninguno aplica, responde null. \
No dejes una variable en null solo porque no la revisaste: si tras comparar contra el \
catálogo genuinamente no hay evidencia, entonces sí responde null.
- Para CADA variable categórica que sí puedas clasificar, además del valor debes llenar \
su campo gemelo "<variable>__evidencia" con el fragmento EXACTO, copiado tal cual (misma \
ortografía, mayúsculas/minúsculas, abreviaturas) de la descripción original que te llevó \
a esa conclusión. Por ejemplo, si clasificas Salida_Fases = "Trifasico" porque el texto \
dice "3F" o "TRIF", el campo de evidencia debe ser exactamente "3F" o "TRIF" (la subcadena \
real del texto), NUNCA la palabra "Trifasico" ni una paráfrasis. Si no encuentras una \
subcadena literal clara que respalde el valor, sé más cauteloso con la clasificación y, \
si aun así no hay evidencia textual concreta, deja tanto el valor como su evidencia en null.
- Para variables numéricas (potencia, voltaje, corriente, etc.), extrae solo el número, \
sin unidades ni texto adicional. Si hay un rango (ej. 110-220), toma el valor más alto.
- No inventes ni "adivines" un valor si el texto es ambiguo: en ese caso responde null, \
pero solo después de haber buscado genuinamente evidencia para esa variable.
- Ignora ruido típico de aduanas: códigos arancelarios, cantidades, pesos, condiciones \
comerciales (DIFERIDO, DIAS, etc.). Ese ruido no es excusa para dejar de buscar características.
- Si la marca no está explícita o el texto indica "S/M", "SIN MARCA", "GENERICO" -> marca = null. \
Recuerda: la marca se resuelve AL FINAL, después de agotar la revisión de características.
- Responde exclusivamente con el ARRAY JSON solicitado, sin texto adicional ni comentarios.
"""


@dataclass
class ResultadoRescateIA:
    valores: dict = field(default_factory=dict)
    exito: bool = False
    error: Optional[str] = None


class RescatadorIA:
    """
    Envuelve la SDK de Gemini con:
      - caché en memoria (capa 0, por ejecución)
      - caché persistente SQLite (capa 1, entre ejecuciones)
      - batching real (N descripciones -> 1 sola llamada)
      - backoff exponencial con jitter ante 429 / cuota
    """

    def __init__(
        self,
        api_key: str,
        maestro,
        modelo: str = MODELO_IA,
        max_reintentos: int = 5,
        temperatura: float = 0.0,
        rpm_limite: int = 12,
        batch_size: int = BATCH_SIZE_DEFAULT,
        ruta_cache_db: str | Path | None = None,
        usar_cache_persistente: bool = True,
    ):
        """
        rpm_limite: peticiones POR LOTE por minuto que el cliente se
        auto-impone. Como cada petición ahora mueve hasta `batch_size`
        descripciones, con rpm_limite=12 y batch_size=15 se pueden
        rescatar hasta ~180 descripciones/min sin tocar el límite de 15
        RPM del tier gratuito (dejamos margen de por medio).
        batch_size: cuántas descripciones se agrupan por petición.
        DEFAULT bajado a 15 (antes 30): lotes más chicos mejoran el
        recall de características técnicas porque el modelo reparte
        mejor su atención por descripción, a costa de más peticiones.
        Si priorizas velocidad sobre precisión de características, puedes
        subirlo hasta 30-50, pero revisa el impacto con el script de
        evaluación (eval_recall_caracteristicas.py).
        """
        if not GENAI_DISPONIBLE:
            raise RuntimeError(
                "El paquete 'google-genai' no está instalado. "
                "Agrega 'google-genai' a requirements.txt para usar el rescate por IA."
            )

        # SDK nuevo: cliente por instancia (sin estado global como el
        # genai.configure() del paquete deprecado).
        self.client = genai.Client(api_key=api_key)

        self.maestro = maestro
        # Clave de caché real: incluye la versión de prompt para que un
        # cambio de PROMPT_SISTEMA invalide automáticamente el caché viejo.
        self._modelo_cache_key = f"{modelo}::{PROMPT_VERSION}"
        self.max_reintentos = max_reintentos
        self.batch_size = max(1, batch_size)
        self.variables_cat = list(maestro.variables_categoricas)
        self.variables_pot = list(maestro.variables_potencia)

        self._schema_lote = _construir_schema_lote(self.variables_cat, self.variables_pot)

        # Config de generación reutilizable: JSON estricto con el schema del
        # lote. En el SDK nuevo el system prompt vive AQUÍ (antes iba en
        # GenerativeModel); perderlo cambiaría por completo la conducta del modelo.
        # AFC (function calling automático) se desactiva: este flujo es de
        # salida JSON estructurada pura y el SDK recomienda no usar AFC directo
        # en Models.generate_content.
        self._generation_config = genai_types.GenerateContentConfig(
            system_instruction=PROMPT_SISTEMA,
            temperature=temperatura,
            response_mime_type="application/json",
            response_schema=self._schema_lote,
            automatic_function_calling=genai_types.AutomaticFunctionCallingConfig(disable=True),
        )

        self.modelo_nombre = modelo
        logger.info("RescateIA inicializado con modelo: %s", modelo)

        # Capa 0: caché en memoria (vida = duración del proceso/sesión Streamlit)
        self._cache_mem: dict[str, ResultadoRescateIA] = {}

        # Capa 1: caché persistente en disco (vive entre corridas y sesiones).
        # Ruta ABSOLUTA anclada a la raíz del proyecto: si la app se lanza
        # desde otra carpeta, la caché relativa al CWD crearía un SQLite nuevo
        # y el usuario perdería todo el ahorro de cuota acumulado.
        self.usar_cache_persistente = usar_cache_persistente
        if usar_cache_persistente:
            if ruta_cache_db is None:
                ruta_cache_db = Path(__file__).resolve().parent.parent / "data" / "cache_ia.sqlite"
            self._cache_db = CacheIA(ruta_cache_db)
        else:
            self._cache_db = None

        # Rate limiting propio: espaciamos llamadas (por LOTE) según rpm_limite
        self._rpm_limite = max(1, rpm_limite)
        self._intervalo_min_seg = 60.0 / self._rpm_limite
        self._ultima_llamada_ts = 0.0

        # Telemetría / UI de Streamlit
        self.llamadas_realizadas = 0          # peticiones HTTP reales a Gemini
        self.lotes_procesados = 0
        self.descripciones_desde_cache_mem = 0
        self.descripciones_desde_cache_db = 0
        self.descripciones_rescatadas_api = 0
        self.errores = 0
        self.errores_cuota = 0
        self.errores_api = 0
        self.errores_formato = 0
        self.errores_otros = 0
        self.ultimo_error_detalle: Optional[str] = None
        self.detalle_errores: list[dict] = []
        self._max_detalle_errores = 50

    # ------------------------------------------------------------------
    # Compatibilidad hacia atrás con la v1 (atributos que tu app.py /
    # dashboards de Streamlit puedan estar leyendo con el nombre viejo)
    # ------------------------------------------------------------------
    @property
    def llamadas_desde_cache(self) -> int:
        """Alias de v1: total de descripciones resueltas sin llamar a la API
        (memoria + SQLite persistente)."""
        return self.descripciones_desde_cache_mem + self.descripciones_desde_cache_db

    # ------------------------------------------------------------------
    # Construcción del prompt de usuario para un LOTE
    # ------------------------------------------------------------------
    def _opciones_validas(self, var: str) -> list[str]:
        reglas = self.maestro.dict_caracteristicas.get(var, [])
        vistos, opciones = set(), []
        for _, valor in reglas:
            if valor not in vistos:
                vistos.add(valor)
                opciones.append(str(valor))
        return opciones

    def _bloque_catalogo(self) -> str:
        bloques = []
        if self.variables_cat:
            bloques.append(
                "=== CHECKLIST OBLIGATORIA: revisa TODAS estas variables antes de mirar la marca ===\n"
                "VARIABLES CATEGÓRICAS Y SUS VALORES PERMITIDOS:"
            )
            for var in self.variables_cat:
                opciones = self._opciones_validas(var)
                if opciones:
                    bloques.append(f"- {var}: {', '.join(opciones)}")
                else:
                    bloques.append(f"- {var}: (sin catálogo cerrado, usa tu criterio o null)")
        if self.variables_pot:
            bloques.append("\nVARIABLES NUMÉRICAS A EXTRAER (solo el número): " + ", ".join(self.variables_pot))
        return "\n".join(bloques)

    def _prompt_lote(self, descripciones: list[str]) -> str:
        lineas = [f"[{i}] {desc}" for i, desc in enumerate(descripciones)]
        bloque_desc = "DESCRIPCIONES A ANALIZAR (formato [id] texto):\n" + "\n".join(lineas)
        # IMPORTANTE: el catálogo de variables va PRIMERO y las descripciones
        # DESPUÉS. Esto no es cosmético: el modelo presta más atención a lo
        # que lee primero, y queremos que llegue a las descripciones ya con
        # la checklist de características en mente, no que la vea como un
        # "además" al final después de haber quedado satisfecho con la marca.
        return (
            f"{self._bloque_catalogo()}\n\n"
            f"Para CADA una de las {len(descripciones)} descripciones de abajo, recorre "
            f"primero la lista de variables categóricas y numéricas de arriba (una por una) "
            f"antes de fijarte en la marca.\n\n"
            f"{bloque_desc}\n\n"
            f"Devuelve el ARRAY JSON con {len(descripciones)} objetos, uno por id."
        )

    # ------------------------------------------------------------------
    # Rate limiting + backoff exponencial con jitter (a nivel de lote)
    # ------------------------------------------------------------------
    def _esperar_rate_limit(self):
        ahora = time.monotonic()
        transcurrido = ahora - self._ultima_llamada_ts
        if transcurrido < self._intervalo_min_seg:
            time.sleep(self._intervalo_min_seg - transcurrido)
        self._ultima_llamada_ts = time.monotonic()

    def _registrar_error_detalle(self, descripciones: list[str], error: str):
        if len(self.detalle_errores) < self._max_detalle_errores:
            self.detalle_errores.append({
                "descripcion": descripciones[0] if descripciones else "",
                "error": error,
                "tamano_lote": len(descripciones),
            })

    def _llamar_gemini_lote(self, descripciones: list[str]) -> dict[int, dict]:
        """
        Llama a Gemini con un lote y devuelve {indice_local: valores_dict}.
        Reintenta con backoff exponencial + jitter ante 429 / errores
        transitorios. Si todo falla, devuelve {} (degradación segura:
        esas filas se quedan con el resultado de las reglas).
        """
        prompt = self._prompt_lote(descripciones)

        for intento in range(1, self.max_reintentos + 1):
            self._esperar_rate_limit()
            try:
                self.llamadas_realizadas += 1
                if self.llamadas_realizadas == 1:
                    logger.info("Primera llamada a Gemini — modelo: %s", self.modelo_nombre)
                respuesta = self.client.models.generate_content(
                    model=self.modelo_nombre,
                    contents=prompt,
                    config=self._generation_config,
                )
                data = json.loads(respuesta.text)

                if not isinstance(data, list):
                    raise ValueError(f"Se esperaba un ARRAY JSON, se recibió: {type(data)}")

                mapeado = {}
                for item in data:
                    idx = item.get("id")
                    if isinstance(idx, int) and 0 <= idx < len(descripciones):
                        mapeado[idx] = item
                return mapeado

            except genai_errors.APIError as e:
                codigo = getattr(e, "code", None)

                if codigo == 429:
                    # 429 - cuota / rate limit: backoff exponencial + jitter
                    espera = min(60, (2 ** intento) + random.uniform(0, 1))
                    logger.warning(
                        "Cuota de Gemini excedida en lote de %s (intento %s/%s). "
                        "Reintentando en %.1fs. Detalle: %s",
                        len(descripciones), intento, self.max_reintentos, espera, e,
                    )
                    self.ultimo_error_detalle = f"429 ResourceExhausted: {e}"
                    time.sleep(espera)
                    continue

                logger.error("Error de API de Gemini (codigo %s): %s", codigo, e)
                self.errores += 1
                self.errores_api += 1
                self.ultimo_error_detalle = f"APIError {codigo}: {e}"
                self._registrar_error_detalle(descripciones, str(e))
                # Errores de API "duros" (no cuota) también ameritan un
                # reintento breve por si es un problema transitorio de red.
                if intento < self.max_reintentos:
                    time.sleep(min(30, 2 ** intento))
                    continue
                return {}

            except (json.JSONDecodeError, ValueError) as e:
                logger.error("Respuesta de Gemini no es JSON de lote válido: %s", e)
                self.errores += 1
                self.errores_formato += 1
                self.ultimo_error_detalle = f"JSON inválido: {e}"
                self._registrar_error_detalle(descripciones, str(e))
                # Si el lote es grande, un reintento partiéndolo a la mitad
                # suele resolver truncamientos por tamaño de respuesta.
                if len(descripciones) > 5 and intento < self.max_reintentos:
                    mitad = len(descripciones) // 2
                    izq = self._llamar_gemini_lote(descripciones[:mitad])
                    der = self._llamar_gemini_lote(descripciones[mitad:])
                    der_reindexado = {i + mitad: v for i, v in der.items()}
                    return {**izq, **der_reindexado}
                return {}

            except Exception as e:  # Red, timeout u otro imprevisto
                logger.error("Error inesperado llamando a Gemini (lote): %s", e)
                self.errores += 1
                self.errores_otros += 1
                self.ultimo_error_detalle = f"{type(e).__name__}: {e}"
                self._registrar_error_detalle(descripciones, str(e))
                if intento < self.max_reintentos:
                    time.sleep(min(30, 2 ** intento))
                    continue
                return {}

        self.errores += 1
        self.errores_cuota += 1
        self.ultimo_error_detalle = self.ultimo_error_detalle or "Cuota excedida tras reintentos"
        self._registrar_error_detalle(descripciones, "Cuota excedida tras reintentos de lote")
        return {}

    # ------------------------------------------------------------------
    # API pública: rescate de una sola descripción (usa el pipeline de lote de tamaño 1)
    # ------------------------------------------------------------------
    def rescatar(self, desc_clean: str) -> ResultadoRescateIA:
        if not desc_clean:
            return ResultadoRescateIA(exito=False, error="Descripción vacía")
        resultados = self.rescatar_lote([desc_clean])
        return resultados.get(desc_clean, ResultadoRescateIA(exito=False, error="Sin resultado"))

    # ------------------------------------------------------------------
    # API pública: rescate por lote real, con caché de 2 capas
    # ------------------------------------------------------------------
    def rescatar_lote(self, descripciones_unicas: list[str], progreso_callback=None) -> dict[str, ResultadoRescateIA]:
        """
        Procesa una lista de descripciones ÚNICAS y devuelve un dict
        {desc_clean: ResultadoRescateIA}.

        Orden de resolución por descripción:
          1. Caché en memoria (capa 0)
          2. Caché persistente en disco (capa 1)
          3. Llamada real a Gemini, agrupada en lotes de `self.batch_size`

        `progreso_callback(i, total)` opcional, cuenta descripciones ya
        resueltas (desde cualquier capa) para alimentar una barra de
        progreso en Streamlit.
        """
        total = len(descripciones_unicas)
        resultados: dict[str, ResultadoRescateIA] = {}
        resueltas = 0

        def _avance():
            nonlocal resueltas
            resueltas += 1
            if progreso_callback:
                progreso_callback(resueltas, total)

        # --- Capa 0: memoria ---
        pendientes = []
        for desc in descripciones_unicas:
            if desc in self._cache_mem:
                resultados[desc] = self._cache_mem[desc]
                self.descripciones_desde_cache_mem += 1
                _avance()
            else:
                pendientes.append(desc)

        # --- Capa 1: SQLite persistente ---
        if pendientes and self.usar_cache_persistente:
            encontrados_db = self._cache_db.obtener_muchos(pendientes, self._modelo_cache_key)
            if encontrados_db:
                nuevos_pendientes = []
                for desc in pendientes:
                    if desc in encontrados_db:
                        r = ResultadoRescateIA(valores=encontrados_db[desc], exito=True)
                        resultados[desc] = r
                        self._cache_mem[desc] = r
                        self.descripciones_desde_cache_db += 1
                        _avance()
                    else:
                        nuevos_pendientes.append(desc)
                pendientes = nuevos_pendientes

        # --- Capa 2: API real, con tamaño de lote dinámico ---
        def _tamano_lote(n_total: int) -> int:
            """Determina el tamaño del lote según la cantidad TOTAL de filas pendientes.
            Se evalúa UNA vez al inicio con el total completo."""
            if n_total <= 100:
                return 10
            elif n_total <= 300:
                return 15
            elif n_total <= 1000:
                return 50
            else:
                return 60

        batch = _tamano_lote(len(pendientes))
        logger.info(
            "Lotes dinámicos: %d pendientes → batch=%d", len(pendientes), batch
        )
        for idx_lote, inicio in enumerate(range(0, len(pendientes), batch), 1):
            lote = pendientes[inicio:inicio + batch]
            logger.info(
                "Lote %d/%d (%d ítems, offset %d)",
                idx_lote, -(-len(pendientes) // batch), len(lote), inicio,
            )
            mapeado = self._llamar_gemini_lote(lote)
            self.lotes_procesados += 1

            for i, desc in enumerate(lote):
                item = mapeado.get(i)
                if item is not None:
                    valores = {k: v for k, v in item.items() if k != "id"}
                    r = ResultadoRescateIA(valores=valores, exito=True)
                    self.descripciones_rescatadas_api += 1
                    if self.usar_cache_persistente:
                        self._cache_db.guardar(desc, self._modelo_cache_key, valores)
                else:
                    r = ResultadoRescateIA(exito=False, error="No incluido en respuesta de lote / lote fallido")

                resultados[desc] = r
                self._cache_mem[desc] = r
                _avance()

        return resultados

    def cerrar(self):
        """Libera la conexión a la caché persistente. Llamar al finalizar la corrida (opcional)."""
        if self._cache_db is not None:
            self._cache_db.cerrar()