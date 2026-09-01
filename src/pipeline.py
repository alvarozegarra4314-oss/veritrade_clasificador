import pandas as pd
from src.texto_utils import limpiar_texto, identificar_columnas_descripcion
from src.maestro.loader import CargarMaestro
from src.maestro.reglas import (
    extraer_marca,
    extraer_producto_y_modelo_desc1,
    evaluar_caracteristica_categorica,
    extraer_potencia_numerica,
    evaluar_condicionales,
)
from src.ia_rescate import RescatadorIA
from src.maestro_optimizer import construir_propuestas_aprendizaje


def _fila_necesita_rescate(marca, cat_vals: dict, var_principal: str) -> bool:
    """
    Un registro es candidato a rescate por IA si la marca no fue resuelta
    por reglas, o si el producto principal (variable clave del negocio)
    quedó sin clasificar. No exigimos que TODAS las categóricas estén
    llenas: muchas son opcionales y no ameritan gastar una llamada de IA.
    """
    if marca is None:
        return True
    if cat_vals.get(var_principal) is None:
        return True
    return False


def procesar_dataframe_dinamico(
    df_raw: pd.DataFrame,
    ruta_maestro,
    rescatador_ia: RescatadorIA = None,
    progreso_callback=None,
) -> pd.DataFrame:
    """
    rescatador_ia: instancia opcional de RescatadorIA (ver src/ia_rescate.py).
    Si es None, el pipeline se comporta exactamente igual que antes
    (100% determinista, sin llamadas de red).
    progreso_callback(fase: str, i: int, total: int): opcional, para alimentar
    una barra de progreso en Streamlit. Se emiten dos fases:
      - "reglas": avance fila a fila del motor determinista (fase 1).
      - "rescate_ia": avance sobre descripciones únicas pendientes (fase 2).
    Lanza ValueError con mensaje accionable si el archivo no tiene columnas
    de descripción reconocibles.
    """
    maestro = ruta_maestro if isinstance(ruta_maestro, CargarMaestro) else CargarMaestro(ruta_maestro)

    cols_desc = identificar_columnas_descripcion(df_raw.columns)
    if not cols_desc:
        # Error accionable para el usuario final: sin columnas de descripción
        # el pipeline no puede trabajar. Nunca debe llegar a un IndexError.
        raise ValueError(
            "El archivo no contiene ninguna columna de descripción reconocible.\n\n"
            "La herramienta busca columnas cuyo nombre incluya: DESCRIPCION, DETALLE, "
            "MERCADERIA o COMMODITY (se ignoran las administrativas como PARTIDA, "
            "ARANCEL, NANDINA, SUBPARTIDA).\n\n"
            f"Columnas encontradas en la hoja: {', '.join(map(str, df_raw.columns[:15]))}"
            + (" ..." if len(df_raw.columns) > 15 else "")
            + "\n\nVerifica que subiste el archivo Veritrade correcto y selecciona la "
              "hoja de datos (no una hoja auxiliar)."
        )
    cols_indices = [df_raw.columns.get_loc(c) for c in cols_desc]

    var_principal = maestro.variable_producto_principal
    valor_principal = maestro.valor_producto_principal
    variables_cat = maestro.variables_categoricas
    variables_pot = maestro.variables_potencia

    resultados = []
    # Guardamos por fila el desc_clean y si quedó pendiente de rescate,
    # para poder hacer una segunda pasada deduplicada sin repetir el
    # trabajo de las reglas.
    pendientes_rescate = []  # lista de (idx_resultado, desc_clean)
    descripciones_unicas_pendientes = set()

    total_filas = len(df_raw)

    # Memoización por (desc_full, desc_1): los archivos Veritrade repiten
    # muchísimo las mismas descripciones comerciales. Calcular las reglas UNA
    # sola vez por combinación única recorta drásticamente la fase 1 en
    # archivos grandes. Todas las funciones involucradas son deterministas
    # (regex sobre el mismo texto), así que el resultado es idéntico fila a fila.
    cache_reglas = {}

    for row in df_raw.itertuples(index=False):
        # 1. Unir todas las columnas de descripción para características y marcas por diccionario
        textos_desc = [str(row[i]) for i in cols_indices if pd.notna(row[i])]
        desc_completa = " ".join(textos_desc)
        desc_clean = limpiar_texto(desc_completa)

        # 2. Aísla únicamente la PRIMERA descripción (Descripcion 1 / Descripcion Comercial)
        desc_1_raw = str(row[cols_indices[0]]) if cols_indices and pd.notna(row[cols_indices[0]]) else ""
        desc_1_clean = limpiar_texto(desc_1_raw)

        clave_cache = (desc_clean, desc_1_clean)
        calculado = cache_reglas.get(clave_cache)
        if calculado is None:
            # 3. Extraer marca pasando desc_clean (para dict/regex) y desc_1_clean (para posición 2 por coma)
            marca, fuente = extraer_marca(desc_clean, maestro, desc_1_clean=desc_1_clean)

            # 3b. Extracción posicional pura desde Descripcion 1 (SOLO esa columna):
            #     posición 1 = producto y specs técnicas, posición 3 = modelo/serie comercial
            producto_texto_desc1, modelo_serie_desc1 = extraer_producto_y_modelo_desc1(desc_1_clean)

            cat_vals = {
                var: evaluar_caracteristica_categorica(desc_clean, var, maestro)
                for var in variables_cat
            }

            num_vals = {
                var: extraer_potencia_numerica(desc_clean, var, maestro)
                for var in variables_pot
            }

            # Condicionales por rango (hoja 5_Condicionales): rellenan variables
            # categóricas que las palabras clave no pudieron resolver, usando los
            # valores numéricos/categóricos ya extraídos (ej. clasificar por
            # rango de kVA). Nunca sobreescriben lo que ya resolvieron las reglas.
            cat_vals = evaluar_condicionales(cat_vals, num_vals, maestro)

            # Normalización cruzada: un equipo interactivo trifásico siempre es
            # "Online". Se aplica ANTES de cachear para que las filas repetidas
            # reciban exactamente el mismo valor que la primera aparición.
            if cat_vals.get("Tipo_Tecnologia") == "Interactivo" and cat_vals.get("Salida_Fases") == "Trifasico":
                cat_vals["Tipo_Tecnologia"] = "Online"

            cache_reglas[clave_cache] = (
                marca, fuente, producto_texto_desc1, modelo_serie_desc1,
                dict(cat_vals), dict(num_vals),
            )
        else:
            marca, fuente, producto_texto_desc1, modelo_serie_desc1, cat_vals, num_vals = calculado
            # Copias propias para esta fila: los dicts cacheados ya vienen con
            # condicionales y normalización aplicadas, pero fases posteriores
            # (rescate IA) mutan los valores de la fila.
            cat_vals = dict(cat_vals)
            num_vals = dict(num_vals)

        val_principal_extracted = cat_vals.get(var_principal)

        if valor_principal:
            es_principal = (val_principal_extracted == valor_principal)
        else:
            es_principal = bool(val_principal_extracted)

        marca_final = marca or maestro.dict_defaults.get(es_principal, "Marca Generica")
        fuente_marca = fuente if marca else "Default"

        res = {
            "Producto_Declarado": val_principal_extracted,
            "Marca_Declarada": marca,
            var_principal: val_principal_extracted,
            "Es_Producto_Principal": es_principal,
            "Marca_Extraida": marca_final,
            "Origen_Marca": fuente_marca,
            "Producto_Texto_Desc1": producto_texto_desc1,
            "Modelo_Serie_Desc1": modelo_serie_desc1,
            "Rescatado_Por_IA": False,
            "_desc_clean_ia": desc_clean,  # columna técnica, se elimina al final
        }

        for var in variables_cat:
            if var != var_principal:
                res[var] = cat_vals.get(var)

        for var in variables_pot:
            res[var] = num_vals.get(var)

        idx_actual = len(resultados)
        resultados.append(res)

        # Progreso de la fase de reglas (fase "reglas"). Se reporta cada fila;
        # el callback en la UI se auto-limita para no repintar la barra en vano.
        if progreso_callback:
            progreso_callback("reglas", idx_actual + 1, total_filas)

        if rescatador_ia is not None and _fila_necesita_rescate(marca, cat_vals, var_principal):
            pendientes_rescate.append((idx_actual, desc_clean))
            if desc_clean:
                descripciones_unicas_pendientes.add(desc_clean)

    # ------------------------------------------------------------------
    # FASE 2 (opcional): Rescate por IA generativa
    # Solo se ejecuta si se pasó un rescatador y hay filas pendientes.
    # Se llama UNA VEZ por descripción única (deduplicación real de costo).
    # ------------------------------------------------------------------
    # Guardamos también un mapa completo {desc_clean: ResultadoRescateIA}
    # de TODO lo que la IA resolvió con éxito, para poder retroalimentar
    # el maestro (ver src/maestro_optimizer.py) desde main.py/app.py.
    resultados_ia_exitosos = {}

    if rescatador_ia is not None and descripciones_unicas_pendientes:
        lista_unicas = sorted(descripciones_unicas_pendientes)

        def _cb(i, total):
            if progreso_callback:
                progreso_callback("rescate_ia", i, total)

        resultados_ia = rescatador_ia.rescatar_lote(lista_unicas, progreso_callback=_cb)
        resultados_ia_exitosos = {
            desc: r for desc, r in resultados_ia.items() if r and r.exito
        }

        for idx_actual, desc_clean in pendientes_rescate:
            resultado_ia = resultados_ia.get(desc_clean)
            if not resultado_ia or not resultado_ia.exito:
                continue  # degrada silenciosamente: se queda con el resultado de reglas

            valores_ia = resultado_ia.valores
            fila = resultados[idx_actual]
            rescato_algo = False

            # Solo completamos campos que las REGLAS dejaron vacíos.
            # La IA nunca sobreescribe un valor ya confirmado por reglas.
            if fila["Marca_Declarada"] is None and valores_ia.get("marca"):
                fila["Marca_Declarada"] = valores_ia["marca"]
                fila["Marca_Extraida"] = valores_ia["marca"]
                fila["Origen_Marca"] = "IA (Gemini)"
                rescato_algo = True

            for var in variables_cat:
                if fila.get(var) is None and valores_ia.get(var):
                    fila[var] = valores_ia[var]
                    if var == var_principal:
                        fila["Producto_Declarado"] = valores_ia[var]
                        if valor_principal:
                            fila["Es_Producto_Principal"] = (valores_ia[var] == valor_principal)
                        else:
                            fila["Es_Producto_Principal"] = True
                    rescato_algo = True

            for var in variables_pot:
                if fila.get(var) is None and valores_ia.get(var) is not None:
                    fila[var] = valores_ia[var]
                    rescato_algo = True

            if rescato_algo:
                fila["Rescatado_Por_IA"] = True

    # Columnas que SIEMPRE se ocultan del resultado final (aunque se calculan
    # internamente): son recortes literales de la descripción original que
    # solo sirven de apoyo interno, no aportan valor al Excel de salida.
    columnas_ocultas = ["_desc_clean_ia", "Marca_Declarada",
                        "Producto_Texto_Desc1", "Modelo_Serie_Desc1",
                        "Rescatado_Por_IA"]

    # Columnas que dependen de la variable principal (definida en 0b_Config_Linea).
    # Solo se muestran si el usuario definió un VALOR_PRODUCTO_PRINCIPAL; si lo
    # dejó vacío, estas columnas no aportan información útil y se ocultan.
    if not valor_principal:
        columnas_ocultas += ["Producto_Declarado", "Es_Producto_Principal", var_principal]

    df_res = pd.DataFrame(resultados).drop(
        columns=columnas_ocultas,
        errors="ignore",
    )
    df_final = pd.concat([df_raw.reset_index(drop=True), df_res.reset_index(drop=True)], axis=1)

    # Exponemos lo aprendido en esta corrida sobre el propio objeto
    # rescatador_ia (no cambiamos la firma de retorno de la función para
    # no romper a quienes ya la llaman esperando solo el DataFrame).
    # main.py / app.py puede leer `rescatador_ia.resultados_ia_exitosos`
    # después de esta llamada para generar el Maestro Optimizado.
    if rescatador_ia is not None:
        rescatador_ia.resultados_ia_exitosos = resultados_ia_exitosos
        # Construimos ya mismo las propuestas de aprendizaje para el maestro
        # (nuevas marcas / características detectadas por la IA), listas
        # para que main.py / app.py las persista con guardar_maestro_optimizado.
        rescatador_ia.propuestas_aprendizaje = construir_propuestas_aprendizaje(
            resultados_ia_exitosos, maestro, variables_cat
        )

    return df_final