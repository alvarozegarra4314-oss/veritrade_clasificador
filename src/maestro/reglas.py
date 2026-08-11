import re


def extraer_marca(desc_clean: str, maestro, desc_1_clean: str = None):
    """
    Evalúa las reglas de marca:
    1. Regex en la descripción completa.
    2. Diccionario de marcas en la descripción completa.
    3. Fallback: Posición 2 (después de la primera coma) EXCLUSIVAMENTE en Descripcion 1.
    """
    if not desc_clean:
        return None, None

    # 1. Reglas Regex (sobre todo el texto)
    for regex_comp in maestro.patrones_regex:
        m = regex_comp.search(desc_clean)
        if m:
            candidato = m.group(1).strip()
            if candidato not in maestro.stopwords:
                return candidato, "Regex Directa"

    # 2. Diccionario de Marcas (sobre todo el texto)
    for patron, estandar in maestro.lista_marcas:
        if re.search(fr'(?:^|(?<=\W)){re.escape(patron)}(?:$|(?=\W))', desc_clean):
            return estandar, "Diccionario Marcas"

    # 3. FALLBACK: Extraer posición 2 de Descripcion 1 (separada por comas)
    texto_comas = desc_1_clean if desc_1_clean else desc_clean
    partes = [p.strip() for p in texto_comas.split(',')]
    
    if len(partes) >= 2:
        candidato_posicion_2 = partes[1].strip()
        # Verificar que el candidato sea válido y no sea una palabra a ignorar
        if candidato_posicion_2 and len(candidato_posicion_2) > 1 and candidato_posicion_2 not in maestro.stopwords:
            return candidato_posicion_2, "Posición 2 (Descripcion 1)"

    return None, None


def evaluar_caracteristica_categorica(desc_clean: str, var_name: str, maestro):
    """Evalúa reglas de variables categóricas recibiendo la descripción YA limpiada."""
    if not desc_clean:
        return None
    reglas = maestro.dict_caracteristicas.get(var_name, [])
    for regex_comp, resultado in reglas:
        if regex_comp.search(desc_clean):
            return resultado
    return None


def extraer_potencia_numerica(desc_clean: str, var_name: str, maestro):
    """Extrae valores numéricos (potencia, voltaje, etc.) recibiendo la descripción YA limpiada."""
    if not desc_clean:
        return None
    patrones = maestro.dict_potencia.get(var_name, [])
    for regex_comp, mult in patrones:
        m = regex_comp.search(desc_clean)
        if m:
            val_str = m.group(1).replace(',', '.')
            try:
                val = float(val_str)
                return round(val * mult, 2)
            except ValueError:
                continue
    return None