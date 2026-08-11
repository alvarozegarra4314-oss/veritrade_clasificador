import re


def extraer_marca(desc_clean: str, maestro):
    """Evalúa las reglas de marca recibiendo la descripción YA limpiada."""
    if not desc_clean:
        return None, None

    for regex_comp in maestro.patrones_regex:
        m = regex_comp.search(desc_clean)
        if m:
            candidato = m.group(1).strip()
            if candidato not in maestro.stopwords:
                return candidato, "Regex Directa"

    for patron, estandar in maestro.lista_marcas:
        if re.search(fr'(?:^|(?<=\W)){re.escape(patron)}(?:$|(?=\W))', desc_clean):
            return estandar, "Diccionario Marcas"

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