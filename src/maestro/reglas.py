import re


# Unidades técnicas y términos que indican especificación, no marca
UNIDADES_Y_SPECS = {
    'V', 'A', 'W', 'KW', 'KVA', 'HZ', 'AMP', 'VOLTIOS', 'AMPERIOS', 
    '2P', '3P', '1P', '2P+T', '3P+T', 'IP20', 'IP65', 'IP66'
}

# Palabras de producto que frecuentemente se cuelan entre comas
PALABRAS_DESCARTE = {
    'TOMA', 'TOMACORRIENTE', 'INTERRUPTOR', 'PLACA', 'MODULO', 'TIPO', 
    'SERIE', 'CABLE', 'TABLERO', 'ENCHUFE', 'CLAVIJA', 'S/M', 'SIN', 'MARCA',
    'ITALIANA', 'EUROAMERICANA', 'SENCILLO', 'DOBLE', 'TRIPLE'
}


def es_candidato_marca_valido(candidato: str, stopwords: set) -> bool:
    """
    Valida si una cadena extraída por posición realmente cumple con el formato de una marca.
    Retorna False si detecta códigos, números, especificaciones técnicas o descripciones.
    """
    if not candidato:
        return False
        
    cand_upper = candidato.upper().strip()

    # 1. Descartar indicadores de "Sin Marca"
    if cand_upper in {"S/M", "S/MARCA", "SIN MARCA", "GENERICO", "S/M.", "S / M"}:
        return False

    # 2. Descartar si contiene NÚMEROS (Regla clave: las marcas no llevan dígitos en posición 2)
    if any(char.isdigit() for char in cand_upper):
        return False

    # 3. Descartar si contiene caracteres de códigos/modelos
    if any(char in cand_upper for char in ['/', '+', '=', '<', '>', '%', '#']):
        return False

    # 4. Descartar si excede 3 palabras o es demasiado largo
    palabras = cand_upper.split()
    if len(palabras) > 3 or len(cand_upper) > 25:
        return False

    # 5. Descartar si alguna palabra es un término de producto, unidad técnica o stopword
    for p in palabras:
        if p in PALABRAS_DESCARTE or p in UNIDADES_Y_SPECS or p in stopwords:
            return False

    return True

def extraer_marca(desc_clean: str, maestro, desc_1_clean: str = None):
    """
    Extrae la marca aplicando validación estricta antes de aceptar fallback posicional.
    """
    if not desc_clean:
        return None, None

    # 1. PRIORIDAD 1: Diccionario de Marcas (1_Marcas)
    for patron, estandar in maestro.lista_marcas:
        if re.search(fr'(?:^|(?<=\W)){re.escape(patron)}(?:$|(?=\W))', desc_clean):
            return estandar, "Diccionario Marcas"

    # 2. PRIORIDAD 2: Reglas Regex (4_Tecnico_RegexMarca_NOEDIT)
    for regex_comp in maestro.patrones_regex:
        m = regex_comp.search(desc_clean)
        if m:
            candidato = m.group(1).strip()
            if candidato not in maestro.stopwords and es_candidato_marca_valido(candidato, maestro.stopwords):
                return candidato, "Regex Directa"

    # 3. PRIORIDAD 3: Fallback por Comas con Validación Estricta
    texto_comas = desc_1_clean if desc_1_clean else desc_clean
    partes = [p.strip() for p in texto_comas.split(',') if p.strip()]

    # Evalúa cada bloque separado por coma a partir de la posición 2
    for candidato in partes[1:]:
        if es_candidato_marca_valido(candidato, maestro.stopwords):
            return candidato, "Posición Coma Validada"

    # Si nada fue válido, retorna None para que pipeline.py aplique el Default
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