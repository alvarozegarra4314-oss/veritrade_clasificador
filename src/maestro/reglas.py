import re

# Indicadores explícitos de productos sin marca
INDICADORES_SIN_MARCA = {
    "S/M", "S/MARCA", "SIN MARCA", "GENERICO", "S/M.", "S / M", 
    "NO APLICA", "N/A", "NO INDICA", "SIN/MARCA", "S/ M"
}

# Unidades técnicas y especificaciones
UNIDADES_Y_SPECS = {
    'V', 'A', 'W', 'KW', 'KVA', 'HZ', 'AMP', 'VOLTIOS', 'AMPERIOS', 
    '2P', '3P', '1P', '2P+T', '3P+T', 'IP20', 'IP65', 'IP66'
}

# Palabras de productos y atributos comerciales a descartar
PALABRAS_DESCARTE = {
    'TOMA', 'TOMACORRIENTE', 'INTERRUPTOR', 'PLACA', 'MODULO', 'TIPO', 
    'SERIE', 'CABLE', 'TABLERO', 'ENCHUFE', 'CLAVIJA', 'SIN', 'MARCA',
    'ITALIANA', 'EUROAMERICANA', 'SENCILLO', 'DOBLE', 'TRIPLE',
    # Atributos comerciales y aduaneros frecuentes
    'DISPOSITIVO', 'ACCESORIO', 'ACABADO', 'COMPOSICION', 'DIMENSIONES',
    'PRESENTACION', 'PESO', 'APLICACION', 'USO', 'COLOR', 'DISEÑO',
    'DIFERIDO', 'DIAS', 'UNIDAD', 'PIEZAS'
}

# Prefijos de negación comunes en descripciones comerciales
PREFIJOS_NEGACION = [
    "NO ", "SIN ", "EXCLUYE ", "NO INCLUYE ", "S/ ", "S/MAPA ", "SIN/ "
]

def es_indicador_sin_marca(candidato: str) -> bool:
    """Retorna True si el texto declara explícitamente que no tiene marca."""
    if not candidato:
        return False
    return candidato.upper().strip() in INDICADORES_SIN_MARCA


def normalizar_numero_extraido(texto_cifra: str) -> float:
    """
    Convierte cadenas numéricas complejas a float limpio:
    - '1,5' -> 1.5
    - '110-220' -> 220.0 (toma el límite superior del rango)
    - '10/15' -> 15.0
    - '1,000' -> 1000.0 (maneja miles)
    """
    if not texto_cifra:
        return None

    cifra_clean = texto_cifra.strip()

    # 1. Si es un rango (ej: '110-220' o '10/15'), tomar el número mayor (límite superior)
    if '-' in cifra_clean or '/' in cifra_clean:
        partes = re.split(r'[-/]', cifra_clean)
        # Extraer solo los números de cada parte
        valores = []
        for p in partes:
            num_str = re.sub(r'[^\d.,]', '', p)
            if num_str:
                val = normalizar_numero_extraido(num_str)
                if val is not None:
                    valores.append(val)
        return max(valores) if valores else None

    # 2. Manejar separadores de miles vs decimales
    # Si tiene punto/coma seguido de exactamente 3 dígitos al final (ej: 1,000 o 1.000)
    if re.search(r'[.,]\d{3}$', cifra_clean) and not re.search(r'[.,]\d{1,2}$', cifra_clean):
        cifra_clean = re.sub(r'[.,]', '', cifra_clean)  # Quitar separador de miles
    else:
        # Reemplazar coma por punto decimal
        cifra_clean = cifra_clean.replace(',', '.')

    # 3. Eliminar cualquier caracter no numérico remanente salvo el punto
    cifra_clean = re.sub(r'[^\d.]', '', cifra_clean)

    try:
        return float(cifra_clean)
    except ValueError:
        return None

def tiene_negacion_previa(texto: str, posicion_inicio_match: int) -> bool:
    """
    Inspecciona el texto justo antes de la coincidencia para determinar
    si está precedido por una palabra de negación (ej. 'NO ONLINE', 'SIN BATERIA').
    """
    # Toma hasta 15 caracteres hacia atrás desde donde inició la coincidencia
    inicio_contexto = max(0, posicion_inicio_match - 15)
    texto_previo = texto[inicio_contexto:posicion_inicio_match].upper()

    # Verifica si alguna palabra de negaciones está al final del texto previo
    for neg in PREFIJOS_NEGACION:
        if texto_previo.endswith(neg) or neg in texto_previo[-8:]:
            return True
            
    return False


def es_candidato_marca_valido(candidato: str, stopwords: set) -> bool:
    """Valida si una cadena realmente puede ser el nombre de una marca."""
    if not candidato:
        return False
        
    cand_upper = candidato.upper().strip()

    # 1. Descartar si es indicador explícito de "Sin Marca"
    if es_indicador_sin_marca(cand_upper):
        return False

    # 2. Descartar si contiene NÚMEROS
    if any(char.isdigit() for char in cand_upper):
        return False

    # 3. Descartar si contiene símbolos de códigos o atributos clave-valor (:)
    if any(char in cand_upper for char in ['/', '+', '=', '<', '>', '%', '#', ':']):
        return False

    # 4. Descartar si excede 3 palabras o es demasiado largo (> 25 caracteres)
    palabras = cand_upper.split()
    if len(palabras) > 3 or len(cand_upper) > 25:
        return False

    # 5. Descartar si alguna palabra es un término descartado o unidad técnica
    for p in palabras:
        p_clean = re.sub(r'[^\w]', '', p)  # Limpia puntuación pegada
        if p_clean in PALABRAS_DESCARTE or p_clean in UNIDADES_Y_SPECS or p_clean in stopwords:
            return False

    return True

def extraer_marca(desc_clean: str, maestro, desc_1_clean: str = None):
    """
    Extrae la marca validando prioridad y deteniéndose si detecta S/M explícito.
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

    # 3. PRIORIDAD 3: Evaluamos únicamente Descripcion 1 por comas
    texto_comas = desc_1_clean if desc_1_clean else desc_clean
    partes = [p.strip() for p in texto_comas.split(',') if p.strip()]

    # Recorremos los bloques desde la Posición 2
    for candidato in partes[1:]:
        # REGLA CLAVE: Si la posición declara explícitamente "S/M" o "SIN MARCA",
        # DETENEMOS LA BÚSQUEDA inmediatamente. No se evalúan bloques posteriores.
        if es_indicador_sin_marca(candidato):
            return None, "Declarado Sin Marca (S/M)"

        if es_candidato_marca_valido(candidato, maestro.stopwords):
            return candidato, "Posición Coma Validada"

    return None, None


def extraer_producto_y_modelo_desc1(desc_1_clean: str):
    """
    Extrae DIRECTAMENTE por posición (sin diccionarios ni regex de negocio)
    dos campos a partir de "Descripcion 1", que sigue el formato típico:

        "Producto y Especificaciones Técnicas, Marca Declarada, Modelo o Serie Comercial"

    Ejemplo:
        "ALLSAIW 10K PRO 3/3 - UPS, ALLSAI, W KPRO"
         -> producto_texto = "ALLSAIW 10K PRO 3/3 - UPS"   (posición 1, antes de la 1a coma)
         -> modelo_serie   = "W KPRO"                      (posición 3, después de la 2a coma)

    Reglas:
    - Solo se usa `desc_1_clean` (Descripcion 1 / Descripcion Comercial). Nunca
      se mezcla con las demás columnas de descripción.
    - `producto_texto` = todo el bloque antes de la primera coma (posición 1).
      Si no hay comas, es la descripción completa.
    - `modelo_serie` = todo lo que queda a partir de la 3ra posición
      (después de la marca, posición 2). Si hay más de 3 bloques,
      se re-unen con ", " porque el modelo/serie puede traer comas
      internas (ej. números de serie con formato raro).
    - Si no existe la posición 3 (p. ej. "PRODUCTO, MARCA" sin modelo),
      `modelo_serie` devuelve None.
    - No aplica validaciones de "marca válida" ni descarta S/M: es una
      extracción posicional pura, independiente de `extraer_marca`.
    """
    if not desc_1_clean:
        return None, None

    partes = [p.strip() for p in desc_1_clean.split(',') if p.strip()]

    producto_texto = partes[0] if len(partes) >= 1 else None
    modelo_serie = ", ".join(partes[2:]) if len(partes) >= 3 else None

    return producto_texto, modelo_serie


def evaluar_potencia_numerica_condicion(valor_actual, operador: str, valor_1, valor_2) -> bool:
    """Compara un valor NUMÉRICO ya extraído contra la condición de una regla condicional."""
    if valor_actual is None:
        return False
    try:
        v = float(valor_actual)
    except (TypeError, ValueError):
        return False

    if operador == ">":
        return v > valor_1 if valor_1 is not None else False
    if operador == ">=":
        return v >= valor_1 if valor_1 is not None else False
    if operador == "<":
        return v < valor_1 if valor_1 is not None else False
    if operador == "<=":
        return v <= valor_1 if valor_1 is not None else False
    if operador == "==":
        return valor_1 is not None and v == float(valor_1)
    if operador == "!=":
        return valor_1 is not None and v != float(valor_1)
    if operador == "BETWEEN":
        if valor_1 is None or valor_2 is None:
            return False
        lo, hi = min(valor_1, valor_2), max(valor_1, valor_2)
        return lo <= v <= hi
    return False


def evaluar_categorica_condicion(valor_actual, operador: str, valor_1) -> bool:
    """Compara un valor CATEGÓRICO ya extraído contra la condición de una regla condicional."""
    if valor_actual is None:
        return False
    actual_norm = str(valor_actual).strip().upper()
    esperado_norm = str(valor_1).strip().upper() if valor_1 is not None else None

    if operador == "==":
        return esperado_norm is not None and actual_norm == esperado_norm
    if operador == "!=":
        return esperado_norm is not None and actual_norm != esperado_norm
    return False


def evaluar_condicionales(cat_vals: dict, num_vals: dict, maestro) -> dict:
    """
    Aplica las reglas condicionales del maestro (hoja 5_Condicionales) sobre
    los valores ya extraídos por palabras clave/regex y por potencia numérica.

    Reglas de comportamiento:
    - Solo actúa sobre una Variable_Resultado si esta sigue en None: nunca
      sobreescribe un valor que ya fue resuelto por coincidencia de palabra
      clave (misma jerarquía que el rescate por IA: reglas > condicionales).
    - Las condiciones de una misma Regla_ID se combinan con AND: todas deben
      cumplirse para que se asigne Valor_Resultado.
    - Se evalúan en orden de Prioridad (ya vienen ordenadas desde el loader);
      la primera regla que cumple para una variable "gana" y las siguientes
      reglas para esa misma variable se ignoran automáticamente porque el
      valor deja de ser None.
    - maestro.condicionales ya viene filtrado en el loader para excluir
      cualquier regla que referencie variables ajenas a esta línea de
      producto (protección de scoping entre líneas, ej. UPS vs Interruptores).
    """
    for regla in getattr(maestro, "condicionales", []):
        var_resultado = regla["variable_resultado"]

        if cat_vals.get(var_resultado) is not None:
            continue  # ya resuelto por reglas de palabra clave, no se toca

        cumple_todas = True
        for cond in regla["condiciones"]:
            var_condicion = cond["variable"]
            operador = cond["operador"]

            if cond["es_numerica"]:
                valor_actual = num_vals.get(var_condicion)
                cumple = evaluar_potencia_numerica_condicion(
                    valor_actual, operador, cond["valor_1"], cond["valor_2"]
                )
            else:
                valor_actual = cat_vals.get(var_condicion)
                cumple = evaluar_categorica_condicion(valor_actual, operador, cond["valor_1"])

            if not cumple:
                cumple_todas = False
                break

        if cumple_todas:
            cat_vals[var_resultado] = regla["valor_resultado"]

    return cat_vals


def evaluar_caracteristica_categorica(desc_clean: str, var_name: str, maestro) -> str:
    """
    Evalúa las reglas ordenadas por prioridad para una variable categórica,
    ignorando coincidencias que estén negadas (ej: 'NO ONLINE', 'SIN BATERIA').
    """
    if not desc_clean or var_name not in maestro.dict_caracteristicas:
        return None

    # Recorre las reglas precompiladas y ordenadas por prioridad
    for regex_comp, valor_resultado in maestro.dict_caracteristicas[var_name]:
        # Busca todas las coincidencias del patrón en la descripción
        for match in regex_comp.finditer(desc_clean):
            pos_inicio = match.start()
            
            # Si la coincidencia NO está negada, es un match válido
            if not tiene_negacion_previa(desc_clean, pos_inicio):
                return valor_resultado

    return None


def extraer_potencia_numerica(desc_clean: str, var_name: str, maestro) -> float:
    """
    Extrae un valor numérico para una variable técnica (kVA, kW, V, A)
    aplicando normalización de decimales, rangos y multiplicadores.
    """
    if not desc_clean or var_name not in maestro.variables_potencia:
        return None

    # Recorre las reglas numéricas precompiladas para esta variable
    for regex_comp, multiplicador in maestro.dict_potencia.get(var_name, []):
        match = regex_comp.search(desc_clean)
        if match:
            # El grupo 1 siempre captura la cifra/rango en el Regex
            cifra_raw = match.group(1)
            num_normalizado = normalizar_numero_extraido(cifra_raw)
            
            if num_normalizado is not None:
                # Aplica multiplicador (ej. 0.001 si viene en W y se requiere en kW)
                return round(num_normalizado * multiplicador, 2)

    return None