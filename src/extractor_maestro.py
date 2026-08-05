import re
import unicodedata
import pandas as pd
from config import (
    PATH_MAESTRO,
    SHEET_CONFIG_LINEA,
    SHEET_MAESTRO_MARCAS,
    SHEET_STOPWORDS,
    SHEET_PATRONES_REGEX,
    SHEET_MARCAS_DEFAULT,
    SHEET_REGLAS_CARACTERISTICAS,
    SHEET_PATRONES_POTENCIA,
    COL_DESCRIPCION
)

VALORES_MARCA_NULA = {'S/M', 'SIN MARCA', 'NO INDICA', 'N/A', 'NONE', 'S/N', '_X0000_'}

# ==========================================
# FUNCIONES AUXILIARES Y PARSERS
# ==========================================

def limpiar_texto(texto):
    """
    Normaliza el texto eliminando tildes, convirtiéndolo a mayúsculas
    y limpiando espacios sobrantes.
    """
    if not texto or pd.isna(texto):
        return ""
    texto_str = str(texto)
    nfkd_form = unicodedata.normalize('NFKD', texto_str)
    sin_tildes = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return sin_tildes.upper().strip()


def construir_patron_desde_palabras(lista_palabras):
    partes = []
    for palabra in lista_palabras:
        palabra_limpia = limpiar_texto(str(palabra))
        if not palabra_limpia:
            continue
        escapada = re.escape(palabra_limpia).replace(r'\ ', r'\s+')
        partes.append(escapada)
    return '|'.join(partes)


def parsear_descripcion1(texto_raw):
    """
    Descompone Descripcion1 por comas según la convención de la DUA/DAM.
    """
    if not texto_raw or pd.isna(texto_raw):
        return {"producto": "", "marca": "", "modelo": "", "producto_limpio": "", "marca_limpia": ""}

    partes = [p.strip() for p in str(texto_raw).split(',')]
    prod = partes[0] if len(partes) > 0 else ""
    marca = partes[1] if len(partes) > 1 else ""
    modelo = ", ".join(partes[2:]) if len(partes) > 2 else ""

    return {
        "producto": prod,
        "marca": marca,
        "modelo": modelo,
        "producto_limpio": limpiar_texto(prod),
        "marca_limpia": limpiar_texto(marca)
    }


# ==========================================
# CARGADOR CENTRALIZADO Y PRE-COMPILADO
# ==========================================

class CargarMaestro:
    """
    Carga el Excel maestro de una línea de producto (UPS, Tableros, etc.)
    y pre-compila todas las estructuras de datos en listas y diccionarios
    nativos de Python con RegEx pre-compilados para máxima velocidad.
    El formato de hojas es el mismo para cualquier línea -- lo que cambia
    entre líneas es el contenido, no la estructura.
    """
    def __init__(self, ruta_excel=PATH_MAESTRO):
        self.ruta_excel = ruta_excel
        self.config_linea = {}
        self.lista_marcas = []           # List[tuple(patron_limpio, marca_estandar)]
        self.stopwords = set()          # Set[str]
        self.patrones_regex = []        # List[re.Pattern]
        self.dict_defaults = {}         # Dict[bool, str]
        self.dict_caracteristicas = {}  # Dict[variable, List[tuple(compiled_regex, valor_resultado)]]
        self.dict_potencia = {}         # Dict[variable, List[tuple(compiled_regex, multiplicador)]]

        self._cargar_y_precompilar()

    def _cargar_y_precompilar(self):
        # Intentar usar calamine si está disponible para lectura ultrarrápida del maestro
        engine = None
        try:
            import python_calamine
            engine = "calamine"
        except ImportError:
            engine = "openpyxl"

        # 0. Config de la línea de producto (variable/valor que define el
        #    "producto principal completo"). Si el maestro no tiene esta
        #    hoja todavía, se usan valores por defecto y se avisa.
        try:
            df_cfg = pd.read_excel(self.ruta_excel, sheet_name=SHEET_CONFIG_LINEA, engine=engine)
            self.config_linea = dict(zip(df_cfg['Parametro'], df_cfg['Valor']))
        except (ValueError, KeyError):
            print(
                f"⚠️  No encontré la hoja '{SHEET_CONFIG_LINEA}' en {self.ruta_excel}. "
                "Usando valores por defecto (Tipo_Producto_Detallado == 'UPS Sistema Completo')."
            )
            self.config_linea = {
                "VARIABLE_PRODUCTO_PRINCIPAL": "Tipo_Producto_Detallado",
                "VALOR_PRODUCTO_PRINCIPAL": "UPS Sistema Completo",
            }

        # 1. Maestro de Marcas
        df_marcas = pd.read_excel(self.ruta_excel, sheet_name=SHEET_MAESTRO_MARCAS, engine=engine)
        if 'Prioridad' in df_marcas.columns:
            df_marcas = df_marcas.sort_values(by='Prioridad')

        for _, row in df_marcas.iterrows():
            patron = limpiar_texto(str(row.get('Patron_Busqueda', '')))
            estandar = row.get('Marca_Estandar', '')
            if patron and pd.notna(estandar):
                self.lista_marcas.append((patron, str(estandar)))

        # 2. Stopwords
        df_stopwords = pd.read_excel(self.ruta_excel, sheet_name=SHEET_STOPWORDS, engine=engine)
        self.stopwords = set(
            limpiar_texto(str(x)) for x in df_stopwords['Palabra_Ignorar'].dropna()
        )

        # 3. Patrones Regex Marcas
        df_regex = pd.read_excel(self.ruta_excel, sheet_name=SHEET_PATRONES_REGEX, engine=engine)
        df_regex = df_regex.sort_values(by='Orden_Prioridad')
        for pat in df_regex['Pattern_Regex'].dropna().astype(str):
            try:
                self.patrones_regex.append(re.compile(pat, re.IGNORECASE))
            except re.error:
                continue

        # 4. Marca por defecto según si la fila es o no el "producto
        #    principal" (columna 'Es_Producto_Principal', antes 'Es_UPS').
        df_defaults = pd.read_excel(self.ruta_excel, sheet_name=SHEET_MARCAS_DEFAULT, engine=engine)
        col_bool = df_defaults.columns[0]
        for _, row in df_defaults.iterrows():
            key = bool(row[col_bool])
            self.dict_defaults[key] = str(row['Marca_Default'])

        # 5. Reglas de Características (100% genérico: cualquier Variable
        #    que exista en la hoja, para cualquier línea de producto)
        df_carac_raw = pd.read_excel(self.ruta_excel, sheet_name=SHEET_REGLAS_CARACTERISTICAS, engine=engine)
        df_carac_raw = df_carac_raw.dropna(subset=['Palabra_Clave'])

        agrupado = (
            df_carac_raw
            .groupby(['Variable', 'Valor_Resultado', 'Prioridad'])['Palabra_Clave']
            .apply(list)
            .reset_index()
        )
        agrupado['Patron_Busqueda'] = agrupado['Palabra_Clave'].apply(construir_patron_desde_palabras)
        agrupado = agrupado.rename(columns={'Prioridad': 'Orden_Prioridad'}).sort_values(by=['Variable', 'Orden_Prioridad'])

        for variable, group in agrupado.groupby('Variable'):
            reglas_var = []
            for _, fila in group.iterrows():
                patron_str = str(fila['Patron_Busqueda']).strip()
                resultado = fila['Valor_Resultado']
                if patron_str:
                    try:
                        regex_compilado = re.compile(fr'\b({patron_str})\b', re.IGNORECASE)
                        reglas_var.append((regex_compilado, resultado))
                    except re.error:
                        continue
            self.dict_caracteristicas[variable] = reglas_var

        # 6. Patrones de atributos numéricos (potencia, o cualquier otro
        #    atributo numérico que otra línea necesite extraer)
        df_pot = pd.read_excel(self.ruta_excel, sheet_name=SHEET_PATRONES_POTENCIA, engine=engine)
        df_pot = df_pot.sort_values(by=['Variable', 'Orden_Prioridad'])

        for variable, group in df_pot.groupby('Variable'):
            patrones_var = []
            for _, fila in group.iterrows():
                patron_str = str(fila['Pattern_Regex']).strip()
                mult = float(fila.get('Multiplicador_kVA', 1.0))
                if patron_str:
                    try:
                        regex_compilado = re.compile(patron_str, re.IGNORECASE)
                        patrones_var.append((regex_compilado, mult))
                    except re.error:
                        continue
            self.dict_potencia[variable] = patrones_var

    @property
    def variable_producto_principal(self) -> str:
        return self.config_linea.get("VARIABLE_PRODUCTO_PRINCIPAL", "Tipo_Producto_Detallado")

    @property
    def valor_producto_principal(self) -> str:
        return self.config_linea.get("VALOR_PRODUCTO_PRINCIPAL", "UPS Sistema Completo")


# Alias por compatibilidad con código/scripts viejos que todavía
# importen el nombre anterior de la clase.
CargarMaestroUPS = CargarMaestro


# RegEx genérico pre-compilado para respaldo de potencia
REGEX_RESPALDO_VA = re.compile(r'\b(\d+[\.,]?\d*)\s*(VA|KVA|KW)\b', re.IGNORECASE)


# ==========================================
# FUNCIONES DE EVALUACIÓN OPTIMIZADAS
# ==========================================

def evaluar_caracteristica_categorica_opt(texto_prep, variable, maestro):
    if not texto_prep:
        return None

    reglas_var = maestro.dict_caracteristicas.get(variable, [])

    # Búsqueda por palabras clave pre-compiladas
    for regex_compilado, resultado in reglas_var:
        if regex_compilado.search(texto_prep):
            return resultado

    # Respaldo por Potencia (Específica para Salida_Fases)
    if variable == 'Salida_Fases':
        pass  # sin cambios: el respaldo real vive en extraer_potencia_numerica_opt / regla de negocio

    return None


def extraer_potencia_numerica_opt(texto_prep, variable, maestro):
    if not texto_prep:
        return None

    patrones_var = maestro.dict_potencia.get(variable, [])
    for regex_compilado, multiplicador in patrones_var:
        match = regex_compilado.search(texto_prep)
        if match:
            try:
                valor_str = match.group(1).replace(',', '.')
                return round(float(valor_str) * multiplicador, 2)
            except (ValueError, IndexError):
                continue

    # Respaldo nativo (VA, KVA, KW)
    if variable == 'Potencia_kVA':
        match_va = REGEX_RESPALDO_VA.search(texto_prep)
        if match_va:
            try:
                val = float(match_va.group(1).replace(',', '.'))
                unidad = match_va.group(2).upper()
                if unidad == 'VA':
                    return round(val / 1000.0, 2)
                return round(val, 2)
            except ValueError:
                pass

    return None


def buscar_en_maestro_opt(texto_prep, lista_marcas):
    if not texto_prep:
        return None
    for patron, marca_estandar in lista_marcas:
        if patron in texto_prep:
            return marca_estandar
    return None


def buscar_marca_regex_opt(texto_prep, patrones_regex, stopwords):
    if not texto_prep:
        return None
    for regex_compilado in patrones_regex:
        match = regex_compilado.search(texto_prep)
        if match:
            candidato = match.group(1).strip()
            if candidato not in stopwords and len(candidato) > 2:
                return candidato
    return None


# ==========================================
# PROCESAMIENTO POR FILA OPTIMIZADO
# ==========================================

def procesar_dict_fila(row_dict, maestro: CargarMaestro):
    desc1_raw = row_dict.get('Descripcion1', row_dict.get(COL_DESCRIPCION, ''))
    p1 = parsear_descripcion1(desc1_raw)

    desc1_clean = limpiar_texto(desc1_raw).replace(',', '.')

    # ------------------------------------------------------------------
    # PASO 1 (PRIORIDAD 1): Análisis sobre Descripcion1
    # ------------------------------------------------------------------
    marca_p1 = None
    fuente_marca = None
    seg2_marca = p1["marca_limpia"]

    if seg2_marca:
        if seg2_marca in VALORES_MARCA_NULA:
            marca_p1 = seg2_marca
            fuente_marca = "P1: Declarado Sin Marca / Nula"
        else:
            m_match = buscar_en_maestro_opt(seg2_marca, maestro.lista_marcas)
            if m_match:
                marca_p1 = m_match
                fuente_marca = "P1: Descripcion1 (Estandarizado)"
            else:
                marca_p1 = p1["marca"].strip().upper()
                fuente_marca = "P1: Descripcion1 (Marca Directa)"

    # Características Paso 1. La variable que define "producto principal"
    # (VARIABLE_PRODUCTO_PRINCIPAL, por defecto Tipo_Producto_Detallado)
    # se lee del maestro, no está escrita a mano -- así sirve para
    # cualquier línea de producto, no solo UPS.
    var_principal = maestro.variable_producto_principal
    valor_principal = maestro.valor_producto_principal

    tipo_prod_p1 = evaluar_caracteristica_categorica_opt(p1["producto_limpio"], var_principal, maestro)
    if not tipo_prod_p1:
        tipo_prod_p1 = evaluar_caracteristica_categorica_opt(desc1_clean, var_principal, maestro)

    fases_p1 = evaluar_caracteristica_categorica_opt(desc1_clean, 'Salida_Fases', maestro)
    montaje_p1 = evaluar_caracteristica_categorica_opt(desc1_clean, 'Formato_Montaje', maestro)
    tecno_p1 = evaluar_caracteristica_categorica_opt(desc1_clean, 'Tipo_Tecnologia', maestro)
    pot_kva_p1 = extraer_potencia_numerica_opt(desc1_clean, 'Potencia_kVA', maestro)
    pot_w_p1 = extraer_potencia_numerica_opt(desc1_clean, 'Potencia_Watts', maestro)

    # ------------------------------------------------------------------
    # PASO 2 (FALLBACK): Solo si Descripcion1 NO traía campo de marca
    # ------------------------------------------------------------------
    descs_secundarias_raw = " ".join([
        str(row_dict.get(f'Descripcion{i}', '') or '') for i in range(2, 6)
    ])
    descs_secundarias_clean = limpiar_texto(descs_secundarias_raw).replace(',', '.')

    embarcador_raw = str(row_dict.get('Naviera', row_dict.get('Embarcador / Exportador', '')) or '')
    embarcador_clean = limpiar_texto(embarcador_raw)

    texto_completo_clean = f"{desc1_clean} {descs_secundarias_clean} {embarcador_clean}".strip()

    tipo_prod_final = (
        tipo_prod_p1
        or evaluar_caracteristica_categorica_opt(descs_secundarias_clean, var_principal, maestro)
        or "Otros Equipos / Accesorios"
    )
    es_principal_final = (tipo_prod_final == valor_principal)

    marca_final = marca_p1
    if not marca_final:
        marca_emb = buscar_en_maestro_opt(embarcador_clean, maestro.lista_marcas)
        if marca_emb:
            marca_final = marca_emb
            fuente_marca = "P2: Embarcador / Naviera"
        else:
            marca_sec = buscar_en_maestro_opt(descs_secundarias_clean, maestro.lista_marcas)
            if marca_sec:
                marca_final = marca_sec
                fuente_marca = "P2: Maestro en Descripciones 2-5"
            else:
                marca_regex = buscar_marca_regex_opt(texto_completo_clean, maestro.patrones_regex, maestro.stopwords)
                if marca_regex:
                    marca_final = f"{marca_regex.upper()} (RegEx)"
                    fuente_marca = "P2: RegEx en Texto Completo"
                else:
                    marca_final = maestro.dict_defaults.get(es_principal_final, "Marca Generica")
                    fuente_marca = "P2: Marca por Defecto"

    fases_final = fases_p1 or evaluar_caracteristica_categorica_opt(descs_secundarias_clean, 'Salida_Fases', maestro)
    montaje_final = montaje_p1 or evaluar_caracteristica_categorica_opt(descs_secundarias_clean, 'Formato_Montaje', maestro)
    tecno_detectada = tecno_p1 or evaluar_caracteristica_categorica_opt(descs_secundarias_clean, 'Tipo_Tecnologia', maestro)

    # ------------------------------------------------------------------
    # REGLA DE NEGOCIO: TECNOLOGÍA SEGÚN FASES (Categorías del Maestro)
    # ------------------------------------------------------------------
    TECNO_ONLINE = "On-Line Doble Conversión"
    TECNO_INTERACTIVO = "Interactivo (Line-Interactive)"

    es_online = bool(tecno_detectada and any(kw in str(tecno_detectada).upper() for kw in ["ONLINE", "ON-LINE", "DOBLE CONVERSION", "DOBLE CONVERSIÓN"]))

    if fases_final == 'Trifásico':
        tecno_final = TECNO_ONLINE
    elif fases_final == 'Monofásico':
        tecno_final = TECNO_ONLINE if es_online else TECNO_INTERACTIVO
    else:
        if es_online:
            tecno_final = TECNO_ONLINE
        elif tecno_detectada:
            tecno_final = tecno_detectada
        else:
            tecno_final = None

    pot_kva_final = pot_kva_p1 or extraer_potencia_numerica_opt(descs_secundarias_clean, 'Potencia_kVA', maestro)
    pot_w_final = pot_w_p1 or extraer_potencia_numerica_opt(descs_secundarias_clean, 'Potencia_Watts', maestro)

    return {
        "Producto_Declarado": p1["producto"],
        "Marca_Declarada": p1["marca"],
        "Modelo_Declarado": p1["modelo"],
        var_principal: tipo_prod_final,
        "Es_Producto_Principal": es_principal_final,
        "Marca_Extraida": marca_final,
        "Origen_Marca": fuente_marca,
        "Salida_Fases": fases_final,
        "Formato_Montaje": montaje_final,
        "Tipo_Tecnologia": tecno_final,
        "Potencia_kVA": pot_kva_final,
        "Potencia_Watts": pot_w_final
    }


# ==========================================
# PROCESO PRINCIPAL
# ==========================================

def procesar_dataframe_dinamico(df_raw, ruta_maestro=PATH_MAESTRO):
    df = df_raw.copy()

    maestro = CargarMaestro(ruta_excel=ruta_maestro)

    registros = df.to_dict('records')
    resultados = [procesar_dict_fila(row, maestro) for row in registros]

    df_res = pd.DataFrame(resultados, index=df.index)
    for col in df_res.columns:
        df[col] = df_res[col]

    return df
