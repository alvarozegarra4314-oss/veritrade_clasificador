import sys
import re
import unicodedata
from pathlib import Path
import pandas as pd
import numpy as np


def limpiar_texto(texto):
    if not texto or pd.isna(texto):
        return ""
    texto_str = str(texto)
    nfkd_form = unicodedata.normalize('NFKD', texto_str)
    sin_tildes = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return sin_tildes.upper().strip()


def parse_bool(val):
    """Convierte texto o números a booleano de forma segura en Python."""
    if pd.isna(val):
        return False
    if isinstance(val, bool):
        return val
    val_str = str(val).strip().upper()
    return val_str in ['TRUE', 'VERDADERO', '1', 'T', 'SI', 'YES']


def parse_float(val, default=1.0):
    """Convierte un multiplicador a float de forma segura sin retornar NaN."""
    if pd.isna(val):
        return default
    try:
        f = float(val)
        return f if not np.isnan(f) else default
    except (ValueError, TypeError):
        return default


def construir_patron_desde_palabras(lista_palabras):
    partes = []
    for palabra in lista_palabras:
        palabra_limpia = limpiar_texto(str(palabra))
        if not palabra_limpia:
            continue
        escapada = re.escape(palabra_limpia).replace(r'\ ', r'\s+')
        partes.append(escapada)
    return '|'.join(partes)


def identificar_columnas_descripcion(df_columns):
    """Detecta de forma dinámica las columnas que contienen descripciones comerciales."""
    cols_desc = []
    for col in df_columns:
        c_upper = str(col).strip().upper()
        if any(k in c_upper for k in ['DESCRIPCION', 'DESC_', 'DESC ', 'DETALLE', 'MERCADERIA', 'COMMODITY']):
            cols_desc.append(col)
    
    # Si no se detectó ninguna por nombre, buscar columnas de tipo texto
    if not cols_desc:
        cols_desc = list(df_columns)
        
    return cols_desc


class CargarMaestro:
    def __init__(self, ruta_excel):
        self.ruta_excel = ruta_excel
        self.config_linea = {}
        self.lista_marcas = []
        self.stopwords = set()
        self.patrones_regex = []
        self.dict_defaults = {True: "Marca Principal", False: "Marca Componentes"}
        self.dict_caracteristicas = {}
        self.dict_potencia = {}
        self.variables_categoricas = []
        self.variables_potencia = []

        self._cargar_y_precompilar()

    def _buscar_hoja(self, sheet_names, palabras_clave):
        for s in sheet_names:
            for kw in palabras_clave:
                if kw.lower() in s.lower():
                    return s
        return None

    def _cargar_y_precompilar(self):
        if hasattr(self.ruta_excel, 'seek'):
            self.ruta_excel.seek(0)

        with pd.ExcelFile(self.ruta_excel) as xls:
            sheets = xls.sheet_names

            s_cfg = self._buscar_hoja(sheets, ["Config_Linea", "0b_Config", "Config"])
            s_marcas = self._buscar_hoja(sheets, ["Maestro_Marcas", "1_Marcas", "Marcas"])
            s_stopwords = self._buscar_hoja(sheets, ["Stopwords", "1b_Palabras_Ignorar", "Palabras_Ignorar", "Ignorar"])
            s_defaults = self._buscar_hoja(sheets, ["Marcas_Default", "1c_Marca_Por_Defecto", "Marca_Por_Defecto", "Default"])
            s_carac = self._buscar_hoja(sheets, ["Reglas_Caracteristicas", "2_Caracteristicas", "Caracteristicas"])
            s_pot = self._buscar_hoja(sheets, ["Patrones_Potencia", "3_Tecnico_Potencia", "Tecnico_Potencia", "Potencia"])
            s_regex = self._buscar_hoja(sheets, ["Patrones_Regex", "4_Tecnico_RegexMarca", "RegexMarca", "Regex"])

            # 0. Config Linea
            if s_cfg:
                df_cfg = pd.read_excel(xls, sheet_name=s_cfg)
                df_cfg.columns = [str(c).strip() for c in df_cfg.columns]
                col_p = [c for c in df_cfg.columns if 'PARAMETRO' in c.upper() or 'PARAM' in c.upper()][0]
                col_v = [c for c in df_cfg.columns if 'VALOR' in c.upper() or 'VAL' in c.upper()][0]
                self.config_linea = dict(zip(df_cfg[col_p].astype(str).str.strip(), df_cfg[col_v].astype(str).str.strip()))

            # 1. Marcas
            if s_marcas:
                df_marcas = pd.read_excel(xls, sheet_name=s_marcas)
                df_marcas.columns = [str(c).strip() for c in df_marcas.columns]
                col_pat = [c for c in df_marcas.columns if 'PATRON' in c.upper() or 'BUSQUEDA' in c.upper()][0]
                col_est = [c for c in df_marcas.columns if 'MARCA' in c.upper() or 'ESTANDAR' in c.upper()][0]
                
                if 'Prioridad' in df_marcas.columns:
                    df_marcas = df_marcas.sort_values(by='Prioridad')
                    
                for _, row in df_marcas.iterrows():
                    patron = limpiar_texto(str(row.get(col_pat, '')))
                    estandar = row.get(col_est, '')
                    if patron and pd.notna(estandar):
                        self.lista_marcas.append((patron, str(estandar).strip()))

            # 2. Stopwords
            if s_stopwords:
                df_sw = pd.read_excel(xls, sheet_name=s_stopwords)
                col_sw = df_sw.columns[0]
                self.stopwords = set(limpiar_texto(str(x)) for x in df_sw[col_sw].dropna())

            # 3. Regex Marcas
            if s_regex:
                df_reg = pd.read_excel(xls, sheet_name=s_regex)
                df_reg.columns = [str(c).strip() for c in df_reg.columns]
                if 'Orden_Prioridad' in df_reg.columns:
                    df_reg = df_reg.sort_values(by='Orden_Prioridad')
                cols_pat = [c for c in df_reg.columns if any(k in c.lower() for k in ['pattern', 'patron', 'regex'])]
                col_pat = cols_pat[0] if cols_pat else df_reg.columns[-1]
                for pat in df_reg[col_pat].dropna().astype(str):
                    try:
                        self.patrones_regex.append(re.compile(pat, re.IGNORECASE))
                    except re.error:
                        continue

            # 4. Defaults
            if s_defaults:
                df_def = pd.read_excel(xls, sheet_name=s_defaults)
                col_bool = df_def.columns[0]
                col_val = df_def.columns[1]
                for _, row in df_def.iterrows():
                    key = parse_bool(row[col_bool])
                    self.dict_defaults[key] = str(row[col_val]).strip()

            # 5. Características Categóricas
            if s_carac:
                df_carac = pd.read_excel(xls, sheet_name=s_carac)
                df_carac.columns = [str(c).strip() for c in df_carac.columns]
                col_kw = [c for c in df_carac.columns if 'PALABRA' in c.upper() or 'CLAVE' in c.upper()][0]
                df_carac = df_carac.dropna(subset=[col_kw])
                
                prio_col = 'Prioridad' if 'Prioridad' in df_carac.columns else 'Orden_Prioridad'
                if prio_col not in df_carac.columns:
                    df_carac[prio_col] = 1

                agrupado = (
                    df_carac
                    .groupby(['Variable', 'Valor_Resultado', prio_col])[col_kw]
                    .apply(list)
                    .reset_index()
                )
                agrupado['Patron_Busqueda'] = agrupado[col_kw].apply(construir_patron_desde_palabras)
                agrupado = agrupado.sort_values(by=['Variable', prio_col])

                for var, group in agrupado.groupby('Variable'):
                    self.variables_categoricas.append(var)
                    reglas_var = []
                    for _, fila in group.iterrows():
                        patron_str = str(fila['Patron_Busqueda']).strip()
                        resultado = fila['Valor_Resultado']
                        if patron_str:
                            try:
                                regex_compilado = re.compile(fr'(?:^|(?<=\W))({patron_str})(?:$|(?=\W))', re.IGNORECASE)
                                reglas_var.append((regex_compilado, resultado))
                            except re.error:
                                continue
                    self.dict_caracteristicas[var] = reglas_var

            # 6. Potencia y Métricas Numéricas
            if s_pot:
                df_pot = pd.read_excel(xls, sheet_name=s_pot)
                df_pot.columns = [str(c).strip() for c in df_pot.columns]
                if 'Orden_Prioridad' in df_pot.columns:
                    df_pot = df_pot.sort_values(by=['Variable', 'Orden_Prioridad'])

                col_pat_pot = [c for c in df_pot.columns if 'PATTRN' in c.upper() or 'PATRON' in c.upper() or 'REGEX' in c.upper()][0]
                col_mult_pot = [c for c in df_pot.columns if 'MULT' in c.upper() or 'KVA' in c.upper()]
                col_mult = col_mult_pot[0] if col_mult_pot else None

                for var, group in df_pot.groupby('Variable'):
                    self.variables_potencia.append(var)
                    patrones_var = []
                    for _, fila in group.iterrows():
                        patron_str = str(fila[col_pat_pot]).strip()
                        mult = parse_float(fila.get(col_mult, 1.0)) if col_mult else 1.0
                        if patron_str:
                            try:
                                regex_compilado = re.compile(patron_str, re.IGNORECASE)
                                patrones_var.append((regex_compilado, mult))
                            except re.error:
                                continue
                    self.dict_potencia[var] = patrones_var

    @property
    def variable_producto_principal(self) -> str:
        return self.config_linea.get("VARIABLE_PRODUCTO_PRINCIPAL", "Tipo_Producto_Detallado")

    @property
    def valor_producto_principal(self) -> str:
        return self.config_linea.get("VALOR_PRODUCTO_PRINCIPAL", "")


def extraer_marca(desc, maestro):
    desc_clean = limpiar_texto(desc)
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


def evaluar_caracteristica_categorica_opt(desc, var_name, maestro):
    desc_clean = limpiar_texto(desc)
    if not desc_clean:
        return None
    reglas = maestro.dict_caracteristicas.get(var_name, [])
    for regex_comp, resultado in reglas:
        if regex_comp.search(desc_clean):
            return resultado
    return None


def extraer_potencia_numerica_opt(desc, var_name, maestro):
    desc_clean = limpiar_texto(desc)
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


def procesar_dataframe_dinamico(df_raw, ruta_maestro):
    maestro = CargarMaestro(ruta_maestro)
    
    # Detección dinámica de las columnas de descripción
    cols_desc = identificar_columnas_descripcion(df_raw.columns)
    
    resultados = []
    for _, row in df_raw.iterrows():
        # Combinar dinámicamente todo el texto descriptivo disponible en la fila
        textos_desc = [str(row[c]) for c in cols_desc if pd.notna(row[c])]
        desc_completa = " ".join(textos_desc)

        # 1. Extracción de marca
        marca, fuente = extraer_marca(desc_completa, maestro)

        # 2. Extracción de características categóricas
        cat_vals = {
            var: evaluar_caracteristica_categorica_opt(desc_completa, var, maestro)
            for var in maestro.variables_categoricas
        }

        # 3. Extracción de métricas numéricas
        num_vals = {
            var: extraer_potencia_numerica_opt(desc_completa, var, maestro)
            for var in maestro.variables_potencia
        }

        var_principal = maestro.variable_producto_principal
        valor_principal = maestro.valor_producto_principal

        val_principal_extracted = cat_vals.get(var_principal)
        
        if valor_principal:
            es_principal = (val_principal_extracted == valor_principal)
        else:
            es_principal = bool(val_principal_extracted)

        marca_final = marca or maestro.dict_defaults.get(es_principal, "Marca Generica")
        fuente_marca = fuente if marca else "Default"

        # Regla especial de tecnología UPS
        if "Tipo_Tecnologia" in cat_vals and "Salida_Fases" in cat_vals:
            if cat_vals["Tipo_Tecnologia"] == "Interactivo" and cat_vals["Salida_Fases"] == "Trifasico":
                cat_vals["Tipo_Tecnologia"] = "Online"

        # Construcción del diccionario de resultados
        res = {
            "Producto_Declarado": val_principal_extracted,
            "Marca_Declarada": marca,
            var_principal: val_principal_extracted,
            "Es_Producto_Principal": es_principal,
            "Marca_Extraida": marca_final,
            "Origen_Marca": fuente_marca,
        }

        for var in maestro.variables_categoricas:
            if var != var_principal:
                res[var] = cat_vals.get(var)

        for var in maestro.variables_potencia:
            res[var] = num_vals.get(var)

        resultados.append(res)

    df_res = pd.DataFrame(resultados)
    df_final = pd.concat([df_raw.reset_index(drop=True), df_res.reset_index(drop=True)], axis=1)
    return df_final


def sanitizar_dataframe_para_excel(df):
    """Limpia caracteres de control XML e invisibles que corrompen archivos Excel en Linux/Cloud."""
    df_clean = df.copy()
    for col in df_clean.columns:
        if df_clean[col].dtype == 'object':
            df_clean[col] = (
                df_clean[col]
                .astype(str)
                .str.replace(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]', '', regex=True)
                .str.replace(r'_x[0-9a-fA-F]{4}_', '', regex=True)
            )
            df_clean[col] = df_clean[col].replace('nan', '')
    return df_clean