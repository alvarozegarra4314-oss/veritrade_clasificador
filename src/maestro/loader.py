import re
import pandas as pd
from src.texto_utils import (
    limpiar_texto,
    parse_bool,
    parse_float,
    construir_patron_desde_palabras,
)

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
        self.condicionales = []
        self.condicionales_descartadas = []  # auditoría: reglas ignoradas por no aplicar a esta línea

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
            s_condicionales = self._buscar_hoja(sheets, ["Condicionales", "5_Condicionales"])

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

            # 7. Condicionales por rango (5_Condicionales) — reglas que
            # clasifican una variable categórica según el valor de OTRA
            # variable ya extraída (típicamente numérica, ej. rangos de
            # kVA), no por palabras clave en el texto.
            if s_condicionales:
                df_cond = pd.read_excel(xls, sheet_name=s_condicionales)
                df_cond.columns = [str(c).strip() for c in df_cond.columns]

                columnas_requeridas = {
                    "Regla_ID", "Variable_Resultado", "Valor_Resultado",
                    "Variable_Condicion", "Operador",
                }
                if columnas_requeridas.issubset(set(df_cond.columns)):
                    variables_conocidas = set(self.variables_categoricas) | set(self.variables_potencia)
                    col_prio = "Prioridad" if "Prioridad" in df_cond.columns else None

                    for regla_id, grupo in df_cond.groupby("Regla_ID"):
                        primera = grupo.iloc[0]
                        variable_resultado = str(primera["Variable_Resultado"]).strip()
                        valor_resultado = str(primera["Valor_Resultado"]).strip()
                        prioridad = parse_float(primera.get(col_prio, 1), default=1.0) if col_prio else 1.0

                        # SEGURIDAD DE SCOPING: si la variable a clasificar no
                        # pertenece a esta línea de producto, la regla completa
                        # se descarta (nunca se aplica una condicional de otra
                        # línea, ej. UPS, sobre un maestro de Interruptores).
                        if variable_resultado not in variables_conocidas:
                            self.condicionales_descartadas.append({
                                "Regla_ID": regla_id,
                                "Variable_Resultado": variable_resultado,
                                "Motivo": "Variable_Resultado no existe en esta línea de producto",
                            })
                            continue

                        condiciones = []
                        regla_valida = True
                        for _, fila in grupo.iterrows():
                            var_condicion = str(fila["Variable_Condicion"]).strip()

                            # Misma seguridad de scoping, ahora sobre cada
                            # condición individual de la regla.
                            if var_condicion not in variables_conocidas:
                                self.condicionales_descartadas.append({
                                    "Regla_ID": regla_id,
                                    "Variable_Resultado": variable_resultado,
                                    "Motivo": f"Variable_Condicion '{var_condicion}' no existe en esta línea de producto",
                                })
                                regla_valida = False
                                break

                            operador = str(fila["Operador"]).strip().upper()
                            valor_1 = parse_float(fila.get("Valor_1"), default=None)
                            valor_2 = parse_float(fila.get("Valor_2"), default=None)
                            # Para operadores de igualdad/desigualdad el valor puede
                            # ser texto (ej. Tipo_Tecnologia == Online), no numérico.
                            if operador in ("==", "!="):
                                valor_1 = str(fila.get("Valor_1")).strip() if pd.notna(fila.get("Valor_1")) else None

                            condiciones.append({
                                "variable": var_condicion,
                                "operador": operador,
                                "valor_1": valor_1,
                                "valor_2": valor_2,
                                "es_numerica": var_condicion in self.variables_potencia,
                            })

                        if regla_valida and condiciones:
                            self.condicionales.append({
                                "regla_id": regla_id,
                                "variable_resultado": variable_resultado,
                                "valor_resultado": valor_resultado,
                                "prioridad": prioridad,
                                "condiciones": condiciones,
                            })

                    self.condicionales.sort(key=lambda r: r["prioridad"])

    @property
    def variable_producto_principal(self) -> str:
        return self.config_linea.get("VARIABLE_PRODUCTO_PRINCIPAL", "Tipo_Producto_Detallado")

    @property
    def valor_producto_principal(self) -> str:
        return self.config_linea.get("VALOR_PRODUCTO_PRINCIPAL", "")