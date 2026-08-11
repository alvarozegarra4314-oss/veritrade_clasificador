import pandas as pd
from src.texto_utils import limpiar_texto, identificar_columnas_descripcion
from src.maestro.loader import CargarMaestro
from src.maestro.reglas import (
    extraer_marca,
    evaluar_caracteristica_categorica,
    extraer_potencia_numerica,
)


def procesar_dataframe_dinamico(df_raw: pd.DataFrame, ruta_maestro) -> pd.DataFrame:
    maestro = ruta_maestro if isinstance(ruta_maestro, CargarMaestro) else CargarMaestro(ruta_maestro)
    
    cols_desc = identificar_columnas_descripcion(df_raw.columns)
    cols_indices = [df_raw.columns.get_loc(c) for c in cols_desc]
    
    var_principal = maestro.variable_producto_principal
    valor_principal = maestro.valor_producto_principal
    variables_cat = maestro.variables_categoricas
    variables_pot = maestro.variables_potencia

    resultados = []

    for row in df_raw.itertuples(index=False):
        # 1. Unir todas las columnas de descripción para características y marcas por diccionario
        textos_desc = [str(row[i]) for i in cols_indices if pd.notna(row[i])]
        desc_completa = " ".join(textos_desc)
        desc_clean = limpiar_texto(desc_completa)

        # 2. Aísla únicamente la PRIMERA descripción (Descripcion 1 / Descripcion Comercial)
        desc_1_raw = str(row[cols_indices[0]]) if cols_indices and pd.notna(row[cols_indices[0]]) else ""
        desc_1_clean = limpiar_texto(desc_1_raw)

        # 3. Extraer marca pasando desc_clean (para dict/regex) y desc_1_clean (para posición 2 por coma)
        marca, fuente = extraer_marca(desc_clean, maestro, desc_1_clean=desc_1_clean)

        cat_vals = {
            var: evaluar_caracteristica_categorica(desc_clean, var, maestro)
            for var in variables_cat
        }

        num_vals = {
            var: extraer_potencia_numerica(desc_clean, var, maestro)
            for var in variables_pot
        }

        val_principal_extracted = cat_vals.get(var_principal)
        
        if valor_principal:
            es_principal = (val_principal_extracted == valor_principal)
        else:
            es_principal = bool(val_principal_extracted)

        marca_final = marca or maestro.dict_defaults.get(es_principal, "Marca Generica")
        fuente_marca = fuente if marca else "Default"

        if cat_vals.get("Tipo_Tecnologia") == "Interactivo" and cat_vals.get("Salida_Fases") == "Trifasico":
            cat_vals["Tipo_Tecnologia"] = "Online"

        res = {
            "Producto_Declarado": val_principal_extracted,
            "Marca_Declarada": marca,
            var_principal: val_principal_extracted,
            "Es_Producto_Principal": es_principal,
            "Marca_Extraida": marca_final,
            "Origen_Marca": fuente_marca,
        }

        for var in variables_cat:
            if var != var_principal:
                res[var] = cat_vals.get(var)

        for var in variables_pot:
            res[var] = num_vals.get(var)

        resultados.append(res)

    df_res = pd.DataFrame(resultados)
    df_final = pd.concat([df_raw.reset_index(drop=True), df_res.reset_index(drop=True)], axis=1)
    return df_final