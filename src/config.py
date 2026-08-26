from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_RAW_DIR = BASE_DIR / "data" / "raw"
DATA_MAESTRO_DIR = BASE_DIR / "data" / "maestro"
OUTPUT_DIR = BASE_DIR / "output"

LINEA_PRODUCTO = "UPS"

LINEAS = {
    "UPS": {
        "file_raw": "Veritrade_ups 2025 VAlf_DATA.xlsx",
        "raw_sheet_name": "Veritrade22_0825",
        "file_maestro": "Maestro_UPS_v2.xlsx",
        "file_output": "Resultado_Clasificacion_UPS.xlsx",
    },
    # "TABLEROS": {
    #     "file_raw": "Veritrade_Tableros_2025.xlsx",
    #     "raw_sheet_name": "Hoja1",
    #     "file_maestro": "Maestro_Tableros_v1.xlsx",
    #     "file_output": "Resultado_Clasificacion_Tableros.xlsx",
    # },
}


def config_linea(linea: str = LINEA_PRODUCTO) -> dict:
    if linea not in LINEAS:
        raise ValueError(
            f"Línea de producto '{linea}' no está registrada en config.LINEAS. "
            f"Líneas disponibles: {list(LINEAS.keys())}"
        )
    cfg = LINEAS[linea]
    return {
        "linea": linea,
        "path_raw": DATA_RAW_DIR / cfg["file_raw"],
        "raw_sheet_name": cfg["raw_sheet_name"],
        "path_maestro": DATA_MAESTRO_DIR / cfg["file_maestro"],
        "path_output": OUTPUT_DIR / cfg["file_output"],
    }


_cfg_activa = config_linea(LINEA_PRODUCTO)
PATH_RAW = _cfg_activa["path_raw"]
RAW_SHEET_NAME = _cfg_activa["raw_sheet_name"]
PATH_MAESTRO = _cfg_activa["path_maestro"]
PATH_OUTPUT = _cfg_activa["path_output"]

SHEET_CONFIG_LINEA = "0b_Config_Linea"
SHEET_MAESTRO_MARCAS = "1_Marcas"
SHEET_STOPWORDS = "1b_Palabras_Ignorar"
SHEET_MARCAS_DEFAULT = "1c_Marca_Por_Defecto"
SHEET_REGLAS_CARACTERISTICAS = "2_Caracteristicas"
SHEET_PATRONES_POTENCIA = "3_Tecnico_Potencia_NOEDIT"
SHEET_PATRONES_REGEX = "4_Tecnico_RegexMarca_NOEDIT"
SHEET_CONDICIONALES = "5_Condicionales"

COL_DESCRIPCION = "Descripcion Comercial"
COL_PATRON = "Patrón detectado en texto"
COL_MARCA_STD = "Marca estandarizada"

# ---------------------------------------------------------------------
# IA de rescate (Gemini) — SDK google-genai
# Si Google retira o renombra un modelo, se cambia aquí (o en la UI)
# sin tocar código del motor.
# ---------------------------------------------------------------------
MODELO_IA_DEFAULT = "gemini-3.5-flash-lite"
MODELOS_IA_DISPONIBLES = [
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
]
