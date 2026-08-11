import sys
from pathlib import Path

# Raíz del proyecto
BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import pandas as pd
from src import config
from src.pipeline import procesar_dataframe_dinamico


def main(linea: str = config.LINEA_PRODUCTO):
    cfg = config.config_linea(linea)

    path_raw = Path(cfg['path_raw'])
    path_maestro = Path(cfg['path_maestro'])
    path_output = Path(cfg['path_output'])

    print(f"Línea de producto: {cfg['linea']}")
    print(f"1. Leyendo datos crudos desde: {path_raw}")
    df_raw = pd.read_excel(path_raw, sheet_name=cfg['raw_sheet_name'])

    print(f"2. Procesando dataset con {path_maestro.name}...")
    df_procesado = procesar_dataframe_dinamico(df_raw, ruta_maestro=path_maestro)

    path_output.parent.mkdir(parents=True, exist_ok=True)

    print(f"3. Guardando resultado en: {path_output}")
    df_procesado.to_excel(path_output, index=False)

    print("✅ Proceso completado con éxito.")
    return df_procesado


if __name__ == "__main__":
    linea_arg = sys.argv[1] if len(sys.argv) > 1 else config.LINEA_PRODUCTO
    main(linea_arg)