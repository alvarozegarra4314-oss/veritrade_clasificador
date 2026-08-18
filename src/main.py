import sys
from pathlib import Path

# Raíz del proyecto
BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import pandas as pd
from src import config
from src.pipeline import procesar_dataframe_dinamico
from src.maestro_optimizer import guardar_maestro_optimizado


def main(linea: str = config.LINEA_PRODUCTO, rescatador_ia=None):
    """
    rescatador_ia: instancia opcional de RescatadorIA (src/ia_rescate.py).
    Si se pasa y hubo aprendizaje útil durante la corrida, además del
    resultado clasificado se genera "Maestro_Optimizado_<linea>.xlsx" en
    la carpeta de output, con las marcas/características que la IA
    encontró y una hoja de auditoría "Log_Aprendizaje_IA".
    """
    cfg = config.config_linea(linea)

    path_raw = Path(cfg['path_raw'])
    path_maestro = Path(cfg['path_maestro'])
    path_output = Path(cfg['path_output'])

    print(f"Línea de producto: {cfg['linea']}")
    print(f"1. Leyendo datos crudos desde: {path_raw}")
    df_raw = pd.read_excel(path_raw, sheet_name=cfg['raw_sheet_name'])

    print(f"2. Procesando dataset con {path_maestro.name}...")
    df_procesado = procesar_dataframe_dinamico(
        df_raw, ruta_maestro=path_maestro, rescatador_ia=rescatador_ia
    )

    path_output.parent.mkdir(parents=True, exist_ok=True)

    print(f"3. Guardando resultado en: {path_output}")
    df_procesado.to_excel(path_output, index=False)

    # 4. Si hubo rescate IA con aprendizaje aprovechable, generamos el
    #    Maestro Optimizado (append de marcas/características + log de
    #    auditoría), sin tocar nunca el maestro original.
    if rescatador_ia is not None:
        propuestas = getattr(rescatador_ia, "propuestas_aprendizaje", None)
        hay_algo_que_guardar = propuestas and (
            propuestas.get("nuevas_marcas")
            or propuestas.get("nuevas_caracteristicas")
            or propuestas.get("revisar_manual")
        )
        if hay_algo_que_guardar:
            path_maestro_optimizado = path_output.parent / f"Maestro_Optimizado_{linea}.xlsx"
            print(f"4. Generando maestro optimizado en: {path_maestro_optimizado}")
            resumen = guardar_maestro_optimizado(
                ruta_maestro_original=path_maestro,
                propuestas=propuestas,
                ruta_salida=path_maestro_optimizado,
            )
            print(
                f"   -> {resumen['marcas_agregadas']} marca(s) nueva(s), "
                f"{resumen['caracteristicas_agregadas']} característica(s) nueva(s) agregada(s), "
                f"{resumen['pendientes_revision']} pendiente(s) de revisión manual "
                f"(ver hoja 'Log_Aprendizaje_IA')."
            )

    print("✅ Proceso completado con éxito.")
    return df_procesado


if __name__ == "__main__":
    linea_arg = sys.argv[1] if len(sys.argv) > 1 else config.LINEA_PRODUCTO
    main(linea_arg)