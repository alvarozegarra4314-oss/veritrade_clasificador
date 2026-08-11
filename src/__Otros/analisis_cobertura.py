import pandas as pd
from collections import Counter
import re
import config
from config import COL_DESCRIPCION

##### ESTE ES EL ANALISIS DE COBERTURA DE MARCAS NO IDENTIFICADAS #####


def analizar_no_identificados(linea: str = config.LINEA_PRODUCTO):
    ruta_salida = config.config_linea(linea)['path_output']
    print(f"Cargando resultados desde: {ruta_salida}")
    df = pd.read_excel(ruta_salida)
    
    # Filtrar solo las filas no identificadas (marcas por defecto del maestro,
    # ver hoja 1c_Marca_Por_Defecto: "Marca Generica" / "Marca Componentes")
    valores_no_identificados = {'Marca Generica', 'Marca Componentes'}
    df_unmapped = df[df['Marca_Extraida'].isin(valores_no_identificados)]
    total_unmapped = len(df_unmapped)
    total_filas = len(df)
    
    print(f"\n--- DIAGNÓSTICO DE COBERTURA ---")
    print(f"Total filas: {total_filas}")
    print(f"No identificados: {total_unmapped} ({(total_unmapped/total_filas)*100:.1f}%)")
    
    # Extraer las primeras palabras de cada descripción (donde suele estar la marca)
    palabras = []
    
    # Palabras comunes a ignorar que no son marcas
    stop_words = {'UPS', 'PARA', 'CON', 'DEL', 'POR', 'LOS', 'LAS', 'SISTEMA', 'UNIDAD', 'EQUIPO', 'MODELO', 'SERIE', 'DE'}
    
    for desc in df_unmapped[COL_DESCRIPCION].dropna():
        # Limpiar caracteres especiales y separar por espacios
        tokens = re.findall(r'\b[A-Z0-9]+\b', str(desc).upper())
        # Filtrar tokens cortos o palabras de parada
        tokens_filtrados = [t for t in tokens if len(t) > 2 and t not in stop_words and not t.isdigit()]
        
        # Tomar los primeros 3 términos útiles de la descripción
        palabras.extend(tokens_filtrados[:3])
        
    conteo = Counter(palabras).most_common(25)
    
    print("\n--- TOP 25 PALABRAS MÁS FRECUENTES EN FILAS NO IDENTIFICADAS ---")
    print(f"{'Candidato a Marca':<25} | {'Frecuencia':<10} | {'Impacto estimado en %'}")
    print("-" * 60)
    for palabra, freq in conteo:
        pct = (freq / total_filas) * 100
        print(f"{palabra:<25} | {freq:<10} | {pct:.2f}%")

if __name__ == "__main__":
    analizar_no_identificados()