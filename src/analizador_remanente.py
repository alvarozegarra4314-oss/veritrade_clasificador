import os
import re
from collections import Counter
import pandas as pd

# ==========================================
# 1. RUTAS ABSOLUTAS AL PROYECTO (data/maestro)
# ==========================================
# Se ancla desde /src -> sube a la raíz -> entra a data/maestro
SRC_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SRC_DIR)
CARPETA_MAESTRO = os.path.join(PROJECT_ROOT, "data", "maestro")

NOMBRE_ARCHIVO = "Descripciones_NO DETECTADO_UPS.xlsx"
RUTA_ENTRADA = os.path.join(CARPETA_MAESTRO, NOMBRE_ARCHIVO)
RUTA_SALIDA = os.path.join(CARPETA_MAESTRO, "Resultado_Analisis_Remanente.xlsx")

# ==========================================
# 2. DICCIONARIO DE LIMPIEZA (STOPWORDS)
# ==========================================
STOPWORDS = {
    # Conectores y gramática
    'DE', 'DEL', 'LA', 'EL', 'LOS', 'LAS', 'EN', 'PARA', 'CON', 'POR', 'SIN', 'UN', 'UNA', 'Y', 'O', 'QUE',
    'FOR', 'WITH', 'WITHOUT', 'TYPE', 'DESCRIPCION', 'ITEMS', 'ITEM', 'TOTAL', 'CANTIDAD', 'VALOR', 'PRECIO',
    'SOLO', 'PERMITE', 'PARTE', 'PARTES', 'SET', 'UNIDADES', 'USO', 'MEDIDA', 'PESO', 'LOTE', 'PAG',
    
    # Términos aduaneros y administrativos
    'MARCA', 'MODELO', 'SERIE', 'S/M', 'S/N', 'NOS', 'SPEC', 'EQUIPO', 'EQUIPOS', 'TIPO', 'CODIGO',
    'NO', 'N/A', 'COMERCIAL', 'INDUSTRIAL', 'FABRICANTE', 'M/N', 'ESTRUC', 'ESTRUCTURA',
    'BULTO', 'BULTOS', 'CAJA', 'CAJAS', 'PIEZA', 'PIEZAS', 'INCLUYE', 'ACCESORIOS', 'METAL', 'ACERADA',
    'PLASTICO', 'CABLE', 'CABLES', 'PANEL', 'CONTROL', 'MUESTRA', 'MUESTRAS', 'DUA', 'DAM',
    'IMPORTACION', 'PARTIDA', 'ARANCELARIA', 'PAIS', 'ORIGEN', 'DECLARACION', 'NRO', 'NUMERO',
    
    # Términos técnicos de UPS / Electrónica
    'UPS', 'SAI', 'ALIMENTACION', 'FUENTE', 'PODER', 'POWER', 'CARGADOR', 'ADAPTADOR', 'ESTABILIZADOR',
    'REGULADOR', 'ENERGIA', 'CONVERTIDOR', 'SUPPLY', 'CORRIENTE', 'VOLTAJE', 'MINI', 'SISTEMA', 'SISTEMAS',
    'RESPALDO', 'ININTERRUMPIDA', 'ININTERRUNPIDA', 'ININTERRUMPIBLE', 'UNINTERRUPTIBLE', 'TENSION',
    'AUTONOMIA', 'CABINET', 'GABINETE', 'DISPOSITIVO', 'PROTECCION', 'ELECTRICA', 'ELECTRICO', 'ELECTRONICO',
    'EQUIPAMIENTO', 'RED', 'REDES', 'DATOS', 'SERVIDOR', 'SERVIDORES', 'COMPUTADORA', 'COMPUTADORAS',
    'ALMACENADOR', 'TARJETA', 'TORRE', 'MODULAR', 'LINEA', 'DOBLE', 'INTERNAL', 'CONVERT', 'ESTABILI',
    'INVERSOR', 'PUERTO', 'SERIAL', 'ONLINE', 'ON-LINE', 'INTERACTIVO', 'INTERACTIVE', 'MONOFASICO',
    'TRIFASICO', 'RACK', 'RACKMOUNT', 'TOWER', 'CARD', 'USB', 'RS232', 'SNMP', 'HIGH', 'FREQUENCY', 'LCD',
    'POTENCIA', 'SUMINISTRO', 'ELECTRONICA', '3PHASES', 'U.C', 'OUTPUT', 'INPUT', 'FACTOR', 'PF', 'LED',
    'VERSION', 'BATTERY', 'BATERIA', 'BATERIAS',
    
    # Unidades de medida técnicas
    'KVA', 'KW', 'VA', 'WATTS', 'VOLTIOS', 'AMPERIOS', 'VOLTS', 'VOLT', 'AMP', 'HZ', '60HZ', '50HZ',
    '220V', '220VAC', '110V', '380V', '480V', '7AH', '9AH', '12V', 'KG', 'USD', 'FOB', 'CIF'
}

# ==========================================
# 3. FUNCIONES PROCESADORAS
# ==========================================
def extraer_marcas_sintacticas(texto):
    """Detecta nombres ubicados tras patrones típicos de aduanas."""
    patrones = [
        r'MARCA[:\s,]+([A-Z0-9\.\-\&]+(?:\s+[A-Z0-9\.\-\&]+)?)',
        r'BRAND[:\s,]+([A-Z0-9\.\-\&]+(?:\s+[A-Z0-9\.\-\&]+)?)',
        r'UPS[,\s]+([A-Z0-9\.\-\&]+(?:\s+[A-Z0-9\.\-\&]+)?)[,\;]',
        r'FABRICANTE[:\s,]+([A-Z0-9\.\-\&]+(?:\s+[A-Z0-9\.\-\&]+)?)',
        r'M/([A-Z0-9\.\-\&]+)',
    ]
    
    marcas = []
    for pat in patrones:
        coincidencias = re.findall(pat, texto)
        for cand in coincidencias:
            cand_clean = cand.strip()
            if (cand_clean not in STOPWORDS and 
                len(cand_clean) >= 3 and 
                not cand_clean.isdigit() and 
                not re.match(r'^\d+[A-Z]+$', cand_clean)):
                marcas.append(cand_clean)
    return marcas

def analizar_descripciones(df):
    """Procesa el DataFrame para extraer palabras y marcas sospechosas."""
    col_desc = 'Descripcion Comercial' if 'Descripcion Comercial' in df.columns else df.columns[0]
    textos = df[col_desc].dropna().astype(str).str.upper().tolist()
    
    todas_palabras = []
    marcas_patrones = []
    
    for desc in textos:
        tokens = re.findall(r'\b[A-Z0-9\.\-\&]{3,}\b', desc)
        for t in tokens:
            if t not in STOPWORDS and not t.isdigit() and not re.match(r'^\d+[A-Z]+$', t):
                todas_palabras.append(t)
                
        marcas_patrones.extend(extraer_marcas_sintacticas(desc))
        
    frec_palabras = Counter(todas_palabras).most_common(50)
    frec_marcas = Counter(marcas_patrones).most_common(30)
    
    return frec_palabras, frec_marcas

# ==========================================
# 4. FLUJO PRINCIPAL
# ==========================================
def main():
    if not os.path.exists(RUTA_ENTRADA):
        print(f"❌ No se encontró el archivo de entrada:")
        print(f"   Ruta esperada: {RUTA_ENTRADA}")
        print("💡 Verifica que el archivo esté guardado dentro de 'data/maestro/'.")
        return

    print(f"📂 Cargando archivo desde: {RUTA_ENTRADA}")
    df = pd.read_excel(RUTA_ENTRADA) if RUTA_ENTRADA.endswith(('.xlsx', '.xls')) else pd.read_csv(RUTA_ENTRADA)
    print(f"✅ Registros cargados: {len(df)}\n")
    
    frec_palabras, frec_marcas = analizar_descripciones(df)
    
    df_marcas = pd.DataFrame(frec_marcas, columns=['Posible Marca / Empresa', 'Apariciones'])
    df_palabras = pd.DataFrame(frec_palabras, columns=['Palabra / Término', 'Frecuencia'])
    
    # Imprimir resumen en consola
    print("=" * 60)
    print("🏷️ TOP POSIBLES MARCAS / EMPRESAS DETECTADAS")
    print("=" * 60)
    print(df_marcas.to_string(index=False))
    
    print("\n" + "=" * 60)
    print("🔤 TOP 25 PALABRAS MÁS FRECUENTES (FILTRADAS)")
    print("=" * 60)
    print(df_palabras.head(25).to_string(index=False))
    print("=" * 60)
    
    # Guardar resultado directamente en data/maestro/
    try:
        with pd.ExcelWriter(RUTA_SALIDA, engine='openpyxl') as writer:
            df_marcas.to_excel(writer, sheet_name='Posibles_Marcas', index=False)
            df_palabras.to_excel(writer, sheet_name='Top_Palabras', index=False)
        print(f"\n💾 Reporte exportado exitosamente a:\n   {RUTA_SALIDA}")
    except Exception as e:
        print(f"\n⚠️ Error al exportar el Excel: {e}")

if __name__ == "__main__":
    main()