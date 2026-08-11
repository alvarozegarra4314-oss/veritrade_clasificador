"""
Test del pipeline real (extractor_maestro.py) contra el maestro de la
línea de producto activa. Cubre tanto las funciones de bajo nivel
(evaluar_caracteristica_categorica_opt, extraer_potencia_numerica_opt)
como el pipeline completo por fila (procesar_dict_fila), incluyendo la
regla de negocio Tecnología-según-Fases.
"""
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import config
from __Otros.extractor_maestro import (
    CargarMaestro,
    limpiar_texto,
    evaluar_caracteristica_categorica_opt,
    extraer_potencia_numerica_opt,
    procesar_dict_fila,
)

CASOS_CARACTERISTICAS = [
    # (descripcion, variable, valor_esperado)
    ("UPS TRIFASICO 10KVA MONTAJE EN RACK", "Tipo_Producto_Detallado", "UPS Sistema Completo"),
    ("UPS TRIFASICO 10KVA MONTAJE EN RACK", "Salida_Fases", "Trifásico"),
    ("UPS TRIFASICO 10KVA MONTAJE EN RACK", "Formato_Montaje", "Rack"),
    ("UPS MONOFASICO 1 FASE TORRE", "Salida_Fases", "Monofásico"),
    ("UPS MONOFASICO 1 FASE TORRE", "Formato_Montaje", "Torre"),
    ("BATERIA DE RESPALDO PARA UPS", "Tipo_Producto_Detallado", "Batería / Banco de Baterías"),
]

CASOS_POTENCIA = [
    # (descripcion, variable, valor_esperado)
    ("UPS 10KVA TRIFASICO", "Potencia_kVA", 10.0),
    ("REGULADOR 1000VA", "Potencia_kVA", 1.0),  # respaldo nativo VA -> kVA
]

# (row_dict, campo_a_revisar, valor_esperado) -- prueba el pipeline por fila
# completo, incluida la regla Tecnología-según-Fases.
CASOS_FILA = [
    (
        {"Descripcion1": "UPS SISTEMA COMPLETO, APC, SMART 3KVA", "Embarcador / Exportador": "APC INC"},
        "Marca_Extraida", "APC",
    ),
    (
        {"Descripcion1": "UPS TRIFASICO 20KVA"},
        "Tipo_Tecnologia", "On-Line Doble Conversión",  # trifásico => siempre On-Line
    ),
    (
        {"Descripcion1": "UPS MONOFASICO 3KVA INTERACTIVO"},
        "Tipo_Tecnologia", "Interactivo (Line-Interactive)",  # monofásico sin online => interactivo
    ),
]


def correr_pruebas():
    print(f"Cargando maestro desde: {config.PATH_MAESTRO}")
    maestro = CargarMaestro(ruta_excel=config.PATH_MAESTRO)

    fallos = 0

    print("\n--- Características (evaluar_caracteristica_categorica_opt) ---")
    for texto, variable, esperado in CASOS_CARACTERISTICAS:
        texto_prep = limpiar_texto(texto).replace(',', '.')
        resultado = evaluar_caracteristica_categorica_opt(texto_prep, variable, maestro)
        ok = resultado == esperado
        fallos += 0 if ok else 1
        estado = "OK " if ok else "FAIL"
        print(f"[{estado}] {variable:25s} <- \"{texto}\" => {resultado!r} (esperado {esperado!r})")

    print("\n--- Potencia (extraer_potencia_numerica_opt) ---")
    for texto, variable, esperado in CASOS_POTENCIA:
        texto_prep = limpiar_texto(texto).replace(',', '.')
        resultado = extraer_potencia_numerica_opt(texto_prep, variable, maestro)
        ok = resultado == esperado
        fallos += 0 if ok else 1
        estado = "OK " if ok else "FAIL"
        print(f"[{estado}] {variable:15s} <- \"{texto}\" => {resultado!r} (esperado {esperado!r})")

    print("\n--- Pipeline por fila (procesar_dict_fila) ---")
    for row_dict, campo, esperado in CASOS_FILA:
        resultado_fila = procesar_dict_fila(row_dict, maestro)
        resultado = resultado_fila.get(campo)
        ok = resultado == esperado
        fallos += 0 if ok else 1
        estado = "OK " if ok else "FAIL"
        desc = row_dict.get("Descripcion1", "")
        print(f"[{estado}] {campo:20s} <- \"{desc}\" => {resultado!r} (esperado {esperado!r})")

    print("\n" + "=" * 60)
    if fallos == 0:
        print("✅ Todas las pruebas pasaron.")
    else:
        print(f"❌ {fallos} prueba(s) fallaron. Revisa las palabras clave en el maestro.")
    print("=" * 60)

    return fallos == 0


if __name__ == "__main__":
    ok = correr_pruebas()
    sys.exit(0 if ok else 1)
