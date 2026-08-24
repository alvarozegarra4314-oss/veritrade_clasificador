"""Configuración de pytest para todo el proyecto.

Garantiza que la raíz del proyecto esté en sys.path para que los tests
puedan importar el paquete `src.*` sin instalación adicional.
"""

import sys
from pathlib import Path

RAIZ_PROYECTO = Path(__file__).resolve().parent
if str(RAIZ_PROYECTO) not in sys.path:
    sys.path.insert(0, str(RAIZ_PROYECTO))
