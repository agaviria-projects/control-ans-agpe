import sys
from pathlib import Path

def get_base_dir():
    """
    Devuelve la carpeta raíz de ejecución:
    - En desarrollo (.py): raíz del proyecto
    - En producción (.exe): carpeta donde está el ejecutable
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).resolve().parents[2]
