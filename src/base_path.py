# src/base_path.py
from pathlib import Path
import sys

def get_base_dir() -> Path:
    """
    Base del proyecto:
    - Dev: carpeta raíz del repo (AGPE_AUTOMATION)
    - EXE (PyInstaller): carpeta donde está el .exe (escribible)
    """
    if getattr(sys, "frozen", False):
        # Carpeta donde está el ejecutable (permite escribir output/)
        return Path(sys.executable).resolve().parent

    # Dev: .../AGPE_AUTOMATION/src/base_path.py -> parents[1] = .../AGPE_AUTOMATION
    return Path(__file__).resolve().parents[1]


def get_resource_path(relative: str) -> Path:
    """
    Para leer recursos empaquetados (assets) en PyInstaller.
    En Dev, apunta al repo normal.
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative
    return get_base_dir() / relative
