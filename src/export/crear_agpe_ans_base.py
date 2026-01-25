from pathlib import Path
import shutil


def crear_agpe_ans_base():
    """
    Crea AGPE_ANS.xlsm copiándolo desde una plantilla base.
    openpyxl NO puede crear .xlsm desde cero.
    """

    base_dir = Path(__file__).resolve().parents[2]

    carpeta_plantillas = base_dir / "plantillas"
    carpeta_salida = base_dir / "data_clean"

    plantilla_base = carpeta_plantillas / "AGPE_ANS_BASE.xlsm"
    archivo_destino = carpeta_salida / "AGPE_ANS.xlsm"

    print("🧱 Verificando AGPE_ANS.xlsm...")

    if archivo_destino.exists():
        print("ℹ️ AGPE_ANS.xlsm ya existe. No se recrea.")
        return

    if not plantilla_base.exists():
        raise FileNotFoundError(
            "❌ No existe la plantilla AGPE_ANS_BASE.xlsm en /plantillas"
        )

    carpeta_salida.mkdir(exist_ok=True)

    shutil.copy(plantilla_base, archivo_destino)

    print("✅ AGPE_ANS.xlsm creado correctamente desde plantilla")


if __name__ == "__main__":
    crear_agpe_ans_base()