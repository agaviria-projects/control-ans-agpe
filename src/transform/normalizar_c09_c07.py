from pathlib import Path
import pandas as pd


def leer_encabezados_plantilla(ruta_plantilla: Path) -> list:
    """
    Lee solo los encabezados (fila 1) de la PLANTILLA_C09_C07.xlsx
    """
    df_headers = pd.read_excel(ruta_plantilla, nrows=0)
    return list(df_headers.columns)


def normalizar_base_para_plantilla():
    base_dir = Path(__file__).resolve().parents[2]

    ruta_base_raw = base_dir / "data" / "processed" / "base_c09_c07_raw.xlsx"
    ruta_plantilla = base_dir / "entrada_diaria" / "PLANTILLA C09_C07.xlsx"
    ruta_salida = base_dir / "data" / "processed" / "base_c09_c07_normalizada.xlsx"

    if not ruta_base_raw.exists():
        raise FileNotFoundError("No existe base_c09_c07_raw.xlsx")

    if not ruta_plantilla.exists():
        raise FileNotFoundError("No existe PLANTILLA C09_C07.xlsx en ENTRADA_DIARIA")

    # Leer base consolidada (CSV unificados)
    df_base = pd.read_excel(ruta_base_raw, dtype=str)

    # Leer encabezados de la plantilla
    columnas_plantilla = leer_encabezados_plantilla(ruta_plantilla)

    print(f"📋 Columnas plantilla ({len(columnas_plantilla)}):")
    print(columnas_plantilla)

    print(f"\n📊 Columnas base CSV ({len(df_base.columns)}):")
    print(list(df_base.columns))

    # Agregar columnas faltantes en la base
    for col in columnas_plantilla:
        if col not in df_base.columns:
            df_base[col] = ""

    # Reordenar exactamente como la plantilla
    df_normalizada = df_base[columnas_plantilla]

    # Guardar resultado
    df_normalizada.to_excel(ruta_salida, index=False)

    print(f"\n💾 Base normalizada generada: {ruta_salida.name}")
    print(f"✅ Registros: {len(df_normalizada)}")

    return df_normalizada
