from pathlib import Path
import pandas as pd


def consolidar_c09_c07():
    print("📥 Consolidando C09 + C07...")

    base_dir = Path(__file__).resolve().parents[2]
    raw_dir = base_dir / "data_raw"
    clean_dir = base_dir / "data_clean"
    output_file = clean_dir / "AGPE_CLEAN.xlsx"

    archivos = list(raw_dir.glob("pendientes_*.csv"))

    if not archivos:
        raise FileNotFoundError("❌ No se encontraron archivos pendientes_*.csv en data_raw")

    dfs = []
    for archivo in archivos:
        print(f"➡️ Leyendo {archivo.name}")
        df = pd.read_csv(archivo, dtype=str, encoding="latin1")
        dfs.append(df)

    df_total = pd.concat(dfs, ignore_index=True)

    # Columnas DEFINIDAS del merge C09 + C07
    columnas_merge = [
        "Pedido",
        "Tipo_Trabajo",
        "Fecha_Concepto",
        "Fecha_Inicio_ANS",
        "ClienteID",
        "Nombre_Cliente",
        "Direccion",
        "Municipio",
        "Subzona",
        "Coordenadax",
        "Coordenaday",
        "Actividad",
        "Tipo_Dirección",
        "Observación_Solicitud",
        "Pedido_CRM",
        "Detalle Visita",
        "Tipo Medidor"
    ]

    # Normalizar estructura
    for col in columnas_merge:
        if col not in df_total.columns:
            df_total[col] = ""

    df_total = df_total[columnas_merge]

    clean_dir.mkdir(exist_ok=True)
    df_total.to_excel(output_file, index=False)

    print(f"✅ AGPE_CLEAN generado correctamente")
    print(f"📊 Registros totales: {len(df_total)}")