import sys
from pathlib import Path
from datetime import datetime
import os

from src.base_path import get_base_dir, get_resource_path

APP_DIR = get_base_dir()

# --------------------------------------------------
# Agregar proyecto al path (compatibilidad)
# --------------------------------------------------
BASE_DIR = APP_DIR
sys.path.append(str(BASE_DIR))
# --------------------------------------------------
# Imports estándar
# --------------------------------------------------
import tkinter as tk
from tkinter import messagebox
from .calendario_ans import abrir_calendario

# --------------------------------------------------
# Imports del proyecto (UNA SOLA VEZ)
# --------------------------------------------------
from src.mapas.generar_mapa_agpe import generar_mapa_leaflet_agpe
from src.export.append_agpe_ans import append_agpe_ans
from src.extract.consolidar_c09_c07 import consolidar_c09_c07
from src.export.preparar_agpe_clean_excel import preparar_agpe_clean_excel

# --------------------------------------------------
# Funciones del formulario (NO TOCAR LÓGICA)
# --------------------------------------------------
def ejecutar_merge():
    try:
        lbl_estado.config(text="⏳ Generando AGPE_CLEAN...")
        ventana.update_idletasks()

        consolidar_c09_c07()
        preparar_agpe_clean_excel()

        lbl_estado.config(text="✅ AGPE_CLEAN listo para edición")
        messagebox.showinfo(
            "AGPE",
            "AGPE_CLEAN fue generado y preparado correctamente para edición."
        )

    except Exception as e:
        lbl_estado.config(text="❌ Error generando AGPE_CLEAN")
        messagebox.showerror("Error", f"Ocurrió un error:\n\n{e}")


def ejecutar_append():
    try:
        lbl_estado.config(text="⏳ Actualizando AGPE_ANS...")
        ventana.update_idletasks()

        append_agpe_ans()

        lbl_estado.config(text="✅ AGPE_ANS actualizado correctamente")
        messagebox.showinfo(
            "AGPE",
            "El archivo AGPE_ANS fue actualizado correctamente."
        )

    except (RuntimeError, ValueError) as e:
        lbl_estado.config(text="ℹ️ No hay registros nuevos para agregar")
        messagebox.showinfo("AGPE", str(e))

    except Exception as e:
        lbl_estado.config(text="❌ Error en el proceso")
        messagebox.showerror(
            "Error",
            f"Ocurrió un error inesperado:\n\n{e}"
        )


def ejecutar_mapa():
    try:
        lbl_estado.config(text="🗺️ Generando mapa AGPE...")
        ventana.update_idletasks()

        generar_mapa_leaflet_agpe()

        lbl_estado.config(text="✅ Mapa AGPE generado correctamente")
        # messagebox.showinfo(
        #     "Mapa de Geolocalización - AGPE",
        #     "El mapa fue generado correctamente."
        # )

    except Exception as e:
        lbl_estado.config(text="❌ Error generando el mapa")
        messagebox.showerror("Error", f"Ocurrió un error:\n\n{e}")


def mostrar_calendario():
    abrir_calendario(ventana)


def salir_panel():
    ventana.destroy()


def abrir_archivo(ruta):
    try:
        os.startfile(ruta)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")


def actualizar_hora():
    hora_actual = datetime.now().strftime("%I:%M:%S %p")
    lbl_hora.config(text=hora_actual)
    ventana.after(1000, actualizar_hora)

# --------------------------------------------------
# Interfaz gráfica – ESTILO FENIX (COMPACTO)
# --------------------------------------------------
ventana = tk.Tk()
ventana.title("AGPE - Elite Ingenieros S.A.S.")
ventana.geometry("600x520")
ventana.resizable(False, False)
ventana.configure(bg="#EAEDED")

# --------------------------------------------------
# Franja verde superior (título centrado)
# --------------------------------------------------
header = tk.Frame(ventana, bg="#1E8449", height=40)
header.pack(fill="x")
header.pack_propagate(False)

# Título centrado
lbl_header = tk.Label(
    header,
    text="CONTROL ANS – AGPE",
    bg="#1E8449",
    fg="white",
    font=("Segoe UI", 11, "bold")
)
lbl_header.place(relx=0.5, rely=0.5, anchor="center")

# Hora (derecha)
lbl_hora = tk.Label(
    header,
    text="",
    bg="#1E8449",
    fg="white",
    font=("Segoe UI", 10, "bold")
)
lbl_hora.pack(side="right", padx=12)

# --------------------------------------------------
# Identidad corporativa (debajo del header)
# --------------------------------------------------
frame_identidad = tk.Frame(ventana, bg="#EAEDED")
frame_identidad.pack(pady=(6, 8))  # 👈 OJO al pady

ruta_logo = get_resource_path("assets/logo.png")
logo_img = tk.PhotoImage(file=str(ruta_logo)).subsample(2, 2)


lbl_logo = tk.Label(
    frame_identidad,
    image=logo_img,
    bg="#EAEDED"
)
lbl_logo.image = logo_img
lbl_logo.pack(side="left", padx=(0, 10))

lbl_empresa = tk.Label(
    frame_identidad,
    text="ELITE Ingenieros S.A.S.",
    bg="#EAEDED",
    fg="#1B263B",
    font=("Segoe UI", 16, "bold")
)
lbl_empresa.pack(side="left")

# --------------------------------------------------
# Contenedor principal
# --------------------------------------------------
contenedor = tk.Frame(ventana, bg="#EAEDED")
contenedor.pack(pady=(12, 15))

tk.Button(
    contenedor,
    text="Generar Merge C09 + C07",
    font=("Segoe UI", 10, "bold"),
    bg="#1E8449",
    fg="white",
    width=26,
    height=2,
    cursor="hand2",
    command=ejecutar_merge
).pack(pady=6)

tk.Button(
    contenedor,
    text="Actualizar AGPE_ANS",
    font=("Segoe UI", 10, "bold"),
    bg="#1E8449",
    fg="white",
    width=26,
    height=2,
    cursor="hand2",
    command=ejecutar_append
).pack(pady=6)

tk.Button(
    contenedor,
    text="Generar Mapa AGPE",
    font=("Segoe UI", 10, "bold"),
    bg="#117A65",
    fg="white",
    width=26,
    height=2,
    cursor="hand2",
    command=ejecutar_mapa
).pack(pady=6)

tk.Button(
    contenedor,
    text="Salir del Panel",
    font=("Segoe UI", 10, "bold"),
    bg="#922B21",
    fg="white",
    width=26,
    height=2,
    cursor="hand2",
    command=salir_panel
).pack(pady=(14, 6))

# --------------------------------------------------
# Separador visual antes de iconos
# --------------------------------------------------
# tk.Frame(ventana, bg="#B3B6B7", height=1).pack(fill="x", pady=(8, 4))

# --------------------------------------------------
# ICONOS DE ACCESO RÁPIDO (ESTILO FENIX)
# --------------------------------------------------

# Separador visual único (solo uno)
tk.Frame(ventana, bg="#B3B6B7", height=1).pack(fill="x", pady=(30, 12))

# Contenedor alineado a la izquierda
frame_iconos = tk.Frame(ventana, bg="#EAEDED")
frame_iconos.pack(fill="x", padx=12, pady=(2, 4), anchor="w")

# Cargar imágenes (tamaño elegante tipo FENIX)
icon_agpe_ans = tk.PhotoImage(
    file=str(get_resource_path("assets/agpe_ans.png"))
).subsample(1, 1)

icon_agpe_clean = tk.PhotoImage(
    file=str(get_resource_path("assets/agpe_clean.png"))
).subsample(1, 1)

icon_bdpcp = tk.PhotoImage(
    file=str(get_resource_path("assets/BDPCP.png"))
).subsample(1, 1)

icon_primer = tk.PhotoImage(
    file=str(get_resource_path("assets/primer_visita.png"))
).subsample(1, 1)


def crear_icono(parent, img, tooltip, archivo, pady=(4, 0)):
    lbl = tk.Label(
        parent,
        image=img,
        bg="#EAEDED",
        cursor="hand2"
    )
    lbl.image = img
    lbl.pack(side="left", padx=(0, 14), pady=pady)  # 👈 pady configurable

    lbl.bind("<Button-1>", lambda e: abrir_archivo(archivo))
    lbl.bind("<Enter>", lambda e: lbl_estado.config(text=tooltip))
    lbl.bind("<Leave>", lambda e: lbl_estado.config(
        text="⚙️ Esperando acción del usuario..."
    ))

crear_icono(
    frame_iconos,
    icon_agpe_ans,
    "📄 Abrir AGPE_ANS",
    APP_DIR / "data_clean" / "AGPE_ANS.xlsm"
)


crear_icono(
    frame_iconos,
    icon_agpe_clean,
    "🧹 Abrir AGPE_CLEAN",
    APP_DIR / "data_clean" / "AGPE_CLEAN.xlsx"
)

crear_icono(
    frame_iconos,
    icon_bdpcp,
    "🗄️ Abrir BDPCP",
    APP_DIR / "data_clean" / "BDPCP.xlsx"
)

crear_icono(
    frame_iconos,
    icon_primer,
    "📝 Abrir Plantilla Primer Visita",
    APP_DIR / "data_clean" / "PLANTILLA_PRIMER VISITAS.xlsm",
    pady=(10, 0)
)

# --------------------------------------------------
# Footer definitivo
# --------------------------------------------------
frame_footer = tk.Frame(ventana, bg="#EAEDED")
frame_footer.pack(side="bottom", fill="x", pady=(0, 4))

tk.Frame(frame_footer, bg="#B3B6B7", height=1).pack(fill="x")

frame_footer_in = tk.Frame(frame_footer, bg="#EAEDED")
frame_footer_in.pack(fill="x", padx=10, pady=2)

lbl_estado = tk.Label(
    frame_footer_in,
    text="⚙️ Esperando acción del usuario...",
    bg="#EAEDED",
    fg="#1B263B",
    font=("Segoe UI", 9, "italic")
)
lbl_estado.pack(side="left")

tk.Button(
    frame_footer_in,
    text="📅 Calendario ANS",
    font=("Segoe UI", 9, "bold"),
    bg="#EAEDED",
    fg="#1E8449",
    relief="flat",
    cursor="hand2",
    command=mostrar_calendario
).pack(side="right", padx=(0, 8))

tk.Label(
    frame_footer_in,
    text="© 2025 Elite Ingenieros S.A.S.",
    bg="#EAEDED",
    fg="#7B7D7D",
    font=("Segoe UI", 9, "italic")
).pack(side="right", padx=(0, 12))

def iniciar_panel():
    actualizar_hora()
    ventana.mainloop()

if __name__ == "__main__":
    iniciar_panel()    