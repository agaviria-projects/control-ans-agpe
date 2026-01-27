# src/ui/calendario_ans.py
from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import date, timedelta
import tkinter as tk
from tkinter import ttk


# ============================================================
# CONFIGURACIÓN FESTIVOS (mantén aquí los festivos del ANS)
# ============================================================
FESTIVOS = {
    # 2025
    date(2025, 1, 1), date(2025, 1, 6), date(2025, 3, 24), date(2025, 4, 17), date(2025, 4, 18),
    date(2025, 5, 1), date(2025, 5, 26), date(2025, 6, 16), date(2025, 6, 23), date(2025, 7, 7),
    date(2025, 8, 7), date(2025, 8, 18), date(2025, 10, 13), date(2025, 11, 3), date(2025, 11, 17),
    date(2025, 12, 8), date(2025, 12, 25),

    # 2026
    date(2026, 1, 1), date(2026, 1, 12), date(2026, 3, 23), date(2026, 4, 2), date(2026, 4, 3),
    date(2026, 5, 1), date(2026, 5, 18), date(2026, 6, 8), date(2026, 6, 15), date(2026, 6, 29),
    date(2026, 7, 20), date(2026, 8, 7), date(2026, 8, 17), date(2026, 10, 12), date(2026, 11, 2),
    date(2026, 11, 16), date(2026, 12, 8), date(2026, 12, 25),
}


@dataclass(frozen=True)
class Theme:
    # Estética similar a tu 2da imagen (compacta)
    bg: str = "#F2F2F2"
    header_bg: str = "#1F1F1F"
    header_fg: str = "#FFFFFF"
    grid_bg: str = "#FFFFFF"
    weekday_fg: str = "#333333"

    # Colores “como la 1ra imagen” (festivos en rojo sólido)
    festivo_bg: str = "#C8102E"
    festivo_fg: str = "#FFFFFF"

    # Sábados y domingos “pintados”
    weekend_bg: str = "#FADADD"    # rosado suave
    weekend_fg: str = "#222222"

    # Día de hoy (borde)
    today_border: str = "#111111"

    # Columna Semana (Sem)
    weeknum_bg: str = "#E6E6E6"   # gris suave
    weeknum_fg: str = "#333333"


class CalendarioANS(tk.Toplevel):
    def __init__(self, master: tk.Misc | None = None, year: int | None = None, month: int | None = None):
        super().__init__(master)
        self.title("Calendario ANS - Elite Ingenieros")
        self.resizable(False, False)

        self.theme = Theme()

        hoy = date.today()
        self.year = year or hoy.year
        self.month = month or hoy.month

        self.configure(bg=self.theme.bg)

        # Header
        header = tk.Frame(self, bg=self.theme.header_bg)
        header.pack(fill="x", padx=10, pady=(10, 6))

        btn_prev_m = ttk.Button(header, text="◀", width=3, command=self._prev_month)
        btn_prev_m.pack(side="left")

        self.lbl_title = tk.Label(
            header,
            text="",
            bg=self.theme.header_bg,
            fg=self.theme.header_fg,
            font=("Segoe UI", 11, "bold"),
            padx=10
        )
        self.lbl_title.pack(side="left")

        btn_next_m = ttk.Button(header, text="▶", width=3, command=self._next_month)
        btn_next_m.pack(side="left")

        ttk.Separator(self).pack(fill="x", padx=10, pady=(0, 8))

        # Contenedor del calendario
        self.grid_frame = tk.Frame(self, bg=self.theme.grid_bg, bd=1, relief="solid")
        self.grid_frame.pack(padx=10, pady=(0, 10))

        # Footer
        footer = tk.Frame(self, bg=self.theme.bg)
        footer.pack(fill="x", padx=10, pady=(0, 10))

        ttk.Button(footer, text="Vista anual", command=self._open_year_view).pack(side="left")
        ttk.Button(footer, text="Cerrar", command=self.destroy).pack(side="right")

        self._render_month()

    # -------------------------
    # Helper: número de semana ISO por fila
    # -------------------------
    def _iso_week_for_row(self, year: int, month: int, week: list[int]) -> int:
        # Busca el primer día real (no cero) y su índice de columna (Lu=0..Do=6)
        first_day = next((d for d in week if d != 0), None)
        if first_day is None:
            return 0

        first_idx = week.index(first_day)  # offset desde lunes
        d0 = date(year, month, first_day)

        # Lunes real de esa fila (puede caer en mes/año anterior)
        monday = d0 - timedelta(days=first_idx)
        return monday.isocalendar().week

    # -------------------------
    # Render mensual
    # -------------------------
    def _render_month(self):
        # Limpia grid
        for w in self.grid_frame.winfo_children():
            w.destroy()

        mes_nombre = calendar.month_name[self.month]
        self.lbl_title.config(text=f"{mes_nombre} {self.year}".title())

        # Encabezados días (Lu..Do)
        # calendar: lunes=0...domingo=6. Queremos Lu..Do
        dias = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sa", "Do"]

        # Columna adicional: número de semana
        lbl_sem = tk.Label(
            self.grid_frame,
            text="Sem",
            bg=self.theme.weeknum_bg,
            fg=self.theme.weeknum_fg,
            font=("Segoe UI", 9, "bold"),
            width=4,
            pady=4
        )
        lbl_sem.grid(row=0, column=0, sticky="nsew")

        # Días arrancan en columna 1
        for c, d in enumerate(dias, start=1):
            lbl = tk.Label(
                self.grid_frame,
                text=d,
                bg=self.theme.grid_bg,
                fg=self.theme.weekday_fg,
                font=("Segoe UI", 9, "bold"),
                width=4,
                pady=4
            )
            lbl.grid(row=0, column=c, sticky="nsew")

        cal = calendar.Calendar(firstweekday=0)  # 0 = lunes
        weeks = cal.monthdayscalendar(self.year, self.month)

        hoy = date.today()

        for r, week in enumerate(weeks, start=1):
            # Semana ISO en columna 0
            wk = self._iso_week_for_row(self.year, self.month, week)
            wk_cell = tk.Label(
                self.grid_frame,
                text=str(wk) if wk else "",
                bg=self.theme.weeknum_bg,
                fg=self.theme.weeknum_fg,
                width=4,
                pady=6,
                font=("Segoe UI", 9, "bold"),
                bd=1,
                relief="solid"
            )
            wk_cell.grid(row=r, column=0, sticky="nsew", padx=1, pady=1)

            # Días: columnas 1..7
            for day_idx, daynum in enumerate(week):
                col = day_idx + 1

                if daynum == 0:
                    # celda vacía
                    cell = tk.Label(self.grid_frame, text="", bg=self.theme.grid_bg, width=4, pady=6)
                    cell.grid(row=r, column=col, sticky="nsew", padx=1, pady=1)
                    continue

                d = date(self.year, self.month, daynum)
                is_weekend = day_idx in (5, 6)  # Sa=5, Do=6
                is_festivo = d in FESTIVOS

                bg = self.theme.grid_bg
                fg = self.theme.weekday_fg

                if is_weekend:
                    bg = self.theme.weekend_bg
                    fg = self.theme.weekend_fg

                if is_festivo:
                    bg = self.theme.festivo_bg
                    fg = self.theme.festivo_fg

                # Borde para hoy
                bd = 1
                relief = "solid"
                highlight = (d == hoy)

                cell = tk.Label(
                    self.grid_frame,
                    text=str(daynum),
                    bg=bg,
                    fg=fg,
                    width=4,
                    pady=6,
                    font=("Segoe UI", 9, "bold" if is_festivo else "normal"),
                    bd=bd,
                    relief=relief
                )

                if highlight:
                    cell.config(highlightthickness=2, highlightbackground=self.theme.today_border)

                cell.grid(row=r, column=col, sticky="nsew", padx=1, pady=1)

        # Ajuste de columnas uniforme (Sem + 7 días = 8 columnas)
        for c in range(8):
            self.grid_frame.grid_columnconfigure(c, weight=1)

    # -------------------------
    # Navegación
    # -------------------------
    def _prev_month(self):
        if self.month == 1:
            self.month = 12
            self.year -= 1
        else:
            self.month -= 1
        self._render_month()

    def _next_month(self):
        if self.month == 12:
            self.month = 1
            self.year += 1
        else:
            self.month += 1
        self._render_month()

    # -------------------------
    # Vista anual (12 meses)
    # -------------------------
    def _open_year_view(self):
        YearView(self, self.year)


class YearView(tk.Toplevel):
    def __init__(self, master: tk.Misc, year: int):
        super().__init__(master)
        self.title(f"Calendario ANS - {year}")
        self.resizable(True, True)

        self.theme = Theme()
        self.year = year
        self.configure(bg=self.theme.bg)

        top = tk.Frame(self, bg=self.theme.bg)
        top.pack(fill="x", padx=10, pady=10)

        ttk.Button(top, text="◀ Año", command=self._prev_year).pack(side="left")
        self.lbl = tk.Label(top, text=str(self.year), bg=self.theme.bg, font=("Segoe UI", 12, "bold"))
        self.lbl.pack(side="left", padx=10)
        ttk.Button(top, text="Año ▶", command=self._next_year).pack(side="left")

        ttk.Button(top, text="Cerrar", command=self.destroy).pack(side="right")

        self.container = tk.Frame(self, bg=self.theme.bg)
        self.container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._render_year()

    def _render_year(self):
        for w in self.container.winfo_children():
            w.destroy()

        self.lbl.config(text=str(self.year))

        # 12 meses en 3 filas x 4 columnas
        meses = list(range(1, 13))
        idx = 0
        for r in range(3):
            for c in range(4):
                m = meses[idx]
                idx += 1

                frame = tk.Frame(self.container, bg=self.theme.grid_bg, bd=1, relief="solid")
                frame.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")

                title = tk.Label(
                    frame, text=calendar.month_name[m].title(),
                    bg=self.theme.grid_bg, font=("Segoe UI", 10, "bold"), pady=4
                )
                title.pack(fill="x")

                mini = tk.Frame(frame, bg=self.theme.grid_bg)
                mini.pack(padx=6, pady=(0, 6))

                # Encabezados mini
                dias = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sa", "Do"]

                # Columna "Sem"
                tk.Label(
                    mini, text="Sem",
                    bg=self.theme.weeknum_bg, fg=self.theme.weeknum_fg,
                    font=("Segoe UI", 7, "bold"), width=3
                ).grid(row=0, column=0)

                # Días arrancan en columna 1
                for cc, dname in enumerate(dias, start=1):
                    tk.Label(mini, text=dname, bg=self.theme.grid_bg, font=("Segoe UI", 7, "bold"), width=3)\
                        .grid(row=0, column=cc)

                cal = calendar.Calendar(firstweekday=0)
                weeks = cal.monthdayscalendar(self.year, m)

                for rr, week in enumerate(weeks, start=1):
                    # Semana ISO (columna 0)
                    if any(dn != 0 for dn in week):
                        first_day = next(dn for dn in week if dn != 0)
                        first_idx = week.index(first_day)
                        monday = date(self.year, m, first_day) - timedelta(days=first_idx)
                        wk = monday.isocalendar().week
                    else:
                        wk = 0

                    tk.Label(
                        mini, text=str(wk) if wk else "",
                        bg=self.theme.weeknum_bg, fg=self.theme.weeknum_fg,
                        width=3, font=("Segoe UI", 7, "bold")
                    ).grid(row=rr, column=0, padx=1, pady=1)

                    # Días (col 1..7)
                    for day_idx, daynum in enumerate(week):
                        col = day_idx + 1

                        if daynum == 0:
                            tk.Label(mini, text="", bg=self.theme.grid_bg, width=3).grid(row=rr, column=col)
                            continue

                        d = date(self.year, m, daynum)
                        is_weekend = day_idx in (5, 6)
                        is_festivo = d in FESTIVOS

                        bg = self.theme.grid_bg
                        fg = self.theme.weekday_fg
                        if is_weekend:
                            bg = self.theme.weekend_bg
                            fg = self.theme.weekend_fg
                        if is_festivo:
                            bg = self.theme.festivo_bg
                            fg = self.theme.festivo_fg

                        tk.Label(
                            mini, text=str(daynum), bg=bg, fg=fg, width=3,
                            font=("Segoe UI", 7, "bold" if is_festivo else "normal")
                        ).grid(row=rr, column=col, padx=1, pady=1)

        for c in range(4):
            self.container.grid_columnconfigure(c, weight=1)
        for r in range(3):
            self.container.grid_rowconfigure(r, weight=1)

    def _prev_year(self):
        self.year -= 1
        self._render_year()

    def _next_year(self):
        self.year += 1
        self._render_year()


def abrir_calendario(master: tk.Misc | None = None):
    CalendarioANS(master)


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    abrir_calendario(root)
    root.mainloop()
