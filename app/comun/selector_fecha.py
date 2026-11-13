# -*- coding: utf-8 -*-
"""
selector_fecha.py — calendario 📅 para elegir una fecha sin escribirla (13 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Petición de Víctor: en «Fechas de validez del Registro Info», un calendario con su desplegable para no tener que
escribir las fechas. La fecha elegida se escribe en la casilla con el mismo formato de siempre («dd.mm.aaaa»), así
que lo que se pega en el RINFO no cambia. La casilla se puede seguir escribiendo a mano.

Sin dependencias nuevas (solo tkinter y el módulo calendar). La lógica (`semanas`, `leer`, `formatear`) se prueba en
tests/test_selector_fecha.py.
"""
import calendar
import re
import tkinter as tk
from datetime import date
from tkinter import ttk

MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
         "noviembre", "diciembre")
DIAS = ("L", "M", "X", "J", "V", "S", "D")
_RE = re.compile(r"^\s*(\d{1,2})[-./, ]?(\d{1,2})[-./, ]?(\d{4})\s*$")


def formatear(d):
    """date → «dd.mm.aaaa» (el formato que se pega en el RINFO)."""
    return d.strftime("%d.%m.%Y")


def leer(texto, por_defecto=None):
    """«01.10.2026», «01-10-2026», «01102026»… → date; si no es una fecha válida, `por_defecto`."""
    m = _RE.match(str(texto or ""))
    if not m:
        return por_defecto
    try:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return por_defecto


def semanas(anio, mes):
    """Las semanas del mes (lunes primero), con 0 en los días de otros meses."""
    return calendar.Calendar(firstweekday=0).monthdayscalendar(anio, mes)


class Calendario(tk.Toplevel):
    """Ventanita con desplegables de mes y año, ◀ ▶ y los días. Al pulsar un día llama a `al_elegir(date)`."""

    C_CAB, C_HOY, C_SEL = "#285FA8", "#FFF3C4", "#1A7A40"

    def __init__(self, parent, fecha, al_elegir, anclaje=None):
        super().__init__(parent)
        self.withdraw()
        self.overrideredirect(False)
        self.title("Elegir fecha")
        self.resizable(False, False)
        self.configure(bg="white")
        self.transient(parent)
        self._al_elegir = al_elegir
        self._sel = fecha or date.today()
        self._var_mes = tk.StringVar(value=MESES[self._sel.month - 1])
        self._var_anio = tk.StringVar(value=str(self._sel.year))
        self._construir()
        self._pintar()
        self.bind("<Escape>", lambda _e: self.destroy())
        self.update_idletasks()
        if anclaje is not None:
            x, y = anclaje.winfo_rootx(), anclaje.winfo_rooty() + anclaje.winfo_height() + 2
            self.geometry(f"+{x}+{y}")
        self.deiconify()
        self.grab_set()
        self.focus_set()

    def _construir(self):
        cab = tk.Frame(self, bg=self.C_CAB, padx=6, pady=6)
        cab.pack(fill="x")
        tk.Button(cab, text="◀", width=3, relief="flat", command=lambda: self._mover(-1)).pack(side="left")
        self._cb_mes = ttk.Combobox(cab, textvariable=self._var_mes, values=MESES, state="readonly", width=11)
        self._cb_mes.pack(side="left", padx=4)
        hoy = date.today().year
        self._cb_anio = ttk.Combobox(cab, textvariable=self._var_anio, state="readonly", width=6,
                                     values=[str(a) for a in range(hoy - 5, hoy + 11)])
        self._cb_anio.pack(side="left")
        tk.Button(cab, text="▶", width=3, relief="flat", command=lambda: self._mover(1)).pack(side="left", padx=(4, 0))
        self._cb_mes.bind("<<ComboboxSelected>>", lambda _e: self._pintar())
        self._cb_anio.bind("<<ComboboxSelected>>", lambda _e: self._pintar())
        self._dias = tk.Frame(self, bg="white", padx=6, pady=6)
        self._dias.pack()
        pie = tk.Frame(self, bg="white", padx=6, pady=(0))
        pie.pack(fill="x", pady=(0, 6))
        tk.Button(pie, text="Hoy", relief="flat", bg="#E8EEF6", command=lambda: self._elegir(date.today())).pack(
            side="left", padx=6)
        tk.Button(pie, text="Cancelar", relief="flat", bg="#E8EEF6", command=self.destroy).pack(side="right", padx=6)

    def _mes_anio(self):
        return MESES.index(self._var_mes.get()) + 1, int(self._var_anio.get())

    def _mover(self, paso):
        mes, anio = self._mes_anio()
        mes += paso
        if mes < 1:
            mes, anio = 12, anio - 1
        elif mes > 12:
            mes, anio = 1, anio + 1
        self._var_mes.set(MESES[mes - 1])
        self._var_anio.set(str(anio))
        self._pintar()

    def _pintar(self):
        for w in self._dias.winfo_children():
            w.destroy()
        mes, anio = self._mes_anio()
        for c, d in enumerate(DIAS):
            tk.Label(self._dias, text=d, bg="white", fg="#666", width=4, font=("Segoe UI", 8, "bold")).grid(
                row=0, column=c)
        hoy = date.today()
        for f, semana in enumerate(semanas(anio, mes), start=1):
            for c, dia in enumerate(semana):
                if not dia:
                    continue
                d = date(anio, mes, dia)
                fondo = self.C_SEL if d == self._sel else self.C_HOY if d == hoy else "white"
                tk.Button(self._dias, text=str(dia), width=3, relief="flat", bg=fondo,
                          fg="white" if d == self._sel else "#1A1A2E", cursor="hand2",
                          command=lambda d=d: self._elegir(d)).grid(row=f, column=c, padx=1, pady=1)

    def _elegir(self, d):
        self._al_elegir(d)
        self.destroy()


def boton_calendario(parent, var, **opciones):
    """Botón 📅 que abre el calendario y escribe la fecha elegida en `var` («dd.mm.aaaa»)."""
    boton = tk.Button(parent, text="📅", relief="flat", cursor="hand2", **opciones)

    def _abrir():
        Calendario(boton.winfo_toplevel(), leer(var.get(), date.today()), lambda d: var.set(formatear(d)),
                   anclaje=boton)
    boton.configure(command=_abrir)
    return boton
