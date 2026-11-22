# -*- coding: utf-8 -*-
"""
ventana_plantilla_rinfo.py — Acción 6, «Generar plantilla Tarifa desde RINFO» (22 nov 2026, petición de Víctor).

Ventana para elegir el RINFO o el Purchasing Info Records: arrastrándolo desde el Explorador de Windows a la ventana
(app/ui/dnd_windows.py, el mismo mecanismo que el chat del Organizador) o con «Examinar…». Se pueden elegir varios.
Al elegirlos, la ventana se cierra y la Consola genera las plantillas en segundo plano; todo lo demás va a la consola
de log (regla de la Consola: ventanas solo para decidir o para errores que impiden seguir).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import tkinter as tk
from tkinter import filedialog

from app.comun.tarifa_constants import C_BG, C_MID, C_TEXT, FONT_BOLD, FONT_LABEL, FONT_SUB, FONT_TITLE
from app.comun.tarifa_core import ajustar_geometria_pantalla, cargar_icono_ventana
from app.ui.dnd_windows import hook_dropfiles

TITULO = "Generar plantilla Tarifa desde RINFO"


def solo_excel(rutas):
    """Las rutas que son libros .xlsx (sin los temporales «~$» de Excel), sin repetir."""
    vistas, salida = set(), []
    for r in rutas or ():
        nombre = os.path.basename(r)
        if nombre.startswith("~$") or not nombre.lower().endswith(".xlsx"):
            continue
        clave = os.path.normcase(os.path.abspath(r))
        if clave not in vistas:
            vistas.add(clave)
            salida.append(r)
    return salida


class VentanaPlantillaRinfo(tk.Toplevel):
    """on_archivos(lista_de_rutas) se llama con los .xlsx elegidos (al menos uno), después de cerrar la ventana."""

    def __init__(self, parent, carpeta_inicial, on_archivos):
        super().__init__(parent)
        self.title(TITULO)
        self.configure(bg=C_BG)
        self.resizable(False, False)
        self._carpeta = carpeta_inicial if carpeta_inicial and os.path.isdir(carpeta_inicial) else None
        self._on_archivos = on_archivos
        cargar_icono_ventana(self)
        self._construir_ui()
        ajustar_geometria_pantalla(self, 560)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.arrastrar_activo = hook_dropfiles(self, self._recibir)
        if not self.arrastrar_activo:
            self._lbl_zona_sub.config(text="(pulse «Examinar…» para elegirlos)")
        self.grab_set()

    def _construir_ui(self):
        hdr = tk.Frame(self, bg=C_MID, pady=10, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📄  " + TITULO, bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr, text="Plantilla para que el proveedor rellene la nueva revisión de precios.",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(3, 0))

        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=16)
        cuerpo.pack(fill="both", expand=True)
        tk.Label(cuerpo, justify="left", wraplength=510, bg=C_BG, fg=C_TEXT, font=FONT_LABEL,
                 text=("Se genera en TARIFAS una tarifa por proveedor con sus artículos, el precio vigente, el "
                       "teléfono y los hoteles. Si un proveedor tiene precios distintos según el hotel, sale una "
                       "tarifa por cada grupo de hoteles con los mismos precios.\n\n"
                       "El proveedor rellena las columnas amarillas PRECIO NUEVO y FECHA INICIO. Cuando la "
                       "devuelva, guárdela en TARIFAS y úsela en «Crear / Actualizar Registro Info SAP».")
                 ).pack(anchor="w", pady=(0, 12))

        self._zona = tk.Frame(cuerpo, bg="#EAF1FA", highlightthickness=2, highlightbackground="#7FA6D6",
                              padx=10, pady=22)
        self._zona.pack(fill="x", pady=(0, 12))
        tk.Label(self._zona, text="⬇", bg="#EAF1FA", fg="#3A6EA5", font=("Segoe UI", 22)).pack()
        tk.Label(self._zona, text="Arrastre aquí el RINFO o el Purchasing Info Records (.xlsx)",
                 bg="#EAF1FA", fg=C_TEXT, font=FONT_BOLD).pack()
        self._lbl_zona_sub = tk.Label(self._zona, text="(uno o varios a la vez)", bg="#EAF1FA", fg="#4A6A8D",
                                      font=FONT_SUB)
        self._lbl_zona_sub.pack()

        self._lbl_error = tk.Label(cuerpo, text="", bg=C_BG, fg="#C0392B", font=("Segoe UI", 8, "italic"),
                                   wraplength=510, justify="left")
        self._lbl_error.pack(anchor="w")

        botones = tk.Frame(cuerpo, bg=C_BG)
        botones.pack(fill="x", pady=(6, 0))
        tk.Button(botones, text="📂  Examinar…", command=self._examinar, bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=14, pady=8, cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(side="left")
        tk.Button(botones, text="Cancelar", command=self.destroy, font=FONT_LABEL, padx=12, pady=8,
                  cursor="hand2", relief="flat", bg="#D8DEE8", activebackground="#C0C8D8").pack(side="right")

    def _examinar(self):
        rutas = filedialog.askopenfilenames(
            parent=self, title="Elegir RINFO o Purchasing Info Records", initialdir=self._carpeta,
            filetypes=[("Libros de Excel", "*.xlsx"), ("Todos los archivos", "*.*")])
        if rutas:
            self._recibir(list(rutas))

    def _recibir(self, rutas):
        validas = solo_excel(rutas)
        if not validas:
            self._lbl_error.config(text="⚠  Solo se aceptan libros de Excel .xlsx (un RINFO o un Purchasing Info "
                                        "Records descargado de SAP).")
            return
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()
        self._on_archivos(validas)
