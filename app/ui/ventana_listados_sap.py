# -*- coding: utf-8 -*-
"""
ventana_listados_sap.py — «Faltan los listados de SAP actualizados» (13 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Sale después de guardar códigos SAP (Acción 8 o códigos SAP faltantes de la Acción 4) cuando alguno no está en
Productos.xlsx. Explica, con imágenes, cómo descargar de SAP Productos.xlsx («Gestionar datos maestros de producto»)
y Listas de productos.xlsx («Lista de productos») y guardarlos en RECURSOS. Opciones:
  - «Ya están en RECURSOS: actualizar ahora» → resultado "actualizar" (la app lanza la Acción 3);
  - «Más tarde» (o cerrar) → avisa de que la web queda parcialmente actualizada; resultado "despues".
La lógica y los textos están en app/comun/listados_sap.py.
"""
import os
import sys
import time
import tkinter as tk
from tkinter import messagebox, ttk

from app.comun import listados_sap as L
from app.comun.tarifa_constants import (
    C_ACCENT, C_ACCENT2, C_BORDER, C_CARD, C_DARK, C_HOVER, C_MID, C_MUTED, C_SUCCESS, C_TEXT_LIGHT, C_WARNING,
    FONT_BOLD, FONT_LABEL, FONT_SMALL, FONT_TITLE, resource_path,
)
from app.comun.tarifa_core import cargar_icono_ventana


class VentanaListadosSap(tk.Toplevel):
    def __init__(self, parent, codigos_fuera, carpeta_recursos, ahora=None):
        super().__init__(parent)
        self.title("Faltan los listados de SAP actualizados")
        self.configure(bg=C_MID)
        self.transient(parent)
        cargar_icono_ventana(self)
        self.resultado = "despues"
        self._codigos = list(codigos_fuera)
        self._recursos = carpeta_recursos
        self._desde = time.time() if ahora is None else ahora
        self._imagenes = []
        self._construir()
        self.protocol("WM_DELETE_WINDOW", self._mas_tarde)
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(980, sw - 40), min(820, sh - 60)
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(5, (sh - h) // 5)}")
        self.grab_set()
        self._vigilar()

    # ── Interfaz ─────────────────────────────────────────────────────────────────────────────────────
    def _imagen(self, ruta, reducir=2):
        try:
            img = tk.PhotoImage(file=resource_path(ruta))
            if reducir > 1:
                img = img.subsample(reducir)
            self._imagenes.append(img)
            return img
        except tk.TclError:
            return None

    def _construir(self):
        hdr = tk.Frame(self, bg=C_DARK, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⚠  Faltan los listados de SAP actualizados", font=FONT_TITLE, fg=C_ACCENT,
                 bg=C_DARK).pack(side="left", padx=14)

        n = len(self._codigos)
        lista = ", ".join(self._codigos[:12]) + (" …" if n > 12 else "")
        tk.Label(self, text=(f"Has asignado {n} código(s) SAP que todavía no están en Productos.xlsx: {lista}.\n"
                             "Los códigos ya se envían a Catálogo Asignaciones, pero para que toda la información "
                             "quede registrada (descripción, unidad,\ngrupos y categoría de valoración) hay que "
                             "descargar de SAP estos dos listados, guardarlos en RECURSOS y actualizar con la "
                             "Acción 3."),
                 font=FONT_LABEL, fg=C_TEXT_LIGHT, bg=C_MID, justify="left").pack(anchor="w", padx=14, pady=(8, 4))

        # Botones (antes que la guía, para que siempre se vean)
        abajo = tk.Frame(self, bg=C_MID, pady=10)
        abajo.pack(fill="x", padx=14, side="bottom")
        self._btn_ya = tk.Button(abajo, text="✔  Ya están en RECURSOS: actualizar ahora (Acción 3)", font=FONT_BOLD,
                                 fg=C_DARK, bg=C_ACCENT, activebackground=C_ACCENT2, relief="flat", cursor="hand2",
                                 command=self._ya_estan)
        self._btn_ya.pack(side="left", ipadx=10, ipady=4)
        tk.Button(abajo, text="📂  Abrir RECURSOS", font=FONT_BOLD, fg=C_TEXT_LIGHT, bg=C_BORDER,
                  activebackground=C_HOVER, relief="flat", cursor="hand2", command=self._abrir_recursos).pack(
            side="left", padx=8, ipadx=8, ipady=4)
        tk.Button(abajo, text="Más tarde", font=FONT_BOLD, fg=C_TEXT_LIGHT, bg=C_BORDER, activebackground=C_HOVER,
                  relief="flat", cursor="hand2", command=self._mas_tarde).pack(side="left", ipadx=10, ipady=4)
        self._lbl_estado = tk.Label(self, text="", font=FONT_LABEL, fg=C_MUTED, bg=C_MID, justify="left")
        self._lbl_estado.pack(fill="x", padx=14, side="bottom")

        # Guía con scroll
        caja = tk.Frame(self, bg=C_MID)
        caja.pack(fill="both", expand=True, padx=10, pady=4)
        canvas = tk.Canvas(caja, bg=C_CARD, highlightthickness=0)
        barra = ttk.Scrollbar(caja, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=barra.set)
        barra.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        guia = tk.Frame(canvas, bg=C_CARD, padx=14, pady=10)
        canvas.create_window((0, 0), window=guia, anchor="nw")
        guia.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Enter>", lambda _e: self.bind_all(
            "<MouseWheel>", lambda ev: canvas.yview_scroll(int(-ev.delta / 120) or 1, "units")))
        canvas.bind("<Leave>", lambda _e: self.unbind_all("<MouseWheel>"))
        self._guia = guia

        def titulo(texto):
            tk.Label(guia, text=texto, font=FONT_BOLD, fg=C_ACCENT2, bg=C_CARD, anchor="w").pack(
                anchor="w", pady=(10, 4))

        titulo("1.  En SAP, la aplicación de cada listado")
        for L_ in L.LISTADOS:
            tk.Label(guia, text=f"   • {L_['archivo']}  ←  «{L_['app_sap']}»", font=FONT_LABEL, fg=C_TEXT_LIGHT,
                     bg=C_CARD, anchor="w").pack(anchor="w")
        img = self._imagen("assets/guia_sap_apps.png")
        if img:
            tk.Label(guia, image=img, bg=C_CARD).pack(anchor="w", pady=(6, 0))

        titulo("2.  Pulsa «Ir» y, a la derecha de la barra de la tabla, el botón ⌄ de exportar (el último, "
               "«Abrir menú») → «Exportar como…»")
        img = self._imagen("assets/guia_sap_barra.png")
        if img:
            tk.Label(guia, image=img, bg=C_CARD).pack(anchor="w")

        titulo("3.  Nombre de fichero y «Exportar» (formato Microsoft Excel *.xlsx)")
        fila = tk.Frame(guia, bg=C_CARD)
        fila.pack(anchor="w")
        for i, L_ in enumerate(L.LISTADOS):
            col = tk.Frame(fila, bg=C_CARD)
            col.grid(row=0, column=i, padx=(0, 18), sticky="n")
            tk.Label(col, text=f"{L_['archivo']}  →  nombre «{L_['nombre']}»", font=FONT_LABEL, fg=C_TEXT_LIGHT,
                     bg=C_CARD).pack(anchor="w")
            img = self._imagen(L_["imagen"])
            if img:
                tk.Label(col, image=img, bg=C_CARD).pack(anchor="w", pady=(4, 0))

        titulo("4.  Guarda los dos archivos en la carpeta RECURSOS del proyecto (sustituyendo a los anteriores)")
        tk.Label(guia, text=f"   {self._recursos}", font=FONT_SMALL, fg=C_MUTED, bg=C_CARD, anchor="w").pack(
            anchor="w")
        tk.Label(guia, text="   Después pulsa «Ya están en RECURSOS»: la Consola lanza sola la Acción 3 y completa "
                            "los datos SAP en el catálogo y en la web.",
                 font=FONT_LABEL, fg=C_TEXT_LIGHT, bg=C_CARD, anchor="w").pack(anchor="w", pady=(4, 0))

    # ── Estado de los archivos ───────────────────────────────────────────────────────────────────────
    def _texto_estado(self):
        estado = L.estado_archivos(self._recursos, self._desde)
        partes = []
        for nombre, (existe, actualizado) in estado.items():
            partes.append(f"{'✔' if actualizado else '○'} {nombre}: "
                          + ("descargado de nuevo" if actualizado else "sin descargar de nuevo" if existe
                             else "no está en RECURSOS"))
        return "    ".join(partes), not L.sin_actualizar(estado)

    def _vigilar(self):
        if not self.winfo_exists():
            return
        texto, listos = self._texto_estado()
        self._lbl_estado.configure(text=texto, fg=C_SUCCESS if listos else C_WARNING)
        self.after(2000, self._vigilar)

    # ── Acciones ─────────────────────────────────────────────────────────────────────────────────────
    def _abrir_recursos(self):
        try:
            if sys.platform.startswith("win"):
                os.startfile(self._recursos)  # noqa: S606  (carpeta del proyecto)
        except OSError as e:
            messagebox.showwarning("RECURSOS", f"No se pudo abrir la carpeta:\n{e}", parent=self)

    def _ya_estan(self):
        faltan = L.sin_actualizar(L.estado_archivos(self._recursos, self._desde))
        if faltan and not messagebox.askyesno(
                "Listados de SAP",
                "No veo descargados de nuevo en RECURSOS:\n\n   " + "\n   ".join(faltan)
                + "\n\n¿Actualizar igualmente con la Acción 3?", parent=self):
            return
        self.resultado = "actualizar"
        self.destroy()

    def _mas_tarde(self):
        messagebox.showwarning("Web parcialmente actualizada", L.AVISO_PARCIAL, parent=self)
        self.resultado = "despues"
        self.destroy()
