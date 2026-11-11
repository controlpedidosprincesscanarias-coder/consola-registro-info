# -*- coding: utf-8 -*-
"""
ventana_consulta_manual.py
«Consultar el manual» (nov 2026): muestra los apartados del manual que mejor encajan con lo que se
escribe, con un fragmento y su página, y permite abrir el manual o guardar una copia. La búsqueda
está en app/comun/manual.py.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import tkinter as tk
from tkinter import filedialog, messagebox

from app.comun import manual
from app.comun.tarifa_constants import (
    C_BG, C_MID, C_MUTED, C_TEXT, FONT_BOLD, FONT_LABEL, FONT_SMALL, FONT_SUB, FONT_TITLE,
)
from app.comun.tarifa_core import cargar_icono_ventana, centrar_sobre_padre

_DESTACADO = "#0A4A7A"


class VentanaConsultaManual(tk.Toplevel):
    def __init__(self, app, consulta="", log=None):
        super().__init__(app)
        self._log = log or (lambda *a, **k: None)
        self._indice = manual.cargar_indice()
        self.title("Consultar el manual")
        self.configure(bg=C_BG)
        self.minsize(620, 460)
        cargar_icono_ventana(self)

        hdr = tk.Frame(self, bg=C_MID, pady=12, padx=18)
        hdr.pack(fill="x")
        tk.Label(hdr, text="\U0001f50e  Consultar el manual de usuario", bg=C_MID, fg="white",
                 font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr, text="Escribe tu duda con tus palabras y pulsa Intro. Se buscan los apartados del manual.",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        barra = tk.Frame(self, bg=C_BG, padx=16, pady=10)
        barra.pack(fill="x")
        self._var = tk.StringVar(value=consulta)
        ent = tk.Entry(barra, textvariable=self._var, font=FONT_LABEL, relief="solid", bd=1)
        ent.pack(side="left", fill="x", expand=True, ipady=4)
        ent.bind("<Return>", lambda e: self._buscar())
        tk.Button(barra, text="Buscar", command=self._buscar, font=FONT_BOLD, bg=_DESTACADO, fg="white",
                  relief="flat", padx=14, pady=4, cursor="hand2").pack(side="left", padx=(8, 0))

        cont = tk.Frame(self, bg=C_BG, padx=16)
        cont.pack(fill="both", expand=True)
        self._lienzo = tk.Canvas(cont, bg=C_BG, highlightthickness=0)
        barra_v = tk.Scrollbar(cont, orient="vertical", command=self._lienzo.yview)
        self._lista = tk.Frame(self._lienzo, bg=C_BG)
        self._lista.bind("<Configure>", lambda e: self._lienzo.configure(scrollregion=self._lienzo.bbox("all")))
        self._ventana_lista = self._lienzo.create_window((0, 0), window=self._lista, anchor="nw")
        self._lienzo.bind("<Configure>", lambda e: self._lienzo.itemconfigure(self._ventana_lista, width=e.width))
        self._lienzo.configure(yscrollcommand=barra_v.set)
        self._lienzo.pack(side="left", fill="both", expand=True)
        barra_v.pack(side="right", fill="y")

        pie = tk.Frame(self, bg=C_BG, padx=16, pady=12)
        pie.pack(fill="x")
        tk.Button(pie, text="\U0001f4d8  Abrir el manual", command=self._abrir, font=FONT_BOLD, bg=_DESTACADO,
                  fg="white", relief="flat", padx=14, pady=6, cursor="hand2").pack(side="right")
        tk.Button(pie, text="\U0001f4be  Guardar una copia…", command=self._guardar, font=FONT_LABEL,
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="right", padx=(0, 8))
        tk.Button(pie, text="Cerrar", command=self.destroy, font=FONT_LABEL, relief="flat", padx=14,
                  pady=6, cursor="hand2").pack(side="left")

        self.geometry("720x560")
        self.transient(app)
        centrar_sobre_padre(self, app)
        ent.focus_set()
        ent.icursor("end")
        self._buscar()

    def _buscar(self):
        for w in self._lista.winfo_children():
            w.destroy()
        consulta = self._var.get().strip()
        if not self._indice:
            self._mensaje("No se ha encontrado el índice del manual dentro de la aplicación. "
                          "Puedes abrir el manual completo con el botón de abajo.")
            return
        if not consulta:
            self._mensaje("Por ejemplo: «crear registro info», «códigos SAP faltantes», «actualizar listado "
                          "SAP», «tarifa incompleta», «contraseña Catálogo Asignaciones».")
            return
        resultados = manual.buscar(self._indice, consulta)
        self._log(f"Consulta al manual: «{consulta}» — {len(resultados)} apartado(s).", "INFO")
        if not resultados:
            self._mensaje("No se ha encontrado ningún apartado con esas palabras. Prueba con otras "
                          "(por ejemplo, el nombre de la acción o de la ventana) o abre el manual completo.")
            return
        for r in resultados:
            self._tarjeta(r)

    def _mensaje(self, texto):
        tk.Label(self._lista, text=texto, bg=C_BG, fg=C_MUTED, font=FONT_LABEL, justify="left",
                 wraplength=640).pack(anchor="w", pady=10)

    def _tarjeta(self, r):
        caja = tk.Frame(self._lista, bg="white", highlightbackground="#B9C7DA", highlightthickness=1,
                        padx=12, pady=8)
        caja.pack(fill="x", pady=(0, 8))
        pagina = f"  ·  página {r['pagina']}" if r.get("pagina") else ""
        tk.Label(caja, text=f"{r['capitulo']}{pagina}", bg="white", fg=C_MUTED, font=FONT_SMALL,
                 anchor="w").pack(fill="x")
        tk.Label(caja, text=r["titulo"], bg="white", fg=C_TEXT, font=FONT_BOLD, anchor="w",
                 justify="left", wraplength=620).pack(fill="x")
        if r.get("fragmento"):
            tk.Label(caja, text=r["fragmento"], bg="white", fg=C_TEXT, font=FONT_LABEL, anchor="w",
                     justify="left", wraplength=620).pack(fill="x", pady=(2, 0))

    def _abrir(self):
        try:
            manual.abrir_manual()
            self._log("Abriendo el manual de usuario...", "INFO")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Manual de usuario", f"No se pudo abrir el manual:\n{e}", parent=self)

    def _guardar(self):
        destino = filedialog.asksaveasfilename(parent=self, title="Guardar una copia del manual",
                                               initialfile=manual.NOMBRE_PDF, defaultextension=".pdf",
                                               filetypes=[("PDF", "*.pdf")])
        if not destino:
            return
        try:
            manual.guardar_copia(destino)
            self._log(f"Manual guardado en: {destino}", "OK")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Manual de usuario", f"No se pudo guardar:\n{e}", parent=self)
