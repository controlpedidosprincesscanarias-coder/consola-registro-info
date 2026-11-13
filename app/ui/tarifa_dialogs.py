# -*- coding: utf-8 -*-
"""
tarifa_dialogs.py
Ventanas modales (Toplevel) de la aplicación TARIFA
Central Compras Princess Canarias
Autor: Victor Alejandro Martin Andrade (VAMA)
(c) VAMA 2026

Clases exportadas:
    VentanaSeleccionXLS          — selección de archivo .xls a renombrar
    VentanaCrearProyecto         — crear estructura PROYECTO/TARIFAS/RECURSOS
    VentanaResolverSAPDuplicados — elegir el SAP correcto cuando hay duplicados
    VentanaSAPFaltantes          — entrada manual de códigos SAP faltantes
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox
import threading

from app.comun.tarifa_constants import (
    C_DARK, C_MID, C_CARD, C_BORDER, C_HOVER, C_ACCENT,
    C_ACCENT2, C_BG, C_PANEL,
    C_TEXT, C_TEXT_LIGHT, C_MUTED, C_SUCCESS, FONT_TITLE,
    FONT_SUB, FONT_LABEL, FONT_BOLD,
    FONT_SMALL, FONT_TINY,
)
from app.comun.tarifa_core import (
    renombrar_a_tarifa, crear_estructura_proyecto, FOLDER_TARIFAS,
    FOLDER_RECURSOS, cargar_icono_ventana, centrar_sobre_padre,
)

class VentanaSeleccionXLS(tk.Toplevel):
    def __init__(self, parent, archivos, carpeta_tarifas, consola):
        super().__init__(parent)
        self.title("Seleccionar archivo de tarifa")
        self.resizable(True, True)
        self.minsize(560, 340)
        self.configure(bg=C_BG)
        cargar_icono_ventana(self)
        self.carpeta_tarifas = carpeta_tarifas
        self.consola         = consola
        self.seleccionado    = None
        self._build_ui(archivos)
        self.grab_set()
        self.transient(parent)
        centrar_sobre_padre(self, parent)


    def _build_ui(self, archivos):
        frm_top = tk.Frame(self, bg=C_MID, pady=12)
        frm_top.pack(fill="x")
        tk.Label(frm_top, text="\u26a0\ufe0f  TARIFA.xls no encontrada",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack()
        tk.Label(frm_top,
                 text="Seleccione el archivo a renombrar como TARIFA.xls:",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(pady=(2, 0))

        frm_tabla = tk.Frame(self, bg=C_BG, padx=14, pady=10)
        frm_tabla.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("Princess.Treeview",
                         background=C_PANEL, fieldbackground=C_PANEL,
                         foreground=C_TEXT, rowheight=22, font=FONT_LABEL)
        style.configure("Princess.Treeview.Heading",
                         background=C_MID, foreground="white",
                         font=FONT_BOLD, relief="flat")
        style.map("Princess.Treeview",
                   background=[("selected", C_MID)],
                   foreground=[("selected", "white")])

        cols = ("nombre", "fecha_mod")
        self.tree = ttk.Treeview(frm_tabla, columns=cols, show="headings",
                                  selectmode="browse", style="Princess.Treeview")
        self.tree.heading("nombre",    text="Nombre del archivo")
        self.tree.heading("fecha_mod", text="\xdaltima modificaci\xf3n")
        self.tree.column("nombre",    width=360, anchor="w")
        self.tree.column("fecha_mod", width=130, anchor="center")

        sb = ttk.Scrollbar(frm_tabla, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.tree.tag_configure("par",   background="#F0F5FB")
        self.tree.tag_configure("impar", background="white")

        if archivos:
            for i, a in enumerate(archivos):
                tag = "par" if i % 2 == 0 else "impar"
                self.tree.insert("", "end", iid=a["ruta"],
                                  values=(a["nombre"], a["fecha_mod"]), tags=(tag,))
            self.tree.selection_set(archivos[0]["ruta"])
        else:
            tk.Label(frm_tabla,
                     text="No se encontraron archivos .xls en la carpeta TARIFAS.",
                     fg=C_MUTED, bg=C_BG, font=FONT_LABEL).pack(expand=True)

        self.tree.bind("<Double-1>", lambda e: self._confirmar())

        frm_btn = tk.Frame(self, bg=C_BG, padx=14, pady=10)
        frm_btn.pack(fill="x", side="bottom")

        tk.Button(frm_btn, text="\u2714  Usar este archivo",
                  command=self._confirmar,
                  bg=C_ACCENT, fg=C_DARK, font=FONT_BOLD,
                  padx=14, pady=6, cursor="hand2", relief="flat",
                  activebackground=C_ACCENT2).pack(side="right", padx=(6, 0))

        tk.Button(frm_btn, text="Cancelar", command=self.destroy,
                  font=FONT_LABEL, padx=14, pady=6, cursor="hand2",
                  relief="flat", bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(side="right")

    def _confirmar(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sin selecci\xf3n",
                                   "Seleccione un archivo de la lista.", parent=self)
            return

        ruta_origen = sel[0]
        nombre = os.path.basename(ruta_origen)

        try:
            destino = renombrar_a_tarifa(ruta_origen, self.carpeta_tarifas)
            self.consola.log(f"Archivo renombrado: {nombre} \u2192 TARIFA.xls", "OK")
            self.seleccionado = destino
            self.destroy()
            self.master.after(100, self.master._buscar)

        except Exception as e:
            self.consola.log(f"Error al renombrar: {e}", "ERROR")
            messagebox.showerror("Error al renombrar",
                                 f"No se pudo completar la operaci\xf3n:\n{e}",
                                 parent=self)


# ---------------------------------------------------------
#  VENTANA: CREAR ESTRUCTURA DE PROYECTO
# ---------------------------------------------------------
class VentanaCrearProyecto(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Crear estructura de proyecto")
        self.resizable(False, False)
        self.configure(bg=C_BG)
        self.ruta_creada = None
        self._build_ui()
        self.grab_set()
        self.transient(parent)
        centrar_sobre_padre(self, parent)


    def _build_ui(self):
        # Cabecera
        hdr = tk.Frame(self, bg=C_MID, pady=14, padx=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text="\U0001f4c2  Crear estructura de carpetas",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Se crear\xe1 la siguiente estructura en el escritorio:",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        # Diagrama de carpetas
        frm_tree = tk.Frame(self, bg="#EEF2F8", padx=20, pady=16)
        frm_tree.pack(fill="x", padx=16, pady=12)

        lineas = [
            ("\U0001f4c1 Escritorio/", C_MID, FONT_BOLD),
            ("  \u2514\u2500 \U0001f4c1 PROYECTO/", C_DARK, FONT_BOLD),
            ("       \u251c\u2500 \U0001f4c2 TARIFAS/    \u2190 aqui se coloca TARIFA.xls", C_SUCCESS, FONT_LABEL),
            ("       \u2514\u2500 \U0001f4c2 RECURSOS/", C_MID, FONT_LABEL),
        ]
        for texto, color, fuente in lineas:
            tk.Label(frm_tree, text=texto, bg="#EEF2F8", fg=color,
                     font=fuente, anchor="w", justify="left").pack(fill="x")

        # Nota
        frm_nota = tk.Frame(self, bg=C_BG, padx=16)
        frm_nota.pack(fill="x")
        tk.Label(frm_nota,
                 text="\u2139\ufe0f  Si la carpeta PROYECTO ya existe, se respetan los archivos existentes.",
                 bg=C_BG, fg=C_MUTED, font=FONT_SMALL, wraplength=380, justify="left").pack(anchor="w")

        # Botones
        frm_btn = tk.Frame(self, bg=C_BG, padx=16, pady=14)
        frm_btn.pack(fill="x")

        tk.Button(frm_btn, text="\u2714  Crear carpetas",
                  command=self._crear,
                  bg=C_SUCCESS, fg="white", font=FONT_BOLD,
                  padx=14, pady=7, cursor="hand2", relief="flat",
                  activebackground=C_DARK).pack(side="right", padx=(6, 0))

        tk.Button(frm_btn, text="Cancelar", command=self.destroy,
                  font=FONT_LABEL, padx=14, pady=7, cursor="hand2",
                  relief="flat", bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(side="right")

    def _crear(self):
        escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
        # Fallback para Windows si 'Desktop' no existe con ese nombre
        if not os.path.isdir(escritorio):
            escritorio = os.path.join(os.path.expanduser("~"), "Escritorio")
        if not os.path.isdir(escritorio):
            escritorio = os.path.expanduser("~")

        try:
            ruta = crear_estructura_proyecto(escritorio)
            self.ruta_creada = ruta
            messagebox.showinfo(
                "Carpetas creadas",
                f"Estructura creada correctamente en:\n{ruta}\n\n"
                f"  \u2022 {os.path.join(ruta, FOLDER_TARIFAS)}\n"
                f"  \u2022 {os.path.join(ruta, FOLDER_RECURSOS)}",
                parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron crear las carpetas:\n{e}", parent=self)


# ---------------------------------------------------------
#  VENTANA PRINCIPAL
# ---------------------------------------------------------

# ============================================================
#  VENTANA ENTRADA CODIGOS SAP FALTANTES
# ============================================================

class VentanaResolverSAPDuplicados(tk.Toplevel):
    """
    Se muestra cuando VentanaSAPFaltantes._guardar() detecta que alguno
    de los códigos SAP recién introducidos ya está asignado a otro
    código DALI -- bien porque ya existía en el catálogo (asignado a un
    DALI distinto), bien porque se ha tecleado el mismo SAP para dos
    DALI distintos dentro del propio lote que se está guardando ahora.

    El usuario marca con casillas cuáles DALI hay que "borrar" (dejar
    sin ese SAP) de cada conflicto -- el que queda sin marcar es el
    que se considera correcto y conserva el código.

    Tras cerrarse, self.resultado es:
      - None si el usuario canceló todo el guardado.
      - Un set de códigos DALI (int) cuyo SAP hay que borrar/omitir,
        si confirmó la resolución.
    """
    def __init__(self, parent, conflictos):
        """
        conflictos: dict {cod_sap: [(cod_dali, es_existente, fila_o_None,
                     descripcion), ...]} -- una entrada por cada DALI que
                     compite por ese mismo código SAP. `es_existente=True`
                     significa que ese DALI ya tenía este SAP asignado en
                     el catálogo antes de este guardado (fila_o_None es su
                     fila real); `es_existente=False` significa que es uno
                     de los códigos que se están tecleando ahora mismo.
        """
        super().__init__(parent)
        self.title("Códigos SAP duplicados — hay que corregir")
        self.configure(bg=C_MID)
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

        self._conflictos = conflictos
        self.resultado = None
        self._checks = {}   # {(cod_sap, cod_dali): tk.BooleanVar}

        self._build_ui()
        cargar_icono_ventana(self)

        # Centrado sobre pantalla (tamaño variable según nº de conflictos)
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = min(760, sw - 60)
        h = min(self.winfo_reqheight() + 20, sh - 100)
        x = (sw - w) // 2
        y = max(20, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ----------------------------------------------------------
    def _build_ui(self):
        frm_hdr = tk.Frame(self, bg=C_DARK, pady=10)
        frm_hdr.pack(fill="x")
        tk.Label(frm_hdr,
                 text="⚠  Códigos SAP duplicados",
                 font=FONT_TITLE, fg=C_ACCENT, bg=C_DARK).pack(side="left", padx=14)

        tk.Label(self,
                 text="Alguno de los códigos SAP introducidos ya está en uso por otro "
                      "código DALI. Marca cuál de los dos (o más) es el incorrecto -- "
                      "se borrará su código SAP, y el que quede sin marcar conservará el suyo.\n"
                      "En cada grupo debe quedar al menos un DALI sin marcar.",
                 font=FONT_LABEL, fg=C_TEXT_LIGHT, bg=C_MID, justify="left",
                 wraplength=700).pack(anchor="w", padx=14, pady=(10, 6))

        # ── Botones (empaquetados antes de la zona con scroll, para que
        #    queden siempre visibles en la parte inferior) ──
        frm_btns = tk.Frame(self, bg=C_MID, pady=10)
        frm_btns.pack(fill="x", padx=14, side="bottom")
        tk.Button(frm_btns,
                  text="\u2714  Confirmar y continuar guardado",
                  font=FONT_BOLD, fg=C_DARK, bg=C_ACCENT,
                  activebackground=C_ACCENT2, relief="flat", cursor="hand2",
                  command=self._confirmar).pack(side="left", ipadx=12, ipady=4)
        tk.Button(frm_btns,
                  text="\u2718  Cancelar todo el guardado",
                  font=FONT_BOLD, fg=C_TEXT_LIGHT, bg=C_BORDER,
                  activebackground=C_HOVER, relief="flat", cursor="hand2",
                  command=self._cancelar).pack(side="left", padx=8, ipadx=10, ipady=4)

        # ── Zona con scroll, un bloque por cada SAP en conflicto ──
        frm_zona = tk.Frame(self, bg=C_MID)
        frm_zona.pack(fill="both", expand=True, padx=10, pady=4)
        canvas  = tk.Canvas(frm_zona, bg=C_CARD, highlightthickness=0)
        scrolly = ttk.Scrollbar(frm_zona, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrolly.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrolly.pack(side="right", fill="y")
        frm_interior = tk.Frame(canvas, bg=C_CARD)
        canvas.create_window((0, 0), window=frm_interior, anchor="nw")
        frm_interior.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        for cod_sap, candidatos in self._conflictos.items():
            frm_grupo = tk.Frame(frm_interior, bg=C_CARD, bd=1,
                                  relief="solid", padx=10, pady=8)
            frm_grupo.pack(fill="x", padx=10, pady=6)
            tk.Label(frm_grupo, text=f"Código SAP {cod_sap}",
                     font=FONT_BOLD, fg=C_ACCENT, bg=C_CARD).pack(anchor="w")
            for cod_dali, es_existente, fila, desc in candidatos:
                var = tk.BooleanVar(value=False)
                self._checks[(cod_sap, cod_dali)] = var
                origen = (f"ya en el catálogo, fila {fila}" if es_existente
                          else "recién introducido en este lote")
                texto = f"Borrar de DALI {int(cod_dali):06d}"
                if desc:
                    texto += f" — {desc}"
                texto += f"  ({origen})"
                tk.Checkbutton(
                    frm_grupo, text=texto, variable=var,
                    font=FONT_LABEL, fg=C_TEXT_LIGHT, bg=C_CARD,
                    selectcolor=C_DARK, activebackground=C_CARD,
                    anchor="w", justify="left",
                ).pack(anchor="w", padx=(16, 0), pady=1)

    # ----------------------------------------------------------
    def _confirmar(self):
        a_borrar = set()
        for cod_sap, candidatos in self._conflictos.items():
            marcados = [cod_dali for (cod_dali, _, _, _) in candidatos
                        if self._checks[(cod_sap, cod_dali)].get()]
            if len(marcados) == len(candidatos):
                messagebox.showwarning(
                    "Selección incompleta",
                    f"Para el SAP {cod_sap} no puede quedar ningún DALI "
                    f"correcto -- deja al menos uno sin marcar.")
                return
            a_borrar.update(marcados)
        self.resultado = a_borrar
        self.destroy()

    def _cancelar(self):
        self.resultado = None
        self.destroy()


class VentanaSAPFaltantes(tk.Toplevel):
    """
    Muestra una tabla con los códigos DALI que no tienen SAP asignado
    y permite al usuario introducir el código SAP correspondiente.
    Una vez confirmado, escribe los valores en la columna «Código SAP» del
    catálogo completo (RECURSOS/catalogo-dali-completo.xlsx), buscando por
    código DALI. Hasta v10.2.50 se escribían en la columna D del LISTADO
    CODIGOS DALI - SAP, protegido con una contraseña escrita en el código;
    el LISTADO ya no existe y esa contraseña se ha quitado.
    Devuelve True si el usuario guardó cambios, False si canceló.
    """

    def __init__(self, parent, faltantes, ruta_listado, consola_fn=None,
                 descripciones=None, ruta_tarifa=None, catalogo_refresh_fn=None):
        """
        faltantes    : lista de (fila_xlsx_tarifa, codigo_dali)
        ruta_listado : ruta al catálogo completo (RECURSOS/catalogo-dali-completo.xlsx;
                       el nombre del parámetro es el de siempre)
        consola_fn   : función log_consola(msg, nivel) opcional
        descripciones: dict {codigo_dali: descripcion_str} opcional
        ruta_tarifa  : ruta al archivo .xlsx de tarifa recién generado
                       (para escribir también los SAP en col A)
        catalogo_refresh_fn: función opcional (lista_codigos_dali) -> None,
                       llamada tras guardar con los DALI cuyo SAP ha cambiado
                       (escritos o borrados por duplicado); en la app envía los
                       cambios a Catálogo Asignaciones (App._tras_cambiar_sap_en_catalogo)
        """
        super().__init__(parent)
        self.title("Códigos SAP faltantes — Introducir manualmente")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.configure(bg=C_MID)

        self._faltantes    = faltantes        # [(fila, dali), ...]
        self._ruta_listado = ruta_listado
        self._ruta_tarifa  = ruta_tarifa      # xlsx tarifa generado
        self._log          = consola_fn or (lambda m, n="INFO": None)
        self._catalogo_refresh_fn = catalogo_refresh_fn
        self._guardado     = False
        self._entradas     = {}               # {fila_xlsx: (cod_dali, tk.StringVar)}
        self._descripciones = descripciones or {}
        self._escritos_en_tarifa = 0          # cuántos SAP se volcaron al xlsx tarifa

        self._build_ui()
        self._centrar(parent)

    # ----------------------------------------------------------
    def _centrar(self, parent):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(1100, sw - 40)
        # Desconta barra de tareas Windows (~48 px) + margen superior (5 px)
        # y un pequeño colchón visual para que la ventana no quede pegada al borde
        h  = min(sh - 80, sh - 80)
        x  = max(0, (sw - w) // 2)
        y  = 5
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ----------------------------------------------------------
    def _build_ui(self):
        from tkinter import ttk

        # ── Cabecera ──────────────────────────────────────────
        frm_hdr = tk.Frame(self, bg=C_DARK, pady=10)
        frm_hdr.pack(fill="x")
        tk.Label(frm_hdr,
                 text="⚠  Códigos DALI sin SAP asignado",
                 font=FONT_TITLE, fg=C_ACCENT, bg=C_DARK).pack(side="left", padx=14)
        tk.Label(frm_hdr,
                 text=f"{len(self._faltantes)} registro(s)",
                 font=FONT_LABEL, fg=C_MUTED, bg=C_DARK).pack(side="right", padx=14)

        # ── Instrucción ───────────────────────────────────────
        tk.Label(self,
                 text="Introduce el código SAP para cada código DALI listado.\n"
                      "Deja en blanco los que no quieras actualizar ahora.",
                 font=FONT_LABEL, fg=C_TEXT_LIGHT, bg=C_MID,
                 justify="left").pack(anchor="w", padx=14, pady=(8, 4))

        # ── Botones (se empaquetan ANTES de frm_tabla para que Tkinter los
        #    reserve siempre en la parte inferior, incluso con listas largas) ─
        frm_btns = tk.Frame(self, bg=C_MID, pady=10)
        frm_btns.pack(fill="x", padx=14, side="bottom")

        tk.Button(frm_btns,
                  text="\u2714  Guardar en el catálogo",
                  font=FONT_BOLD, fg=C_DARK, bg=C_ACCENT,
                  activebackground=C_ACCENT2, relief="flat", cursor="hand2",
                  command=self._guardar).pack(side="left", ipadx=12, ipady=4)

        tk.Button(frm_btns,
                  text="\u2718  Cancelar",
                  font=FONT_BOLD, fg=C_TEXT_LIGHT, bg=C_BORDER,
                  activebackground=C_HOVER, relief="flat", cursor="hand2",
                  command=self.destroy).pack(side="left", padx=8, ipadx=10, ipady=4)

        tk.Label(frm_btns,
                 text="Los cambios se escriben directamente en\nRECURSOS/catalogo-dali-completo.xlsx",
                 font=FONT_TINY, fg=C_MUTED, bg=C_MID,
                 justify="left").pack(side="right")

        # ── Tabla con scroll ──────────────────────────────────
        frm_tabla = tk.Frame(self, bg=C_MID)
        frm_tabla.pack(fill="both", expand=True, padx=10, pady=4)

        canvas  = tk.Canvas(frm_tabla, bg=C_CARD, highlightthickness=0)
        scrolly = ttk.Scrollbar(frm_tabla, orient="vertical",   command=canvas.yview)
        scrollx = ttk.Scrollbar(frm_tabla, orient="horizontal",  command=canvas.xview)
        canvas.configure(yscrollcommand=scrolly.set, xscrollcommand=scrollx.set)

        scrolly.pack(side="right",  fill="y")
        scrollx.pack(side="bottom", fill="x")
        canvas.pack(side="left",    fill="both", expand=True)

        frm_inner = tk.Frame(canvas, bg=C_CARD)
        _win_id = canvas.create_window((0, 0), window=frm_inner, anchor="nw")

        def _on_config(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_resize(e):
            canvas.itemconfig(_win_id, width=e.width)
        frm_inner.bind("<Configure>", _on_config)
        canvas.bind("<Configure>",    _on_canvas_resize)

        # Col 0 y 2 y 3: ancho fijo; col 1 (DESCRIPCIÓN): se expande con la ventana
        frm_inner.columnconfigure(0, weight=0, minsize=110)
        frm_inner.columnconfigure(1, weight=1, minsize=260)
        frm_inner.columnconfigure(2, weight=0, minsize=140)
        frm_inner.columnconfigure(3, weight=0, minsize=220)

        # Cabecera tabla — 4 columnas
        hdrs  = ["CÓDIGO DALI", "DESCRIPCIÓN", "FORMATO", "CÓDIGO SAP  (introducir)"]
        col_w = [110, 360, 140, 220]
        for ci, (h, w) in enumerate(zip(hdrs, col_w)):
            tk.Label(frm_inner, text=h,
                     font=FONT_BOLD, fg=C_TEXT, bg="#D9D9D9",
                     width=w//7, anchor="center",
                     relief="flat", bd=0).grid(
                row=0, column=ci, padx=1, pady=1, sticky="nsew")

        # Filas de datos
        _vistos_dali = {}   # {cod_dali: StringVar} — comparte var entre filas del mismo DALI
        for idx, (fila_xlsx, cod_dali) in enumerate(self._faltantes, start=1):
            bg = C_PANEL if idx % 2 == 0 else "#EFF3F8"

            # Col 0 — CÓDIGO DALI (seleccionable/copiable)
            ent_dali = tk.Entry(frm_inner,
                                font=FONT_LABEL, fg=C_TEXT, bg=bg,
                                relief="flat", bd=0,
                                width=col_w[0]//7,
                                justify="center",
                                readonlybackground=bg,
                                state="normal")
            ent_dali.insert(0, str(cod_dali).zfill(6))
            ent_dali.config(state="readonly")
            ent_dali.grid(row=idx, column=0, padx=1, pady=1, sticky="nsew")

            # Col 1 — DESCRIPCIÓN (seleccionable/copiable)
            _desc_tuple = self._descripciones.get(cod_dali, ("", ""))
            _desc   = _desc_tuple[0] if isinstance(_desc_tuple, tuple) else str(_desc_tuple)
            _format = _desc_tuple[1] if isinstance(_desc_tuple, tuple) else ""
            ent_desc = tk.Entry(frm_inner,
                                font=FONT_LABEL, fg=C_TEXT, bg=bg,
                                relief="flat", bd=0,
                                width=1,
                                readonlybackground=bg,
                                state="normal")
            ent_desc.insert(0, _desc)
            ent_desc.config(state="readonly")
            ent_desc.grid(row=idx, column=1, padx=(6, 1), pady=1, sticky="nsew")

            # Col 2 — FORMATO/UNIDAD (seleccionable/copiable)
            ent_fmt = tk.Entry(frm_inner,
                               font=FONT_LABEL, fg=C_TEXT, bg=bg,
                               relief="flat", bd=0,
                               width=col_w[2]//7,
                               justify="center",
                               readonlybackground=bg,
                               state="normal")
            ent_fmt.insert(0, _format)
            ent_fmt.config(state="readonly")
            ent_fmt.grid(row=idx, column=2, padx=1, pady=1, sticky="nsew")

            # Col 3 — ENTRADA código SAP
            # Si este cod_dali ya apareció antes, reutilizar su StringVar
            # para que todas sus filas compartan el mismo valor introducido.
            if cod_dali in _vistos_dali:
                var = _vistos_dali[cod_dali]
            else:
                var = tk.StringVar()
                _vistos_dali[cod_dali] = var
            self._entradas[fila_xlsx] = (cod_dali, var)
            ent = tk.Entry(frm_inner, textvariable=var,
                           font=FONT_LABEL, fg=C_TEXT, bg=bg,
                           relief="flat", bd=0,
                           width=col_w[3]//7)
            ent.grid(row=idx, column=3, padx=1, pady=1, sticky="nsew")

    # ----------------------------------------------------------
    def _mostrar_resolucion_duplicados_y_esperar(self, conflictos):
        """
        Muestra VentanaResolverSAPDuplicados de forma segura para Tkinter
        y bloquea el hilo que llama hasta que el usuario resuelve los
        conflictos o cancela.

        Se llama desde _tarea_guardar(), que corre en un hilo de fondo
        -- Tkinter no es thread-safe, así que no se puede crear la
        ventana directamente ahí (mismo hallazgo que _ejecutar_en_hilo_
        principal_y_esperar en tarifa_app.py, 2026-09-24). Se programa
        con self.after(0, ...) para que la ventana se cree en el hilo
        principal, y se bloquea con threading.Event hasta que se cierre.

        Devuelve None si el usuario canceló todo el guardado, o un set
        de códigos DALI cuyo SAP hay que borrar/omitir si confirmó.
        """
        resultado = {}
        evento = threading.Event()

        def _mostrar():
            dlg = VentanaResolverSAPDuplicados(self, conflictos)
            self.wait_window(dlg)
            resultado["valor"] = dlg.resultado
            evento.set()

        self.after(0, _mostrar)
        evento.wait()
        return resultado.get("valor")

    def _guardar(self):
        """
        Recoge las entradas, valida que sean numéricas y lanza el guardado
        en un hilo secundario para no bloquear la UI.
        """
        # ── Validar entradas en el hilo principal (UI) ────────
        # _entradas: {fila_xlsx: (cod_dali, StringVar)}
        # Un mismo cod_dali puede aparecer en varias filas; todas comparten
        # la misma StringVar, así que consolidamos por cod_dali.
        nuevos = {}   # {cod_dali: cod_sap}
        for fila_xlsx, (cod_dali, var) in self._entradas.items():
            if cod_dali in nuevos:
                continue   # ya procesado (filas duplicadas del mismo DALI)
            txt = var.get().strip()
            if txt:
                try:
                    nuevos[cod_dali] = int(txt)
                except ValueError:
                    from tkinter import messagebox
                    messagebox.showwarning(
                        "Valor no válido",
                        f"El código SAP para DALI {str(cod_dali).zfill(6)} debe ser un número entero.\n"
                        f"Valor introducido: '{txt}'")
                    return

        if not nuevos:
            from tkinter import messagebox
            messagebox.showinfo("Sin cambios", "No se introdujo ningún código SAP.")
            self.destroy()
            return

        # Deshabilitar botón mientras se guarda
        for w in self.winfo_children():
            try:
                w.config(state="disabled")
            except Exception:
                pass

        def _tarea_guardar():
            """Ejecutada en hilo secundario — toda la E/S con openpyxl aquí."""
            _bloqueo = getattr(self.master, "_catalogo_lock", None)
            _tomado = [False]          # si este hilo tiene el bloqueo del catálogo (se suelta pase lo que pase)

            def _soltar():
                if _tomado[0]:
                    _tomado[0] = False
                    _bloqueo.release()
            try:
                from openpyxl import load_workbook
                from app.comun import catalogo_local

                # v10.2.50: el código SAP se escribe en el catálogo completo, buscando por código DALI.
                # Con el bloqueo del catálogo de la app, para no chocar con otra escritura a la vez.
                if _bloqueo is not None:
                    _bloqueo.acquire()
                    _tomado[0] = True
                wb, ws, _mapa_cat = catalogo_local.abrir_para_editar(self._ruta_listado)
                _col_sap = _mapa_cat["cod_sap"]

                # ── Mapa DALI→fila y SAP→(DALI, fila): el segundo permite
                #    detectar códigos SAP duplicados más abajo.
                _mapa_fila, _mapa_sap_existente = catalogo_local.mapas_dali_sap(ws, _mapa_cat)

                # ── Detectar códigos SAP duplicados: ya asignados a otro
                #    DALI en el listado, o repetidos dentro del propio
                #    lote que se está guardando ahora. Si hay alguno, se
                #    pide al usuario que indique cuál es el correcto
                #    antes de escribir nada.
                from collections import defaultdict
                _por_sap_nuevo = defaultdict(list)
                for _cd, _cs in nuevos.items():
                    _por_sap_nuevo[_cs].append(_cd)

                conflictos = {}
                for cod_sap, dalis_nuevos in _por_sap_nuevo.items():
                    candidatos = []
                    if cod_sap in _mapa_sap_existente:
                        dali_existente, fila_existente = _mapa_sap_existente[cod_sap]
                        if dali_existente not in dalis_nuevos:
                            candidatos.append((
                                dali_existente, True, fila_existente,
                                self._descripciones.get(dali_existente, "")))
                    if len(dalis_nuevos) > 1 or candidatos:
                        for _d in dalis_nuevos:
                            candidatos.append((
                                _d, False, None, self._descripciones.get(_d, "")))
                    if len(candidatos) > 1:
                        conflictos[cod_sap] = candidatos

                _dali_limpiados = []   # siempre definida, aunque no haya conflictos
                if conflictos:
                    self._log(
                        f"  ⚠ {len(conflictos)} código(s) SAP en conflicto "
                        f"con otro DALI -- pidiendo resolución al usuario...",
                        "WARN")
                    a_borrar = self._mostrar_resolucion_duplicados_y_esperar(conflictos)
                    if a_borrar is None:
                        wb.close()
                        _soltar()
                        self._log(
                            "  ✘ Guardado cancelado por el usuario "
                            "(duplicados sin resolver).", "WARN")

                        def _restaurar_botones():
                            for w in self.winfo_children():
                                try:
                                    w.config(state="normal")
                                except Exception:
                                    pass
                        self.after(0, _restaurar_botones)
                        return

                    _filas_a_limpiar = []
                    _dali_limpiados  = []
                    for cod_sap, candidatos in conflictos.items():
                        for cod_dali, es_existente, fila, _desc in candidatos:
                            if cod_dali in a_borrar:
                                if es_existente:
                                    _filas_a_limpiar.append(fila)
                                    _dali_limpiados.append(cod_dali)
                                nuevos.pop(cod_dali, None)
                    for fila in _filas_a_limpiar:
                        ws.cell(fila, _col_sap).value = None
                    if _filas_a_limpiar:
                        self._log(
                            f"  ✔ {len(_filas_a_limpiar)} código(s) SAP "
                            f"duplicados borrados del catálogo (ya existentes "
                            f"en conflicto).", "OK")

                # ── Escribir SAP en col D ──────────────────────
                escritos       = 0
                no_encontrados = []
                for cod_dali, cod_sap in nuevos.items():
                    if cod_dali in _mapa_fila:
                        ws.cell(_mapa_fila[cod_dali], _col_sap).value = cod_sap
                        escritos += 1
                        self._log(
                            f"  ✔ DALI {cod_dali:06d} → SAP {cod_sap} "
                            f"(fila {_mapa_fila[cod_dali]})", "OK")
                    else:
                        no_encontrados.append(cod_dali)
                        self._log(
                            f"  ✘ DALI {cod_dali:06d} no encontrado en el catálogo", "WARN")

                try:
                    wb.save(self._ruta_listado)
                finally:
                    wb.close()
                    _soltar()
                self._guardado = True
                self._log(f"  └ Catálogo guardado: {escritos} SAP escritos.", "OK")

                # ── Enviar a Catálogo Asignaciones los DALI cuyo SAP ha cambiado
                #    (escritos ahora o borrados por duplicado).
                _cambiados = list(nuevos) + list(_dali_limpiados)
                if _cambiados and self._catalogo_refresh_fn:
                    try:
                        self._catalogo_refresh_fn(_cambiados)
                    except Exception as _e_cat_ref:
                        self._log(
                            f"  ⚠  No se pudieron enviar los cambios a Catálogo "
                            f"Asignaciones: {_e_cat_ref}", "WARN")

                # ── Escribir también en el xlsx de tarifa (col A) ─
                _escritos_tarifa = 0
                if self._ruta_tarifa and os.path.isfile(self._ruta_tarifa):
                    try:
                        wb_tar = load_workbook(self._ruta_tarifa, rich_text=False)
                        ws_tar = wb_tar.active
                        # Desproteger si está protegida
                        ws_tar.protection.sheet   = False
                        ws_tar.protection.enabled = False
                        for fila_xlsx, cod_dali in self._faltantes:
                            if cod_dali in nuevos:
                                ws_tar.cell(fila_xlsx, 1).value = nuevos[cod_dali]
                                _escritos_tarifa += 1
                                self._log(
                                    f"  ✔ TARIFA fila {fila_xlsx}: "
                                    f"DALI {cod_dali:06d} → SAP {nuevos[cod_dali]}", "OK")
                        wb_tar.save(self._ruta_tarifa)
                        wb_tar.close()
                        self._log(
                            f"  └ Tarifa actualizada: {_escritos_tarifa} SAP escritos.", "OK")
                    except Exception as _e_tar:
                        self._log(
                            f"  ✘ No se pudo actualizar la tarifa: {_e_tar}", "WARN")
                self._escritos_en_tarifa = _escritos_tarifa

                # ── Notificar en el hilo principal ────────────
                def _fin_ok(e=escritos, nf=no_encontrados,
                            et=_escritos_tarifa):
                    from tkinter import messagebox
                    msg = f"✔  {e} código(s) SAP guardados en el catálogo."
                    if et:
                        msg += f"\n✔  {et} código(s) escritos también en la tarifa."
                    if nf:
                        msg += f"\n\n⚠  {len(nf)} código(s) DALI no encontrados en el catálogo:\n"
                        msg += ", ".join(str(d) for d in nf)
                    if not et and not self._ruta_tarifa:
                        msg += ("\n\n📋  Recuerda:\n"
                                "  1. Actualiza el archivo Productos en RECURSOS.\n"
                                "  2. Ejecuta de nuevo Acción 3 (Actualizar materiales SAP).\n"
                                "  3. Ejecuta de nuevo Acción 4 (Procesar y generar TARIFA).")
                    messagebox.showinfo("Guardado correctamente", msg)
                    self.destroy()

                try:
                    self.after(0, _fin_ok)
                except Exception:
                    pass  # ventana ya destruida; _guardado y _escritos_en_tarifa ya están puestos

            except PermissionError:
                _soltar()

                def _err_perm():
                    from tkinter import messagebox
                    messagebox.showerror(
                        "Archivo bloqueado",
                        f"No se puede escribir en:\n{self._ruta_listado}\n\n"
                        "Cierra el archivo en Excel e inténtalo de nuevo.")
                    for w in self.winfo_children():
                        try:
                            w.config(state="normal")
                        except Exception:
                            pass
                self.after(0, _err_perm)

            except Exception as ex:
                _soltar()

                def _err_gen(e=str(ex)):
                    from tkinter import messagebox
                    messagebox.showerror("Error al guardar", e)
                    for w in self.winfo_children():
                        try:
                            w.config(state="normal")
                        except Exception:
                            pass
                self._log(f"  ✘ Error al guardar: {ex}", "ERROR")
                self.after(0, _err_gen)

        import threading
        threading.Thread(target=_tarea_guardar, daemon=True).start()

    # ----------------------------------------------------------
    @property
    def guardado(self):
        return self._guardado

    @property
    def escritos_en_tarifa(self):
        return self._escritos_en_tarifa


