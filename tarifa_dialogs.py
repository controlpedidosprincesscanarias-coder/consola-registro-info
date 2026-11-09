# -*- coding: utf-8 -*-
"""
tarifa_dialogs.py
Ventanas modales (Toplevel) de la aplicación TARIFA
Central Compras Princess Canarias
Autor: Victor Alejandro Martin Andrade (VAMA)
(c) VAMA 2026

Clases exportadas:
    VentanaSeleccionXLS          — selección de archivo .xls a renombrar
    VentanaAbrirTarifaEdicion    — abrir tarifa [INCOMPLETO] en Excel (Acción 4)
    VentanaCrearProyecto         — crear estructura PROYECTO/TARIFAS/RECURSOS
    VentanaSeleccionTarifaRI     — selección tarifa + búsqueda Purchasing (Acción 6)
    VentanaSAPFaltantes          — entrada manual de códigos SAP faltantes
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox
import threading

from tarifa_constants import (
    C_DARK, C_MID, C_CARD, C_BORDER, C_HOVER, C_ACCENT,
    C_ACCENT2, C_BG, C_PANEL,
    C_TEXT, C_TEXT_LIGHT, C_MUTED, C_SUCCESS, FONT_TITLE,
    FONT_SUB, FONT_LABEL, FONT_BOLD,
    FONT_SMALL, FONT_TINY,
)
from tarifa_core import (
    resource_path, renombrar_a_tarifa, crear_estructura_proyecto,
    FOLDER_TARIFAS, FOLDER_RECURSOS, cargar_icono_ventana, centrar_sobre_padre,
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
#  VENTANA: ABRIR TARIFA PARA EDICIÓN MANUAL (ACCIÓN 4)
# ---------------------------------------------------------
class VentanaAbrirTarifaEdicion(tk.Toplevel):
    """
    Se muestra cuando la Acción 4 no encuentra TARIFA.xls.
    Lista todos los archivos *[dd-mm-yyyy].xlsx de TARIFAS/ (incluyendo
    los marcados como [INCOMPLETO]) para que el usuario elija uno y lo
    abra directamente en Excel para corregirlo manualmente.
    Recuerda que, al guardar tras corregir los errores, debe eliminar
    '[INCOMPLETO]' del nombre del archivo.
    """

    def __init__(self, parent, archivos, carpeta_tarifas, consola):
        super().__init__(parent)
        self.title("Acción 4 — Abrir tarifa para edición")
        self.resizable(True, True)
        self.minsize(620, 420)
        self.configure(bg=C_BG)
        cargar_icono_ventana(self)
        self.carpeta_tarifas = carpeta_tarifas
        self.consola         = consola
        self._build_ui(archivos)
        self.grab_set()
        self.transient(parent)
        centrar_sobre_padre(self, parent)


    def _build_ui(self, archivos):
        # ── Cabecera ────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=C_MID, pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⚠️  TARIFA.xls no encontrada",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Seleccione un archivo de tarifa para abrirlo y editarlo manualmente:",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(3, 0))

        # ── Tabla de archivos ────────────────────────────────────────
        frm_tabla = tk.Frame(self, bg=C_BG, padx=14, pady=10)
        frm_tabla.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("Edit.Treeview",
                         background=C_PANEL, fieldbackground=C_PANEL,
                         foreground=C_TEXT, rowheight=24, font=FONT_LABEL)
        style.configure("Edit.Treeview.Heading",
                         background=C_MID, foreground="white",
                         font=FONT_BOLD, relief="flat")
        style.map("Edit.Treeview",
                   background=[("selected", C_MID)],
                   foreground=[("selected", "white")])

        cols = ("nombre", "estado", "fecha_mod")
        self.tree = ttk.Treeview(frm_tabla, columns=cols, show="headings",
                                  selectmode="browse", style="Edit.Treeview")
        self.tree.heading("nombre",    text="Nombre del archivo")
        self.tree.heading("estado",    text="Estado")
        self.tree.heading("fecha_mod", text="Última modificación")
        self.tree.column("nombre",    width=330, anchor="w")
        self.tree.column("estado",    width=110, anchor="center")
        self.tree.column("fecha_mod", width=140, anchor="center")

        sb = ttk.Scrollbar(frm_tabla, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.tree.tag_configure("par",        background="#F0F5FB")
        self.tree.tag_configure("impar",      background="white")
        self.tree.tag_configure("incompleto", background="#FFF3CD", foreground="#7A5000")

        if archivos:
            for i, a in enumerate(archivos):
                incompleto = "[INCOMPLETO]" in a["nombre_archivo"].upper()
                estado     = "⚠ INCOMPLETO" if incompleto else "✔ Procesable"
                if incompleto:
                    tag = "incompleto"
                else:
                    tag = "par" if i % 2 == 0 else "impar"
                self.tree.insert("", "end", iid=a["ruta"],
                                  values=(a["nombre_display"], estado, a["fecha_mod"]),
                                  tags=(tag,))
            self.tree.selection_set(archivos[0]["ruta"])
        else:
            tk.Label(frm_tabla,
                     text="No se encontraron archivos de tarifa en la carpeta TARIFAS.",
                     fg=C_MUTED, bg=C_BG, font=FONT_LABEL).pack(expand=True)

        self.tree.bind("<Double-1>", lambda e: self._abrir())

        # ── Aviso recordatorio ───────────────────────────────────────
        frm_aviso = tk.Frame(self, bg="#FFF8E6", bd=0)
        frm_aviso.pack(fill="x", padx=14, pady=(0, 6))
        tk.Label(
            frm_aviso,
            text=(
                "📝  Recuerde: al corregir los errores marcados en amarillo y rojo,\n"
                "guarde el archivo eliminando  [INCOMPLETO]  de su nombre\n"
                "para que pueda ser procesado en la Acción 6."
            ),
            bg="#FFF8E6", fg="#7A5000", font=FONT_SMALL,
            justify="left", padx=10, pady=8,
            wraplength=560, anchor="w"
        ).pack(fill="x")

        # ── Botones ──────────────────────────────────────────────────
        frm_btn = tk.Frame(self, bg=C_BG, padx=14, pady=10)
        frm_btn.pack(fill="x", side="bottom")

        tk.Button(frm_btn, text="📂  Abrir en Excel",
                  command=self._abrir,
                  bg=C_ACCENT, fg=C_DARK, font=FONT_BOLD,
                  padx=14, pady=6, cursor="hand2", relief="flat",
                  activebackground=C_ACCENT2).pack(side="right", padx=(6, 0))

        tk.Button(frm_btn, text="Cancelar", command=self.destroy,
                  font=FONT_LABEL, padx=14, pady=6, cursor="hand2",
                  relief="flat", bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(side="right")

    def _abrir(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sin selección",
                                   "Seleccione un archivo de la lista.", parent=self)
            return

        ruta = sel[0]
        nombre = os.path.basename(ruta)
        incompleto = "[INCOMPLETO]" in nombre.upper()

        # Mostrar recordatorio adicional si es INCOMPLETO
        if incompleto:
            if not messagebox.askyesno(
                    "Archivo incompleto",
                    f"📂  {nombre}\n\n"
                    "Este archivo contiene errores marcados en amarillo o rojo.\n\n"
                    "Recuerde: al guardar tras corregir los errores,\n"
                    "elimine  [INCOMPLETO]  del nombre del archivo\n"
                    "para que pueda procesarse en Acción 6.\n\n"
                    "¿Abrir el archivo ahora?",
                    parent=self):
                return

        try:
            if os.name == "nt":
                os.startfile(ruta)
            else:
                import subprocess
                subprocess.call(["xdg-open", ruta])
            self.consola.log(f"Abierto para edición: {nombre}", "INFO")
        except Exception as e:
            self.consola.log(f"Error al abrir: {e}", "ERROR")
            messagebox.showerror("Error al abrir",
                                 f"No se pudo abrir el archivo:\n{e}", parent=self)


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
#  VENTANA: SELECCIÓN TARIFA PARA REGISTRO INFO (ACCIÓN 6)
# ---------------------------------------------------------
class VentanaSeleccionTarifaRI(tk.Toplevel):
    """
    Paso A — presenta al usuario la lista de archivos * [dd-mm-yyyy].xlsx
    encontrados en la carpeta TARIFAS para que seleccione con cuál trabajar.
    Tras la selección ejecuta los pasos B/C/D de la Acción 6.
    """

    def __init__(self, parent, archivos, carpeta_tarifas, consola):
        super().__init__(parent)
        self.title("Acción 6 — Seleccionar tarifa para Registro Info")
        self.resizable(True, True)
        self.minsize(620, 360)
        self.configure(bg=C_BG)
        cargar_icono_ventana(self)
        self.carpeta_tarifas = carpeta_tarifas
        self.consola         = consola
        self._archivos       = archivos
        self._build_ui()
        self.grab_set()
        self.transient(parent)
        centrar_sobre_padre(self, parent)


    def _build_ui(self):
        # ── Cabecera ─────────────────────────────────────────────
        hdr = tk.Frame(self, bg=C_MID, pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="💾  Crear nuevo Registro Info SAP",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Seleccione el archivo de tarifa con el que desea trabajar:",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        # ── Tabla ─────────────────────────────────────────────────
        frm_tabla = tk.Frame(self, bg=C_BG, padx=14, pady=10)
        frm_tabla.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("RI.Treeview",
                         background=C_PANEL, fieldbackground=C_PANEL,
                         foreground=C_TEXT, rowheight=23, font=FONT_LABEL)
        style.configure("RI.Treeview.Heading",
                         background=C_MID, foreground="white",
                         font=FONT_BOLD, relief="flat")
        style.map("RI.Treeview",
                   background=[("selected", "#7B4FA6")],
                   foreground=[("selected", "white")])

        cols = ("nombre", "fecha_mod")
        self.tree = ttk.Treeview(frm_tabla, columns=cols, show="headings",
                                  selectmode="browse", style="RI.Treeview")
        self.tree.heading("nombre",    text="Archivo de tarifa")
        self.tree.heading("fecha_mod", text="Última modificación")
        self.tree.column("nombre",    width=410, anchor="w")
        self.tree.column("fecha_mod", width=150, anchor="center")

        sb = ttk.Scrollbar(frm_tabla, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.tree.tag_configure("par",   background="#F0F5FB")
        self.tree.tag_configure("impar", background="white")

        for i, a in enumerate(self._archivos):
            tag = "par" if i % 2 == 0 else "impar"
            self.tree.insert("", "end", iid=a["ruta"],
                              values=(a["nombre"], a["fecha_mod"]),
                              tags=(tag,))
        if self._archivos:
            self.tree.selection_set(self._archivos[0]["ruta"])

        self.tree.bind("<Double-1>", lambda e: self._confirmar())

        # ── Botones ───────────────────────────────────────────────
        frm_btn = tk.Frame(self, bg=C_BG, padx=14, pady=10)
        frm_btn.pack(fill="x", side="bottom")

        tk.Button(frm_btn, text="✔  Seleccionar",
                  command=self._confirmar,
                  bg="#7B4FA6", fg="white", font=FONT_BOLD,
                  padx=14, pady=6, cursor="hand2", relief="flat",
                  activebackground="#9B6FC6").pack(side="right", padx=(6, 0))

        tk.Button(frm_btn, text="Cancelar", command=self.destroy,
                  font=FONT_LABEL, padx=14, pady=6, cursor="hand2",
                  relief="flat", bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(side="right")

    # ── Lógica de negocio pasos B, C, D ───────────────────────────
    def _confirmar(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sin selección",
                                   "Seleccione un archivo de la lista.", parent=self)
            return

        ruta_tarifa_sel = sel[0]
        nombre_display  = self.tree.item(ruta_tarifa_sel, "values")[0]
        self.consola.log(f"Tarifa seleccionada: {nombre_display}", "INFO")

        # ── PASO B: leer nombre del proveedor desde la tarifa ────────
        # La tarifa procesada tiene una única hoja llamada "TARIFA".
        # El nombre del proveedor está en C2 (etiqueta "Proveedor:" en B2).
        nombre_proveedor = None
        try:
            import openpyxl
            wb = openpyxl.load_workbook(ruta_tarifa_sel,
                                         read_only=True, data_only=True)
            # Intentar primero hoja TARIFA (tarifa procesada), luego hoja activa
            ws = None
            if "TARIFA" in wb.sheetnames:
                ws = wb["TARIFA"]
            elif "Info general p.reg.info" in wb.sheetnames:
                ws = wb["Info general p.reg.info"]
            else:
                ws = wb.active

            if ws is not None:
                # Buscar fila donde B = "Proveedor:" y tomar C de esa fila
                for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
                    etiq = str(row[1]).strip().lower() if row[1] else ""
                    if "proveedor" in etiq:
                        val = row[2] if len(row) > 2 else None
                        nombre_proveedor = str(val).strip() if val else None
                        break
                # Fallback: celda C2 directamente
                if not nombre_proveedor:
                    val = ws["C2"].value
                    nombre_proveedor = str(val).strip() if val else None
            wb.close()
        except Exception as e:
            self.consola.log(f"Error leyendo tarifa: {e}", "ERROR")
            messagebox.showerror(
                "Error al leer el archivo",
                f"No se pudo leer el nombre de proveedor de la tarifa:\n\n{e}",
                parent=self)
            return

        if not nombre_proveedor:
            self.consola.log("Proveedor no identificado en la tarifa.", "WARN")
            self._preguntar_sin_purchasing("(proveedor no identificado)", ruta_tarifa_sel)
            return

        self.consola.log(f"Proveedor identificado (tarifa): {nombre_proveedor}", "OK")

        # ── PASO C: buscar Purchasing Info Records*.xlsx compatible ──
        # Compara el nombre del proveedor de la tarifa contra los nombres
        # únicos que aparecen en la columna "Nombre del proveedor" (col D)
        # de la hoja "Info general p.reg.info" de cada Purchasing encontrado.
        self.consola.log("Buscando Purchasing compatible (comparando datos reales)...", "INFO")
        purchasing_encontrado = self._buscar_purchasing(nombre_proveedor)

        if purchasing_encontrado:
            arch_p, tipo, score, prov_purch = purchasing_encontrado
            tipo_txt = {"exacto": "Exacta", "alta": "Alta", "media": "Media"}.get(tipo, tipo)
            self.consola.log(
                f"Purchasing encontrado: {os.path.basename(arch_p)} "
                f"(coincidencia {tipo_txt}, score {score}%)", "OK")
            messagebox.showinfo(
                "Purchasing Info Records localizado",
                f"✔  Se encontró el archivo Purchasing Info Records\n"
                f"   compatible con el proveedor:\n\n"
                f"   Proveedor (tarifa)    : {nombre_proveedor}\n"
                f"   Proveedor (Purchasing): {prov_purch}\n"
                f"   Archivo encontrado   : {os.path.basename(arch_p)}\n"
                f"   Tipo coincidencia    : {tipo_txt} ({score}%)\n\n"
                "Continuando con la generación del Registro Info…",
                parent=self)
            # TODO: continuar con los siguientes pasos de la acción 6
            self.destroy()

        else:
            # ── PASO D: no existe Purchasing compatible ───────────
            self.consola.log(
                "No se encontró Purchasing Info Records compatible.", "WARN")
            self._preguntar_sin_purchasing(nombre_proveedor, ruta_tarifa_sel)

    def _buscar_purchasing(self, nombre_proveedor):
        """
        Busca en carpeta_tarifas todos los archivos Purchasing Info Records*.xlsx
        y compara el nombre del proveedor de la tarifa contra los nombres únicos
        que aparecen en la columna D ("Nombre del proveedor") de la hoja
        "Info general p.reg.info" de cada archivo Purchasing (los datos reales,
        no un encabezado fijo).

        Algoritmo de puntuación (0-100):
          - Exacta (100) : cadenas idénticas ignorando mayúsculas/tildes/puntuación
          - Alta   (≥70) : todas las palabras significativas presentes + bigramas
          - Media  (≥45) : mitad de palabras presentes o similitud de secuencia

        Devuelve (ruta, tipo, score, nombre_proveedor_purchasing) o None.
        """
        import unicodedata
        import re as _re
        import openpyxl

        def _norm(txt):
            if not txt:
                return ""
            txt = str(txt).upper()
            txt = unicodedata.normalize("NFD", txt)
            txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
            txt = _re.sub(r"[^A-Z0-9 ]", " ", txt)
            return _re.sub(r"\s+", " ", txt).strip()

        def _palabras(txt_norm):
            STOP = {"SA", "SL", "SLU", "SAU", "LTD", "SAS", "BV", "NV",
                    "THE", "AND", "LOS", "LAS", "DEL", "DE", "LA", "EL",
                    "CON", "SIN", "POR", "SPA", "AG", "INC", "CO", "SAC"}
            return [w for w in txt_norm.split() if len(w) >= 3 and w not in STOP]

        def _score(ref_norm, cand_norm):
            if not ref_norm or not cand_norm:
                return 0
            if ref_norm == cand_norm:
                return 100
            if ref_norm in cand_norm or cand_norm in ref_norm:
                return 95
            pal_ref = _palabras(ref_norm)
            if not pal_ref:
                return 0
            presentes = sum(1 for p in pal_ref if p in cand_norm)
            ratio_pal = presentes / len(pal_ref)
            def _bg(s):
                return set(s[i:i+2] for i in range(len(s) - 1))
            bg_r, bg_c = _bg(ref_norm), _bg(cand_norm)
            jaccard = len(bg_r & bg_c) / len(bg_r | bg_c) if (bg_r or bg_c) else 0.0
            return round(ratio_pal * 60 + jaccard * 40)

        ref_norm = _norm(nombre_proveedor)
        HOJA_P   = "Info general p.reg.info"
        COL_PROV = 3   # columna D = índice 3 (0-based) en iter_rows values_only

        mejor_ruta  = None
        mejor_score = 0
        mejor_tipo  = None
        mejor_prov  = None

        try:
            for entry in os.scandir(self.carpeta_tarifas):
                if not entry.is_file():
                    continue
                nombre_arch = entry.name.lower()
                if not (nombre_arch.startswith("purchasing info records") and
                        nombre_arch.endswith(".xlsx")):
                    continue

                # Extraer nombres únicos de proveedores de la columna D (datos)
                proveedores_purch = set()
                try:
                    wb_p = openpyxl.load_workbook(
                        entry.path, read_only=True, data_only=True)
                    if HOJA_P in wb_p.sheetnames:
                        ws_p = wb_p[HOJA_P]
                        for i, row in enumerate(
                                ws_p.iter_rows(min_row=2, values_only=True)):
                            if len(row) > COL_PROV and row[COL_PROV]:
                                prov_val = str(row[COL_PROV]).strip()
                                if prov_val:
                                    proveedores_purch.add(prov_val)
                            if i > 500:   # limitar lectura en archivos muy grandes
                                break
                    wb_p.close()
                except Exception:
                    proveedores_purch = set()

                if not proveedores_purch:
                    self.consola.log(
                        f"  └ {entry.name}: sin datos de proveedor legibles", ">>")
                    continue

                self.consola.log(
                    f"  └ {entry.name}: {len(proveedores_purch)} proveedor(es) único(s)", ">>")

                # Calcular mejor score contra todos los nombres únicos del archivo
                for pv in proveedores_purch:
                    s = _score(ref_norm, _norm(pv))
                    if s > mejor_score:
                        mejor_score = s
                        mejor_ruta  = entry.path
                        mejor_prov  = pv

        except PermissionError:
            pass

        if mejor_ruta is None or mejor_score < 45:
            return None

        if mejor_score >= 95:
            tipo = "exacto"
        elif mejor_score >= 70:
            tipo = "alta"
        else:
            tipo = "media"

        return (mejor_ruta, tipo, mejor_score, mejor_prov)

    def _preguntar_sin_purchasing(self, nombre_proveedor, ruta_tarifa):
        """
        PASO D — informa al usuario que no existe Purchasing compatible
        y le pregunta si desea descargarlo de SAP o crear un Registro Info Nuevo.
        """
        ventana = tk.Toplevel(self)
        ventana.title("Registro Info — sin Purchasing para este proveedor")
        ventana.resizable(False, False)
        ventana.configure(bg=C_BG)
        ventana.grab_set()
        ventana.transient(self)
        try:
            ventana.iconbitmap(resource_path("assets/logo.ico"))
        except Exception:
            pass

        # ── Cabecera ───────────────────────────────────────────────
        hdr = tk.Frame(ventana, bg="#8B3A3A", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⚠   No se encontró Registro Info para este proveedor",
                 bg="#8B3A3A", fg="white", font=FONT_TITLE).pack(anchor="w")

        # ── Cuerpo ────────────────────────────────────────────────
        cuerpo = tk.Frame(ventana, bg=C_BG, padx=20, pady=16)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text="No se encontró ningún archivo\n"
                      "Purchasing Info Records\n"
                      "en la carpeta TARIFAS para el proveedor:",
                 bg=C_BG, fg=C_TEXT, font=FONT_LABEL,
                 justify="left").pack(anchor="w")

        tk.Label(cuerpo, text=f"  {nombre_proveedor}",
                 bg="#EEF2F8", fg="#7B4FA6",
                 font=("Segoe UI", 10, "bold"),
                 padx=10, pady=6, relief="flat").pack(
            fill="x", pady=(6, 14))

        tk.Label(cuerpo,
                 text="¿Cómo desea proceder?",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 8))

        # ── Botones de decisión ───────────────────────────────────
        frm_btn = tk.Frame(cuerpo, bg=C_BG)
        frm_btn.pack(fill="x")

        def _descargar_sap():
            ventana.destroy()
            self.destroy()
            self.consola.log(
                "Usuario eligió: descargar Purchasing desde SAP.", "INFO")

            instrucciones = tk.Toplevel()
            instrucciones.title("Cómo descargar el Purchasing Info Records desde SAP")
            instrucciones.resizable(True, True)
            instrucciones.configure(bg=C_BG)
            instrucciones.minsize(580, 520)
            try:
                instrucciones.iconbitmap(resource_path("assets/logo.ico"))
            except Exception:
                pass

            # ── Cabecera ──────────────────────────────────────────
            hdr2 = tk.Frame(instrucciones, bg=C_MID, pady=12, padx=16)
            hdr2.pack(fill="x")
            tk.Label(hdr2, text="📥  Descargar Purchasing Info Records desde SAP",
                     bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
            tk.Label(hdr2,
                     text="Siga los pasos indicados a continuación en SAP S/4HANA:",
                     bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

            # ── Cuerpo scrollable ─────────────────────────────────
            frame_outer = tk.Frame(instrucciones, bg=C_BG)
            frame_outer.pack(fill="both", expand=True, padx=14, pady=10)

            canvas_i = tk.Canvas(frame_outer, bg=C_BG, highlightthickness=0)
            sb_i = ttk.Scrollbar(frame_outer, orient="vertical",
                                 command=canvas_i.yview)
            canvas_i.configure(yscrollcommand=sb_i.set)
            sb_i.pack(side="right", fill="y")
            canvas_i.pack(side="left", fill="both", expand=True)

            frame_steps = tk.Frame(canvas_i, bg=C_BG)
            canvas_win_id = canvas_i.create_window((0, 0), window=frame_steps,
                                                    anchor="nw")

            def _on_frame_config(e):
                canvas_i.configure(scrollregion=canvas_i.bbox("all"))
            def _on_canvas_resize(e):
                canvas_i.itemconfig(canvas_win_id, width=e.width)
            frame_steps.bind("<Configure>", _on_frame_config)
            canvas_i.bind("<Configure>", _on_canvas_resize)

            PASOS = [
                ("1", "Acceder a la aplicación TARIFAS en SAP",
                 "En el menú de favoritos de SAP S/4HANA, localice el grupo\n"
                 "TARIFAS  y pulse en:\n\n"
                 "   «Modificaciones en masa de registros info de compras»"),
                ("2", "Pantalla de búsqueda",
                 "Se abrirá la pantalla  «Estándar COMPRAS».\n"
                 "Localice el campo  Proveedor:  en la parte superior izquierda."),
                ("3", "Abrir el selector de Proveedor",
                 "Pulse el icono de cuadraditos (⊞) a la derecha del campo\n"
                 "Proveedor:  para abrir el buscador de proveedores."),
                ("4", "Buscar y seleccionar el proveedor",
                 f"En el buscador que se abre:\n\n"
                 f"  1. Escriba el nombre del proveedor en el campo de búsqueda:\n"
                 f"       «{nombre_proveedor}»\n\n"
                 f"  2. Pulse la lupa (🔍) o el botón  Ir.\n\n"
                 f"  3. Marque la casilla junto al Acreedor (✔) para seleccionar\n"
                 f"     todos los registros del proveedor.\n\n"
                 f"  4. Pulse  OK."),
                ("5", "Ejecutar la búsqueda",
                 "De vuelta en la pantalla principal, pulse el botón  Ir  (azul)\n"
                 "para cargar todos los Registros Info del proveedor."),
                ("6", "Seleccionar todos y descargar",
                 "Una vez cargados los registros en la tabla:\n\n"
                 "  1. Marque la casilla de la cabecera de la columna (junto a\n"
                 "     «Registro info de compras») para SELECCIONAR TODOS.\n\n"
                 "  2. Pulse el botón  Descargar  (parte superior derecha)."),
                ("7", "Guardar y relanzar la Acción 6",
                 "Se descargará automáticamente el archivo:\n\n"
                 "   📄  Purchasing Info Records.xlsx\n\n"
                 "Guárdelo en la carpeta  TARIFAS  del proyecto y, a continuación,\n"
                 "vuelva a ejecutar la  ▶ Acción 6  para continuar el proceso."),
            ]

            for num, titulo, desc in PASOS:
                frm_paso = tk.Frame(frame_steps, bg=C_BG, pady=3)
                frm_paso.pack(fill="x", padx=6, pady=2)

                tk.Label(frm_paso,
                         text=f" {num} ",
                         bg=C_MID, fg="white",
                         font=("Segoe UI", 10, "bold"),
                         padx=8, pady=4, relief="flat").pack(
                             side="left", anchor="n", padx=(0, 10))

                frm_txt = tk.Frame(frm_paso, bg="#EEF2F8", padx=10, pady=8)
                frm_txt.pack(side="left", fill="x", expand=True)

                tk.Label(frm_txt, text=titulo,
                         bg="#EEF2F8", fg="#1A2E4A",
                         font=("Segoe UI", 9, "bold"),
                         anchor="w", justify="left").pack(anchor="w")
                tk.Label(frm_txt, text=desc,
                         bg="#EEF2F8", fg=C_TEXT,
                         font=FONT_LABEL,
                         anchor="w", justify="left").pack(anchor="w", pady=(4, 0))

            # ── Botón cerrar ──────────────────────────────────────
            frm_cerrar = tk.Frame(instrucciones, bg=C_BG, pady=8, padx=14)
            frm_cerrar.pack(fill="x")
            tk.Button(frm_cerrar, text="Cerrar",
                      command=instrucciones.destroy,
                      font=FONT_LABEL, padx=14, pady=6,
                      cursor="hand2", relief="flat",
                      bg="#D8DEE8",
                      activebackground="#C0C8D8").pack(side="right")

            instrucciones.update_idletasks()
            sw = instrucciones.winfo_screenwidth()
            sh = instrucciones.winfo_screenheight()
            w  = max(instrucciones.winfo_width(),  580)
            h  = min(instrucciones.winfo_height(), sh - 80)
            instrucciones.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        def _generar_nuevo():
            ventana.destroy()
            self.consola.log(
                "Usuario eligió: crear Registro Info Nuevo.", "INFO")
            self._crear_rinfo_nuevo(nombre_proveedor, ruta_tarifa)

        def _cancelar():
            ventana.destroy()
            self.consola.log("Acción 6 cancelada por el usuario.", "WARN")

        tk.Button(frm_btn,
                  text="📥  Descargar Purchasing desde SAP",
                  command=_descargar_sap,
                  bg=C_MID, fg=C_TEXT_LIGHT,
                  font=FONT_BOLD, padx=12, pady=8,
                  cursor="hand2", relief="flat",
                  activebackground=C_HOVER,
                  activeforeground=C_ACCENT2).pack(fill="x", pady=(0, 4))

        tk.Button(frm_btn,
                  text="🆕  Crear Registro Info Nuevo para este proveedor",
                  command=_generar_nuevo,
                  bg="#7B4FA6", fg="white",
                  font=FONT_BOLD, padx=12, pady=8,
                  cursor="hand2", relief="flat",
                  activebackground="#9B6FC6").pack(fill="x", pady=(0, 4))

        tk.Button(frm_btn, text="Cancelar", command=_cancelar,
                  font=FONT_LABEL, padx=12, pady=6,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(fill="x")

        # Centrar sobre el padre
        ventana.update_idletasks()
        pw = self.winfo_rootx() + self.winfo_width()  // 2
        ph = self.winfo_rooty() + self.winfo_height() // 2
        w, h = ventana.winfo_width(), ventana.winfo_height()
        ventana.geometry(f"+{pw - w // 2}+{ph - h // 2}")

    def _crear_rinfo_nuevo(self, nombre_proveedor, ruta_tarifa):
        """
        Crea un nuevo archivo  RINFO (PROVEEDOR) [dd-mm-yyyy].xlsx
        copiando la estructura del Purchasing Info Records*.xlsx de TARIFAS/
        y eliminando todos los datos desde la fila 2 en las 4 hojas:
          - Info general p.reg.info
          - Datos org.compr.p.reg.info
          - Det.precios p.registros info
          - Escalas p.cond.precio PIR
        La fila 1 (cabecera) se conserva intacta en todas las hojas.

        Nota: el archivo SAP exporta un atributo 's' en las columnas XML
        que openpyxl no soporta; se elimina antes de cargar el workbook.
        """
        import openpyxl
        import zipfile
        import re
        from datetime import datetime

        # ── Buscar Purchasing Info Records*.xlsx en TARIFAS/ ──────
        plantilla = None
        try:
            for entry in os.scandir(self.carpeta_tarifas):
                if (entry.is_file() and
                        entry.name.lower().startswith("purchasing info records") and
                        entry.name.lower().endswith(".xlsx")):
                    plantilla = entry.path
                    break
        except Exception:
            pass

        if not plantilla:
            messagebox.showerror(
                "Plantilla no encontrada",
                "❌  No se encontró ningún archivo\n"
                "«Purchasing Info Records*.xlsx» en la carpeta TARIFAS\n"
                "para usar como plantilla.\n\n"
                "Descargue primero el archivo desde SAP y\n"
                "colóquelo en la carpeta TARIFAS.",
                parent=self)
            self.consola.log(
                "No se encontró Purchasing Info Records como plantilla.", "ERROR")
            return

        self.consola.log(
            f"Plantilla encontrada: {os.path.basename(plantilla)}", "INFO")

        # ── Nombre del archivo destino ─────────────────────────────
        prov_sano = nombre_proveedor.strip()
        for ch in r'\/:*?"<>|':
            prov_sano = prov_sano.replace(ch, "")
        prov_sano = prov_sano.strip()
        fecha_hoy    = datetime.now().strftime("%d-%m-%Y")
        nombre_nuevo = f"RINFO ({prov_sano}) [{fecha_hoy}].xlsx"
        ruta_nuevo   = os.path.join(self.carpeta_tarifas, nombre_nuevo)

        # ── Copiar plantilla saneando XML (atributo 's' en <col>) ─
        # Los archivos exportados por SAP incluyen s="N" en elementos <col>
        # que openpyxl no reconoce. Se elimina al reescribir el ZIP.
        try:
            with zipfile.ZipFile(plantilla, "r") as zin:
                with zipfile.ZipFile(ruta_nuevo, "w", zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        data = zin.read(item.filename)
                        if (item.filename.startswith("xl/worksheets/sheet") and
                                item.filename.endswith(".xml")):
                            text = data.decode("utf-8")
                            text = re.sub(
                                r'(<col\b[^>]*?)\s+s=\"\d+\"', r'\1', text)
                            data = text.encode("utf-8")
                        zout.writestr(item, data)
        except Exception as e:
            messagebox.showerror(
                "Error al copiar plantilla",
                f"❌  No se pudo preparar el archivo plantilla:\n\n{e}",
                parent=self)
            self.consola.log(f"Error al copiar/sanear plantilla: {e}", "ERROR")
            return

        # ── Abrir y eliminar datos desde fila 2 en las 4 hojas ────
        HOJAS_PURCHASING = [
            "Info general p.reg.info",
            "Datos org.compr.p.reg.info",
            "Det.precios p.registros info",
            "Escalas p.cond.precio PIR",
        ]
        try:
            wb = openpyxl.load_workbook(ruta_nuevo)
        except Exception as e:
            messagebox.showerror(
                "Error al abrir archivo",
                f"❌  No se pudo abrir el archivo preparado:\n\n{e}",
                parent=self)
            self.consola.log(f"Error al abrir RINFO copiado: {e}", "ERROR")
            try:
                os.remove(ruta_nuevo)
            except Exception:
                pass
            return

        hojas_limpiadas = []
        for nombre_hoja in HOJAS_PURCHASING:
            if nombre_hoja not in wb.sheetnames:
                self.consola.log(
                    f"  └ Hoja «{nombre_hoja}» no encontrada — omitida", "WARN")
                continue
            ws = wb[nombre_hoja]
            max_r = ws.max_row
            if max_r >= 2:
                ws.delete_rows(2, max_r - 1)
                filas_eliminadas = max_r - 1
            else:
                filas_eliminadas = 0
            hojas_limpiadas.append(nombre_hoja)
            self.consola.log(
                f"  └ Hoja «{nombre_hoja}»: {filas_eliminadas} fila(s) eliminada(s)", "INFO")

        # ── Guardar ────────────────────────────────────────────────
        try:
            wb.save(ruta_nuevo)
            wb.close()
        except Exception as e:
            messagebox.showerror(
                "Error al guardar",
                f"❌  No se pudo guardar el archivo:\n\n"
                f"   {nombre_nuevo}\n\n{e}",
                parent=self)
            self.consola.log(f"Error al guardar RINFO nuevo: {e}", "ERROR")
            try:
                os.remove(ruta_nuevo)
            except Exception:
                pass
            return

        self.consola.log(f"Registro Info Nuevo creado: {nombre_nuevo}", "OK")

        # ── Preguntar si abrir el archivo ──────────────────────────
        def _abrir():
            try:
                os.startfile(ruta_nuevo)
            except AttributeError:
                import subprocess
                subprocess.call(["xdg-open", ruta_nuevo])
            except Exception as ex:
                self.consola.log(f"No se pudo abrir automáticamente: {ex}", "WARN")

        resp = messagebox.askyesno(
            "Registro Info Nuevo creado",
            f"✔  Se ha creado el archivo:\n\n"
            f"   {nombre_nuevo}\n\n"
            f"   Hojas vaciadas ({len(hojas_limpiadas)}):\n"
            + "".join(f"      • {h}\n" for h in hojas_limpiadas) +
            f"\n📁  Guardado en:\n   {self.carpeta_tarifas}\n\n"
            "¿Desea abrir el archivo ahora?",
            parent=self)

        if resp:
            _abrir()

        self.destroy()


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
    código DALI -- bien porque ya existía en el LISTADO (asignado a un
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
                     el LISTADO antes de este guardado (fila_o_None es su
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
                origen = (f"ya en el listado, fila {fila}" if es_existente
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
    Una vez confirmado, escribe los valores en la columna D del archivo
    'LISTADO CODIGOS DALI - SAP.xlsx' (desbloquea con Princess2026 y
    vuelve a bloquear al guardar).
    Devuelve True si el usuario guardó cambios, False si canceló.
    """
    CLAVE_HOJA = "Princess2026"

    def __init__(self, parent, faltantes, ruta_listado, consola_fn=None,
                 descripciones=None, ruta_tarifa=None, catalogo_refresh_fn=None):
        """
        faltantes    : lista de (fila_xlsx_tarifa, codigo_dali)
        ruta_listado : ruta al archivo LISTADO CODIGOS DALI - SAP.xlsx
        consola_fn   : función log_consola(msg, nivel) opcional
        descripciones: dict {codigo_dali: descripcion_str} opcional
        ruta_tarifa  : ruta al archivo .xlsx de tarifa recién generado
                       (para escribir también los SAP en col A)
        catalogo_refresh_fn: función opcional (lista_codigos_dali) -> None,
                       llamada cuando la resolución de duplicados borra el
                       SAP de un DALI existente -- para que catalogo-dali-
                       completo.xlsx se mantenga sincronizado incluso para
                       códigos DALI ajenos a la tarifa en curso (ver
                       App._refrescar_sap_catalogo_para_dali)
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
                  text="\u2714  Guardar en LISTADO",
                  font=FONT_BOLD, fg=C_DARK, bg=C_ACCENT,
                  activebackground=C_ACCENT2, relief="flat", cursor="hand2",
                  command=self._guardar).pack(side="left", ipadx=12, ipady=4)

        tk.Button(frm_btns,
                  text="\u2718  Cancelar",
                  font=FONT_BOLD, fg=C_TEXT_LIGHT, bg=C_BORDER,
                  activebackground=C_HOVER, relief="flat", cursor="hand2",
                  command=self.destroy).pack(side="left", padx=8, ipadx=10, ipady=4)

        tk.Label(frm_btns,
                 text="Los cambios se escriben directamente en\nLISTADO CODIGOS DALI - SAP.xlsx",
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
            try:
                from openpyxl import load_workbook

                # ── Carga rápida: read_only=False pero keep_vba=False ──
                # rich_text=False evita parsear formato de texto enriquecido
                wb = load_workbook(self._ruta_listado, rich_text=False)
                ws = wb.active

                # Desproteger
                ws.protection.sheet   = False
                ws.protection.enabled = False

                # ── Construir mapa DALI→fila y SAP→(DALI, fila) con
                #    iter_rows (mucho más rápido que ws.cell() fila a
                #    fila porque lee en bloque). El segundo mapa es lo
                #    que permite detectar códigos SAP duplicados más
                #    abajo.
                _mapa_fila = {}
                _mapa_sap_existente = {}
                for idx, row in enumerate(
                        ws.iter_rows(min_row=8, min_col=1, max_col=4,
                                     values_only=True),
                        start=8):
                    v = row[0]
                    if v is None:
                        continue
                    try:
                        _dali_int = int(float(v))
                    except (ValueError, TypeError):
                        continue
                    _mapa_fila[_dali_int] = idx
                    v_sap = row[3]
                    if v_sap not in (None, ""):
                        try:
                            _mapa_sap_existente[int(float(v_sap))] = (_dali_int, idx)
                        except (ValueError, TypeError):
                            pass

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
                        ws.cell(fila, 4).value = None
                    if _filas_a_limpiar:
                        self._log(
                            f"  ✔ {len(_filas_a_limpiar)} código(s) SAP "
                            f"duplicados borrados del listado (ya existentes "
                            f"en conflicto).", "OK")

                # ── Escribir SAP en col D ──────────────────────
                escritos       = 0
                no_encontrados = []
                for cod_dali, cod_sap in nuevos.items():
                    if cod_dali in _mapa_fila:
                        ws.cell(_mapa_fila[cod_dali], 4).value = cod_sap
                        escritos += 1
                        self._log(
                            f"  ✔ DALI {cod_dali:06d} → SAP {cod_sap} "
                            f"(fila {_mapa_fila[cod_dali]})", "OK")
                    else:
                        no_encontrados.append(cod_dali)
                        self._log(
                            f"  ✘ DALI {cod_dali:06d} no encontrado en listado", "WARN")

                # Volver a proteger col D
                ws.protection.password = self.CLAVE_HOJA
                ws.protection.sheet    = True
                ws.protection.enable()

                wb.save(self._ruta_listado)
                self._guardado = True
                self._log(f"  └ Listado guardado: {escritos} SAP escritos.", "OK")

                # ── Refrescar catalogo-dali-completo.xlsx para los DALI ──
                #    cuyo SAP se acaba de borrar por duplicado -- pueden
                #    ser ajenos a esta tarifa, así que el enganche normal
                #    del catálogo (más abajo, en tarifa_app.py) nunca los
                #    tocaría por su cuenta.
                if _dali_limpiados and self._catalogo_refresh_fn:
                    try:
                        self._catalogo_refresh_fn(_dali_limpiados)
                    except Exception as _e_cat_ref:
                        self._log(
                            f"  ⚠  No se pudo refrescar el catálogo tras "
                            f"limpiar duplicados: {_e_cat_ref}", "WARN")

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
                    msg = f"✔  {e} código(s) SAP guardados en el LISTADO."
                    if et:
                        msg += f"\n✔  {et} código(s) escritos también en la tarifa."
                    if nf:
                        msg += f"\n\n⚠  {len(nf)} código(s) DALI no encontrados en el listado:\n"
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


