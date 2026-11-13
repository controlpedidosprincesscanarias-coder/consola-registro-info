# -*- coding: utf-8 -*-
"""
accion_tarifa.py
Acciones 4 y 6 — Generar la TARIFA final, procesar los archivos [INCOMPLETO]/INC
(transformación, SAP sin descripción, reinscripción SAP) y crear el Registro Info SAP.

Mixin de App (tarifa_app.py): sus métodos se usan como self.metodo(). Separado de tarifa_app.py
en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import threading
from app.comun.tarifa_core import (
    TARIFA_NAME,
    FOLDER_TARIFAS,
)
from app.ui.tarifa_dialogs import (
    VentanaSAPFaltantes,
)


class AccionesTarifa:
    # ── ACCIÓN 6 — Crear nuevo Registro Info SAP ──────────────
    def _crear_registro_sap(self):
        """
        Acción 6 — Muestra todos los archivos [INCOMPLETO] de la carpeta TARIFAS/
        (excluyendo los que empiecen por RINFO) y abre el seleccionado en Excel
        para edición manual.
        """
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            messagebox.showwarning("Sin proyecto",
                                   "Seleccione primero la carpeta del Proyecto.")
            return
        carpeta_tarifas = os.path.join(carpeta, FOLDER_TARIFAS)
        if not os.path.isdir(carpeta_tarifas):
            messagebox.showwarning("Sin carpeta TARIFAS",
                                   f"No se encontró la carpeta TARIFAS/ en:\n{carpeta}")
            return

        # ── Buscar archivos [INCOMPLETO] (excluyendo RINFO) ──────────────
        import re as _re6
        _pat = _re6.compile(r'.+\[\d{2}-\d{2}-\d{4}\].*\.xlsx$', _re6.IGNORECASE)
        archivos_inc = []
        try:
            for entry in sorted(os.scandir(carpeta_tarifas), key=lambda e: e.name.lower()):
                nombre = entry.name
                nu = nombre.upper()
                if (not entry.is_file()
                        or not _pat.match(nombre)
                        or nu.startswith("RINFO")
                        or "[INCOMPLETO]" not in nu):
                    continue
                ts = entry.stat().st_mtime
                fecha = datetime.fromtimestamp(ts).strftime("%d-%m-%Y %H:%M")
                archivos_inc.append({
                    "nombre": nombre,
                    "ruta":   entry.path,
                    "fecha":  fecha,
                })
        except PermissionError:
            pass

        if not archivos_inc:
            messagebox.showinfo(
                "Sin archivos incompletos",
                "No se encontraron archivos [INCOMPLETO] en la carpeta TARIFAS/.\n\n"
                "Cuando existan tarifas con códigos SAP pendientes o líneas duplicadas,\n"
                "aparecerán aquí para su revisión manual.")
            return

        self._ventana_accion6_incompletos(archivos_inc)

    def _ventana_accion6_incompletos(self, archivos):
        """
        Ventana de Acción 6: lista archivos [INCOMPLETO] y abre el seleccionado en Excel.
        """
        SEL_PURPLE = "#7B4FA6"
        SEL_FG     = "#FFFFFF"
        LIST_BG    = "#FFFFFF"
        LIST_FG    = "#1A1A2E"
        LIST_ALT   = "#F5F5FA"
        HDR_BG     = "#1E1E2E"
        INFO_BG    = "#FFF8E7"
        INFO_FG    = "#7A5800"
        INFO_BDR   = "#E8C84B"

        top = tk.Toplevel(self)
        top.title("Acción 6 — Archivos incompletos para edición manual")
        top.configure(bg=HDR_BG)
        top.resizable(True, True)
        top.minsize(700, 460)
        top.grab_set()

        # ── Cabecera ──────────────────────────────────────────────────────
        frm_hdr = tk.Frame(top, bg=HDR_BG, padx=18, pady=14)
        frm_hdr.pack(fill="x")
        frm_icon = tk.Frame(frm_hdr, bg=HDR_BG)
        frm_icon.pack(fill="x")
        tk.Label(frm_icon, text="✏️",
                 bg=HDR_BG, fg="#FFFFFF",
                 font=("Segoe UI", 20)).pack(side="left", padx=(0, 10))
        frm_tit = tk.Frame(frm_icon, bg=HDR_BG)
        frm_tit.pack(side="left")
        tk.Label(frm_tit,
                 text="Edición manual de tarifas incompletas",
                 bg=HDR_BG, fg="#FFFFFF",
                 font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(frm_tit,
                 text="Seleccione un archivo para abrirlo en Excel y editarlo manualmente.",
                 bg=HDR_BG, fg="#AAAACC",
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 0))

        # ── Panel informativo ─────────────────────────────────────────────
        frm_info = tk.Frame(top, bg=INFO_BG,
                            highlightbackground=INFO_BDR,
                            highlightthickness=1,
                            padx=14, pady=10)
        frm_info.pack(fill="x", padx=14, pady=(8, 0))
        tk.Label(frm_info,
                 text="ℹ  Instrucciones de uso",
                 bg=INFO_BG, fg=INFO_FG,
                 font=("Segoe UI", 9, "bold"),
                 anchor="w").pack(anchor="w")
        tk.Label(frm_info,
                 text=(
                     "1.  Seleccione el archivo y pulse  «Abrir en Excel».\n"
                     "2.  Revise el archivo: elimine líneas duplicadas e introduzca\n"
                     "     los códigos SAP que estén pendientes.\n"
                     "3.  Guarde y cierre el archivo en Excel.\n"
                     "4.  Ejecute de nuevo  Acción 4  para comprobación y validación final."
                 ),
                 bg=INFO_BG, fg=INFO_FG,
                 font=("Segoe UI", 8),
                 justify="left",
                 anchor="w").pack(anchor="w", pady=(4, 0))

        # ── Botones (side=bottom) ─────────────────────────────────────────
        frm_btn = tk.Frame(top, bg=HDR_BG, padx=14, pady=10)
        frm_btn.pack(fill="x", side="bottom")

        def _abrir(_evt=None):
            sel = tv.selection()
            if not sel:
                messagebox.showwarning("Sin selección",
                                       "Seleccione un archivo para continuar.",
                                       parent=top)
                return
            idx = _iids.index(sel[0])
            ruta = archivos[idx]["ruta"]
            self.log_consola(
                f"Acción 6 — Abriendo en Excel: {os.path.basename(ruta)}", "INFO")
            try:
                os.startfile(ruta)
            except AttributeError:
                import subprocess
                subprocess.Popen(["xdg-open", ruta])
            except Exception as _e:
                messagebox.showerror("Error al abrir",
                                     f"No se pudo abrir el archivo:\n{_e}",
                                     parent=top)

        def _cerrar(_evt=None):
            top.destroy()

        tk.Button(frm_btn,
                  text="Cerrar",
                  bg="#3A3A55", fg="#DDDDEE",
                  font=("Segoe UI", 10),
                  relief="flat", padx=16, pady=6,
                  cursor="hand2",
                  command=_cerrar).pack(side="right", padx=(6, 0))
        tk.Button(frm_btn,
                  text="✏️  Abrir en Excel",
                  bg=SEL_PURPLE, fg="#FFFFFF",
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=18, pady=6,
                  cursor="hand2",
                  command=_abrir).pack(side="right")

        top.bind("<Return>", _abrir)
        top.bind("<Escape>", _cerrar)

        # ── Lista ─────────────────────────────────────────────────────────
        frm_list = tk.Frame(top, bg=LIST_BG, bd=1, relief="solid")
        frm_list.pack(fill="both", expand=True, padx=14, pady=(8, 0))

        style = ttk.Style(top)
        _sn = "Ac6Inc.Treeview"
        style.theme_use("clam")
        style.configure(_sn,
                        background=LIST_BG,
                        foreground=LIST_FG,
                        fieldbackground=LIST_BG,
                        rowheight=28,
                        font=("Segoe UI", 10),
                        borderwidth=0, relief="flat")
        style.configure(f"{_sn}.Heading",
                        background="#EEEEF5",
                        foreground="#555566",
                        font=("Segoe UI", 9, "bold"),
                        relief="flat", borderwidth=0)
        style.map(_sn,
                  background=[("selected", SEL_PURPLE)],
                  foreground=[("selected", SEL_FG)])

        tv = ttk.Treeview(frm_list,
                          columns=("nombre", "fecha"),
                          show="headings",
                          selectmode="browse",
                          style=_sn)
        tv.heading("nombre", text="Archivo [INCOMPLETO]")
        tv.heading("fecha",  text="Última modificación")
        tv.column("nombre", stretch=True,  anchor="w", minwidth=360)
        tv.column("fecha",  stretch=False, anchor="e", width=170)

        vsb = ttk.Scrollbar(frm_list, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        tv.pack(side="left", fill="both", expand=True)

        tv.tag_configure("even", background=LIST_BG,  foreground=LIST_FG)
        tv.tag_configure("odd",  background=LIST_ALT, foreground=LIST_FG)

        _iids = []
        for i, a in enumerate(archivos):
            tag = "odd" if i % 2 else "even"
            iid = tv.insert("", "end",
                            values=(f"⚠  {a['nombre']}", a["fecha"]),
                            tags=(tag,))
            _iids.append(iid)

        if _iids:
            tv.selection_set(_iids[0])
            tv.focus(_iids[0])

        tv.bind("<Double-Button-1>", _abrir)

        # ── Centrar ───────────────────────────────────────────────────────
        _w, _h = 740, 490
        _sw = top.winfo_screenwidth()
        _sh = top.winfo_screenheight()
        top.geometry(f"{_w}x{_h}+{(_sw-_w)//2}+{max(60, (_sh-_h)//3)}")

    def _procesar(self):
        carpeta = self._carpeta_proyecto.get().strip()
        carpeta_tarifas = os.path.join(carpeta, FOLDER_TARIFAS) if carpeta else ""

        # ── PRIORIDAD 1: TARIFA.xls disponible → flujo normal ───────
        if self._ruta_tarifa and os.path.isfile(self._ruta_tarifa):
            self._ejecutar_transformacion()
            return

        # ── PRIORIDAD 2: archivos *[dd-mm-yyyy].xlsx con [INCOMPLETO] ──
        # Reglas de visualización:
        #   • Si empieza por RINFO  → NO mostrar
        #   • Si empieza por INC    → mostrar SÍ, pero omitir "INC " en la descripción
        #   • Si NO contiene [INCOMPLETO] → NO mostrar
        # El archivo seleccionado se trata siempre como flujo INC (reinscripción SAP).
        import re as _re_proc
        _patron_fecha = _re_proc.compile(r'.+\[\d{2}-\d{2}-\d{4}\].*\.xlsx$', _re_proc.IGNORECASE)
        _archivos_inc = []
        if os.path.isdir(carpeta_tarifas):
            try:
                for _entry in sorted(os.scandir(carpeta_tarifas),
                                     key=lambda e: e.name.lower()):
                    _nombre = _entry.name
                    _nu     = _nombre.upper()
                    if (not _entry.is_file()
                            or not _patron_fecha.match(_nombre)
                            or _nu.startswith("RINFO")
                            or "[INCOMPLETO]" not in _nu):
                        continue
                    _ts    = _entry.stat().st_mtime
                    _fecha = datetime.fromtimestamp(_ts).strftime("%d-%m-%Y %H:%M")
                    # Nombre para mostrar: sin extensión y sin prefijo "INC "
                    _nd = os.path.splitext(_nombre)[0]
                    if _nd.upper().startswith("INC "):
                        _nd = _nd[4:].strip()
                    _archivos_inc.append({
                        "nombre_archivo": _nombre,
                        "nombre_display":  _nd,
                        "ruta":            _entry.path,
                        "fecha_mod":       _fecha,
                    })
            except PermissionError:
                pass

        if _archivos_inc:
            self.log_consola(
                f"Acción 4: {len(_archivos_inc)} archivo(s) [INCOMPLETO] disponibles "
                f"para procesar.", "INFO")
            _ruta_sel = self._elegir_inc_para_procesar(_archivos_inc)
            if _ruta_sel:
                self._ruta_resultado = _ruta_sel
                self._ejecutar_reinscripcion_sap_desde_inc(_ruta_sel)
            return

        # ── Sin archivos procesables → avisar ────────────────────────
        messagebox.showwarning(
            "Sin archivo",
            f"No se encontró {TARIFA_NAME} en la carpeta TARIFAS/\n\n"
            "Tampoco se encontraron archivos de tarifa [INCOMPLETO] con formato\n"
            "   * [dd-mm-yyyy].xlsx\n\n"
            "Coloca el archivo .xls del proveedor en:\n"
            f"  {carpeta_tarifas}\n\n"
            "Luego pulsa  🔍 Buscar TARIFA.xls  para localizarlo.")

    def _elegir_inc_para_procesar(self, archivos):
        """
        Diálogo modal de selección de tarifa [INCOMPLETO] para Acción 4.
        Estilo idéntico a VentanaSeleccionTarifaRI (Acción 6): lista blanca,
        selección en morado, botones en esquina inferior derecha.
        Devuelve la ruta del archivo elegido o None si se cancela.
        """
        SEL_PURPLE = "#7B4FA6"   # morado Acción 6
        SEL_FG     = "#FFFFFF"
        LIST_BG    = "#FFFFFF"
        LIST_FG    = "#1A1A2E"
        LIST_ALT   = "#F5F5FA"
        HDR_BG     = "#1E1E2E"   # cabecera oscura igual que Acción 6

        sel = {"ruta": None}

        top = tk.Toplevel(self)
        top.title("Acción 4 — Seleccionar tarifa a procesar")
        top.configure(bg=HDR_BG)
        top.resizable(True, True)
        top.minsize(620, 380)
        top.grab_set()

        # ── Cabecera (estilo Acción 6) ────────────────────────────
        frm_hdr = tk.Frame(top, bg=HDR_BG, padx=18, pady=14)
        frm_hdr.pack(fill="x")

        frm_icon_title = tk.Frame(frm_hdr, bg=HDR_BG)
        frm_icon_title.pack(fill="x")
        tk.Label(frm_icon_title, text="📋",
                 bg=HDR_BG, fg="#FFFFFF",
                 font=("Segoe UI", 20)).pack(side="left", padx=(0, 10))
        frm_titles = tk.Frame(frm_icon_title, bg=HDR_BG)
        frm_titles.pack(side="left")
        tk.Label(frm_titles,
                 text="Procesar tarifa [INCOMPLETO]",
                 bg=HDR_BG, fg="#FFFFFF",
                 font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(frm_titles,
                 text="Seleccione el archivo de tarifa que desea reinscribir con códigos SAP:",
                 bg=HDR_BG, fg="#AAAACC",
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 0))

        # ── Botones (esquina inferior derecha, estilo Acción 6) ───
        # Se empaquetan ANTES que la lista para que siempre queden visibles
        frm_btn = tk.Frame(top, bg=HDR_BG, padx=14, pady=10)
        frm_btn.pack(fill="x", side="bottom")

        def _confirmar(_evt=None):
            focused = tv.selection()
            if not focused:
                messagebox.showwarning(
                    "Sin selección",
                    "Selecciona un archivo de la lista para continuar.",
                    parent=top)
                return
            idx = _iids.index(focused[0])
            sel["ruta"] = archivos[idx]["ruta"]
            top.destroy()

        def _cancelar(_evt=None):
            top.destroy()

        tk.Button(frm_btn,
                  text="Cancelar",
                  bg="#3A3A55", fg="#DDDDEE",
                  font=("Segoe UI", 10),
                  relief="flat", padx=16, pady=6,
                  cursor="hand2",
                  command=_cancelar).pack(side="right", padx=(6, 0))
        tk.Button(frm_btn,
                  text="✔  Seleccionar",
                  bg=SEL_PURPLE, fg="#FFFFFF",
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=18, pady=6,
                  cursor="hand2",
                  command=_confirmar).pack(side="right")

        top.bind("<Return>",  _confirmar)
        top.bind("<Escape>",  _cancelar)

        # ── Contenedor lista (fondo blanco) — se empaqueta después de botones ──
        frm_list = tk.Frame(top, bg=LIST_BG,
                            bd=1, relief="solid",
                            padx=0, pady=0)
        frm_list.pack(fill="both", expand=True,
                      padx=14, pady=(10, 0))

        # Treeview con dos columnas: nombre | fecha_mod
        style = ttk.Style(top)
        _style_name = "AccINC.Treeview"
        style.theme_use("clam")
        style.configure(_style_name,
                        background=LIST_BG,
                        foreground=LIST_FG,
                        fieldbackground=LIST_BG,
                        rowheight=28,
                        font=("Segoe UI", 10),
                        borderwidth=0,
                        relief="flat")
        style.configure(f"{_style_name}.Heading",
                        background=LIST_BG,
                        foreground="#888899",
                        font=("Segoe UI", 9),
                        relief="flat",
                        borderwidth=0)
        style.map(_style_name,
                  background=[("selected", SEL_PURPLE)],
                  foreground=[("selected", SEL_FG)])

        tv = ttk.Treeview(frm_list,
                          columns=("nombre", "fecha"),
                          show="headings",
                          selectmode="browse",
                          style=_style_name)
        tv.heading("nombre", text="Archivo")
        tv.heading("fecha",  text="Modificado")
        tv.column("nombre", stretch=True,  anchor="w", minwidth=300)
        tv.column("fecha",  stretch=False, anchor="e", width=160)

        vsb = ttk.Scrollbar(frm_list, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        tv.pack(side="left",   fill="both", expand=True)

        # Rellenar filas — nombre_display ya tiene INC eliminado
        _iids = []
        for i, a in enumerate(archivos):
            iid = tv.insert("", "end",
                            values=(a["nombre_display"], a["fecha_mod"]),
                            tags=("odd" if i % 2 else "even",))
            _iids.append(iid)
            # Guardar índice en el iid para recuperar ruta
        style.configure(_style_name, rowheight=28)
        tv.tag_configure("odd",  background=LIST_ALT)
        tv.tag_configure("even", background=LIST_BG)

        # Pre-seleccionar primera fila
        if _iids:
            tv.selection_set(_iids[0])
            tv.focus(_iids[0])

        # bind doble clic aquí, tras crear tv
        tv.bind("<Double-Button-1>", _confirmar)

        # ── Centrar: horizontal centrado, vertical en tercio superior ─
        _w, _h = 680, 420
        self.update_idletasks()
        _sw = top.winfo_screenwidth()
        _sh = top.winfo_screenheight()
        _px = (_sw - _w) // 2
        _py = max(60, (_sh - _h) // 3)   # tercio superior de la pantalla
        top.geometry(f"{_w}x{_h}+{_px}+{_py}")

        self.wait_window(top)
        return sel["ruta"]

    def _ejecutar_transformacion(self):
        self._habilitar_btn_procesar(False)
        self._btn_buscar.config(state="disabled")
        self._habilitar_btn_resultado(False)

        def proceso():
            try:
                from app.tarifas.transformar_tarifa import transformar
            except ImportError:
                self.after(0, lambda: messagebox.showerror(
                    "M\xf3dulo no encontrado",
                    "No se encontr\xf3 transformar_tarifa.py junto al ejecutable."))
                self.after(0, lambda: self._habilitar_btn_procesar(True))
                self.after(0, lambda: self._btn_buscar.config(state="normal"))
                return

            self.consola.separador("INICIO DE TRANSFORMACI\xd3N")
            self.log_consola(f"Archivo fuente: {self._ruta_tarifa}", "INFO")
            self.after(0, lambda: self._barra.indeterminate("Procesando..."))

            TOTAL_PASOS  = 22
            pasos_hechos = [0]

            def log_con_progreso(msg, nivel=">>"):
                self.log_consola(msg, nivel)
                if msg.startswith("Paso "):
                    pasos_hechos[0] += 1
                    p = pasos_hechos[0]
                    self.after(0, lambda pp=p: self._barra.set(
                        pp, TOTAL_PASOS, f"Paso {pp} de {TOTAL_PASOS}..."))

            try:
                resultado = transformar(self._ruta_tarifa, log_fn=log_con_progreso)
                # V10+: transformar devuelve (ruta_salida, _analisis)
                if isinstance(resultado, tuple):
                    ruta_salida, _analisis = resultado
                else:
                    ruta_salida, _analisis = resultado, {}

                self._ruta_resultado = ruta_salida
                self.log_consola(f"Archivo .xlsx generado: {os.path.basename(ruta_salida)}", "OK")

                # ── Informe de duplicados ────────────────────
                _dup_sap  = _analisis.get("dup_sap",  [])
                _dup_dali = _analisis.get("dup_dali", [])
                _falt_sap = _analisis.get("faltantes_sap", [])
                _sin_desc_und_prev = _analisis.get("sin_desc_und", [])

                # ── Enganche del catálogo DALI completo (2026-09-22) ──
                # Mismo criterio que decide si el archivo lleva [INCOMPLETO]
                # en _finalizar_y_guardar(): solo se actualiza el catálogo
                # cuando la tarifa queda lista a la primera (sin duplicados,
                # sin descripciones/unidades SAP pendientes) -- si la
                # transformación sale [INCOMPLETO], el catálogo se actualizará
                # más tarde, al corregirla y procesarla vía Acción 4/6.
                if not (_dup_sap or _dup_dali or _sin_desc_und_prev):
                    if self._catalogo_reciente_o_confirmar():
                        self._actualizar_codigo_proveedor_catalogo(ruta_salida)
                self.after(0, lambda: self._barra.completado("Transformaci\xf3n completada"))

                if _dup_sap or _dup_dali:
                    _lineas_dup = []
                    if _dup_sap:
                        self.log_consola(
                            f"  \u26a0  {len(_dup_sap)} c\xf3digo(s) SAP duplicados: "
                            + ", ".join(str(d) for d in _dup_sap[:10])
                            + (" ..." if len(_dup_sap) > 10 else ""), "WARN")
                        _lineas_dup.append(
                            f"\u2022  SAP: {len(_dup_sap)} c\xf3digo(s) duplicados \u00fanicos\n"
                            "   " + ", ".join(str(d) for d in _dup_sap[:10])
                            + (" ..." if len(_dup_sap) > 10 else ""))
                    if _dup_dali:
                        self.log_consola(
                            f"  \u26a0  {len(_dup_dali)} c\xf3digo(s) DALI duplicados: "
                            + ", ".join(str(d) for d in _dup_dali[:10])
                            + (" ..." if len(_dup_dali) > 10 else ""), "WARN")
                        _lineas_dup.append(
                            f"\u2022  DALI: {len(_dup_dali)} c\xf3digo(s) duplicados \u00fanicos\n"
                            "   " + ", ".join(str(d) for d in _dup_dali[:10])
                            + (" ..." if len(_dup_dali) > 10 else ""))
                else:
                    self.log_consola("  \u2714  Sin duplicados en SAP ni DALI.", "OK")

                self.consola.separador("FIN DE TRANSFORMACI\xd3N")
                self._ruta_tarifa = None
                self.after(0, lambda: self._set_estado(False, reset_resultado=False))
                self.after(10, lambda: self._habilitar_btn_resultado(True))

                # ── Mensaje resumen con duplicados ───────────
                _hay_warn = bool(_dup_sap or _dup_dali or _sin_desc_und_prev)

                # Ruta abreviada: mostrar solo las últimas 3 partes de la ruta
                _partes_ruta = ruta_salida.replace("\\", "/").split("/")
                _ruta_corta  = ("...\\" + "\\".join(_partes_ruta[-3:])
                                if len(_partes_ruta) > 3 else ruta_salida)
                _nombre_arch = os.path.basename(ruta_salida)

                _msg_ok = (
                    f"✔  Archivo generado correctamente:\n"
                    f"   {_nombre_arch}\n"
                    f"   {_ruta_corta}"
                )

                if _dup_sap or _dup_dali:
                    _msg_ok += "\n\n─────────────────────────────────────────────"
                    _msg_ok += "\n⚠  Se detectaron referencias duplicadas:"
                    if _dup_sap:
                        _msg_ok += f"\n   •  SAP:   {len(_dup_sap)} código(s) único(s) duplicados"
                    if _dup_dali:
                        _msg_ok += f"\n   •  DALI:  {len(_dup_dali)} código(s) único(s) duplicados"
                    _msg_ok += (
                        "\n\nLas celdas afectadas están resaltadas en ROJO en el archivo.\n"
                        "Mientras existan duplicados, el Registro Info no se podrá\n"
                        "tramitar correctamente.\n\n"
                        "Puede revisar y corregir el archivo:\n"
                        "  →  Ahora:       cierre esta ventana y pulse  Acción 5\n"
                        "                  para abrirlo directamente.\n"
                        "  →  Más tarde:   abra el archivo desde  Acción 6  en cualquier\n"
                        "                  momento para retomar la corrección.\n\n"
                        "Una vez eliminadas todas las duplicidades, renombre el\n"
                        "archivo suprimiendo  [INCOMPLETO]  del nombre para que\n"
                        "quede preparada como  TARIFA PREPARADA PARA CREAR O MODIFICAR REGISTRO INFO SAP."
                    )

                if _sin_desc_und_prev:
                    _n_sdu   = len(_sin_desc_und_prev)
                    _lst_sdu = (", ".join(str(d) for d in _sin_desc_und_prev[:10])
                                + (" ..." if _n_sdu > 10 else ""))
                    _msg_ok += (
                        "\n\n─────────────────────────────────────────────"
                        f"\n⚠  {_n_sdu} artículo(s) sin descripción ni unidad SAP:\n"
                        f"   Códigos DALI afectados:  {_lst_sdu}\n"
                        "   (celdas marcadas en AMARILLO en el archivo)\n\n"
                        "El listado de Productos SAP en RECURSOS está incompleto\n"
                        "o desactualizado. Para resolverlo:\n\n"
                        "  📥  1. Descargue el listado actualizado desde SAP:\n"
                        "          https://my405727.s4hana.cloud.sap/ui#Shell-home\n"
                        "          Guárdelo en la carpeta  RECURSOS/\n\n"
                        "  ⚙️   2. Ejecute  Acción 3  (Actualizar listado materiales SAP)\n"
                        "          para incorporar los nuevos datos al catálogo.\n\n"
                        "  🔄  3. Ejecute de nuevo  Acción 4  (Procesar y generar TARIFA)\n"
                        "          para que las descripciones queden completas.\n\n"
                        "El archivo permanecerá como [INCOMPLETO] hasta que se\n"
                        "resuelva esta incidencia."
                    )

                _titulo_ok = "Completado con advertencias ⚠" if _hay_warn else "Procesado correctamente"
                _fn_msg    = messagebox.showwarning if _hay_warn else messagebox.showinfo
                self.after(0, lambda m=_msg_ok, t=_titulo_ok, fn=_fn_msg: fn(t, m))

                # ── SAP faltantes → preguntar al usuario ─────
                if _falt_sap:
                    self.log_consola(
                        f"  \u26a0  {len(_falt_sap)} c\xf3digo(s) DALI sin SAP asignado.", "WARN")
                    for _f, _d in _falt_sap[:5]:
                        self.log_consola(f"      DALI {str(_d).zfill(6)}", "WARN")
                    if len(_falt_sap) > 5:
                        self.log_consola(
                            f"      ... y {len(_falt_sap) - 5} m\xe1s", "WARN")

                    # Leer descripcion y formato de cada fila faltante desde el xlsx generado
                    # Estructura del xlsx: A=SAP, B=COD_DALI, C=COD_PROV, D=DESCRIPCION, E=FORMATO
                    _descripciones = {}   # {cod_dali: (descripcion_str, formato_str)}
                    try:
                        from openpyxl import load_workbook as _lw
                        _wb_tmp = _lw(ruta_salida, read_only=True, data_only=True)
                        _ws_tmp = _wb_tmp.active
                        for _frow, _dcod in _falt_sap:
                            _desc_val   = _ws_tmp.cell(_frow, 4).value   # col D = DESCRIPCION
                            _format_val = _ws_tmp.cell(_frow, 5).value   # col E = FORMATO
                            _descripciones[_dcod] = (
                                str(_desc_val).strip()   if _desc_val   else "",
                                str(_format_val).strip() if _format_val else ""
                            )
                        _wb_tmp.close()
                    except Exception as _e_desc:
                        self.log_consola(f"  [INFO] No se pudo leer descripciones: {_e_desc}", "INFO")

                    # Catálogo completo (v10.2.50: ya no hay LISTADO CODIGOS DALI - SAP): RECURSOS del
                    # proyecto o, si no, el de la carpeta padre de la del .xls.
                    from app.comun import catalogo_local as _cl
                    _carpeta_xls = (os.path.dirname(os.path.abspath(self._ruta_resultado))
                                    if self._ruta_resultado else None)
                    _ruta_listado = _cl.localizar_catalogo(self._carpeta_proyecto.get().strip() or None,
                                                           _carpeta_xls)
                    if _ruta_listado:
                        self.log_consola(f"  [DIAG] ✔ Catálogo: {_ruta_listado}", "OK")
                    else:
                        self.log_consola("  [DIAG] ✘ catalogo-dali-completo.xlsx no encontrado.", "WARN")

                    # Construir mensaje para el diálogo de elección
                    _msg_falt = (
                        f"Se encontraron {len(_falt_sap)} c\xf3digo(s) DALI sin c\xf3digo SAP asignado:\n\n"
                        + "\n".join(
                            f"  \u2022  DALI {str(d).zfill(6)}  {_descripciones.get(d, ('',''))[0][:40]}"
                            for _, d in _falt_sap[:10])
                        + (f"\n  ... y {len(_falt_sap)-10} m\xe1s" if len(_falt_sap) > 10 else "")
                        + "\n\n\xbfC\xf3mo deseas introducir los c\xf3digos SAP?"
                    )

                    if _ruta_listado:
                        def _mostrar_dialogo_sap(
                                falt=_falt_sap, rl=_ruta_listado, msg=_msg_falt,
                                descs=_descripciones, rs=ruta_salida,
                                dups=bool(_dup_sap or _dup_dali)):
                            from tkinter import messagebox
                            resp = messagebox.askyesnocancel(
                                "C\xf3digos SAP faltantes",
                                msg + "\n\n"
                                "[S\xcd]       \u2192  Introducir ahora desde la aplicaci\xf3n\n"
                                "[NO]       \u2192  Modificar directamente en el Excel\n"
                                "[CANCELAR] \u2192  Ignorar por ahora",
                                icon="warning")
                            if resp is True:
                                # Abrir ventana de entrada SAP
                                ven = VentanaSAPFaltantes(
                                    self, falt, rl,
                                    consola_fn=self.log_consola,
                                    descripciones=descs,
                                    ruta_tarifa=rs,
                                    catalogo_refresh_fn=self._tras_cambiar_sap_en_catalogo)
                                self.wait_window(ven)
                                if ven.guardado:   # (13 nov 2026) ¿faltan los listados de SAP actualizados?
                                    self._avisar_listados_sap(list(ven.nuevos_sap.values()))

                                # ── Tras cerrar la ventana: comprobar si
                                #    se completaron todos y no hay duplicados
                                if ven.guardado:
                                    _todos_completados = (ven.escritos_en_tarifa == len(falt))
                                    _sin_duplicados    = not dups
                                    if _todos_completados and _sin_duplicados:
                                        # Renombrar eliminando [INCOMPLETO]
                                        try:
                                            _dir  = os.path.dirname(rs)
                                            _base = os.path.basename(rs)
                                            import re as _re
                                            _nuevo = _re.sub(
                                                r'\[INCOMPLETO\]\s*', '', _base,
                                                flags=_re.IGNORECASE).strip()
                                            _ruta_nueva = os.path.join(_dir, _nuevo)
                                            if _nuevo != _base:
                                                if os.path.exists(_ruta_nueva):
                                                    os.remove(_ruta_nueva)
                                                    self.log_consola(
                                                        f"  🗑 Sustituido archivo final anterior: {_nuevo}",
                                                        "INFO")
                                                os.rename(rs, _ruta_nueva)
                                                self._ruta_resultado = _ruta_nueva
                                                self.log_consola(
                                                    f"  ✔ Archivo renombrado → {_nuevo}", "OK")
                                                # ── Enganche del catálogo DALI completo (2026-09-23) ──
                                                # Mismo criterio que en el flujo principal de
                                                # transformar(): al quedar la tarifa lista aquí
                                                # (todos los SAP completados, sin duplicados) hay
                                                # que actualizar también el código proveedor del
                                                # catálogo -- antes este camino (VentanaSAPFaltantes
                                                # dentro de Acción 4) no lo hacía.
                                                if self._catalogo_reciente_o_confirmar():
                                                    self._actualizar_codigo_proveedor_catalogo(_ruta_nueva)
                                                messagebox.showinfo(
                                                    "Tarifa completada",
                                                    f"✔  Todos los códigos SAP han sido cumplimentados"
                                                    f" y no existen duplicados.\n\n"
                                                    f"El archivo ha sido renombrado a:\n{_nuevo}\n\n"
                                                    f"\n✔  TARIFA PREPARADA PARA CREAR O MODIFICAR REGISTRO INFO SAP.")
                                        except Exception as _e_ren:
                                            self.log_consola(
                                                f"  ✘ Error al renombrar: {_e_ren}", "WARN")
                                    else:
                                        _motivos = []
                                        if not _todos_completados:
                                            _pendientes = len(falt) - ven.escritos_en_tarifa
                                            _motivos.append(
                                                f"{_pendientes} código(s) SAP aún pendiente(s)")
                                        if not _sin_duplicados:
                                            _motivos.append("existen duplicados en el archivo")
                                        self.log_consola(
                                            f"  ⚠ Archivo mantiene [INCOMPLETO]: "
                                            + "; ".join(_motivos), "WARN")
                            elif resp is False:
                                # Abrir el Excel directamente
                                try:
                                    os.startfile(rl)
                                except Exception:
                                    import subprocess
                                    subprocess.Popen(["start", "", rl], shell=True)
                                messagebox.showinfo(
                                    "Modificar en Excel",
                                    f"Se abri\xf3 el archivo:\n{rl}\n\n"
                                    "Cuando hayas a\xf1adido los c\xf3digos SAP:\n"
                                    "  1. Actualiza el archivo Productos en RECURSOS.\n"
                                    "  2. Ejecuta de nuevo Acci\xf3n 3 (Actualizar materiales SAP).\n"
                                    "  3. Ejecuta de nuevo Acci\xf3n 4 (Procesar y generar TARIFA).")
                            # resp is None → ignorar
                        self.after(200, _mostrar_dialogo_sap)
                    else:
                        # No se encontró el listado → solo avisar
                        def _avisar_sin_listado(falt=_falt_sap):
                            from tkinter import messagebox
                            messagebox.showwarning(
                                "C\xf3digos SAP faltantes",
                                f"{len(falt)} c\xf3digo(s) DALI no tienen SAP asignado.\n\n"
                                "No se encontr\xf3 RECURSOS/catalogo-dali-completo.xlsx.\n"
                                "Ejecuta la Acci\xf3n 2 (o desc\xe1rgalo de Cat\xe1logo Asignaciones) y vuelve a intentarlo.")
                        self.after(200, _avisar_sin_listado)



            except Exception as e:
                self.log_consola(f"ERROR durante la transformación: {e}", "ERROR")
                self.after(0, lambda: self._barra.reset("Error en la transformación"))
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error en el procesado",
                    f"Se produjo un error:\n{err}"))
            finally:
                self.after(0, lambda: self._habilitar_btn_procesar(True))
                self.after(0, lambda: self._btn_buscar.config(state="normal"))

        threading.Thread(target=proceso, daemon=True).start()

    def _ventana_sap_sin_desc(self, ssd_list):
        """
        Ventana modal personalizada para avisar de artículos con CÓDIGO SAP
        pero sin DESCRIPCIÓN ni UNIDAD en el LISTADO DALI-SAP.
        Incluye enlace clicable a SAP S/4HANA Cloud.
        """
        import webbrowser

        URL_SAP   = "https://my405727.s4hana.cloud.sap/ui#Shell-home"
        HDR_BG    = "#1E1E2E"
        BODY_BG   = "#F7F7FB"
        WARN_BG   = "#FFF8E7"
        WARN_BDR  = "#E8C84B"
        WARN_FG   = "#7A5800"
        LINK_FG   = "#1565C0"
        LINK_HOV  = "#0D47A1"
        BTN_BG    = "#C0392B"
        BTN_FG    = "#FFFFFF"
        TEXT_FG   = "#1A1A2E"
        MUTED_FG  = "#555577"

        top = tk.Toplevel(self)
        top.title("Acción 4 — Productos SAP sin descripción/unidad")
        top.configure(bg=HDR_BG)
        top.resizable(False, False)
        top.grab_set()

        # ── Cabecera ──────────────────────────────────────────────────
        frm_hdr = tk.Frame(top, bg=HDR_BG, padx=20, pady=14)
        frm_hdr.pack(fill="x")

        tk.Label(frm_hdr, text="⚠", bg=HDR_BG, fg="#E8C84B",
                 font=("Segoe UI", 22)).pack(side="left", padx=(0, 12))

        frm_titles = tk.Frame(frm_hdr, bg=HDR_BG)
        frm_titles.pack(side="left")
        tk.Label(frm_titles,
                 text="Productos SAP sin descripción ni unidad",
                 bg=HDR_BG, fg="#FFFFFF",
                 font=("Segoe UI", 12, "bold")).pack(anchor="w")
        tk.Label(frm_titles,
                 text="El listado de Productos SAP en RECURSOS está incompleto o desactualizado",
                 bg=HDR_BG, fg="#AAAACC",
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))

        # ── Cuerpo ────────────────────────────────────────────────────
        frm_body = tk.Frame(top, bg=BODY_BG, padx=20, pady=16)
        frm_body.pack(fill="both", expand=True)

        # Lista de artículos afectados
        _n = len(ssd_list)
        tk.Label(frm_body,
                 text=f"Se encontraron {_n} artículo(s) con CÓDIGO SAP asignado\n"
                      "pero sin DESCRIPCIÓN ni UNIDAD SAP en el catálogo:",
                 bg=BODY_BG, fg=TEXT_FG,
                 font=("Segoe UI", 9), justify="left", anchor="w").pack(anchor="w")

        frm_lista = tk.Frame(frm_body, bg="#EEEEF5", padx=10, pady=8,
                             highlightbackground="#CCCCDD", highlightthickness=1)
        frm_lista.pack(fill="x", pady=(6, 12))
        for _, cd, cs in ssd_list[:10]:
            tk.Label(frm_lista,
                     text=f"  •  DALI {str(cd).zfill(6)}   →   SAP {cs}",
                     bg="#EEEEF5", fg=TEXT_FG,
                     font=("Courier New", 9), anchor="w").pack(anchor="w")
        if _n > 10:
            tk.Label(frm_lista,
                     text=f"  ... y {_n - 10} artículo(s) más",
                     bg="#EEEEF5", fg=MUTED_FG,
                     font=("Segoe UI", 8, "italic"), anchor="w").pack(anchor="w")

        # ── Panel de descarga ─────────────────────────────────────────
        frm_warn = tk.Frame(frm_body, bg=WARN_BG, padx=14, pady=10,
                            highlightbackground=WARN_BDR, highlightthickness=1)
        frm_warn.pack(fill="x", pady=(0, 10))

        tk.Label(frm_warn,
                 text="📥  Descargue el listado actualizado de Productos desde SAP:",
                 bg=WARN_BG, fg=WARN_FG,
                 font=("Segoe UI", 9, "bold"), anchor="w").pack(anchor="w")

        # Enlace clicable
        lbl_link = tk.Label(frm_warn,
                            text=URL_SAP,
                            bg=WARN_BG, fg=LINK_FG,
                            font=("Segoe UI", 9, "underline"),
                            cursor="hand2", anchor="w")
        lbl_link.pack(anchor="w", pady=(4, 6))

        def _abrir_url(e=None):
            webbrowser.open(URL_SAP)

        lbl_link.bind("<Button-1>", _abrir_url)
        lbl_link.bind("<Enter>", lambda e: lbl_link.config(fg=LINK_HOV))
        lbl_link.bind("<Leave>", lambda e: lbl_link.config(fg=LINK_FG))

        tk.Label(frm_warn,
                 text="Guarde el archivo descargado en la carpeta  RECURSOS/  del proyecto.",
                 bg=WARN_BG, fg=WARN_FG,
                 font=("Segoe UI", 8), anchor="w").pack(anchor="w")

        # ── Pasos a seguir ────────────────────────────────────────────
        tk.Frame(frm_body, bg="#DDDDEE", height=1).pack(fill="x", pady=(4, 10))

        tk.Label(frm_body,
                 text="⚙️  Si ya ha descargado y guardado el archivo en RECURSOS/:",
                 bg=BODY_BG, fg=TEXT_FG,
                 font=("Segoe UI", 9, "bold"), anchor="w").pack(anchor="w")

        for paso, txt in [
            ("1.", "Ejecute  Acción 3  — Actualizar listado materiales SAP"),
            ("2.", "Ejecute  Acción 4  — Procesar y generar TARIFA"),
        ]:
            frm_paso = tk.Frame(frm_body, bg=BODY_BG)
            frm_paso.pack(anchor="w", padx=(10, 0), pady=1)
            tk.Label(frm_paso, text=paso, bg=BODY_BG, fg=LINK_FG,
                     font=("Segoe UI", 9, "bold"), width=3).pack(side="left")
            tk.Label(frm_paso, text=txt, bg=BODY_BG, fg=TEXT_FG,
                     font=("Segoe UI", 9)).pack(side="left")

        # ── Botón cerrar ──────────────────────────────────────────────
        frm_btn = tk.Frame(top, bg=HDR_BG, padx=16, pady=10)
        frm_btn.pack(fill="x", side="bottom")

        tk.Button(frm_btn, text="  Entendido  ",
                  bg=BTN_BG, fg=BTN_FG,
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=20, pady=6,
                  cursor="hand2",
                  command=top.destroy).pack(side="right")

        # ── Centrar en pantalla ───────────────────────────────────────
        top.update_idletasks()
        _w = 560
        _h = top.winfo_reqheight()
        _sw = top.winfo_screenwidth()
        _sh = top.winfo_screenheight()
        _px = (_sw - _w) // 2
        _py = max(60, (_sh - _h) // 3)
        top.geometry(f"{_w}x{_h}+{_px}+{_py}")

        top.bind("<Return>", lambda e: top.destroy())
        top.bind("<Escape>", lambda e: top.destroy())
        self.wait_window(top)

    def _leer_tabla_cruce_listado_sap(self):
        """
        PASOS 1+2 (extraídos de _ejecutar_reinscripcion_sap_desde_inc,
        Acción 4): localiza LISTADO CODIGOS DALI-SAP (raíz o RECURSOS)
        y construye la tabla de cruce DALI → (SAP, desc_sap, unidad_sap).

        Devuelve (ruta_listado, tabla_listado) si todo va bien, o
        (None, None) si el LISTADO no se encuentra o no se puede leer
        -- en ambos casos ya deja el aviso hecho (log + barra.reset +
        messagebox.showerror); el llamador solo necesita comprobar
        `if ruta_listado is None: return` sin repetir esa lógica.
        """
        from app.comun import catalogo_local as _cl

        # ── 1. Localizar el catálogo (v10.2.50: ya no hay LISTADO CODIGOS DALI-SAP) ──
        _ruta_listado = _cl.localizar_catalogo(self._carpeta_proyecto.get().strip() or None)
        if not _ruta_listado:
            self.log_consola("catalogo-dali-completo.xlsx no encontrado.", "ERROR")
            self.after(0, lambda: self._barra.reset("Catálogo no encontrado"))
            self.after(0, lambda: messagebox.showerror(
                "Catálogo no encontrado",
                "❌  No se encontró el archivo\n"
                "   RECURSOS/catalogo-dali-completo.xlsx\n\n"
                "Ejecute la Acción 2 (o descárguelo de Catálogo Asignaciones)\n"
                "y vuelva a intentarlo."))
            return None, None

        self.log_consola(f"  └ Catálogo: {os.path.basename(_ruta_listado)}", "OK")

        # ── 2. Leer tabla DALI → (SAP, desc_sap, unidad_sap) ─────
        self.log_consola("Construyendo tabla de cruce desde el catálogo...", "INFO")
        _tabla_listado = {}   # {cod_dali_int: (cod_sap, desc_sap, unidad_sap)}
        try:
            _n_lst = 0
            for _row in _cl.filas_formato_listado(_ruta_listado):   # misma forma que el antiguo LISTADO
                _v_a = _row[0] if _row else None   # col A: DALI
                _v_d = _row[3] if len(_row) > 3 else None  # col D: SAP
                _v_e = _row[4] if len(_row) > 4 else None  # col E: DESC SAP
                _v_f = _row[6] if len(_row) > 6 else None  # col G: UNIDAD SAP
                if _v_a is None:
                    continue
                try:
                    _k = int(float(str(_v_a).strip()))
                except (ValueError, TypeError):
                    continue
                _cod_sap  = None
                if _v_d is not None and str(_v_d).strip() not in ("", "0", "None"):
                    try:
                        _cod_sap = int(float(str(_v_d).strip()))
                    except (ValueError, TypeError):
                        _cod_sap = str(_v_d).strip()
                _desc_sap  = str(_v_e).strip() if _v_e is not None else ""
                _unid_sap  = str(_v_f).strip() if _v_f is not None else ""
                _tabla_listado[_k] = (_cod_sap, _desc_sap, _unid_sap)
                _n_lst += 1
            self.log_consola(f"  └ {_n_lst} entradas leídas del catálogo.", "OK")
        except Exception as _e_lst:
            self.log_consola(f"Error leyendo el catálogo: {_e_lst}", "ERROR")
            self.after(0, lambda: self._barra.reset("Error leyendo el catálogo"))
            self.after(0, lambda err=str(_e_lst): messagebox.showerror(
                "Error al leer el catálogo", err))
            return None, None

        return _ruta_listado, _tabla_listado


    def _log_resumen_reinscripcion(self, _actualizados, _sin_sap,
                                   _sin_en_lista, _sap_sin_desc,
                                   _dup_sap, _dup_dali):
        """
        PASO 6 (extraído de _ejecutar_reinscripcion_sap_desde_inc,
        Acción 4): vuelca en el log el resumen de la reinscripción SAP
        (aplicados, sin SAP, no encontrados en el LISTADO, SAP sin
        descripción/unidad, duplicados SAP/DALI), marca la barra como
        completada y habilita el botón de ver resultado. Los mismos
        parámetros que antes eran variables sueltas del bucle -- se
        mantienen los nombres tal cual para no arriesgar una fuga al
        extraer (ver HISTORIAL, filas_prod en el punto 2 del roadmap).
        Efecto de solo lectura sobre esos datos; no devuelve nada.
        """
        # ── 6. Log resumen ─────────────────────────────────────────
        self.log_consola(
            f"  └ {_actualizados} SAP aplicados  |  "
            f"{len(_sin_sap)} DALI sin SAP  |  "
            f"{len(_sin_en_lista)} DALI no en listado  |  "
            f"{len(_sap_sin_desc)} SAP sin desc/unidad", "INFO")
        if _dup_sap:
            self.log_consola(
                f"  ⚠  {len(_dup_sap)} SAP duplicados: "
                + ", ".join(str(d) for d in _dup_sap[:10])
                + (" ..." if len(_dup_sap) > 10 else ""), "WARN")
        if _dup_dali:
            self.log_consola(
                f"  ⚠  {len(_dup_dali)} DALI duplicados: "
                + ", ".join(str(d) for d in _dup_dali[:10])
                + (" ..." if len(_dup_dali) > 10 else ""), "WARN")
        if _sin_en_lista:
            self.log_consola(
                f"  ⚠  {len(_sin_en_lista)} DALI no encontrados en el catálogo: "
                + ", ".join(str(d).zfill(6) for d in _sin_en_lista[:10])
                + (" ..." if len(_sin_en_lista) > 10 else ""), "WARN")

        self.after(0, lambda: self._barra.indeterminate("Finalizando..."))
        self.after(0, lambda: self._habilitar_btn_resultado(True))
        self.consola.separador("FIN REINSCRIPCIÓN SAP")

        if _sap_sin_desc:
            self.log_consola(
                f"  ⚠  {len(_sap_sin_desc)} SAP con código pero sin descripción/unidad: "
                + ", ".join(f"SAP {str(s).zfill(6)}" for _, _, s in _sap_sin_desc[:10])
                + (" ..." if len(_sap_sin_desc) > 10 else ""), "WARN")


    def _reaplicar_sap_y_guardar_inc(self, ruta_inc, _tabla_listado,
                                     _barra_finalizada):
        """
        PASOS 3+4+5 (extraídos de _ejecutar_reinscripcion_sap_desde_inc,
        Acción 4): abre el archivo INC, reaplica el código SAP y la
        descripción concatenada fila a fila según `_tabla_listado`
        (construida por _leer_tabla_cruce_listado_sap(), PASOS 1+2),
        reevalúa duplicados SAP/DALI sobre el resultado, y guarda el
        archivo -- exactamente lo mismo que hacía el bloque inline
        original, solo que ahora vive en un método propio.

        `_barra_finalizada` es el mismo dict mutable de proceso() en
        _ejecutar_reinscripcion_sap_desde_inc (ver _barra_finalizada["v"]
        en ese método) -- se pasa por referencia para que los dos
        caminos de error de aquí (abrir/guardar) lo marquen igual que
        antes, sin duplicar esa lógica de seguridad en dos sitios.

        Devuelve (_actualizados, _sin_sap, _sin_en_lista, _sap_sin_desc,
        _dup_sap, _dup_dali) si todo va bien, o None si no se pudo
        abrir o guardar el archivo -- en ambos casos ya deja el aviso
        hecho (log + barra.reset + messagebox.showerror); el llamador
        solo necesita comprobar `if resultado is None: return`.
        """
        from openpyxl import load_workbook as _lw
        from openpyxl.styles import PatternFill

        # ── 3. Abrir INC en escritura y re-aplicar SAP + DESC ────
        self.log_consola("Aplicando SAP y descripciones al archivo INC...", "INFO")
        self.after(0, lambda: self._barra.indeterminate("Reescribiendo SAP y DESC..."))

        try:
            wb_inc = _lw(ruta_inc, rich_text=False)
            ws_inc = wb_inc.active
            # Desproteger por si acaso
            ws_inc.protection.sheet   = False
            ws_inc.protection.enabled = False
        except Exception as _e_open:
            self.log_consola(f"Error abriendo INC: {_e_open}", "ERROR")
            self.after(0, lambda: self._barra.reset("Error abriendo INC"))
            _barra_finalizada["v"] = True
            self.after(0, lambda err=str(_e_open): messagebox.showerror(
                "Error al abrir INC", err))
            return None

        _actualizados  = 0   # filas donde SAP se escribió (tenía valor en LISTADO)
        _sin_sap       = []  # [(fila, cod_dali)] sin SAP en LISTADO
        _sin_en_lista  = []  # códigos DALI del INC no encontrados en el LISTADO
        _sap_sin_desc  = []  # [(fila, cod_dali, cod_sap)] con SAP pero sin desc+unidad

        # Estilos para marcar faltantes en amarillo (igual que transformar_tarifa)
        _AMARILLO = PatternFill(fill_type="solid", fgColor="FFFF00")
        _SIN_FILL = PatternFill(fill_type=None)

        for _row in ws_inc.iter_rows(min_row=8):
            if not _row or len(_row) < 2:
                continue
            _cell_a = _row[0]   # col A: CODIGO SAP
            _cell_b = _row[1]   # col B: CODIGO DALI
            _cell_d = _row[3] if len(_row) > 3 else None  # col D: DESCRIPCION

            _v_b = _cell_b.value if _cell_b else None
            if _v_b is None:
                continue
            try:
                _cod_dali = int(float(str(_v_b).strip()))
            except (ValueError, TypeError):
                continue

            _fila = _cell_a.row

            if _cod_dali not in _tabla_listado:
                _sin_en_lista.append(_cod_dali)
                # Borrar SAP por si tenía un valor anterior incorrecto
                _cell_a.value = None
                continue

            _cod_sap, _desc_sap, _unid_sap = _tabla_listado[_cod_dali]

            # ── Sobreescribir col A (SAP) ─────────────────────────
            _cell_a.value = _cod_sap  # None si no tiene SAP en el listado

            # ── Sobreescribir col D (DESCRIPCION concatenada) ─────
            if _cell_d is not None:
                if _desc_sap or _unid_sap:
                    if _desc_sap and _unid_sap:
                        _nueva_desc = _desc_sap + " - " + _unid_sap
                    else:
                        _nueva_desc = (_desc_sap or _unid_sap).strip()
                else:
                    # Sin datos SAP: mantener descripción DALI original
                    _nueva_desc = None
                if _nueva_desc is not None:
                    _cell_d.value = _nueva_desc

            # ── Marcar faltantes en amarillo en col A ─────────────
            if _cod_sap is None:
                _sin_sap.append((_fila, _cod_dali))
                _cell_a.fill = _AMARILLO
            else:
                _actualizados += 1
                _cell_a.fill  = _SIN_FILL  # quitar amarillo previo
                # ── Detectar SAP presente pero sin desc ni unidad ─
                if not _desc_sap and not _unid_sap:
                    _sap_sin_desc.append((_fila, _cod_dali, _cod_sap))

        # ── 4. Re-evaluar duplicados SAP ──────────────────────────
        _vistos_sap  = {}
        _dup_sap     = []
        _vistos_dali = {}
        _dup_dali    = []
        for _row2 in ws_inc.iter_rows(min_row=8, values_only=False):
            if not _row2 or len(_row2) < 2:
                continue
            _va = _row2[0].value if _row2[0] else None
            _vb = _row2[1].value if len(_row2) > 1 else None
            if _vb is None:
                continue
            try:
                _cd = int(float(str(_vb).strip()))
            except (ValueError, TypeError):
                continue
            if _va is not None and str(_va).strip() not in ("", "0", "None"):
                try:
                    _cs = int(float(str(_va).strip()))
                    if _cs in _vistos_sap:
                        if _cs not in _dup_sap:
                            _dup_sap.append(_cs)
                    else:
                        _vistos_sap[_cs] = _row2[0].row
                except (ValueError, TypeError):
                    pass
            if _cd in _vistos_dali:
                if _cd not in _dup_dali:
                    _dup_dali.append(_cd)
            else:
                _vistos_dali[_cd] = _row2[0].row if _row2[0] else 0

        # ── 5. Guardar INC actualizado ─────────────────────────────
        try:
            wb_inc.save(ruta_inc)
            wb_inc.close()
        except Exception as _e_save:
            self.log_consola(f"Error guardando INC: {_e_save}", "ERROR")
            self.after(0, lambda: self._barra.reset("Error al guardar"))
            _barra_finalizada["v"] = True
            self.after(0, lambda err=str(_e_save): messagebox.showerror(
                "Error al guardar INC", err))
            return None


        return (_actualizados, _sin_sap, _sin_en_lista, _sap_sin_desc,
                _dup_sap, _dup_dali)


    def _gestionar_faltantes_o_finalizar_inc(self, ruta_inc, _ruta_listado,
                                             _actualizados, _sin_sap,
                                             _sap_sin_desc, _dup_sap,
                                             _dup_dali, _barra_finalizada):
        """
        PASO 7 (extraído de _ejecutar_reinscripcion_sap_desde_inc,
        Acción 4, el más delicado de trocear por las closures anidadas
        que dependen de variables de esta misma llamada -- por eso se
        extrajo el bloque COMPLETO tal cual estaba, en vez de partirlo
        en piezas más pequeñas, para no arriesgar una fuga de alguna
        de esas capturas): si quedan códigos DALI sin SAP asignado,
        abre (con self.after, no bloqueante) el diálogo
        VentanaSAPFaltantes para introducirlos o abre el Excel para
        edición manual; si no hay faltantes, comprueba duplicados y
        descripciones vacías, renombra el archivo quitando
        [INCOMPLETO]/"INC " cuando todo está limpio, y engancha la
        actualización del catálogo DALI completo
        (_actualizar_codigo_proveedor_catalogo()) en ese caso.

        `_barra_finalizada` es el mismo dict mutable de proceso() en
        _ejecutar_reinscripcion_sap_desde_inc -- se pasa por
        referencia para que los distintos caminos de este método
        (con/sin faltantes, con/sin duplicados, error al renombrar)
        lo marquen igual que antes; la red de seguridad del finally
        exterior sigue cubriendo cualquier camino no contemplado aquí
        (p.ej. diálogo cancelado, "abrir en Excel" sin guardar).

        No devuelve nada -- efecto por log/mensajes/renombrado/barra,
        igual que el bloque original.
        """
        from openpyxl import load_workbook as _lw

        # ── 7. Si aún hay faltantes → abrir VentanaSAPFaltantes ──
        _hay_dups = bool(_dup_sap or _dup_dali)

        if _sin_sap:
            self.log_consola(
                f"  ⚠  {len(_sin_sap)} código(s) DALI sin SAP asignado en el catálogo.",
                "WARN")

            # Leer descripciones para el diálogo
            _descripciones = {}
            try:
                wb_d2 = _lw(ruta_inc, read_only=True, data_only=True)
                ws_d2 = wb_d2.active
                for _frow, _dcod in _sin_sap:
                    _descripciones[_dcod] = (
                        str(ws_d2.cell(_frow, 4).value or "").strip(),
                        str(ws_d2.cell(_frow, 5).value or "").strip(),
                    )
                wb_d2.close()
            except Exception:
                pass

            def _mostrar_dialogo_inc(
                    falt=_sin_sap, rl=_ruta_listado,
                    descs=_descripciones, ri=ruta_inc,
                    dups=_hay_dups, n_act=_actualizados):

                _msg = (
                    f"✔  {n_act} código(s) SAP aplicados correctamente.\n\n"
                    f"⚠  {len(falt)} código(s) DALI aún sin SAP en el catálogo:\n\n"
                    + "\n".join(
                        f"  •  DALI {str(d).zfill(6)}  "
                        f"{descs.get(d, ('',''))[0][:40]}"
                        for _, d in falt[:10])
                    + (f"\n  ... y {len(falt)-10} más" if len(falt) > 10 else "")
                    + "\n\n¿Cómo deseas introducir los códigos SAP faltantes?"
                )
                resp = messagebox.askyesnocancel(
                    "Códigos SAP faltantes",
                    _msg + "\n\n"
                    "[SÍ]       →  Introducir ahora desde la aplicación\n"
                    "[NO]       →  Modificar directamente en el Excel\n"
                    "[CANCELAR] →  Ignorar por ahora",
                    icon="warning")

                if resp is True:
                    ven = VentanaSAPFaltantes(
                        self, falt, rl,
                        consola_fn=self.log_consola,
                        descripciones=descs,
                        ruta_tarifa=ri,
                        catalogo_refresh_fn=self._tras_cambiar_sap_en_catalogo)
                    self.wait_window(ven)
                    if ven.guardado:   # (13 nov 2026) ¿faltan los listados de SAP actualizados?
                        self._avisar_listados_sap(list(ven.nuevos_sap.values()))

                    if ven.guardado and ven.escritos_en_tarifa >= 1:
                        _todos = (ven.escritos_en_tarifa == len(falt))
                        try:
                            import re as _re2
                            _dir  = os.path.dirname(ri)
                            _base = os.path.basename(ri)
                            if _todos and not dups:
                                # Completado: eliminar [INCOMPLETO] Y prefijo INC
                                _nuevo = _re2.sub(
                                    r'\[INCOMPLETO\]\s*', '', _base,
                                    flags=_re2.IGNORECASE).strip()
                                _nuevo = _re2.sub(
                                    r'^INC\s+', '', _nuevo,
                                    flags=_re2.IGNORECASE).strip()
                            else:
                                _nuevo = _base
                                if not _nuevo.upper().startswith("INC "):
                                    _nuevo = "INC " + _nuevo
                            _ruta_nueva = os.path.join(_dir, _nuevo)
                            if _nuevo != _base:
                                if os.path.exists(_ruta_nueva):
                                    os.remove(_ruta_nueva)
                                    self.log_consola(
                                        f"  🗑 Sustituido archivo anterior: {_nuevo}", "INFO")
                                os.rename(ri, _ruta_nueva)
                                self._ruta_resultado = _ruta_nueva
                                self.log_consola(
                                    f"  ✔ Archivo renombrado → {_nuevo}", "OK")
                            if _todos and not dups:
                                if self._catalogo_reciente_o_confirmar():
                                    self._actualizar_codigo_proveedor_catalogo(_ruta_nueva)
                                self.after(0, lambda: self._barra.completado(
                                    "Reinscripción SAP completada"))
                                _barra_finalizada["v"] = True
                                messagebox.showinfo(
                                    "Tarifa completada",
                                    f"✔  Todos los códigos SAP completados "
                                    f"y sin duplicados.\n\n"
                                    f"Archivo renombrado a:\n{_nuevo}\n\n"
                                    f"\n✔  TARIFA PREPARADA PARA CREAR O MODIFICAR REGISTRO INFO SAP.")
                            else:
                                # Aún quedan duplicados o códigos sin completar:
                                # el catálogo NO se toca todavía (no está lista),
                                # pero la barra debe cerrar aquí -- si no, se queda
                                # colgada en "Finalizando..." indefinidamente.
                                self.after(0, lambda: self._barra.reset(
                                    "Revisar duplicados/pendientes"))
                                _barra_finalizada["v"] = True
                        except Exception as _e_ren:
                            self.log_consola(
                                f"  ✘ Error al renombrar: {_e_ren}", "WARN")
                            self.after(0, lambda: self._barra.reset("Error al renombrar"))
                            _barra_finalizada["v"] = True

                elif resp is False:
                    try:
                        os.startfile(ri)
                    except Exception:
                        import subprocess
                        subprocess.Popen(["start", "", ri], shell=True)

            self.after(200, _mostrar_dialogo_inc)

            # ── Avisar sobre SAP con código pero sin desc/unidad ─
            if _sap_sin_desc:
                def _avisar_sap_sin_desc(
                        ssd=_sap_sin_desc, rl=_ruta_listado):
                    self._ventana_sap_sin_desc(ssd)
                self.after(500, _avisar_sap_sin_desc)

        else:
            # Sin faltantes SAP — comprobar duplicados, desc vacías y renombrar si procede
            _hay_inc = bool(_dup_sap or _dup_dali or _sap_sin_desc)
            _titulo  = "Completado con advertencias ⚠" if _hay_inc else "Reinscripción completada ✔"
            _msg_fin = (
                f"✔  {_actualizados} código(s) SAP aplicados.\n"
                f"✔  Descripciones SAP actualizadas desde el catálogo.\n"
            )
            if _dup_sap:
                _msg_fin += (
                    f"\n⚠  {len(_dup_sap)} SAP duplicado(s): "
                    + ", ".join(str(d) for d in _dup_sap[:10]) + "\n"
                )
            if _dup_dali:
                _msg_fin += (
                    f"\n⚠  {len(_dup_dali)} DALI duplicado(s): "
                    + ", ".join(str(d) for d in _dup_dali[:10]) + "\n"
                )
            if _sap_sin_desc:
                _msg_fin += (
                    f"\n⚠  {len(_sap_sin_desc)} SAP con código pero sin descripción/unidad:\n"
                    + "\n".join(
                        f"   DALI {str(d).zfill(6)}  →  SAP {str(s).zfill(8)}"
                        for _, d, s in _sap_sin_desc[:10])
                    + (" ..." if len(_sap_sin_desc) > 10 else "")
                    + "\n\nLa columna D no ha podido actualizarse para estos artículos.\n"
                    "Ejecute la Acción 3 para completar la descripción y la unidad SAP en el catálogo y de nuevo la Acción 4.\n"
                )
            if not _hay_inc:
                # Sin faltantes, sin duplicados y sin desc vacías → renombrar eliminando INC y [INCOMPLETO]
                try:
                    import re as _re_sf
                    _dir_sf  = os.path.dirname(ruta_inc)
                    _base_sf = os.path.basename(ruta_inc)
                    _nuevo_sf = _re_sf.sub(
                        r'\[INCOMPLETO\]\s*', '', _base_sf,
                        flags=_re_sf.IGNORECASE).strip()
                    _nuevo_sf = _re_sf.sub(
                        r'^INC\s+', '', _nuevo_sf,
                        flags=_re_sf.IGNORECASE).strip()
                    _ruta_sf = os.path.join(_dir_sf, _nuevo_sf)
                    if _nuevo_sf != _base_sf:
                        if os.path.exists(_ruta_sf):
                            os.remove(_ruta_sf)
                            self.log_consola(
                                f"  🗑 Sustituido archivo final anterior: {_nuevo_sf}",
                                "INFO")
                        os.rename(ruta_inc, _ruta_sf)
                        self._ruta_resultado = _ruta_sf
                        self.log_consola(
                            f"  ✔ Archivo renombrado → {_nuevo_sf}", "OK")
                        _msg_fin += (
                            f"\n✔  Archivo renombrado a:\n   {_nuevo_sf}\n\n"
                            "\n✔  TARIFA PREPARADA PARA CREAR O MODIFICAR REGISTRO INFO SAP."
                        )
                        if self._catalogo_reciente_o_confirmar():
                            self._actualizar_codigo_proveedor_catalogo(_ruta_sf)
                    else:
                        _msg_fin += "\n✔  TARIFA PREPARADA PARA CREAR O MODIFICAR REGISTRO INFO SAP."
                        if self._catalogo_reciente_o_confirmar():
                            self._actualizar_codigo_proveedor_catalogo(ruta_inc)
                except Exception as _e_ren_sf:
                    self.log_consola(
                        f"  ✘ Error al renombrar: {_e_ren_sf}", "WARN")
                    _msg_fin += "\n✔  TARIFA PREPARADA PARA CREAR O MODIFICAR REGISTRO INFO SAP."
                self.after(0, lambda: self._barra.completado(
                    "Reinscripción SAP completada"))
                _barra_finalizada["v"] = True
            else:
                # Duplicados o SAP sin desc/unidad detectados desde el
                # principio (sin necesidad de abrir VentanaSAPFaltantes):
                # el archivo se queda [INCOMPLETO], pero la barra debe
                # cerrar aquí igualmente -- si no, se queda colgada en
                # "Finalizando..." indefinidamente.
                self.after(0, lambda: self._barra.reset(
                    "Revisar duplicados/pendientes"))
                _barra_finalizada["v"] = True
            _fn_msg = messagebox.showwarning if _hay_inc else messagebox.showinfo
            self.after(0, lambda t=_titulo, m=_msg_fin, fn=_fn_msg: fn(t, m))

            # ── Avisar sobre SAP con código pero sin desc/unidad ─
            if _sap_sin_desc:
                def _avisar_sap_sin_desc_else(ssd=_sap_sin_desc):
                    self._ventana_sap_sin_desc(ssd)
                self.after(300, _avisar_sap_sin_desc_else)


    def _ejecutar_reinscripcion_sap_desde_inc(self, ruta_inc):
        """
        Acción 4 — Opción INC (sobreescritura completa).

        Re-aplica desde el LISTADO CODIGOS DALI-SAP todos los códigos SAP
        (col A del xlsx tarifa) y la columna de descripción concatenada
        (col D = DESCRIPCION SAP + " " + UNIDAD SAP del listado) como si
        fuera la primera vez que se procesa el archivo.

        Esto garantiza que:
          • Cualquier código SAP que el usuario haya actualizado en el
            LISTADO quede reflejado en el INC.
          • Las descripciones SAP actualizadas en el LISTADO sustituyan
            a las anteriores en el INC.
          • Se detecten nuevos faltantes (filas donde el LISTADO no tiene
            SAP asignado para el DALI correspondiente).

        Estructura del xlsx INC (generado por transformar_tarifa):
          Col A → CODIGO SAP   (se sobreescribe desde col D del LISTADO)
          Col B → CODIGO DALI  (clave de cruce)
          Col C → COD PROVEEDOR
          Col D → DESCRIPCION  (se sobreescribe: desc_sap + " " + unidad_sap)
          Col E → FORMATO
          Col F → PRECIO
          Col G → DESCUENTO
          Datos desde fila 8; filas 1-7 = cabecera.

        Estructura del LISTADO CODIGOS DALI-SAP:
          Col A → CODIGO DALI  (clave de cruce)
          Col D → CODIGO SAP
          Col E → DESCRIPCION SAP
          Col G → UNIDAD SAP
          Datos desde fila 8.
        (mapeo actualizado 2026-08: tras insertar las 3 columnas nuevas
        F, K, L en LISTADO CODIGOS DALI-SAP, la antigua col F (UNIDAD SAP)
        se desplazó a la col G; este método se quedó leyendo la F antigua
        -- que ahora es "Texto UM (expl.)", no la unidad -- corregido
        2026-09 al auditar el punto 4 del roadmap)

        NOTA (2026-09-21): esta acción NO llama a _confirmar_catalogo_
        actualizado() -- a petición del usuario, Acción 4 nunca debe
        ofrecer descargar/actualizar el catálogo, porque es un proceso
        de CARGA de datos (escribe en catalogo-dali-completo.xlsx el
        código proveedor de la tarifa concreta que se está procesando);
        descargar una copia fresca aquí borraría esa actualización que
        Acciones 2/3 acaban de dejar hecha en la misma sesión de trabajo.
        """
        self._habilitar_btn_procesar(False)
        self._btn_buscar.config(state="disabled")
        self._habilitar_btn_resultado(False)

        def proceso():
            _barra_finalizada = {"v": False}   # red de seguridad, ver finally: más abajo
            try:
                self.consola.separador("REINSCRIPCIÓN SAP DESDE EL CATÁLOGO (archivo INC)")
                self.log_consola(f"Archivo INC: {os.path.basename(ruta_inc)}", "INFO")
                self.after(0, lambda: self._barra.indeterminate("Leyendo el catálogo..."))

                _ruta_listado, _tabla_listado = self._leer_tabla_cruce_listado_sap()
                if _ruta_listado is None:
                    return


                _resultado_reaplicar = self._reaplicar_sap_y_guardar_inc(
                    ruta_inc, _tabla_listado, _barra_finalizada)
                if _resultado_reaplicar is None:
                    return
                (_actualizados, _sin_sap, _sin_en_lista, _sap_sin_desc,
                 _dup_sap, _dup_dali) = _resultado_reaplicar


                self._log_resumen_reinscripcion(
                    _actualizados, _sin_sap, _sin_en_lista, _sap_sin_desc,
                    _dup_sap, _dup_dali)


                self._gestionar_faltantes_o_finalizar_inc(
                    ruta_inc, _ruta_listado, _actualizados, _sin_sap,
                    _sap_sin_desc, _dup_sap, _dup_dali, _barra_finalizada)


            except Exception as e:
                self.log_consola(f"ERROR en reinscripción SAP: {e}", "ERROR")
                self.after(0, lambda: self._barra.reset("Error en reinscripción SAP"))
                _barra_finalizada["v"] = True
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error", f"Se produjo un error:\n{err}"))
            finally:
                if not _barra_finalizada["v"]:
                    # Red de seguridad: ningún punto de salida anterior marcó
                    # un estado final de la barra (p.ej. diálogo de faltantes
                    # cancelado, o "abrir en Excel" sin guardar) -- se cierra
                    # aquí para no dejarla colgada indefinidamente en
                    # "Finalizando...".
                    self.after(0, lambda: self._barra.reset("Proceso finalizado"))
                self.after(0, lambda: self._habilitar_btn_procesar(True))
                self.after(0, lambda: self._btn_buscar.config(state="normal"))

        threading.Thread(target=proceso, daemon=True).start()
