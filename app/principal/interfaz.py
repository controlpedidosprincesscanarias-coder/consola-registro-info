# -*- coding: utf-8 -*-
"""
interfaz.py
Construcción de la ventana (panel lateral, consola, botones), consola de log,
estado de los botones y acciones sencillas (carpeta del proyecto, buscar TARIFA.xls, abrir resultado).

Mixin de App (tarifa_app.py): sus métodos se usan como self.metodo(). Separado de tarifa_app.py
en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""

import os
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime
import threading
import time
from app.comun.tarifa_constants import (
    C_DARK, C_MID, C_SIDEBAR, C_CARD, C_BORDER, C_HOVER,
    C_ACCENT, C_ACCENT2, C_ACCENT_DIM,
    C_TEXT_LIGHT, C_MUTED, C_SUCCESS, C_ERROR, C_WARNING,
    C_LOG_BG, SAP_LOGO_PNG,
)
from app.comun.tarifa_core import (
    guardar_config,
    TARIFA_NAME, FOLDER_TARIFAS, FOLDER_RECURSOS,
    buscar_tarifa, listar_xls, ConsolaLog, CanvasProgressBar, abrir_con_su_programa,
)
from app.ui.tarifa_dialogs import (
    VentanaSeleccionXLS, VentanaCrearProyecto,
)
from app.rinfo.crear_registro_info import iniciar_crear_registro_info


class InterfazApp:
    def _centrar(self):
        self.update_idletasks()
        # En Windows 'zoomed' maximiza respetando la barra de tareas (100% útil)
        try:
            self.state("zoomed")
        except Exception:
            # Fallback para Linux/Mac: 100% de la pantalla centrado
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            self.geometry(f"{sw}x{sh}+0+0")

    # ─────────────────────────────────────────────────────────────────────
    #  UI CONSTRUCTION  (rediseño V2 — VAMA 2026)
    # ─────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_header()
        self._build_main()
        self._build_footer()

    def _tick_clock(self):
        now = datetime.now()
        try:
            self._lbl_clock.config(text=now.strftime("%H:%M:%S"))
            self._lbl_clock_date.config(text=now.strftime("%A, %d %b %Y").upper())
        except Exception:
            pass
        self.after(1000, self._tick_clock)

    # ── MAIN (sidebar + consola) ──────────────────────────────────────────
    def _build_main(self):
        main = tk.Frame(self, bg=C_SIDEBAR)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, minsize=295, weight=0)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        # Sidebar
        sidebar = tk.Frame(main, bg=C_SIDEBAR)
        sidebar.grid(row=0, column=0, sticky="nsew")
        tk.Frame(main, bg=C_BORDER, width=1).grid(row=0, column=0, sticky="nse")
        self._build_sidebar(sidebar)

        # Panel derecho (consola + accesos rápidos)
        right = tk.Frame(main, bg=C_LOG_BG)
        right.grid(row=0, column=1, sticky="nsew")
        self._build_console(right)

    # ── SIDEBAR ───────────────────────────────────────────────────────────
    def _build_sidebar(self, sidebar):
        pad = tk.Frame(sidebar, bg=C_SIDEBAR, padx=10, pady=4)
        pad.pack(fill="both", expand=True)

        # ── 00. CARPETA DEL PROYECTO ──────────────────────────────────────
        self._section(pad, "CARPETA DEL PROYECTO")

        frm_ruta = tk.Frame(pad, bg=C_SIDEBAR)
        frm_ruta.pack(fill="x", pady=(2, 0))

        self._ent_proyecto = tk.Entry(
            frm_ruta, textvariable=self._carpeta_proyecto,
            font=("Segoe UI", 8), state="readonly", relief="flat",
            bg=C_CARD, fg="#5A8AB0", readonlybackground=C_CARD,
            insertbackground=C_ACCENT, bd=0)
        self._ent_proyecto.pack(side="left", fill="x", expand=True,
                                 padx=(0, 4), ipady=4)

        tk.Button(frm_ruta, text="📁",
                  command=self._seleccionar_proyecto,
                  font=("Segoe UI", 10), padx=6, pady=2,
                  cursor="hand2", relief="flat",
                  bg=C_ACCENT, fg=C_DARK,
                  activebackground=C_ACCENT2).pack(side="right")

        # Indicadores de subcarpetas como pill-badges
        frm_subs = tk.Frame(pad, bg=C_SIDEBAR)
        frm_subs.pack(fill="x", pady=(5, 0))

        # Pill TARIFAS
        self._frm_sub_tarifas = tk.Frame(
            frm_subs, bg="#0D1F35",
            highlightbackground="#162840", highlightthickness=1)
        self._frm_sub_tarifas.pack(side="left")
        self._lbl_sub_tarifas = tk.Label(
            self._frm_sub_tarifas, text="TARIFAS  —",
            bg="#0D1F35", fg=C_MUTED,
            font=("Segoe UI", 7, "bold"), padx=6, pady=2, cursor="arrow")
        self._lbl_sub_tarifas.pack()

        tk.Label(frm_subs, text="·", bg=C_SIDEBAR,
                  fg=C_BORDER, font=("Segoe UI", 8)).pack(side="left", padx=4)

        # Pill RECURSOS
        self._frm_sub_recursos = tk.Frame(
            frm_subs, bg="#0D1F35",
            highlightbackground="#162840", highlightthickness=1)
        self._frm_sub_recursos.pack(side="left")
        self._lbl_sub_recursos = tk.Label(
            self._frm_sub_recursos, text="RECURSOS  —",
            bg="#0D1F35", fg=C_MUTED,
            font=("Segoe UI", 7, "bold"), padx=6, pady=2, cursor="arrow")
        self._lbl_sub_recursos.pack()

        # Link RAÍZ
        self._lbl_abrir_raiz = tk.Label(
            frm_subs, text="↗ RAÍZ",
            bg=C_SIDEBAR, fg=C_MUTED,
            font=("Segoe UI", 7), cursor="arrow")
        self._lbl_abrir_raiz.pack(side="right")

        self._ruta_sub_tarifas  = None
        self._ruta_sub_recursos = None

        # Botón crear estructura
        tk.Button(pad, text="+  Crear estructura PROYECTO",
                  command=self._crear_proyecto,
                  font=("Segoe UI", 8), padx=6, pady=3,
                  cursor="hand2", relief="flat",
                  bg=C_MID, fg=C_ACCENT,
                  activebackground=C_HOVER,
                  activeforeground=C_ACCENT2).pack(fill="x", pady=(5, 0))

        # ── 01. ESTADO DE TARIFA.XLS ──────────────────────────────────────
        self._section(pad, "ESTADO DE TARIFA.XLS")

        # Tarjeta de estado unificada
        _STC = "#0D1F35"
        frm_status = tk.Frame(pad, bg=_STC,
                               highlightbackground=C_BORDER, highlightthickness=1)
        frm_status.pack(fill="x", pady=(2, 0))

        # Indicador (●)
        self._lbl_estado_icon = tk.Label(
            frm_status, text="●",
            bg=_STC, fg=C_MUTED,
            font=("Segoe UI", 18),
            padx=8, pady=6)
        self._lbl_estado_icon.pack(side="left")

        # Texto de estado
        frm_est_txt = tk.Frame(frm_status, bg=_STC)
        frm_est_txt.pack(side="left", fill="both", expand=True, pady=6)

        self._lbl_estado = tk.Label(
            frm_est_txt, text="Sin verificar",
            bg=_STC, fg=C_MUTED,
            font=("Segoe UI", 8, "bold"), anchor="w")
        self._lbl_estado.pack(fill="x")

        self._lbl_ruta_resultado = tk.Label(
            frm_est_txt, text="Seleccione carpeta",
            bg=_STC, fg=C_MUTED,
            font=("Segoe UI", 7), anchor="w",
            wraplength=130, justify="left")
        self._lbl_ruta_resultado.pack(fill="x")

        # Separador vertical
        tk.Frame(frm_status, bg=C_BORDER, width=1).pack(
            side="left", fill="y", padx=(8, 0))

        # Botón buscar
        frm_buscar = tk.Frame(frm_status, bg=_STC, padx=8)
        frm_buscar.pack(side="left", fill="both")

        tk.Label(frm_buscar, text="Localizar",
                  bg=_STC, fg=C_MUTED,
                  font=("Segoe UI", 7)).pack(pady=(5, 2))

        self._btn_buscar = tk.Button(
            frm_buscar, text="🔍  Buscar",
            command=self._buscar,
            font=("Segoe UI", 8), padx=4, pady=4,
            cursor="hand2", relief="flat",
            bg=C_MID, fg=C_TEXT_LIGHT,
            activebackground=C_HOVER,
            activeforeground=C_ACCENT2,
            justify="center")
        self._btn_buscar.pack(fill="x", pady=(0, 5))

        # ── 02. PROGRESO ──────────────────────────────────────────────────
        self._section(pad, "PROGRESO")
        self._barra = CanvasProgressBar(pad)
        self._barra.pack(fill="x", pady=(2, 0))

        # ── 03. ACCIONES ──────────────────────────────────────────────────
        self._section(pad, "ACCIONES")
        frm_acc = tk.Frame(pad, bg=C_SIDEBAR)
        frm_acc.pack(fill="x", pady=(3, 0))

        self._boton_accion(
            frm_acc, "1", "Organizar artículos DALI - Generar ARTICULOS_DALI", "📋",
            self._organizar_dali, strip_color="#2E8B57"
        ).pack(fill="x", pady=(0, 2))

        self._boton_accion(
            frm_acc, "2", "Actualizar listado artículos DALI", "🔄",
            self._actualizar_listado_dali, strip_color="#1A9B7C"
        ).pack(fill="x", pady=(0, 2))

        self._boton_accion(
            frm_acc, "3", "Actualizar listado materiales SAP", "🔄",
            self._actualizar_listado_sap, strip_color="#1A7A9B"
        ).pack(fill="x", pady=(0, 2))

        self._btn_procesar = self._boton_accion(
            frm_acc, "4", "Generar TARIFA final y Procesar Tarifas Incompletas", "⚙️",
            self._procesar, strip_color=C_ACCENT, state="normal"
        )
        self._btn_procesar.pack(fill="x", pady=(0, 2))

        self._btn_abrir_resultado = self._boton_accion(
            frm_acc, "5", "Abrir archivo ultima TARIFA procesada", "🗂️",
            self._abrir_resultado, strip_color="#2A6B9B", state="disabled"
        )
        self._btn_abrir_resultado.pack(fill="x", pady=(0, 2))

        self._boton_accion(
            frm_acc, "6", "Actualizar manualmente archivos [INCOMPLETOS]", "💾",
            self._crear_registro_sap, strip_color="#7B4FA6"
        ).pack(fill="x", pady=(0, 2))

        self._boton_accion(
            frm_acc, "7", "Generar Listado Tarifa Proveedor Vacia", "📄",
            self._wip, strip_color="#5A4A7A", wip=True
        ).pack(fill="x", pady=(0, 2))

        # (v10.2.50) Acción 8: todos los códigos DALI activos sin SAP, para asignarlos (accion_asignar_sap.py)
        self._boton_accion(
            frm_acc, "8", "Asignar SAP pendientes (DALI activos sin SAP)", "🧩",
            self._asignar_sap_pendientes, strip_color="#1A6B5A"
        ).pack(fill="x", pady=(0, 0))

    # ── CONSOLA + ACCESOS RÁPIDOS ─────────────────────────────────────────
    def _build_console(self, right):
        # Cabecera de consola
        _HDR_BG = "#040B15"
        frm_con_hdr = tk.Frame(right, bg=_HDR_BG)
        frm_con_hdr.pack(fill="x")
        tk.Frame(frm_con_hdr, bg="#0F2A45", height=1).pack(fill="x")
        frm_con_inner = tk.Frame(frm_con_hdr, bg=_HDR_BG, padx=12, pady=5)
        frm_con_inner.pack(fill="x")

        tk.Label(frm_con_inner, text="CONSOLA  DE  SISTEMA",
                  bg=_HDR_BG, fg=C_MUTED,
                  font=("Segoe UI", 7, "bold")).pack(side="left")

        # (nov 2026) Consultar el manual: escribe una duda y pulsa Intro
        _CON_BG = "#071322"
        frm_consulta = tk.Frame(right, bg=_CON_BG, padx=12, pady=6)
        frm_consulta.pack(fill="x")
        tk.Label(frm_consulta, text="\U0001f50e  Consultar el manual", bg=_CON_BG, fg=C_ACCENT2,
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=(0, 8))
        self._var_consulta = tk.StringVar()
        _ent_consulta = tk.Entry(frm_consulta, textvariable=self._var_consulta, font=("Segoe UI", 9),
                                 bg="#0D2236", fg="#FFFFFF", insertbackground="#FFFFFF", relief="flat")
        _ent_consulta.pack(side="left", fill="x", expand=True, ipady=3)
        _PISTA = "Escribe tu duda y pulsa Intro (p. ej. «códigos SAP faltantes»)"

        def _pista(mostrar):
            if mostrar and not self._var_consulta.get():
                _ent_consulta.configure(fg="#5A7A9A"); self._var_consulta.set(_PISTA)
            elif not mostrar and self._var_consulta.get() == _PISTA:
                _ent_consulta.configure(fg="#FFFFFF"); self._var_consulta.set("")
        _pista(True)
        _ent_consulta.bind("<FocusIn>", lambda e: _pista(False))
        _ent_consulta.bind("<FocusOut>", lambda e: _pista(True))

        def _consultar(e=None):
            texto = self._var_consulta.get().strip()
            self._abrir_consulta_manual("" if texto == _PISTA else texto)
        _ent_consulta.bind("<Return>", _consultar)
        tk.Button(frm_consulta, text="Buscar", command=_consultar, bg="#0A4A7A", fg="#FFFFFF",
                  activebackground="#0D5F9A", activeforeground="#FFFFFF", relief="flat",
                  font=("Segoe UI", 8, "bold"), padx=10, cursor="hand2").pack(side="left", padx=(6, 0))

        # ConsolaLog
        self.consola = ConsolaLog(right)
        self.consola.pack(fill="both", expand=True)

        # ── Accesos rápidos ───────────────────────────────────────────────
        _ACC_BG    = "#04090F"
        _SAP_BG    = "#0A4A7A"
        _SAP_HOV   = "#0D5F9A"
        _WIP_BG    = "#1E1A2E"
        _WIP_FG    = "#5A4A7A"

        frm_accesos = tk.Frame(right, bg=_ACC_BG)
        frm_accesos.pack(fill="x", side="bottom")
        tk.Frame(frm_accesos, bg="#0F2A45", height=1).pack(fill="x")

        frm_btns = tk.Frame(frm_accesos, bg=_ACC_BG, padx=10, pady=7)
        frm_btns.pack(fill="x")
        frm_btns.columnconfigure(0, weight=1)
        frm_btns.columnconfigure(1, weight=1)
        frm_btns.columnconfigure(2, weight=1)

        tk.Label(frm_btns, text="ACCESOS RÁPIDOS",
                  bg=_ACC_BG, fg="#1A3A5A",
                  font=("Segoe UI", 7, "bold"),
                  anchor="w").grid(row=0, column=0, columnspan=3,
                                   sticky="w", padx=2, pady=(0, 5))

        # ── Botón SAP ─────────────────────────────────────────────────────
        _URL_SAP = "https://my405727.s4hana.cloud.sap/ui#Shell-home"

        frm_sap = tk.Frame(frm_btns, bg=_SAP_BG, cursor="hand2")
        frm_sap.grid(row=1, column=0, sticky="nsew", padx=(0, 4))
        tk.Frame(frm_sap, bg="#1A9BDA", width=3).place(relx=0, rely=0, relheight=1)

        frm_sap_inner = tk.Frame(frm_sap, bg=_SAP_BG, padx=10, pady=7)
        frm_sap_inner.pack(fill="both", expand=True)

        try:
            from PIL import Image as _Image, ImageTk as _ImageTk
            _sap_pil = _Image.open(SAP_LOGO_PNG).resize(
                (72, 26), _Image.LANCZOS)
            self._sap_logo_img = _ImageTk.PhotoImage(_sap_pil)
            _sap_icon = tk.Label(frm_sap_inner, image=self._sap_logo_img,
                                  bg=_SAP_BG, cursor="hand2")
        except Exception:
            _sap_icon = tk.Label(frm_sap_inner, text="SAP",
                                  bg=_SAP_BG, fg="#FFFFFF",
                                  font=("Segoe UI", 11, "bold"))
        _sap_icon.pack(side="left", padx=(6, 10))

        frm_sap_txt = tk.Frame(frm_sap_inner, bg=_SAP_BG)
        frm_sap_txt.pack(side="left", fill="x", expand=True)
        tk.Label(frm_sap_txt, text="SAP  S/4HANA Cloud",
                  bg=_SAP_BG, fg="#FFFFFF",
                  font=("Segoe UI", 9, "bold"), anchor="w").pack(anchor="w")
        tk.Label(frm_sap_txt, text="Abrir portal web",
                  bg=_SAP_BG, fg="#7BBCDD",
                  font=("Segoe UI", 7), anchor="w").pack(anchor="w")

        def _abrir_sap(e=None):
            import webbrowser
            webbrowser.open(_URL_SAP)
            self.log_consola("Abriendo SAP S/4HANA Cloud en el navegador...", "INFO")

        def _sap_enter(e):
            for w in [frm_sap, frm_sap_inner, _sap_icon, frm_sap_txt] + \
                      list(frm_sap_inner.winfo_children()) + \
                      list(frm_sap_txt.winfo_children()):
                try: w.configure(bg=_SAP_HOV)
                except Exception: pass

        def _sap_leave(e):
            for w in [frm_sap, frm_sap_inner, _sap_icon, frm_sap_txt] + \
                      list(frm_sap_inner.winfo_children()) + \
                      list(frm_sap_txt.winfo_children()):
                try: w.configure(bg=_SAP_BG)
                except Exception: pass

        for _w in [frm_sap, frm_sap_inner, _sap_icon, frm_sap_txt] + \
                   list(frm_sap_inner.winfo_children()) + \
                   list(frm_sap_txt.winfo_children()):
            _w.bind("<Button-1>", _abrir_sap)
            _w.bind("<Enter>",    _sap_enter)
            _w.bind("<Leave>",    _sap_leave)

        # ── Botón Crear Registro Info (acceso rápido) ─────────────────────
        _RI_BG  = "#1A0F2E"
        _RI_HOV = "#2A1A45"
        _RI_FG  = "#C8A84B"
        _RI_SUB = "#8A6FBB"

        frm_ri = tk.Frame(frm_btns, bg=_RI_BG, cursor="hand2")
        frm_ri.grid(row=1, column=1, sticky="nsew", padx=(4, 4))
        tk.Frame(frm_ri, bg="#7B4FA6", width=3).place(relx=0, rely=0, relheight=1)

        frm_ri_inner = tk.Frame(frm_ri, bg=_RI_BG, padx=10, pady=7)
        frm_ri_inner.pack(fill="both", expand=True)
        _ri_icon = tk.Label(frm_ri_inner, text="💾",
                             bg=_RI_BG, fg=_RI_FG,
                             font=("Segoe UI", 14))
        _ri_icon.pack(side="left", padx=(6, 8))
        frm_ri_txt = tk.Frame(frm_ri_inner, bg=_RI_BG)
        frm_ri_txt.pack(side="left", fill="x", expand=True)
        _ri_lbl_title = tk.Label(frm_ri_txt, text="Crear / Actualizar Registro Info SAP",
                                  bg=_RI_BG, fg="#FFFFFF",
                                  font=("Segoe UI", 9, "bold"), anchor="w")
        _ri_lbl_title.pack(anchor="w")
        _ri_lbl_sub = tk.Label(frm_ri_txt, text="Nuevo registro SAP",
                                bg=_RI_BG, fg=_RI_SUB,
                                font=("Segoe UI", 7), anchor="w")
        _ri_lbl_sub.pack(anchor="w")

        def _abrir_ri(e=None):
            self.log_consola("Iniciando: Crear Registro Info\u2026", "INFO")
            iniciar_crear_registro_info(self, self._carpeta_proyecto.get())

        def _ri_enter(e):
            for w in [frm_ri, frm_ri_inner, _ri_icon, frm_ri_txt,
                      _ri_lbl_title, _ri_lbl_sub]:
                try: w.configure(bg=_RI_HOV)
                except Exception: pass

        def _ri_leave(e):
            for w in [frm_ri, frm_ri_inner, _ri_icon, frm_ri_txt,
                      _ri_lbl_title, _ri_lbl_sub]:
                try: w.configure(bg=_RI_BG)
                except Exception: pass

        for _w in [frm_ri, frm_ri_inner, _ri_icon, frm_ri_txt,
                   _ri_lbl_title, _ri_lbl_sub]:
            _w.bind("<Button-1>", _abrir_ri)
            _w.bind("<Enter>",    _ri_enter)
            _w.bind("<Leave>",    _ri_leave)

        # ── (nov 2026) Manual de usuario (acceso rápido) ──────────────────
        _MAN_BG, _MAN_HOV = "#0E2A1E", "#14402C"
        frm_man = tk.Frame(frm_btns, bg=_MAN_BG, cursor="hand2")
        frm_man.grid(row=1, column=2, sticky="nsew", padx=(4, 0))
        tk.Frame(frm_man, bg="#2E8B57", width=3).place(relx=0, rely=0, relheight=1)
        frm_man_inner = tk.Frame(frm_man, bg=_MAN_BG, padx=10, pady=7)
        frm_man_inner.pack(fill="both", expand=True)
        _man_icon = tk.Label(frm_man_inner, text="\U0001f4d8", bg=_MAN_BG, fg="#FFFFFF", font=("Segoe UI", 14))
        _man_icon.pack(side="left", padx=(6, 8))
        frm_man_txt = tk.Frame(frm_man_inner, bg=_MAN_BG)
        frm_man_txt.pack(side="left", fill="x", expand=True)
        _man_t = tk.Label(frm_man_txt, text="Manual de usuario", bg=_MAN_BG, fg="#FFFFFF",
                          font=("Segoe UI", 9, "bold"), anchor="w")
        _man_t.pack(anchor="w")
        _man_s = tk.Label(frm_man_txt, text="Abrir el PDF", bg=_MAN_BG, fg="#7FC8A0",
                          font=("Segoe UI", 7), anchor="w")
        _man_s.pack(anchor="w")
        _man_ws = [frm_man, frm_man_inner, _man_icon, frm_man_txt, _man_t, _man_s]
        for _w in _man_ws:
            _w.bind("<Button-1>", lambda e: self._abrir_manual())
            _w.bind("<Enter>", lambda e: [w.configure(bg=_MAN_HOV) for w in _man_ws])
            _w.bind("<Leave>", lambda e: [w.configure(bg=_MAN_BG) for w in _man_ws])

    # ── (nov 2026) Manual de usuario y consulta ───────────────────────────
    def _abrir_manual(self):
        from app.comun import manual as _manual
        try:
            _manual.abrir_manual()
            self.log_consola("Abriendo el manual de usuario...", "INFO")
        except Exception as e:  # noqa: BLE001
            self.log_consola(f"No se pudo abrir el manual: {e}", "ERROR")
            messagebox.showerror("Manual de usuario", f"No se pudo abrir el manual:\n{e}")

    def _abrir_consulta_manual(self, consulta=""):
        from app.ui.ventana_consulta_manual import VentanaConsultaManual
        VentanaConsultaManual(self, consulta, self.log_consola)

    # ─────────────────────────────────────────────────────────────────────
    #  HELPERS DE UI
    # ─────────────────────────────────────────────────────────────────────

    def _section(self, parent, titulo):
        """Encabezado de sección con línea separadora."""
        frm = tk.Frame(parent, bg=C_SIDEBAR)
        frm.pack(fill="x", pady=(9, 3))
        tk.Label(frm, text=titulo, bg=C_SIDEBAR, fg=C_ACCENT,
                  font=("Segoe UI", 7, "bold")).pack(side="left")
        tk.Frame(frm, bg=C_BORDER, height=1).pack(
            side="left", fill="x", expand=True, padx=(6, 0), pady=(1, 0))

    def _boton_accion(self, parent, numero, texto, icono, cmd,
                      strip_color=C_MID, fg="white", state="normal", wip=False,
                      color=None):
        """Botón de acción con franja de color, número, icono y texto."""
        if color is not None:
            strip_color = color

        BASE_BG  = "#0D1F35"
        HOVER_BG = "#142840"
        DIM_BG   = "#070D18"
        enabled  = (state == "normal")
        bg = BASE_BG if enabled else DIM_BG

        frm = tk.Frame(parent, bg=bg,
                        cursor="hand2" if enabled else "arrow",
                        highlightbackground="#1A2D45" if enabled else "#0E1A28",
                        highlightthickness=1)
        frm._base_bg      = BASE_BG
        frm._hover_bg     = HOVER_BG
        frm._dim_bg       = DIM_BG
        frm._strip_col_on = strip_color
        frm._cmd          = cmd
        frm._enabled      = enabled

        # Franja de color izquierda (4 px)
        strip = tk.Frame(frm, bg=strip_color if enabled else "#1A2535", width=4)
        strip.pack(side="left", fill="y")
        strip.pack_propagate(False)
        frm._strip_w = strip

        # Número de acción
        badge = tk.Label(frm, text=f" {numero} ", bg=bg,
                          fg=C_ACCENT if enabled else "#2A3A4A",
                          font=("Segoe UI", 8, "bold"), padx=5, pady=5)
        badge.pack(side="left")
        frm._badge_w = badge

        # Icono + texto
        lbl = tk.Label(frm, text=f"{icono}  {texto}", bg=bg,
                        fg=C_TEXT_LIGHT if enabled else "#2A3A4A",
                        font=("Segoe UI", 9),
                        padx=6, pady=5, anchor="w")
        lbl.pack(side="left", fill="x", expand=True)
        frm._label_w = lbl

        # Badge WIP
        if wip:
            wip_lbl = tk.Label(frm, text="PRÓXIMO",
                                bg=C_ACCENT_DIM if not enabled else C_ACCENT,
                                fg=C_DARK,
                                font=("Segoe UI", 7, "bold"), padx=6, pady=2)
            wip_lbl.pack(side="right", padx=6)
            frm._wip_w = wip_lbl

        # Hover y click
        if enabled:
            _widgets = [frm, badge, lbl]

            def _enter(e):
                frm.configure(bg=HOVER_BG, highlightbackground="#2A4A6A")
                badge.configure(bg=HOVER_BG)
                lbl.configure(bg=HOVER_BG)

            def _leave(e):
                frm.configure(bg=BASE_BG, highlightbackground="#1A2D45")
                badge.configure(bg=BASE_BG)
                lbl.configure(bg=BASE_BG)

            def _click(e=None):
                frm._cmd()

            for w in _widgets:
                w.bind("<Enter>",    _enter)
                w.bind("<Leave>",    _leave)
                w.bind("<Button-1>", lambda e, f=_click: f())

        return frm

    def _set_boton_estado(self, frm, enabled, new_cmd=None):
        """Activa o desactiva un botón de acción preservando el estilo."""
        if new_cmd is not None:
            frm._cmd = new_cmd
        frm._enabled = enabled

        BASE_BG  = frm._base_bg
        HOVER_BG = frm._hover_bg
        DIM_BG   = frm._dim_bg
        bg = BASE_BG if enabled else DIM_BG

        frm.configure(bg=bg,
                       cursor="hand2" if enabled else "arrow",
                       highlightbackground="#1A2D45" if enabled else "#0E1A28")
        frm._strip_w.configure(bg=frm._strip_col_on if enabled else "#1A2535")
        frm._badge_w.configure(bg=bg, fg=C_ACCENT if enabled else "#2A3A4A")
        frm._label_w.configure(bg=bg, fg=C_TEXT_LIGHT if enabled else "#2A3A4A")

        if enabled:
            def _enter(e):
                frm.configure(bg=HOVER_BG, highlightbackground="#2A4A6A")
                frm._badge_w.configure(bg=HOVER_BG)
                frm._label_w.configure(bg=HOVER_BG)

            def _leave(e):
                frm.configure(bg=BASE_BG, highlightbackground="#1A2D45")
                frm._badge_w.configure(bg=BASE_BG)
                frm._label_w.configure(bg=BASE_BG)

            def _click(e=None):
                frm._cmd()

            for w in [frm, frm._badge_w, frm._label_w]:
                w.bind("<Enter>",    _enter)
                w.bind("<Leave>",    _leave)
                w.bind("<Button-1>", lambda e, f=_click: f())
        else:
            for w in [frm, frm._badge_w, frm._label_w]:
                w.bind("<Enter>",    lambda e: None)
                w.bind("<Leave>",    lambda e: None)
                w.bind("<Button-1>", lambda e: None)

    def _abrir_carpeta(self, ruta):
        if ruta and os.path.isdir(ruta):
            try:
                os.startfile(ruta)
            except AttributeError:
                import subprocess
                subprocess.call(["xdg-open", ruta])

    def _actualizar_indicadores_sub(self):
        """Actualiza los pill-badges de TARIFAS y RECURSOS."""
        cp = self._carpeta_proyecto.get()
        if cp and os.path.isdir(cp):
            t = os.path.join(cp, FOLDER_TARIFAS)
            r = os.path.join(cp, FOLDER_RECURSOS)

            # ── TARIFAS ───────────────────────────────────────────────────
            if os.path.isdir(t):
                self._ruta_sub_tarifas = t
                _pill_bg = "#0E2A1A"
                _pill_bd = "#1A4A2A"
                self._frm_sub_tarifas.configure(bg=_pill_bg,
                                                 highlightbackground=_pill_bd)
                self._lbl_sub_tarifas.config(
                    text="TARIFAS ✔", fg=C_SUCCESS,
                    bg=_pill_bg, cursor="hand2",
                    font=("Segoe UI", 7, "bold"))
                self._lbl_sub_tarifas.bind(
                    "<Button-1>",
                    lambda e: self._abrir_carpeta(self._ruta_sub_tarifas))
            else:
                self._ruta_sub_tarifas = None
                _pill_bg = "#2A0E0E"
                _pill_bd = "#4A1A1A"
                self._frm_sub_tarifas.configure(bg=_pill_bg,
                                                 highlightbackground=_pill_bd)
                self._lbl_sub_tarifas.config(
                    text="TARIFAS ✘", fg=C_ERROR,
                    bg=_pill_bg, cursor="arrow",
                    font=("Segoe UI", 7, "bold"))
                self._lbl_sub_tarifas.unbind("<Button-1>")

            # ── RECURSOS ──────────────────────────────────────────────────
            if os.path.isdir(r):
                self._ruta_sub_recursos = r
                _pill_bg = "#0E2A1A"
                _pill_bd = "#1A4A2A"
                self._frm_sub_recursos.configure(bg=_pill_bg,
                                                  highlightbackground=_pill_bd)
                self._lbl_sub_recursos.config(
                    text="RECURSOS ✔", fg=C_SUCCESS,
                    bg=_pill_bg, cursor="hand2",
                    font=("Segoe UI", 7, "bold"))
                self._lbl_sub_recursos.bind(
                    "<Button-1>",
                    lambda e: self._abrir_carpeta(self._ruta_sub_recursos))
            else:
                self._ruta_sub_recursos = None
                _pill_bg = "#2A0E0E"
                _pill_bd = "#4A1A1A"
                self._frm_sub_recursos.configure(bg=_pill_bg,
                                                  highlightbackground=_pill_bd)
                self._lbl_sub_recursos.config(
                    text="RECURSOS ✘", fg=C_ERROR,
                    bg=_pill_bg, cursor="arrow",
                    font=("Segoe UI", 7, "bold"))
                self._lbl_sub_recursos.unbind("<Button-1>")

            # ── RAÍZ ──────────────────────────────────────────────────────
            _cp = cp
            self._lbl_abrir_raiz.config(
                text="↗ RAÍZ", fg=C_ACCENT, cursor="hand2")
            self._lbl_abrir_raiz.bind(
                "<Button-1>", lambda e, p=_cp: self._abrir_carpeta(p))
        else:
            # Sin carpeta válida: resetear todo a estado neutro
            _pill_bg = "#0D1F35"
            _pill_bd = "#162840"
            for frm_pill, lbl_pill, txt in [
                (self._frm_sub_tarifas,  self._lbl_sub_tarifas,  "TARIFAS  —"),
                (self._frm_sub_recursos, self._lbl_sub_recursos, "RECURSOS  —"),
            ]:
                frm_pill.configure(bg=_pill_bg, highlightbackground=_pill_bd)
                lbl_pill.config(text=txt, fg=C_MUTED, bg=_pill_bg,
                                 cursor="arrow", font=("Segoe UI", 7, "bold"))
                lbl_pill.unbind("<Button-1>")
            self._lbl_abrir_raiz.config(
                text="↗ RAÍZ", fg=C_MUTED, cursor="arrow")
            self._lbl_abrir_raiz.unbind("<Button-1>")
            self._ruta_sub_tarifas  = None
            self._ruta_sub_recursos = None

    def log_consola(self, msg, nivel=">>"):
        """
        Escribe una línea en la consola. Se llama miles de veces desde los
        hilos de fondo de proceso() de cada Acción -- con diferencia, el
        punto de la app que más toca Tkinter desde fuera del hilo principal.

        Fix (2026-09-24, mismo hallazgo que _ejecutar_en_hilo_principal_y_
        esperar): tocar el widget de texto directamente desde un hilo de
        fondo no solo arriesga colgar la app -- también puede hacer que el
        log se procese fuera de orden respecto a otras actualizaciones ya
        correctamente programadas con self.after() (como la barra de
        progreso), dando la sensación de que la barra se pone en verde
        antes de que el log haya terminado de escribirse. Si quien llama
        es un hilo de fondo, se programa con self.after(0, ...) -- sin
        bloquear (a diferencia de _ejecutar_en_hilo_principal_y_esperar,
        aquí no hace falta esperar el resultado, y bloquear en cada línea
        de log sería carísimo) -- así todas las actualizaciones de UI
        Fix adicional (2026-09-25): el timestamp mostrado en cada línea lo
        generaba ConsolaLog.log() con datetime.now() en el momento de
        RENDERIZAR, no en el que se LLAMÓ aquí -- si hay cola pendiente en
        self.after() (varias líneas seguidas desde un hilo de fondo), la
        hora mostrada podía quedar varios segundos por detrás del momento
        real, dando la sensación de que el proceso seguía en marcha después
        de que la barra ya marcara completado. Se captura la hora aquí, en
        el instante real de la llamada, y se pasa ya formada.
        """
        ts = datetime.now().strftime("%H:%M:%S")

        def _escribir():
            try:
                self.consola.log(msg, nivel, ts=ts)
            except tk.TclError:
                # La ventana principal (o su consola) ya no existe -- p.ej. se
                # cerró justo antes de que este log, en cola desde un hilo de
                # fondo o un callback tardío, llegara a ejecutarse. No hay
                # consola donde escribir; se ignora en vez de crashear la app.
                pass

        if threading.current_thread() is threading.main_thread():
            _escribir()
        else:
            self.after(0, _escribir)

    def log_consola_batch(self, lineas_niveles):
        """
        Versión thread-safe de self.consola.log_batch() -- mismo fix que
        log_consola() (2026-09-24/25): _actualizar_listado_sap() (Acción 3)
        llamaba a self.consola.log_batch(...) directamente desde su hilo
        de fondo, sin ninguna protección, saltándose por completo el fix
        de thread-safety. Captura también el timestamp en el momento real
        de la llamada, por el mismo motivo que log_consola().
        """
        ts = datetime.now().strftime("%H:%M:%S")

        def _escribir():
            try:
                self.consola.log_batch(lineas_niveles, ts=ts)
            except tk.TclError:
                pass

        if threading.current_thread() is threading.main_thread():
            _escribir()
        else:
            self.after(0, _escribir)

    def _set_estado(self, encontrado, ruta="", fecha="", reset_resultado=True):
        if encontrado:
            self._lbl_estado_icon.config(fg=C_SUCCESS, text="●")
            self._lbl_estado.config(fg=C_SUCCESS,
                                     text="TARIFA.xls\nENCONTRADA \u2714")
            txt = f"Mod.: {fecha}" if fecha else ruta
            self._lbl_ruta_resultado.config(text=txt, fg=C_SUCCESS)
            # Acción 4 siempre activa; Acción 5 solo si hay resultado generado
        else:
            self._lbl_estado_icon.config(fg=C_ERROR, text="●")
            self._lbl_estado.config(fg=C_ERROR,
                                     text="TARIFA.xls\nNO ENCONTRADA")
            self._lbl_ruta_resultado.config(
                text="Puede abrir tarifa\nexistente o renombrar.", fg=C_WARNING)
            # Acción 4 permanece activa; solo resetear Acción 5 si procede
            if reset_resultado:
                self._habilitar_btn_resultado(False)

    def _habilitar_btn_procesar(self, habilitar):
        cmd = self._procesar if habilitar else None
        self._set_boton_estado(self._btn_procesar, habilitar, cmd)

    def _habilitar_btn_resultado(self, habilitar):
        cmd = self._abrir_resultado if habilitar else None
        self._set_boton_estado(self._btn_abrir_resultado, habilitar, cmd)

    # ── Acciones ──────────────────────────────────────────────

    def _wip(self):
        messagebox.showinfo(
            "En construcci\xf3n",
            "\U0001f6a7  Esta funci\xf3n est\xe1 actualmente en construcci\xf3n.\n\n"
            "Estar\xe1 disponible en una pr\xf3xima versi\xf3n.")

    def _seleccionar_proyecto(self):
        carpeta = filedialog.askdirectory(
            title="Seleccionar carpeta del Proyecto",
            initialdir=self._carpeta_proyecto.get() or os.path.expanduser("~"))
        if carpeta:
            self._carpeta_proyecto.set(carpeta)
            self.config_data["carpeta_proyecto"] = carpeta
            guardar_config(self.config_data)
            self.consola.separador("NUEVA CARPETA DE PROYECTO")
            self.log_consola(f"Proyecto: {carpeta}", "INFO")
            self._actualizar_indicadores_sub()
            self._set_estado(False)
            self._barra.reset()
            self._ruta_tarifa = None
            self._buscar()

    def _crear_proyecto(self):
        vent = VentanaCrearProyecto(self)
        self.wait_window(vent)
        if vent.ruta_creada:
            self._carpeta_proyecto.set(vent.ruta_creada)
            self.config_data["carpeta_proyecto"] = vent.ruta_creada
            guardar_config(self.config_data)
            self.consola.separador("PROYECTO CREADO")
            self.log_consola(f"Proyecto creado en: {vent.ruta_creada}", "OK")
            self._actualizar_indicadores_sub()
            self._set_estado(False)
            self._barra.reset()
            self._ruta_tarifa = None

    def _buscar(self):
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            messagebox.showwarning("Sin proyecto",
                                   "Seleccione primero la carpeta del Proyecto.")
            return
        if not os.path.isdir(carpeta):
            messagebox.showerror("Carpeta no v\xe1lida",
                                  f"La ruta no existe:\n{carpeta}")
            return

        carpeta_tarifas = os.path.join(carpeta, FOLDER_TARIFAS)

        def proceso():
            self.consola.separador("INICIO DE B\xdaSQUEDA")
            self.log_consola("Verificando estructura del proyecto...", "INFO")
            self._barra.set(10, 100, "Verificando carpetas...")
            time.sleep(0.3)

            # Verificar subcarpetas
            for sub in [FOLDER_TARIFAS, FOLDER_RECURSOS]:
                ruta_sub = os.path.join(carpeta, sub)
                if os.path.isdir(ruta_sub):
                    self.log_consola(f"Carpeta {sub}: OK", "OK")
                else:
                    self.log_consola(f"Carpeta {sub}: NO ENCONTRADA", "WARN")

            self._barra.set(30, 100, "Leyendo directorio TARIFAS...")
            time.sleep(0.3)
            self.log_consola(f"Buscando {TARIFA_NAME} en TARIFAS/...", "INFO")
            self._barra.set(60, 100, f"Buscando {TARIFA_NAME}...")
            time.sleep(0.4)
            ruta = buscar_tarifa(carpeta)
            self._barra.set(90, 100, "Procesando resultado...")
            time.sleep(0.2)

            if ruta:
                ts    = os.path.getmtime(ruta)
                fecha = datetime.fromtimestamp(ts).strftime("%d-%m-%Y %H:%M")
                self.log_consola(f"Archivo encontrado: {ruta}", "OK")
                self.log_consola(f"\xdaltima modificaci\xf3n: {fecha}", "OK")
                self._barra.completado("TARIFA.xls localizada correctamente")
                self._ruta_tarifa = ruta
                self.after(0, lambda: self._set_estado(True, ruta, fecha))
            else:
                self.log_consola("TARIFA.xls NO encontrada en TARIFAS/.", "WARN")
                archivos = listar_xls(carpeta_tarifas) if os.path.isdir(carpeta_tarifas) else []
                self.log_consola(f"Archivos .xls disponibles: {len(archivos)}", "INFO")
                for a in archivos:
                    self.log_consola(f"  \u2514 {a['nombre']}  [{a['fecha_mod']}]", ">>")
                self._barra.reset("TARIFA.xls no encontrada")
                self._ruta_tarifa = None
                self.after(0, lambda: self._set_estado(False))
                if archivos:
                    self.after(0, lambda: self._abrir_seleccion(archivos, carpeta_tarifas))

        threading.Thread(target=proceso, daemon=True).start()

    def _abrir_seleccion(self, archivos, carpeta_tarifas):
        vent = VentanaSeleccionXLS(self, archivos, carpeta_tarifas, self.consola)
        self.wait_window(vent)
        if vent.seleccionado:
            self.after(200, self._buscar)


    def _abrir_resultado(self):
        if self._ruta_resultado and os.path.isfile(self._ruta_resultado):
            self.log_consola(f"Abriendo resultado: {self._ruta_resultado}", "INFO")
            try:
                abrir_con_su_programa(self._ruta_resultado)    # (13 nov 2026) que Excel quede delante
            except Exception as e:
                self.log_consola(f"Error al abrir: {e}", "ERROR")
                messagebox.showerror("Error", f"No se pudo abrir:\n{e}")
        else:
            messagebox.showwarning("Sin resultado",
                                   "A\xfan no se ha generado ning\xfan archivo procesado.")

    # ── Actualización automática (nov 2026) ───────────────────────────────
    def _comprobar_actualizaciones(self, version_local):
        """Al arrancar: consulta en segundo plano la última Release publicada y, si es más nueva,
        avisa (app/ui/ventana_actualizacion.py). Desde el código («python tarifa_app.py») no se
        actualiza: solo el .exe. Un fallo de red solo deja una línea en la consola."""
        import sys as _sys
        from app.comun import actualizaciones as _act
        if getattr(_sys, "frozen", False):
            self._limpiar_version_anterior(os.path.abspath(_sys.executable))

        def hilo():
            try:
                info = _act.consultar_ultima()
            except Exception as e:  # noqa: BLE001
                msg = f"No se pudo comprobar si hay versión nueva ({type(e).__name__})."
                self.after(0, lambda: self.log_consola(msg, "WARN"))
                return
            self.after(0, lambda: self._resultado_actualizacion(info, version_local))
        threading.Thread(target=hilo, daemon=True, name="ComprobarActualizacion").start()

    def _limpiar_version_anterior(self, ruta_exe, intento=1):
        """Tras actualizarse, borra el .exe de la versión anterior y lo dice en la consola. Si la versión
        vieja todavía se está cerrando, lo reintenta (hasta 6 veces, cada 10 s)."""
        from app.comun import actualizaciones as _act
        borrados, pendientes = _act.limpiar_restos(ruta_exe)
        for n in borrados:
            self.log_consola(f"Versión anterior borrada: {n}", "OK")
        if pendientes and intento < 6:
            self.after(10000, lambda: self._limpiar_version_anterior(ruta_exe, intento + 1))
        elif pendientes:
            self.log_consola(f"No se pudo borrar todavía {', '.join(pendientes)}; se volverá a intentar "
                             f"en el próximo arranque (se puede borrar a mano).", "WARN")

    def _resultado_actualizacion(self, info, version_local):
        from app.comun import actualizaciones as _act
        if not info:
            return
        if not _act.es_mas_nueva(info["version"], version_local):
            self.log_consola(f"Versión al día ({version_local.strip()}).", "OK")
            return
        self.log_consola(f"Nueva versión disponible: v {info['version']} (tienes la {version_local.strip()}).", "WARN")
        from app.ui.ventana_actualizacion import VentanaActualizacion
        VentanaActualizacion(self, info, version_local, self.log_consola)

