# -*- coding: utf-8 -*-
"""
tarifa_app.py  —  Ventana principal y punto de entrada
CREACIÓN REGISTRO INFO SAP — Central Compras Princess Canarias
Autor: Victor Alejandro Martin Andrade (VAMA)
(c) VAMA 2026

Estructura del proyecto:
    tarifa_design.py   — paleta de colores y tipografía
    tarifa_core.py     — recursos, config, lógica de negocio, widgets UI
    tarifa_dialogs.py  — ventanas modales (Toplevel)
    tarifa_app.py      — ventana principal App + entry point  ← este archivo
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import base64
import threading
import time

from tarifa_constants import (
    C_DARK, C_MID, C_SIDEBAR, C_CARD, C_BORDER, C_HOVER,
    C_ACCENT, C_ACCENT2, C_ACCENT_DIM,
    C_BG, C_TEXT, C_TEXT_LIGHT, C_MUTED, C_SUCCESS,
    C_ERROR, C_WARNING, C_LOG_BG,
    FONT_SUB, FONT_LABEL, FONT_SMALL, SAP_LOGO_B64,
    TABLA_SUBFAMILIA_GRUPO_COMPRAS, TABLA_IGIC_GRUPO_COMPRAS,
)
from tarifa_core import (
    resource_path, LOGO_B64,
    cargar_config, guardar_config, TARIFA_NAME,
    FOLDER_TARIFAS, FOLDER_RECURSOS, buscar_tarifa, listar_xls,
    ConsolaLog, CanvasProgressBar,
)
from credenciales import (
    leer_credenciales_dali, guardar_credenciales_dali,
    borrar_credenciales_dali, migrar_credenciales_dali,
)
from tarifa_dialogs import (
    VentanaSeleccionXLS, VentanaCrearProyecto,
    VentanaSAPFaltantes, VentanaResolverSAPDuplicados,
)
from crear_registro_info import iniciar_crear_registro_info
from acciones import organizar_dali as organizar_dali_datos
from sap_helpers import (
    CATALOGO_COLUMNAS, CATALOGO_CLAVES_SAP,
    CATALOGO_FILA_CABECERA, CATALOGO_PRIMERA_FILA_DATOS,
    mapa_columnas_catalogo, datos_sap_desde_fila_listado, es_precio_valido,
)

# ─── Versión de la aplicación ────────────────────────────────────────────────
APP_VERSION = "v 10.2.36"   # ← actualiza aquí cada vez que termines una fase

# URL del worker de Cloudflare que hace de proxy inverso hacia el backend
# de la app web "Catálogo Asignaciones" (dali-sap-articulos-app). Usada
# solo por _descargar_catalogo_dali_web() para la descarga automática del
# catálogo completo (ver _confirmar_catalogo_actualizado()). Si esta URL
# cambia algún día (redespliegue del Worker con otro nombre), solo hay
# que actualizarla aquí.
DALI_APP_API_BASE = "https://dali-proxy-api.centralcompras1-canarias.workers.dev"
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    # SAP_LOGO_B64 centralizado en tarifa_constants.py

    def __init__(self):
        super().__init__()
        self.title("CREACI\xd3N REGISTRO INFO SAP \u2014 Central Compras Princess Canarias")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        self.minsize(820, 640)
        try:
            self.iconbitmap(resource_path("assets/logo.ico"))
        except Exception:
            pass
        self.config_data       = cargar_config()
        # (nov 2026) Una contraseña de Catálogo Asignaciones antigua en claro se cifra
        # con DPAPI y se borra la copia en claro (ver credenciales.py).
        if migrar_credenciales_dali(self.config_data):
            guardar_config(self.config_data)
        # Memoria de sesión: una vez resuelto el aviso de catalogo-dali-completo.xlsx
        # desactualizado (Acción 2/3), no se repregunta hasta reiniciar la app.
        self._catalogo_verificado_sesion = False
        # Bloqueo global de catalogo-dali-completo.xlsx (2026-09-24): más de
        # una acción puede querer escribir este archivo a la vez (regenerarlo
        # tras Acciones 2/3, actualizar código proveedor al completar una
        # tarifa, o descargarlo fresco desde Catálogo Asignaciones) -- sin
        # protección, dos escrituras simultáneas pueden chocar y perder una
        # de las dos en silencio. Caso real que lo destapó: el usuario lanzó
        # Acción 3 justo cuando el cierre de una transformación aún estaba
        # escribiendo el código proveedor de esa tarifa en el catálogo.
        self._catalogo_lock = threading.Lock()
        self._carpeta_proyecto = tk.StringVar(
            value=self.config_data.get("carpeta_proyecto", ""))
        self._ruta_tarifa      = None
        self._ruta_resultado   = None
        self._logo_img         = None
        self._build_ui()
        self._centrar()
        self.log_consola("Sistema iniciado. Listo para operar.", "SISTEMA")
        self.log_consola("Cargando configuraci\xf3n guardada...", "INFO")
        if self._carpeta_proyecto.get():
            self.log_consola(f"Carpeta PROYECTO: {self._carpeta_proyecto.get()}", "INFO")
            self._actualizar_indicadores_sub()
            self.after(300, self._buscar)
        else:
            self.log_consola("No hay carpeta de proyecto configurada. Seleccione o cree una.", "WARN")

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

    # ── HEADER ────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=C_DARK)
        hdr.pack(fill="x")

        # Franja dorada superior (3 px)
        tk.Frame(hdr, bg=C_ACCENT, height=3).pack(fill="x")

        contenido = tk.Frame(hdr, bg=C_DARK, pady=8, padx=16)
        contenido.pack(fill="x")

        # ── Logo ──────────────────────────────────────────────────────────
        try:
            from io import BytesIO
            from PIL import Image, ImageTk
            ruta_logo = resource_path("assets/logo.png")
            if os.path.isfile(ruta_logo):
                img = Image.open(ruta_logo)
            else:
                img = Image.open(BytesIO(base64.b64decode(LOGO_B64)))
            img    = img.resize((54, 58), Image.LANCZOS).convert("RGBA")
            bg_img = Image.new("RGBA", img.size, (6, 13, 24, 255))
            bg_img.paste(img, mask=img.split()[3])
            self._logo_img = ImageTk.PhotoImage(bg_img)
            tk.Label(contenido, image=self._logo_img,
                     bg=C_DARK, cursor="arrow").pack(side="left", padx=(0, 14))
        except Exception:
            tk.Label(contenido, text="P", bg=C_DARK, fg=C_ACCENT,
                     font=("Segoe UI", 22, "bold italic")).pack(side="left", padx=(0, 14))

        # ── Títulos ───────────────────────────────────────────────────────
        frm_titles = tk.Frame(contenido, bg=C_DARK)
        frm_titles.pack(side="left", fill="y")

        tk.Label(frm_titles,
                 text="CREACIÓN  REGISTRO  INFO  SAP",
                 bg=C_DARK, fg=C_ACCENT,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(frm_titles,
                 text="Central Compras  ·  Princess Canarias",
                 bg=C_DARK, fg="#6A90B8",
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(1, 0))
        tk.Label(frm_titles,
                 text="Módulo: Gestión de Tarifas  —  DALI · SAP",
                 bg=C_DARK, fg=C_MUTED,
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(1, 0))

        # ── Lado derecho: logo SAP + reloj ────────────────────────────────
        try:
            from io import BytesIO as _BIO
            from PIL import Image as _Img, ImageTk as _ImgTk
            _sap_pil = _Img.open(_BIO(base64.b64decode(SAP_LOGO_B64))).resize((72, 26), _Img.LANCZOS)
            _sap_bg  = _Img.new("RGBA", _sap_pil.size, (6, 13, 24, 255))
            (_sap_bg.paste(_sap_pil, mask=_sap_pil.split()[3])
             if _sap_pil.mode == "RGBA" else _sap_bg.paste(_sap_pil))
            self._sap_header_img = _ImgTk.PhotoImage(_sap_bg)
            tk.Label(contenido, image=self._sap_header_img,
                     bg=C_DARK, cursor="arrow").pack(side="right", padx=(14, 0))
        except Exception:
            tk.Label(contenido, text="SAP", bg="#1B4F8A", fg="#FFFFFF",
                     font=("Segoe UI", 10, "bold"), padx=8, pady=2).pack(
                         side="right", padx=(14, 0))

        # Reloj en vivo
        frm_clock = tk.Frame(contenido, bg=C_DARK)
        frm_clock.pack(side="right", fill="y", padx=(0, 6))
        self._lbl_clock = tk.Label(frm_clock, text="", bg=C_DARK, fg=C_ACCENT2,
                                    font=("Segoe UI", 18, "bold"), anchor="e")
        self._lbl_clock.pack(anchor="e")
        self._lbl_clock_date = tk.Label(frm_clock, text="", bg=C_DARK, fg=C_MUTED,
                                         font=("Segoe UI", 7), anchor="e")
        self._lbl_clock_date.pack(anchor="e")
        tk.Label(frm_clock, text=APP_VERSION, bg=C_DARK, fg=C_ACCENT2,
                 font=("Segoe UI", 10, "bold"), anchor="e").pack(anchor="e", pady=(3, 0))

        # Franja dorada inferior (2 px)
        tk.Frame(hdr, bg=C_ACCENT, height=2).pack(fill="x")
        self._tick_clock()

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

        tk.Label(frm_btns, text="ACCESOS RÁPIDOS",
                  bg=_ACC_BG, fg="#1A3A5A",
                  font=("Segoe UI", 7, "bold"),
                  anchor="w").grid(row=0, column=0, columnspan=2,
                                   sticky="w", padx=2, pady=(0, 5))

        # ── Botón SAP ─────────────────────────────────────────────────────
        _URL_SAP = "https://my405727.s4hana.cloud.sap/ui#Shell-home"

        frm_sap = tk.Frame(frm_btns, bg=_SAP_BG, cursor="hand2")
        frm_sap.grid(row=1, column=0, sticky="nsew", padx=(0, 4))
        tk.Frame(frm_sap, bg="#1A9BDA", width=3).place(relx=0, rely=0, relheight=1)

        frm_sap_inner = tk.Frame(frm_sap, bg=_SAP_BG, padx=10, pady=7)
        frm_sap_inner.pack(fill="both", expand=True)

        try:
            from io import BytesIO as _BytesIO
            from PIL import Image as _Image, ImageTk as _ImageTk
            _sap_pil = _Image.open(_BytesIO(base64.b64decode(SAP_LOGO_B64))).resize(
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
        frm_ri.grid(row=1, column=1, sticky="nsew", padx=(4, 0))
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

    # ── FOOTER ────────────────────────────────────────────────────────────
    def _build_footer(self):
        footer = tk.Frame(self, bg=C_DARK)
        footer.pack(fill="x", side="bottom")
        tk.Frame(footer, bg=C_ACCENT_DIM, height=1).pack(fill="x")
        frm_pie = tk.Frame(footer, bg=C_DARK, pady=4)
        frm_pie.pack(fill="x", padx=16)
        tk.Label(frm_pie,
                  text="🏨  Princess Hotels & Resorts  ·  Central de Compras Canarias",
                  bg=C_DARK, fg=C_ACCENT,
                  font=("Segoe UI", 8, "bold")).pack(side="left")
        tk.Label(frm_pie,
                  text=f"© VAMA 2026  ·  Víctor Alejandro Martín Andrade  ·  {APP_VERSION}",
                  bg=C_DARK, fg=C_MUTED,
                  font=("Segoe UI", 8)).pack(side="right")

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

    def _proteger_hoja_destino(self, ws_dest, nombre_archivo="destino"):
        """
        Protege la hoja de 'LISTADO CODIGOS DALI - SAP.xlsx' contra edición,
        independientemente de si se encontró bloqueada o desbloqueada al
        abrirla. Se llama justo antes de cada guardado que toca este archivo
        (Acción 2 y Acción 3), para garantizar que SIEMPRE quede protegida
        al terminar el proceso.

        Con sheet=True, openpyxl deja por defecto BLOQUEADAS todas las
        acciones de edición (formatear celdas, insertar/borrar filas y
        columnas, ordenar, autofiltro, etc.) — es el comportamiento
        equivalente a "Proteger hoja" en Excel con todas las casillas
        desmarcadas. Protegida con contraseña ("Princess2026", la misma
        que se usa siempre para este archivo) — openpyxl la hashea
        automáticamente al asignarla, no se guarda en texto plano en el
        propio .xlsx.
        """
        try:
            ws_dest.protection.sheet = True
            ws_dest.protection.password = "Princess2026"
            self.log_consola(f"  └ Hoja protegida (edición bloqueada) en {nombre_archivo}", "OK")
        except Exception as _e_prot:
            self.log_consola(
                f"  ⚠  No se pudo proteger la hoja en {nombre_archivo}: {_e_prot}",
                "WARN")

    def _actualizar_listado_dali(self):
        """
        PUNTO 2 — Actualizar listado artículos DALI.
        Origen : RECURSOS/ARTICULOS_DALI.xlsx  (generado por Punto 1)
        Destino: <raíz proyecto>/LISTADO CODIGOS DALI - SAP.xlsx
        Mapeo  : A→A, B→B, C→C (desde fila 8), col I (ACTIVO) → col O (ACTIVO DALI)
        (mapeo actualizado 2026-08: tras insertar las 3 columnas nuevas
        F, K, L en LISTADO CODIGOS DALI-SAP, la antigua col L (ACTIVO)
        se desplazó a la col O -- ADEMÁS, escribir en L causaba que
        Acción 3 lo borrara en su "Limpieza previa" (rango E:N) para
        cualquier material sin match SAP, dejando ACTIVO vacío)
        """
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            messagebox.showwarning("Sin proyecto",
                                   "Seleccione primero la carpeta del Proyecto.")
            return

        if not self._confirmar_catalogo_actualizado():
            self.log_consola(
                "Actualización DALI cancelada por el usuario (catálogo desactualizado).", "WARN")
            return

        # ── Rutas ────────────────────────────────────────────────
        ruta_origen  = os.path.join(carpeta, FOLDER_RECURSOS, "ARTICULOS_DALI.xlsx")
        ruta_destino = os.path.join(carpeta, "LISTADO CODIGOS DALI - SAP.xlsx")

        # ── Verificar ARTICULOS_DALI.xlsx ────────────────────────
        if not os.path.isfile(ruta_origen):
            messagebox.showwarning(
                "Archivo no encontrado",
                "\u274c  No se encontró el archivo:\n\n"
                f"   {ruta_origen}\n\n"
                "Debe ejecutar primero la opción\n"
                "\u25b6  1 — Organizar artículos DALI\n\n"
                "para generar este archivo antes de continuar.")
            return

        # ── Leer fecha y hora del encabezado (celdas B3 y B5) ───────
        try:
            import openpyxl
            wb_orig = openpyxl.load_workbook(ruta_origen, read_only=True, data_only=True)
            ws_orig = wb_orig.active
            celda_b3 = ws_orig["B3"].value or ""
            celda_b5 = ws_orig["B5"].value or ""
            wb_orig.close()
        except Exception as e:
            messagebox.showerror("Error al leer origen",
                                 f"No se pudo leer ARTICULOS_DALI.xlsx:\n{e}")
            return

        # Extraer solo el valor limpio eliminando el prefijo de texto
        # Ej: "Actualizado: 08/04/2026" -> "08/04/2026" ; "Hora: 08:47" -> "08:47"
        def _extraer_valor(texto):
            raw = str(texto).strip() if texto else ""
            return raw.split(":", 1)[-1].strip() if ":" in raw else raw

        texto_fecha    = str(celda_b3).strip() if celda_b3 else "(fecha no disponible)"
        valor_fecha_c4 = _extraer_valor(celda_b3) or "(no disponible)"
        valor_hora_c5  = _extraer_valor(celda_b5) or "(no disponible)"

        # ── Confirmación al usuario ───────────────────────────────
        if not messagebox.askyesno(
                "Confirmar actualización DALI",
                f"Archivo origen encontrado:\n"
                f"   ARTICULOS_DALI.xlsx\n\n"
                f"   {texto_fecha}\n\n"
                f"Destino:\n"
                f"   LISTADO CODIGOS DALI - SAP.xlsx\n\n"
                "¿Desea continuar?"):
            self.log_consola("Actualización DALI cancelada por el usuario.", "WARN")
            return

        # ── Verificar destino ─────────────────────────────────────
        if not os.path.isfile(ruta_destino):
            messagebox.showerror(
                "Destino no encontrado",
                "\u274c  No se encontró el archivo destino:\n\n"
                f"   {ruta_destino}\n\n"
                "Verifique que el archivo 'LISTADO CODIGOS DALI - SAP.xlsx'\n"
                "esté en la raíz de la carpeta del proyecto.")
            return

        # ── Proceso en hilo ──────────────────────────────────────
        def proceso():
            self.consola.separador("INICIO ACTUALIZACIÓN DALI")
            self.log_consola("Iniciando actualización listado artículos DALI...", "INFO")
            self.after(0, lambda: self._barra.indeterminate("Leyendo ARTICULOS_DALI..."))

            try:
                import openpyxl

                # -- Leer origen (solo valores, sin fórmulas) --
                self.log_consola(f"Leyendo origen: {os.path.basename(ruta_origen)}", "INFO")
                wb_orig = openpyxl.load_workbook(ruta_origen, read_only=True, data_only=True)
                ws_orig = wb_orig.active

                # Recorrer desde fila 8 hasta la última con valor en col A
                FILA_INI = 8
                datos = []  # lista de (cod_a, val_b, val_c, activo_i)
                for fila in ws_orig.iter_rows(min_row=FILA_INI, values_only=True):
                    val_a = fila[0]  # col A (índice 0)
                    if val_a is None or str(val_a).strip() == "":
                        break        # primera fila vacía en A → fin de datos
                    val_b    = fila[1]  # col B
                    val_c    = fila[2]  # col C
                    activo_i = fila[8] if len(fila) > 8 else None  # col I (índice 8)
                    datos.append((val_a, val_b, val_c, activo_i))

                wb_orig.close()
                n_filas = len(datos)
                self.log_consola(f"Filas leídas desde origen: {n_filas}", "OK")

                if n_filas == 0:
                    self.after(0, lambda: self._barra.reset("Sin datos en origen"))
                    self.after(0, lambda: messagebox.showwarning(
                        "Sin datos",
                        "No se encontraron filas con datos a partir de la fila 8\n"
                        "en el archivo ARTICULOS_DALI.xlsx."))
                    return

                # -- Abrir destino (lectura/escritura, conservando formato) --
                self.log_consola(f"Abriendo destino: {os.path.basename(ruta_destino)}", "INFO")
                self.after(0, lambda: self._barra.indeterminate("Escribiendo destino..."))

                wb_dest = openpyxl.load_workbook(ruta_destino)
                ws_dest = wb_dest.active  # primera hoja

                # -- Escribir datos fila a fila --
                for i, (val_a, val_b, val_c, activo_i) in enumerate(datos):
                    fila_dest = FILA_INI + i
                    ws_dest.cell(row=fila_dest, column=1,  value=val_a)    # A → A
                    ws_dest.cell(row=fila_dest, column=2,  value=val_b)    # B → B
                    ws_dest.cell(row=fila_dest, column=3,  value=val_c)    # C → C
                    ws_dest.cell(row=fila_dest, column=15, value=activo_i) # I → O (col 15)

                # -- Eliminar físicamente filas sobrantes en destino (si había más que ahora) --
                #    ANTES se vaciaban solo las columnas A/B/C/L celda a celda, pero eso
                #    dejaba intactas las columnas D-K de ejecuciones anteriores y, además,
                #    ws.max_row seguía «viendo» esas filas como ocupadas (arrastraba el
                #    recuento a Acción 3, que volvía a cruzar filas fantasma). Se eliminan
                #    las filas por completo con delete_rows(), que limpia TODAS las columnas
                #    de una vez y corrige max_row para las acciones posteriores.
                ultima_fila_dest = ws_dest.max_row
                fila_fin_nuevo   = FILA_INI + n_filas - 1
                if ultima_fila_dest > fila_fin_nuevo:
                    n_sobrantes = ultima_fila_dest - fila_fin_nuevo
                    ws_dest.delete_rows(fila_fin_nuevo + 1, n_sobrantes)
                    self.log_consola(
                        f"Eliminadas físicamente {n_sobrantes} fila(s) sobrante(s) en "
                        f"destino (todas las columnas, no solo A/B/C/L).",
                        "WARN")

                # -- Ajustar Tabla1: rango fila 7 (cabecera) hasta última fila con datos --
                # -- y aplicar formato uniforme a todas las filas de datos (fila 8 en adelante) --
                try:
                    from openpyxl.styles import Font, Alignment, Border, Side

                    # ── 1. Redimensionar Tabla1 ────────────────────────────────
                    _fila_cab   = 7                    # cabecera siempre en fila 7
                    _fila_datos = fila_fin_nuevo       # última fila con datos reales

                    if "Tabla1" in ws_dest.tables:
                        tbl = ws_dest.tables["Tabla1"]
                        # Obtener número de columnas actual del rango de la tabla
                        import re as _re_tbl
                        _m = _re_tbl.match(r"([A-Z]+)\d+:([A-Z]+)\d+", tbl.ref)
                        if _m:
                            _col_ini_ltr = _m.group(1)
                            _col_fin_ltr = _m.group(2)
                        else:
                            _col_ini_ltr = "A"
                            _col_fin_ltr = "L"   # fallback: 12 columnas (A..L)
                        _nuevo_ref = (
                            f"{_col_ini_ltr}{_fila_cab}:"
                            f"{_col_fin_ltr}{_fila_datos}"
                        )
                        tbl.ref = _nuevo_ref
                        # El AutoFilter interno de la tabla no se actualiza solo al
                        # cambiar tbl.ref; si se deja con el rango antiguo, Excel
                        # sigue "recordando" filas que ya no existen en la tabla.
                        if tbl.autoFilter is not None:
                            tbl.autoFilter.ref = _nuevo_ref
                        self.log_consola(
                            f"  └ Tabla1 redimensionada → {_nuevo_ref}", "OK")
                    else:
                        self.log_consola(
                            "  ⚠  'Tabla1' no encontrada en el destino; "
                            "se omite el ajuste de tabla.", "WARN")
                        _col_fin_ltr = "L"   # fallback para el formato

                    # ── 2. Aplicar formato uniforme a filas de datos (fila 8 → última) ──
                    # Obtener el número de columna final de la tabla para acotar el rango
                    from openpyxl.utils import column_index_from_string
                    try:
                        _n_col_fin = column_index_from_string(_col_fin_ltr)
                    except Exception:
                        _n_col_fin = 12   # fallback

                    _linea_fmt  = Side(border_style="thin", color="000000")
                    _borde_fmt  = Border(
                        left   = _linea_fmt, right  = _linea_fmt,
                        top    = _linea_fmt, bottom = _linea_fmt,
                    )
                    _font_dato  = Font(name="Arial", size=11)
                    _alin_dato  = Alignment(horizontal="left", vertical="center",
                                            wrap_text=False)

                    _filas_fmt = 0
                    for _r in ws_dest.iter_rows(
                            min_row=FILA_INI, max_row=_fila_datos,
                            min_col=1, max_col=_n_col_fin):
                        for _c in _r:
                            _c.font      = _font_dato
                            _c.alignment = _alin_dato
                            _c.border    = _borde_fmt
                        _filas_fmt += 1

                    self.log_consola(
                        f"  └ Formato aplicado a {_filas_fmt} fila(s) "
                        f"(filas {FILA_INI}–{_fila_datos})", "OK")

                except Exception as _e_tbl:
                    self.log_consola(
                        f"  ⚠  No se pudo ajustar Tabla1 / formato: {_e_tbl}", "WARN")

                # ── 3. Ampliar el formato condicional de D (duplicados) y O ─────────
                #    (NO activos) a toda la tabla. Ambas reglas están fijadas a un
                #    rango concreto en la plantilla original (ej. D8:D30070) y no se
                #    extienden solas cuando la tabla crece — así que con el tiempo
                #    dejan de cubrir las filas nuevas. Se reutilizan las mismas
                #    reglas (mismo color/estilo) sobre el rango actualizado.
                try:
                    _cf = ws_dest.conditional_formatting
                    _rango_d = f"D{FILA_INI}:D{fila_fin_nuevo}"
                    _rango_o = f"O{FILA_INI}:O{fila_fin_nuevo}"
                    _cf_ampliados = []
                    for _cf_range in list(_cf._cf_rules.keys()):
                        _sqref = str(_cf_range.sqref)
                        if _sqref.startswith(f"D{FILA_INI}:D") and _sqref != _rango_d:
                            _reglas = _cf._cf_rules[_cf_range]
                            del _cf._cf_rules[_cf_range]
                            for _regla in _reglas:
                                _cf.add(_rango_d, _regla)
                            _cf_ampliados.append(f"D ({_sqref} → {_rango_d})")
                        elif _sqref.startswith(f"O{FILA_INI}:O") and _sqref != _rango_o:
                            _reglas = _cf._cf_rules[_cf_range]
                            del _cf._cf_rules[_cf_range]
                            for _regla in _reglas:
                                _cf.add(_rango_o, _regla)
                            _cf_ampliados.append(f"O ({_sqref} → {_rango_o})")
                    if _cf_ampliados:
                        self.log_consola(
                            f"  └ Formato condicional ampliado: {', '.join(_cf_ampliados)}",
                            "OK")
                except Exception as _e_cf:
                    self.log_consola(
                        f"  ⚠  No se pudo ampliar el formato condicional: {_e_cf}", "WARN")

                # -- Actualizar celdas de encabezado con fecha y hora de origen --
                try:
                    ws_dest["C4"] = valor_fecha_c4
                    ws_dest["C5"] = valor_hora_c5
                    self.log_consola(f"  └ Celda C4 actualizada: {valor_fecha_c4}", "OK")
                    self.log_consola(f"  └ Celda C5 actualizada: {valor_hora_c5}", "OK")
                except Exception as e_cab:
                    self.log_consola(f"  ⚠  No se pudo actualizar C4/C5: {e_cab}", "WARN")

                self._proteger_hoja_destino(ws_dest, os.path.basename(ruta_destino))

                wb_dest.save(ruta_destino)
                wb_dest.close()

                self.log_consola(f"Destino guardado: {os.path.basename(ruta_destino)}", "OK")
                self.log_consola(f"  └ {n_filas} artículos actualizados (A, B, C, O)", "OK")
                self.log_consola(f"  └ {texto_fecha}", "INFO")
                self.consola.separador("FIN ACTUALIZACIÓN DALI")

                try:
                    self._regenerar_catalogo_dali_completo()
                except Exception as _e_cat:
                    self.log_consola(
                        f"  ⚠  No se pudo regenerar catalogo-dali-completo.xlsx: {_e_cat}", "WARN")

                self.after(0, lambda: self._barra.completado(
                    f"DALI actualizado — {n_filas} artículos ✔"))

                self.after(0, lambda n=n_filas: messagebox.showinfo(
                    "Actualizacion completada",
                    f"\u2714  Listado DALI actualizado correctamente\n\n"
                    f"   Art\u00edculos DALI actualizados : {n}\n"
                    f"   {texto_fecha}\n\n"
                    f"\U0001f4c1  {ruta_destino}"))

            except Exception as e:
                self.log_consola(f"ERROR: {e}", "ERROR")
                self.consola.separador("FIN ACTUALIZACIÓN DALI (con error)")
                self.after(0, lambda: self._barra.reset("Error en actualización DALI"))
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error en actualización DALI",
                    f"\u2718  El proceso no se completó correctamente.\n\n"
                    f"Detalle:\n{err}\n\n"
                    "Revise la ventana de log para más información."))

        threading.Thread(target=proceso, daemon=True).start()

    def _leer_productos_sap(self, ruta_productos):
        """
        LECTURA 1 (extraida de _actualizar_listado_sap -> proceso()):
        lee Productos.xlsx y construye tabla_productos, clave=codigo SAP
        (str), valor=(desc_prod, texto_um, unidad_base, grp_prod, desc_grp,
        tipo_producto).

        Mapeo de columnas Productos.xlsx -> destino (actualizado 2026-08:
        se insertaron 3 columnas nuevas F, K, L en LISTADO CODIGOS DALI-SAP,
        desplazando las que había):
            C (Descripción producto)              -> E  (DESCRIPCION SAP)
            J (Texto UM expl.)                     -> F  (nueva)
            I (Unidad medida base)                 -> G  (UNIDAD SAP)
            F (Grupo de productos)                 -> H  (GRUPO PRODUCTOS SAP)
            G (Descripción de grupo de productos)  -> I  (Desc. grupo productos SAP)
            D (Tipo de producto)                   -> L  (nueva)

        Devuelve (tabla_productos, filas_prod) o (None, 0) si el archivo
        no tiene filas validas -- en ese caso ya deja la barra reseteada
        y el aviso mostrado al usuario, igual que hacia el bloque inline.
        filas_prod se devuelve por separado (no basta con len(tabla_productos))
        porque si Productos.xlsx trae códigos SAP duplicados, el dict los
        deduplica pero el recuento de filas real debe conservarse tal cual
        -- se usa más adelante en la celda H5 y en el resumen final.
        """
        import openpyxl
        # ── LECTURA 1: Productos.xlsx ─────────────────────────────────
        # Clave cruce: col B (índice 1) → Código producto SAP
        self.log_consola("Leyendo Productos.xlsx...", "INFO")
        wb_prod = openpyxl.load_workbook(ruta_productos, read_only=True, data_only=True)
        ws_prod = wb_prod.active

        tabla_productos = {}
        filas_prod = 0
        for fila in ws_prod.iter_rows(min_row=2, values_only=True):
            if all(v is None for v in fila):
                continue
            clave = fila[1] if len(fila) > 1 else None
            if clave is None or str(clave).strip() == "":
                continue
            clave = str(clave).strip()
            tipo_producto = fila[3] if len(fila) > 3 else None   # D
            desc_prod     = fila[2] if len(fila) > 2 else None   # C
            grp_prod      = fila[5] if len(fila) > 5 else None   # F
            desc_grp      = fila[6] if len(fila) > 6 else None   # G
            unidad_base   = fila[8] if len(fila) > 8 else None   # I
            texto_um      = fila[9] if len(fila) > 9 else None   # J
            tabla_productos[clave] = (desc_prod, texto_um, unidad_base,
                                       grp_prod, desc_grp, tipo_producto)
            filas_prod += 1

        wb_prod.close()
        self.log_consola(f"  └ {filas_prod} productos cargados (Productos.xlsx)", "OK")

        if filas_prod == 0:
            self.after(0, lambda: self._barra.reset("Sin datos en Productos.xlsx"))
            self.after(0, lambda: messagebox.showwarning(
                "Sin datos",
                "No se encontraron filas con datos\n"
                "en el archivo Productos.xlsx."))
            return None, 0

        return tabla_productos, filas_prod

    def _leer_categoria_valoracion(self, ruta_listas):
        """
        LECTURA (nueva, 2026-08): lee "Listas de productos.xlsx" y
        construye tabla_categoria, clave=código SAP (str), valor=
        Categoría valoración -> destino columna K (nueva).

        Clave cruce: col A (índice 0) → Código SAP.
        Valor: col F (índice 5) → Categoría valoración.

        El archivo trae una fila por combinación código+centro, así que
        un mismo código puede repetirse con el mismo valor (normal) o,
        en casos raros, con valores DISTINTOS según el centro. Para esos
        casos se toma el valor más frecuente (moda), asumiendo que los
        demás son errores de introducción, y se devuelve la lista de
        conflictos para avisar al usuario.

        Devuelve (tabla_categoria, conflictos), donde conflictos es una
        lista de (clave, {valor: nº_apariciones}) para los códigos con
        más de un valor distinto.
        """
        import openpyxl
        from collections import Counter

        self.log_consola("Leyendo Listas de productos.xlsx...", "INFO")
        wb_listas = openpyxl.load_workbook(ruta_listas, read_only=True, data_only=True)
        ws_listas = wb_listas.active

        contadores = {}   # clave -> Counter({valor: n})
        filas_listas = 0
        for fila in ws_listas.iter_rows(min_row=2, values_only=True):
            if all(v is None for v in fila):
                continue
            clave = fila[0] if len(fila) > 0 else None
            if clave is None or str(clave).strip() == "":
                continue
            clave = str(clave).strip()
            valor = fila[5] if len(fila) > 5 else None   # F
            if valor is None or str(valor).strip() == "":
                continue
            valor = str(valor).strip()
            contadores.setdefault(clave, Counter())[valor] += 1
            filas_listas += 1
        wb_listas.close()

        tabla_categoria = {}
        conflictos = []
        for clave, contador in contadores.items():
            if len(contador) > 1:
                conflictos.append((clave, dict(contador)))
            # most_common()[0] = (valor, n) más frecuente; en empate,
            # Counter conserva el orden de primera aparición.
            tabla_categoria[clave] = contador.most_common(1)[0][0]

        self.log_consola(
            f"  └ {filas_listas} fila(s) leídas, {len(tabla_categoria)} "
            f"código(s) SAP con Categoría valoración (Listas de productos.xlsx)", "OK")
        if conflictos:
            self.log_consola(
                f"  ⚠  {len(conflictos)} código(s) con Categoría valoración "
                f"distinta según el centro — se tomó el valor más frecuente:", "WARN")
            for clave, valores in conflictos:
                self.log_consola(f"      · {clave}: {valores}", "WARN")

        return tabla_categoria, conflictos

    def _leer_articulos_dali_sap(self, ruta_dali):
        """
        LECTURA 2 (extraida de _actualizar_listado_sap -> proceso()):
        lee ARTICULOS_DALI.xlsx y construye tabla_dali, clave=codigo DALI
        normalizado (str), valor=(subfamilia, igic).
        """
        import openpyxl
        # ── LECTURA 2: ARTICULOS_DALI.xlsx ───────────────────────────
        # Clave cruce: col A (índice 0) → Código DALI
        # Col E (índice 4) → SUBFAMILIA  → lookup en TABLA_SUBFAMILIA
        # Col G (índice 6)→ IGIC        → lookup en TABLA_IGIC
        # (Datos desde fila 8, igual que el resto del sistema)
        self.after(0, lambda: self._barra.indeterminate("Leyendo ARTICULOS_DALI..."))
        self.log_consola("Leyendo ARTICULOS_DALI.xlsx...", "INFO")
        wb_dali = openpyxl.load_workbook(ruta_dali, read_only=True, data_only=True)
        ws_dali = wb_dali.active

        # tabla_dali: clave = CODIGO DALI (col A), valor = (subfamilia, igic)
        # Estructura de ARTICULOS_DALI.xlsx (generado por Acción 1):
        #   A(0)=CODIGO  B(1)=DESCRIPCION  C(2)=UNIDAD  D(3)=FAMILIA
        #   E(4)=SUBFAMILIA  F(5)=NATURALEZA  G(6)=IGIC  H(7)=PROVEEDOR  I(8)=ACTIVO
        #
        # Cruce I   : A destino <-> A ARTICULOS_DALI -> leer G(IGIC)       -> col I destino
        # Cruce J/K : A destino <-> A ARTICULOS_DALI -> leer E(SUBFAMILIA) -> cols J+K destino
        tabla_dali = {}   # { cod_dali_norm: (subfamilia, igic) }
        filas_dali = 0
        FILA_INI_DALI = 8
        for fila in ws_dali.iter_rows(min_row=FILA_INI_DALI, values_only=True):
            if all(v is None for v in fila):
                continue
            cod_raw = fila[0] if len(fila) > 0 else None   # col A — CODIGO
            if cod_raw is None or str(cod_raw).strip() == "":
                continue
            # Normalizar clave: ARTICULOS_DALI almacena CODIGO como entero;
            # el destino puede tenerlo como "001234" (con ceros) -> unificar a "1234"
            try:
                cod_norm = str(int(float(str(cod_raw).strip())))
            except (ValueError, TypeError):
                cod_norm = str(cod_raw).strip()

            subfamilia = str(fila[4]).strip() if len(fila) > 4 and fila[4] is not None else ""  # col E — SUBFAMILIA
            igic       = str(fila[6]).strip() if len(fila) > 6 and fila[6] is not None else ""  # col G — IGIC
            tabla_dali[cod_norm] = (subfamilia, igic)
            filas_dali += 1

        wb_dali.close()
        self.log_consola(f"  └ {filas_dali} artículos cargados (ARTICULOS_DALI.xlsx)", "OK")

        return tabla_dali

    def _auditar_sap_duplicados_listado(self, ws):
        """
        Escanea TODO el LISTADO CODIGOS DALI - SAP.xlsx (no solo lo que
        se esté tecleando en ese momento) buscando códigos SAP asignados
        a más de un código DALI distinto -- el problema real reportado
        por el usuario no es un DALI concreto, sino que dos DALI
        diferentes compartan el mismo SAP en cualquier parte del listado
        (caso real: DALI 15660 y 26046, ambos con SAP 10101038).

        Se llama desde _cruzar_y_escribir_listado_sap() (Acción 3), justo
        antes de proteger/guardar, con la hoja ya actualizada y aún
        editable. Si encuentra duplicados, reutiliza la misma ventana de
        resolución que VentanaSAPFaltantes (VentanaResolverSAPDuplicados)
        -- mostrada de forma thread-safe, ya que _actualizar_listado_sap
        corre en el hilo de fondo de su propio proceso().

        Los códigos marcados por el usuario para borrar se limpian
        directamente en `ws` (col D = None); el guardado posterior de
        _cruzar_y_escribir_listado_sap ya incluye este cambio.

        Devuelve el número de códigos SAP limpiados (0 si no había
        ninguno, o si el usuario canceló la resolución).
        """
        from collections import defaultdict
        _por_sap = defaultdict(list)   # {sap: [(dali, fila), ...]}
        for idx, row in enumerate(
                ws.iter_rows(min_row=8, min_col=1, max_col=4, values_only=True),
                start=8):
            v_dali = row[0]
            v_sap  = row[3]
            if v_dali is None or v_sap in (None, ""):
                continue
            try:
                dali_int = int(float(v_dali))
                sap_int  = int(float(v_sap))
            except (ValueError, TypeError):
                continue
            _por_sap[sap_int].append((dali_int, idx))

        conflictos = {
            cod_sap: [(dali, True, fila, "") for dali, fila in entradas]
            for cod_sap, entradas in _por_sap.items()
            if len(entradas) > 1
        }
        if not conflictos:
            return 0

        self.log_consola(
            f"  ⚠ {len(conflictos)} código(s) SAP duplicados detectados en "
            f"el LISTADO completo -- pidiendo resolución al usuario...",
            "WARN")

        def _mostrar():
            dlg = VentanaResolverSAPDuplicados(self, conflictos)
            self.wait_window(dlg)
            return dlg.resultado

        a_borrar = self._ejecutar_en_hilo_principal_y_esperar(_mostrar)

        if not a_borrar:
            self.log_consola(
                "  ⚠ Auditoría de duplicados del LISTADO cancelada o sin "
                "cambios -- quedan sin resolver.", "WARN")
            return 0

        limpiados = 0
        for cod_sap, candidatos in conflictos.items():
            for cod_dali, _es, fila, _desc in candidatos:
                if cod_dali in a_borrar:
                    ws.cell(fila, 4).value = None
                    limpiados += 1
        self.log_consola(
            f"  ✔ {limpiados} código(s) SAP duplicados limpiados del "
            f"LISTADO.", "OK")
        return limpiados

    def _cruzar_y_escribir_listado_sap(self, ruta_destino, tabla_productos,
                                        tabla_dali, tabla_categoria, conflictos_categoria,
                                        filas_prod,
                                        fecha_productos_solo, hora_productos_solo):
        """
        CRUCE Y ESCRITURA EN DESTINO + Resumen en consola (extraido de
        _actualizar_listado_sap -> proceso()). Es literalmente el resto
        del cuerpo del try de proceso() tras las lecturas: abre el
        destino, limpia columnas E:N, cruza Productos.xlsx,
        ARTICULOS_DALI.xlsx y Listas de productos.xlsx (Cruce 1, Cruce 1b
        y Cruce 2), actualiza cabecera y contadores H1-H3/H5, protege y
        guarda, y muestra el resumen final en consola + mensaje al usuario.

        Mapeo de columnas destino (actualizado 2026-08, tras insertar las
        3 columnas nuevas F, K, L en LISTADO CODIGOS DALI-SAP -- las que
        ya existían antes se desplazaron: antigua F/G/H/I -> G/H/I/J,
        antigua J/K/L -> M/N/O):
            E, F, G, H, I, L  <- Productos.xlsx (Cruce 1, ver _leer_productos_sap)
            K                 <- Listas de productos.xlsx (Cruce 1b, Categoría valoración)
            J                 <- IGIC (Cruce 2, vía ARTICULOS_DALI + TABLA_IGIC_GRUPO_COMPRAS)
            M, N              <- Grupo de Compras (Cruce 2, vía ARTICULOS_DALI + TABLA_SUBFAMILIA_GRUPO_COMPRAS)
            O                 <- ACTIVO DALI (no se escribe aquí; solo se lee para H1/H2/H3)

        No devuelve nada -- es el ultimo tramo del try de proceso(),
        nada se usa despues (solo el except que envuelve a proceso()).
        """
        import openpyxl
        # ── CRUCE Y ESCRITURA EN DESTINO ─────────────────────────────
        self.log_consola("Abriendo destino: LISTADO CODIGOS DALI - SAP.xlsx", "INFO")
        self.after(0, lambda: self._barra.indeterminate("Cruzando y escribiendo..."))

        # Pre-normalizar tablas de lookup a uppercase para acceso O(1).
        # Se hace una sola vez aquí en lugar de comparar k.upper() en cada fila.
        TABLA_SUBFAMILIA_U = {k.upper(): v for k, v in TABLA_SUBFAMILIA_GRUPO_COMPRAS.items()}
        TABLA_IGIC_U       = {k.upper(): v for k, v in TABLA_IGIC_GRUPO_COMPRAS.items()}

        wb_dest = openpyxl.load_workbook(ruta_destino)
        ws_dest = wb_dest.active

        # ── LIMPIEZA PREVIA: borrar columnas E:N desde fila 8 hasta la última con datos ──
        # Se eliminan los valores de E(5) a N(14) antes de rellenarlos con los
        # datos cruzados actualizados, para evitar valores residuales de
        # ejecuciones anteriores. No se toca O (15, ACTIVO DALI): esa columna
        # no se escribe en este proceso, solo se lee para H1/H2/H3.
        self.log_consola("Limpiando columnas E:N (fila 8 → última fila)...", "INFO")
        self.after(0, lambda: self._barra.indeterminate("Limpiando columnas E:N..."))
        _ultima_fila_dest = ws_dest.max_row
        _COLS_LIMPIAR = (5, 6, 7, 8, 9, 10, 11, 12, 13, 14)   # E, F, G, H, I, J, K, L, M, N
        _celdas_borradas = 0
        for _fila_lim in range(8, _ultima_fila_dest + 1):
            for _col_lim in _COLS_LIMPIAR:
                _celda = ws_dest.cell(row=_fila_lim, column=_col_lim)
                if _celda.value is not None:
                    _celda.value = None
                    _celdas_borradas += 1
        self.log_consola(
            f"  └ {_celdas_borradas} celda(s) borradas en E:N "
            f"(filas 8–{_ultima_fila_dest})", "OK")
        # ─────────────────────────────────────────────────────────────────

        act_productos       = 0   # filas con match en Productos.xlsx (E,F,G,H,I,L)
        sin_productos       = 0
        lista_sin_productos = []  # códigos SAP sin match (para log detallado)
        act_categoria       = 0   # filas con match en Listas de productos.xlsx (K)
        sin_categoria       = 0
        lista_sin_categoria = []  # códigos SAP sin match en Listas de productos
        act_dali            = 0   # filas con match en ARTICULOS_DALI (J, M, N)
        sin_dali            = 0
        lista_sin_dali      = []  # códigos DALI sin match (para log detallado)
        sin_subfam          = 0   # naturalezas no encontradas en el diccionario
        dict_sin_subfam     = {} # {naturaleza: [(cod_dali, cod_sap)]} para log agrupado
        sin_igic            = 0   # IGIC no encontrado en el diccionario
        # Contadores para celdas de resumen H1–H3
        h1 = 0   # D con valor  y L = "SI"
        h2 = 0   # D con valor  y L = "NO"
        h3 = 0   # D sin valor  y L = "SI"

        FILA_INI = 8
        for fila_num in range(FILA_INI, ws_dest.max_row + 1):

            # ── Cruce 1: col D destino ↔ col B Productos ──────────────
            cod_sap = ws_dest.cell(row=fila_num, column=4).value   # col D
            if cod_sap is not None and str(cod_sap).strip() != "":
                clave_sap = str(cod_sap).strip()
                if clave_sap in tabla_productos:
                    (desc_prod, texto_um, unidad_base,
                     grp_prod, desc_grp, tipo_producto) = tabla_productos[clave_sap]
                    ws_dest.cell(row=fila_num, column=5,  value=desc_prod)      # E
                    ws_dest.cell(row=fila_num, column=6,  value=texto_um)       # F
                    ws_dest.cell(row=fila_num, column=7,  value=unidad_base)    # G
                    ws_dest.cell(row=fila_num, column=8,  value=grp_prod)       # H
                    ws_dest.cell(row=fila_num, column=9,  value=desc_grp)       # I
                    ws_dest.cell(row=fila_num, column=12, value=tipo_producto)  # L
                    act_productos += 1
                else:
                    sin_productos += 1
                    lista_sin_productos.append(clave_sap)

                # ── Cruce 1b: col D destino ↔ col A Listas de productos ──
                if clave_sap in tabla_categoria:
                    ws_dest.cell(row=fila_num, column=11,
                                 value=tabla_categoria[clave_sap])   # K
                    act_categoria += 1
                else:
                    sin_categoria += 1
                    lista_sin_categoria.append(clave_sap)

            # ── Cruce 2: col A destino <-> col A ARTICULOS_DALI ─────────
            # Rellena col J (IGIC) y cols M+N (GRUPO COMPRAS) del destino.
            # La clave se normaliza igual que en tabla_dali (int -> str sin ceros).
            cod_dali = ws_dest.cell(row=fila_num, column=1).value   # col A destino
            if cod_dali is None or str(cod_dali).strip() == "":
                continue   # fila sin codigo DALI, pasar a la siguiente

            try:
                clave_dali = str(int(float(str(cod_dali).strip())))
            except (ValueError, TypeError):
                clave_dali = str(cod_dali).strip()

            if clave_dali not in tabla_dali:
                sin_dali += 1
                lista_sin_dali.append(clave_dali)
                continue

            subfamilia_raw, igic_raw = tabla_dali[clave_dali]
            act_dali += 1

            # ── Condición: sólo rellenar I/J/K si col D tiene valor ──────
            # El Cruce 2 escribe datos SAP derivados (IGIC, Grupo Compras)
            # únicamente en filas que ya disponen de código SAP en col D.
            val_d_cruce2 = ws_dest.cell(row=fila_num, column=4).value
            if val_d_cruce2 is None or str(val_d_cruce2).strip() == "":
                # Sin código SAP: no se escriben I, J ni K
                pass
            else:
                # col G de ARTICULOS_DALI (IGIC) -> col J destino
                igic_match = TABLA_IGIC_U.get(igic_raw.upper().strip())
                if igic_match:
                    ws_dest.cell(row=fila_num, column=10, value=igic_match[0])
                else:
                    sin_igic += 1

                # col E de ARTICULOS_DALI (SUBFAMILIA) -> cols M y N destino
                subfam_match = TABLA_SUBFAMILIA_U.get(subfamilia_raw.upper().strip())
                if subfam_match:
                    ws_dest.cell(row=fila_num, column=13, value=subfam_match[0])  # M
                    ws_dest.cell(row=fila_num, column=14, value=subfam_match[1])  # N
                else:
                    sin_subfam += 1
                    if subfamilia_raw:
                        cod_sap_str = str(val_d_cruce2).strip()
                        if subfamilia_raw not in dict_sin_subfam:
                            dict_sin_subfam[subfamilia_raw] = []
                        dict_sin_subfam[subfamilia_raw].append((clave_dali, cod_sap_str))

            # ── Contadores resumen I1/I2/I3/I5 ───────────────────────────
            # Se evalúan aquí, tras escribir, para usar los valores finales de D y O
            val_d = ws_dest.cell(row=fila_num, column=4).value    # col D (SAP)
            val_o = ws_dest.cell(row=fila_num, column=15).value   # col O (ACTIVO DALI)
            d_tiene_valor = val_d is not None and str(val_d).strip() != ""
            o_es_si  = str(val_o).strip().upper() == "SI" if val_o is not None else False
            o_es_no  = str(val_o).strip().upper() == "NO" if val_o is not None else False
            if d_tiene_valor and o_es_si:
                h1 += 1
            elif d_tiene_valor and o_es_no:
                h2 += 1
            elif not d_tiene_valor and o_es_si:
                h3 += 1
        # ═══════════════════════════════════════════════════════════════
        #  ACTUALIZAR CELDAS DE CABECERA
        # ═══════════════════════════════════════════════════════════════
        try:
            ws_dest["D4"] = fecha_productos_solo
            ws_dest["D5"] = hora_productos_solo
            ws_dest["I1"] = h1          # D con valor  y O = "SI"
            ws_dest["I2"] = h2          # D con valor  y O = "NO"
            ws_dest["I3"] = h3          # D sin valor  y O = "SI"
            ws_dest["I5"] = filas_prod  # total filas con valor en col B de Productos.xlsx
            self.log_consola(f"  └ Celda D4 actualizada: {fecha_productos_solo}", "OK")
            self.log_consola(f"  └ Celda D5 actualizada: {hora_productos_solo}", "OK")
            self.log_consola(f"  └ I1={h1}  I2={h2}  I3={h3}  I5={filas_prod}", "OK")
        except Exception as e_d4:
            self.log_consola(f"  ⚠  No se pudo actualizar celdas de cabecera: {e_d4}", "WARN")

        # ── Auditoría completa de códigos SAP duplicados en el LISTADO ──
        # (pedida por el usuario tras el caso real de DALI 15660/26046
        # compartiendo el mismo SAP) -- a diferencia del control de
        # VentanaSAPFaltantes (que solo mira los códigos que se están
        # tecleando en ese momento), esto revisa TODO el LISTADO ya
        # actualizado, buscando cualquier SAP asignado a más de un DALI.
        self._auditar_sap_duplicados_listado(ws_dest)

        self._proteger_hoja_destino(ws_dest, os.path.basename(ruta_destino))

        try:
            wb_dest.save(ruta_destino)
            wb_dest.close()
        except PermissionError:
            wb_dest.close()
            self.log_consola(
                "ERROR: No se pudo guardar — el archivo está abierto en Excel.", "ERROR")
            self.after(0, lambda: self._barra.reset("Error al guardar"))
            self.after(0, lambda: messagebox.showerror(
                "Archivo bloqueado por Excel",
                "❌  No se pudo guardar el archivo destino:\n\n"
                "   LISTADO CODIGOS DALI - SAP.xlsx\n\n"
                "El archivo está abierto en Excel.\n\n"
                "📋  Cierre el archivo en Excel y vuelva a ejecutar\n"
                "    la Acción 3."))
            return

        # ── Resumen en consola ────────────────────────────────────────
        # Constante: máximo de entradas detalle a mostrar por categoría.
        # El resto se resume en "… y N más" para evitar miles de repaints.
        _CAP = 30

        self.log_consola(f"Destino guardado: LISTADO CODIGOS DALI - SAP.xlsx", "OK")
        self.log_consola(f"  └ {act_productos} materiales actualizados desde Productos.xlsx", "OK")
        self.log_consola(f"  └ {act_categoria} materiales con Categoría valoración desde Listas de productos.xlsx", "OK")
        self.log_consola(f"  └ {act_dali} artículos cruzados con ARTICULOS_DALI", "OK")

        if lista_sin_categoria:
            lineas_batch = [
                (f"  └ ⚠  {sin_categoria} código(s) SAP sin coincidencia en Listas de productos.xlsx:", "WARN")
            ]
            for cod in lista_sin_categoria[:_CAP]:
                lineas_batch.append((f"      · {cod}", "WARN"))
            if len(lista_sin_categoria) > _CAP:
                lineas_batch.append(
                    (f"      … y {len(lista_sin_categoria) - _CAP} más (omitidos del log)", "WARN"))
            self.log_consola_batch(lineas_batch)

        if lista_sin_productos:
            lineas_batch = [
                (f"  └ ⚠  {sin_productos} código(s) SAP sin coincidencia en Productos.xlsx:", "WARN")
            ]
            for cod in lista_sin_productos[:_CAP]:
                lineas_batch.append((f"      · {cod}", "WARN"))
            if len(lista_sin_productos) > _CAP:
                lineas_batch.append(
                    (f"      … y {len(lista_sin_productos) - _CAP} más (omitidos del log)", "WARN"))
            self.log_consola_batch(lineas_batch)

        if lista_sin_dali:
            lineas_batch = [
                (f"  └ ⚠  {sin_dali} código(s) DALI sin coincidencia en ARTICULOS_DALI:", "WARN")
            ]
            for cod in lista_sin_dali[:_CAP]:
                lineas_batch.append((f"      · {cod}", "WARN"))
            if len(lista_sin_dali) > _CAP:
                lineas_batch.append(
                    (f"      … y {len(lista_sin_dali) - _CAP} más (omitidos del log)", "WARN"))
            self.log_consola_batch(lineas_batch)

        if sin_subfam:
            lineas_batch = [
                (f"  └ ⚠  {sin_subfam} artículo(s) sin mapeo de SUBFAMILIA "
                 f"({len(dict_sin_subfam)} subfamilia(s) distintas):", "WARN")
            ]
            total_mostrados = 0
            for nat, pares in dict_sin_subfam.items():
                if total_mostrados >= _CAP:
                    restantes = sum(len(v) for k, v in dict_sin_subfam.items()
                                    if k != nat) + len(pares)
                    lineas_batch.append(
                        (f"      … y {restantes} artículo(s) más en otras subfamilias (omitidos)", "WARN"))
                    break
                lineas_batch.append(
                    (f"      SUBFAMILIA '{nat}'  ({len(pares)} artículo(s)):", "WARN"))
                for cod_dali, cod_sap in pares:
                    if total_mostrados >= _CAP:
                        lineas_batch.append(
                            (f"          … y {len(pares) - (total_mostrados - 1)} más en esta subfamilia", "WARN"))
                        break
                    lineas_batch.append(
                        (f"          DALI: {cod_dali}   SAP: {cod_sap}", "WARN"))
                    total_mostrados += 1
            self.log_consola_batch(lineas_batch)

        if sin_igic:
            self.log_consola(f"  └ ⚠  {sin_igic} IGIC(s) sin mapeo en diccionario — revise log", "WARN")

        self.consola.separador("FIN ACTUALIZACIÓN SAP")
        self.after(0, lambda: self._barra.completado(
            f"SAP actualizado — {act_dali} filas ✔"))

        def _mostrar_resultado(ap=act_productos, ad=act_dali,
                               sp=sin_productos, sd=sin_dali,
                               ssf=sin_subfam, si=sin_igic):
            hay_warn = sp or sd or ssf or si
            lineas = [
                "\u2714  Actualizaci\u00f3n SAP completada correctamente",
                "",
                f"   Art\u00edculos SAP actualizados  : {ap}",
                f"   Art\u00edculos DALI actualizados : {ad}",
            ]
            if sp:
                lineas.append(f"   \u26a0 C\u00f3digos SAP sin coincidencia  : {sp}")
            if sd:
                lineas.append(f"   \u26a0 C\u00f3digos DALI sin coincidencia : {sd}")
            if ssf:
                lineas.append(f"   \u26a0 Subfamilia(s) sin mapeo       : {ssf}")
            if si:
                lineas.append(f"   \u26a0 IGIC(s) sin mapeo            : {si}")
            if hay_warn:
                lineas += [
                    "",
                    "Revise el log de consola para ver el detalle.",
                ]
            lineas += ["", f"\U0001f4c1  {ruta_destino}"]
            titulo = "SAP \u2014 Actualizaci\u00f3n completada" + (" con advertencias" if hay_warn else "")
            if hay_warn:
                messagebox.showwarning(titulo, "\n".join(lineas))
            else:
                messagebox.showinfo(titulo, "\n".join(lineas))

        try:
            self._regenerar_catalogo_dali_completo()
        except Exception as _e_cat:
            self.log_consola(
                f"  ⚠  No se pudo regenerar catalogo-dali-completo.xlsx: {_e_cat}", "WARN")

        if conflictos_categoria:
            def _mostrar_conflictos_categoria(conf=conflictos_categoria):
                lineas = [
                    "\u26a0  Se han encontrado códigos SAP con Categoría",
                    "   valoración distinta según el centro en",
                    "   \"Listas de productos.xlsx\".",
                    "",
                    "   Se ha usado el valor más frecuente en cada caso,",
                    "   asumiendo que el resto son errores. Revise y",
                    "   corrija estos códigos en el origen si procede:",
                    "",
                ]
                for clave, valores in conf[:20]:
                    detalle = ", ".join(f"{v}×{n}" for v, n in valores.items())
                    lineas.append(f"   • {clave}:  {detalle}")
                if len(conf) > 20:
                    lineas.append(f"   … y {len(conf) - 20} código(s) más (ver log de consola).")
                messagebox.showwarning(
                    "Categoría valoración ambigua — revisar origen", "\n".join(lineas))
            self.after(0, _mostrar_conflictos_categoria)

        self.after(0, _mostrar_resultado)


    def _obtener_credenciales_dali_app(self):
        """
        Devuelve (email, password) para la app web «Catálogo Asignaciones».
        Si ya están guardadas en tarifa_config.json (email y contraseña
        CIFRADA con DPAPI de Windows, ver credenciales.py), las devuelve
        directamente sin preguntar nada. Si no, abre un diálogo modal para
        pedirlas una vez. (nov 2026: antes se guardaban en texto plano y el
        archivo llegó a subirse a GitHub; ahora nunca se escriben en claro.)

        Devuelve (None, None) si el usuario cancela el diálogo.
        """
        email_guardado, pass_guardado = leer_credenciales_dali(self.config_data)
        if email_guardado and pass_guardado:
            return email_guardado, pass_guardado

        resultado = {"email": None, "password": None}

        win = tk.Toplevel(self)
        win.title("Catálogo Asignaciones — credenciales")
        win.configure(bg=C_DARK)
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        frm = tk.Frame(win, bg=C_DARK, padx=24, pady=20)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text="Credenciales de «Catálogo Asignaciones»",
                 bg=C_DARK, fg=C_ACCENT2, font=FONT_SUB).pack(anchor="w")
        tk.Label(
            frm,
            text="Se usan para descargar automáticamente el catálogo\n"
                 "actualizado antes de esta acción. La contraseña se guarda\n"
                 "cifrada: solo tu usuario de Windows, en este equipo,\n"
                 "puede leerla. No se envía a ningún otro sitio.",
            bg=C_DARK, fg=C_MUTED, font=FONT_SMALL, justify="left",
        ).pack(anchor="w", pady=(4, 16))

        tk.Label(frm, text="Email", bg=C_DARK, fg=C_TEXT, font=FONT_LABEL).pack(anchor="w")
        entry_email = tk.Entry(frm, font=FONT_LABEL, width=34)
        entry_email.pack(fill="x", pady=(2, 10))

        tk.Label(frm, text="Contraseña", bg=C_DARK, fg=C_TEXT, font=FONT_LABEL).pack(anchor="w")
        entry_pass = tk.Entry(frm, font=FONT_LABEL, width=34, show="•")
        entry_pass.pack(fill="x", pady=(2, 16))

        def _guardar():
            email = entry_email.get().strip()
            pwd   = entry_pass.get().strip()
            if not email or not pwd:
                messagebox.showwarning("Datos incompletos",
                                       "Indique email y contraseña.", parent=win)
                return
            resultado["email"] = email
            resultado["password"] = pwd
            win.destroy()

        def _cancelar():
            win.destroy()

        frm_btns = tk.Frame(frm, bg=C_DARK)
        frm_btns.pack(fill="x")
        tk.Button(frm_btns, text="Cancelar", command=_cancelar,
                  font=FONT_LABEL).pack(side="right", padx=(8, 0))
        tk.Button(frm_btns, text="Guardar y continuar", command=_guardar,
                  font=FONT_LABEL, bg=C_ACCENT, fg=C_DARK).pack(side="right")

        entry_email.focus_set()
        win.bind("<Return>", lambda e: _guardar())

        # ── Centrar (mismo patrón que el resto de ventanas del proyecto) ──
        win.update_idletasks()
        _w = win.winfo_reqwidth()
        _h = win.winfo_reqheight()
        _sw = win.winfo_screenwidth()
        _sh = win.winfo_screenheight()
        win.geometry(f"{_w}x{_h}+{(_sw-_w)//2}+{max(60, (_sh-_h)//3)}")

        win.wait_window()

        if resultado["email"] and resultado["password"]:
            if not guardar_credenciales_dali(self.config_data, resultado["email"], resultado["password"]):
                self.log_consola(
                    "No se pudo cifrar la contraseña de Catálogo Asignaciones: se usa solo "
                    "en esta sesión y se volverá a pedir la próxima vez.", "WARN")
            guardar_config(self.config_data)
        return resultado["email"], resultado["password"]

    def _adquirir_lock_catalogo_con_aviso(self):
        """
        Adquiere self._catalogo_lock antes de que cualquiera de los tres
        métodos que escriben catalogo-dali-completo.xlsx empiece a tocar
        el archivo (_descargar_catalogo_dali_web, _regenerar_catalogo_
        dali_completo, _actualizar_codigo_proveedor_catalogo). Si el
        lock ya está ocupado por otra de estas operaciones (p.ej. el
        usuario lanzó Acción 3 mientras el cierre de una transformación
        aún estaba escribiendo el código proveedor de otra tarifa), avisa
        en el log antes de quedarse esperando -- así la segunda operación
        se pone en cola en vez de chocar con la primera y perder datos en
        silencio. No hace falta liberar aquí: cada método libera el lock
        él mismo (try/finally) al terminar.
        """
        if not self._catalogo_lock.acquire(blocking=False):
            self.log_consola(
                "Esperando a que termine otra operación sobre el catálogo "
                "DALI completo antes de continuar...", "INFO")
            self._catalogo_lock.acquire(blocking=True)

    def _descargar_catalogo_dali_web(self, email, password):
        """
        Wrapper con bloqueo global (ver _adquirir_lock_catalogo_con_aviso):
        adquiere self._catalogo_lock antes de tocar catalogo-dali-completo.xlsx
        y lo libera siempre al terminar, delegando en _descargar_catalogo_
        dali_web_impl() para la lógica real (sin tocarla).
        """
        self._adquirir_lock_catalogo_con_aviso()
        try:
            return self._descargar_catalogo_dali_web_impl(email, password)
        finally:
            self._catalogo_lock.release()

    def _descargar_catalogo_dali_web_impl(self, email, password):
        """
        Descarga catalogo-dali-completo.xlsx directamente desde la app
        web «Catálogo Asignaciones» (POST /auth/login + GET
        /admin/export-completo, vía DALI_APP_API_BASE), sobrescribiendo
        RECURSOS/catalogo-dali-completo.xlsx. Autenticación por sesión
        (cookie), no por token -- se usa un CookieJar para mantenerla
        entre las dos peticiones.

        Recibe `email`/`password` ya resueltos (ver
        _obtener_credenciales_dali_app(), que crea una ventana Tkinter y
        por eso debe llamarse SIEMPRE desde el hilo principal, nunca
        desde aquí) -- este método hace solo peticiones de red y E/S de
        fichero, sin tocar Tkinter directamente salvo via self.after(),
        por lo que es seguro llamarlo desde un hilo de fondo (ver
        _descargar_catalogo_dali_web_bloqueante(), que es quien realmente
        lo invoca, en su propio hilo, para no congelar la interfaz
        mientras esperamos respuesta del servidor).

        Devuelve True si la descarga se completó correctamente, False
        en cualquier otro caso (credenciales incorrectas, verificación
        en dos pasos pendiente, sin conexión, permiso insuficiente...)
        -- siempre dejando un mensaje claro en el log y, si procede, una
        ventana de aviso, antes de devolver False.
        """
        import json as _json
        import urllib.request as _urlreq
        import urllib.error as _urlerr
        import http.cookiejar as _cookiejar

        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            return False

        self.log_consola("Descargando catálogo actualizado desde Catálogo Asignaciones...", "INFO")
        self.after(0, lambda: self._barra.indeterminate("Descargando catálogo actualizado..."))

        cookie_jar = _cookiejar.CookieJar()
        opener = _urlreq.build_opener(_urlreq.HTTPCookieProcessor(cookie_jar))
        # Cloudflare (WAF/Bot Fight Mode) puede rechazar con 403 el
        # User-Agent por defecto de urllib ("Python-urllib/3.x"),
        # identificable como tráfico de script -- se sustituye por uno
        # de navegador normal para que la petición pase igual que la
        # de cualquier persona usando la app desde Chrome/Edge/Firefox.
        opener.addheaders = [("User-Agent",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")]

        # ── 1. Login ────────────────────────────────────────────────
        try:
            cuerpo = _json.dumps({"email": email, "password": password}).encode("utf-8")
            peticion = _urlreq.Request(
                f"{DALI_APP_API_BASE}/auth/login", data=cuerpo, method="POST",
                headers={"Content-Type": "application/json"})
            with opener.open(peticion, timeout=20) as resp:
                datos_login = _json.loads(resp.read().decode("utf-8"))
        except _urlerr.HTTPError as e:
            try:
                detalle = _json.loads(e.read().decode("utf-8")).get("error", "")
            except Exception:
                detalle = ""
            if e.code in (401, 403):
                self.log_consola(
                    f"Login en Catálogo Asignaciones rechazado (HTTP {e.code}"
                    f"{': ' + detalle if detalle else ''}) — se han borrado las "
                    f"credenciales guardadas, vuelva a intentarlo.", "ERROR")
                borrar_credenciales_dali(self.config_data)
                guardar_config(self.config_data)
            else:
                self.log_consola(f"Error al iniciar sesión en Catálogo Asignaciones "
                                 f"(HTTP {e.code}): {detalle or e.reason}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False
        except Exception as e:
            self.log_consola(f"No se pudo conectar con Catálogo Asignaciones: {e}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False

        if datos_login.get("requiereVerificacion"):
            self.log_consola(
                "Catálogo Asignaciones pide un código de verificación (email "
                "enviado) — inicie sesión una vez desde el navegador para "
                "completarlo, y vuelva a intentar esta acción después.", "WARN")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            self.after(0, lambda: messagebox.showwarning(
                "Verificación pendiente",
                "Catálogo Asignaciones ha pedido un código de verificación "
                "por email (por inactividad prolongada de la cuenta).\n\n"
                "Inicie sesión una vez desde el navegador para completarlo, "
                "y vuelva a intentar esta acción después."))
            return False

        if not datos_login.get("ok"):
            self.log_consola(
                f"Login en Catálogo Asignaciones sin éxito: "
                f"{datos_login.get('error', '(sin detalle)')}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False

        # ── 2. Descargar el Excel completo ─────────────────────────────
        try:
            peticion = _urlreq.Request(f"{DALI_APP_API_BASE}/admin/export-completo")
            with opener.open(peticion, timeout=60) as resp:
                contenido = resp.read()
        except _urlerr.HTTPError as e:
            if e.code == 403:
                self.log_consola(
                    "Su cuenta no tiene permiso para exportar el catálogo completo "
                    "en Catálogo Asignaciones — pídaselo al administrador principal.",
                    "ERROR")
            else:
                self.log_consola(f"Error al descargar el catálogo "
                                 f"(HTTP {e.code}): {e.reason}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False
        except Exception as e:
            self.log_consola(f"No se pudo descargar el catálogo: {e}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False

        # ── 3. Guardar (escritura atómica: fichero temporal + reemplazo) ──
        # Fix (2026-09-25, error real: WinError 5 Acceso denegado en
        # os.replace): el archivo destino puede estar bloqueado un
        # instante por OneDrive sincronizando, un antivirus escaneándolo,
        # o Excel con el archivo abierto -- un bloqueo casi siempre
        # transitorio. Se reintenta unas pocas veces con una pequeña
        # espera antes de dar el guardado por fallido de verdad.
        try:
            recursos_path = os.path.join(carpeta, FOLDER_RECURSOS)
            os.makedirs(recursos_path, exist_ok=True)
            ruta_final = os.path.join(recursos_path, "catalogo-dali-completo.xlsx")
            ruta_tmp = ruta_final + ".tmp"
            with open(ruta_tmp, "wb") as f:
                f.write(contenido)

            _INTENTOS_REEMPLAZO = 5
            _ESPERA_ENTRE_INTENTOS = 1.0  # segundos
            for _intento in range(1, _INTENTOS_REEMPLAZO + 1):
                try:
                    os.replace(ruta_tmp, ruta_final)
                    break
                except PermissionError:
                    if _intento == _INTENTOS_REEMPLAZO:
                        raise
                    self.log_consola(
                        f"  ⚠ catalogo-dali-completo.xlsx está en uso "
                        f"(OneDrive, antivirus o Excel abierto) -- "
                        f"reintentando ({_intento}/{_INTENTOS_REEMPLAZO})...",
                        "WARN")
                    time.sleep(_ESPERA_ENTRE_INTENTOS)
        except Exception as e:
            self.log_consola(f"Catálogo descargado pero no se pudo guardar: {e}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False

        self.log_consola(
            f"  └ catalogo-dali-completo.xlsx actualizado desde Catálogo "
            f"Asignaciones ({len(contenido):,} bytes).".replace(",", "."), "OK")
        self.after(0, lambda: self._barra.completado("Catálogo actualizado"))
        return True

    def _descargar_catalogo_dali_web_bloqueante(self, email, password):
        """
        Ejecuta _descargar_catalogo_dali_web(email, password) en un
        hilo de fondo, mostrando una ventana modal de "Descargando..."
        mientras tanto, y no devuelve el control hasta que termina --
        pero SIN congelar la interfaz mientras espera (la GUI sigue
        respondiendo: la espera se hace con wait_window(), que corre su
        propio bucle de eventos local, y el hilo se comprueba por
        sondeo con after() en vez de con join() bloqueante).

        Corrige un bug real (2026-09-20): antes, _descargar_catalogo_
        dali_web() se llamaba directamente desde el hilo principal --
        las peticiones de red (hasta 20s+60s de timeout) bloqueaban
        Tkinter entero, mostrando "(No responde)" en la barra de
        título de Windows durante toda la espera.
        """
        resultado = {"ok": False}

        def _trabajo():
            resultado["ok"] = self._descargar_catalogo_dali_web(email, password)

        win = tk.Toplevel(self)
        win.title("Descargando catálogo")
        win.configure(bg=C_DARK)
        win.resizable(False, False)
        win.transient(self)
        win.protocol("WM_DELETE_WINDOW", lambda: None)  # no cerrable a mano
        tk.Label(
            win, text="Descargando catálogo actualizado desde\n"
                      "Catálogo Asignaciones, espere un momento...",
            bg=C_DARK, fg=C_TEXT, font=FONT_LABEL, justify="center",
            padx=36,
        ).pack(pady=(28, 4))
        lbl_tiempo = tk.Label(
            win, text="Llevamos 0s...", width=24,
            bg=C_DARK, fg=C_MUTED, font=FONT_SMALL, justify="center",
        )
        lbl_tiempo.pack(pady=(0, 20))
        win.update_idletasks()
        win.grab_set()

        # ── Centrar (mismo patrón que el resto de ventanas del proyecto) ──
        _w = win.winfo_reqwidth()
        _h = win.winfo_reqheight()
        _sw = win.winfo_screenwidth()
        _sh = win.winfo_screenheight()
        win.geometry(f"{_w}x{_h}+{(_sw-_w)//2}+{max(60, (_sh-_h)//3)}")

        hilo = threading.Thread(target=_trabajo, daemon=True)
        _inicio = time.monotonic()
        hilo.start()

        def _comprobar():
            if hilo.is_alive():
                _segundos = int(time.monotonic() - _inicio)
                lbl_tiempo.config(text=f"Llevamos {_segundos}s... (suele tardar ~1 min)")
                win.after(150, _comprobar)
            else:
                win.grab_release()
                win.destroy()

        win.after(150, _comprobar)
        win.wait_window()
        return resultado["ok"]

    def _ejecutar_en_hilo_principal_y_esperar(self, funcion):
        """
        Ejecuta `funcion` (sin argumentos) de forma segura para Tkinter y
        devuelve su resultado -- tanto si quien llama es el hilo principal
        como si es un hilo de fondo (el de proceso() de cualquier Acción).

        Bug real corregido (2026-09-24, caso real: aplicación totalmente
        congelada, hubo que matarla desde el Administrador de tareas):
        Tkinter NO es thread-safe -- crear o manipular cualquier widget
        (messagebox, tk.Toplevel, wait_window...) desde un hilo que no es
        el principal puede colgar la aplicación entera, pero de forma
        INTERMITENTE (a veces "funciona" por pura suerte de sincronización
        con lo que esté haciendo el hilo principal en ese instante, hasta
        que un día no). _catalogo_reciente_o_confirmar() y _confirmar_
        catalogo_actualizado() hacían exactamente esto -- se llaman desde
        dentro de proceso() en los 6+ puntos de enganche del catálogo, y
        proceso() siempre corre en un hilo de fondo (threading.Thread).

        Si ya estamos en el hilo principal, ejecuta `funcion` directamente
        (evita un interbloqueo real: bloquear el hilo principal esperando
        un after() que solo el propio hilo principal puede procesar sería
        un punto muerto inmediato). Si no, programa `funcion` con
        self.after(0, ...) para que corra en el hilo principal, y bloquea
        (de forma segura, con threading.Event) el hilo que llama hasta que
        termine -- devolviendo su resultado, o relanzando su excepción si
        la hubo.
        """
        if threading.current_thread() is threading.main_thread():
            return funcion()

        resultado = {}
        evento = threading.Event()

        def _envoltorio():
            try:
                resultado["valor"] = funcion()
            except Exception as e:
                resultado["excepcion"] = e
            finally:
                evento.set()

        self.after(0, _envoltorio)
        evento.wait()
        if "excepcion" in resultado:
            raise resultado["excepcion"]
        return resultado.get("valor")

    def _catalogo_reciente_o_confirmar(self):
        """
        Wrapper thread-safe (ver _ejecutar_en_hilo_principal_y_esperar):
        delega en _catalogo_reciente_o_confirmar_impl() para la lógica
        real (sin tocarla), garantizando que se ejecuta en el hilo
        principal aunque quien llame sea el hilo de fondo de proceso()
        de cualquier Acción.
        """
        return self._ejecutar_en_hilo_principal_y_esperar(
            self._catalogo_reciente_o_confirmar_impl)

    def _catalogo_reciente_o_confirmar_impl(self):
        """
        Comprobación previa, específica de Acción 4 ("Generar TARIFA final
        y Procesar Tarifas Incompletas"), llamada justo antes de cada uno
        de sus 4 puntos de enganche con _actualizar_codigo_proveedor_
        catalogo() (_ejecutar_transformacion camino limpio, y los 3 caminos
        de _gestionar_faltantes_o_finalizar_inc). A diferencia del aviso de
        Acciones 2/3 (que se muestra SIEMPRE al principio), este NO
        interrumpe nada si todo está reciente -- solo aparece cuando de
        verdad puede haber un problema, para no repetir el aviso constante
        que Acción 4 dejó de mostrar a propósito (ver _confirmar_catalogo_
        actualizado, que Acción 4 ya no llama).

        Revisa la antigüedad de catalogo-dali-completo.xlsx Y de sus dos
        ficheros de origen (ARTICULOS_DALI.xlsx, LISTADO CODIGOS DALI -
        SAP.xlsx) -- si Acciones 1-3 no se han corrido hoy, estos también
        pueden estar desactualizados aunque el catálogo en sí no se haya
        tocado hace mucho.

        Umbral: 30 minutos. Por debajo, se actualiza sin más (devuelve
        True en silencio). Por encima de alguno de los tres, ofrece el
        mismo diálogo de 3 vías que Acciones 2/3 (Sí=descargar, No=
        continuar igualmente, Cancelar=no actualizar el catálogo esta vez
        -- la tarifa en sí ya quedó completada, solo se salta este paso).

        Devuelve True si se debe proceder a actualizar el catálogo con
        _actualizar_codigo_proveedor_catalogo(), False si no.
        """
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            return True

        from datetime import datetime as _dt, timedelta as _td
        UMBRAL = _td(minutes=30)
        ahora = _dt.now()

        rutas_a_revisar = [
            ("catalogo-dali-completo.xlsx",
             os.path.join(carpeta, FOLDER_RECURSOS, "catalogo-dali-completo.xlsx")),
            ("ARTICULOS_DALI.xlsx",
             os.path.join(carpeta, FOLDER_RECURSOS, "ARTICULOS_DALI.xlsx")),
            ("LISTADO CODIGOS DALI - SAP.xlsx",
             os.path.join(carpeta, "LISTADO CODIGOS DALI - SAP.xlsx")),
        ]

        if not os.path.isfile(rutas_a_revisar[0][1]):
            return True   # catálogo nunca generado: nada que comprobar

        anticuados = []
        for nombre, ruta in rutas_a_revisar:
            if not os.path.isfile(ruta):
                continue
            try:
                mtime = _dt.fromtimestamp(os.path.getmtime(ruta))
            except Exception:
                continue
            antiguedad = ahora - mtime
            if antiguedad > UMBRAL:
                minutos = int(antiguedad.total_seconds() // 60)
                txt_edad = f"{minutos} min" if minutos < 60 \
                    else f"{minutos // 60}h {minutos % 60}min"
                anticuados.append((nombre, mtime.strftime("%d/%m/%Y %H:%M"), txt_edad))

        if not anticuados:
            return True

        detalle = "\n".join(
            f"   • {n}: actualizado {f} (hace {e})" for n, f, e in anticuados)
        respuesta = messagebox.askyesnocancel(
            "Catálogo DALI completo — puede estar desactualizado",
            "Antes de actualizar el catálogo con esta tarifa, se detectó "
            "que alguno de los archivos que lo alimentan tiene más de 30 "
            "minutos de antigüedad:\n\n"
            f"{detalle}\n\n"
            "El catálogo también se mantiene desde la aplicación «Catálogo "
            "Asignaciones» — puede convenir refrescarlo antes de escribir "
            "en él.\n\n"
            "  SÍ       → descargar ahora una copia actualizada desde "
            "Catálogo Asignaciones y continuar con ella\n"
            "  NO       → actualizar igualmente con la copia actual\n"
            "  CANCELAR → no actualizar el catálogo esta vez (la tarifa "
            "ya quedó completada igualmente)")

        if respuesta is None:
            self.log_consola(
                "Actualización del catálogo omitida por decisión del "
                "usuario (catálogo/origen desactualizado).", "WARN")
            return False
        if respuesta is False:
            return True

        email, password = self._obtener_credenciales_dali_app()
        if not email or not password:
            self.log_consola("Descarga del catálogo cancelada: sin credenciales.", "WARN")
            return False
        return self._descargar_catalogo_dali_web_bloqueante(email, password)

    def _confirmar_catalogo_actualizado(self):
        """
        Wrapper thread-safe (ver _ejecutar_en_hilo_principal_y_esperar):
        delega en _confirmar_catalogo_actualizado_impl() para la lógica
        real (sin tocarla), garantizando que se ejecuta en el hilo
        principal aunque quien llame sea el hilo de fondo de proceso()
        de Acción 2 o Acción 3.
        """
        return self._ejecutar_en_hilo_principal_y_esperar(
            self._confirmar_catalogo_actualizado_impl)

    def _confirmar_catalogo_actualizado_impl(self):
        """
        Aviso previo, común a las Acciones 2 y 3 (las que acaban
        modificando RECURSOS/catalogo-dali-completo.xlsx durante la
        actualización de plantillas DALI/SAP -- Acción 4 no la llama en
        absoluto, ver nota en _ejecutar_reinscripcion_sap_desde_inc):
        antes de empezar nada, informa de la fecha de última
        actualización de ese archivo -- el catálogo también se mantiene
        desde la app web «Catálogo Asignaciones» (asignación manual de
        Código Proveedor por otras personas), así que la copia local de
        este equipo puede quedarse desactualizada sin que esta
        aplicación tenga forma de saberlo.

        Diálogo de 3 vías (Sí / No / Cancelar):
          Sí       -> descargar ahora una copia fresca desde Catálogo
                      Asignaciones (_descargar_catalogo_dali_web()) y
                      continuar con ella.
          No       -> continuar sin descargar, con la copia local tal
                      cual está.
          Cancelar -> no continuar con la acción.

        Memoria de sesión (2026-09-21, a petición del usuario): una vez
        resuelto (Sí o No, no Cancelar) dentro de esta misma ejecución
        de la aplicación, NO se vuelve a preguntar -- Acción 2 y Acción
        3 suelen correr una justo detrás de la otra como parte del
        mismo proceso de actualización de plantillas, y repetir la
        pregunta sería puro ruido. Se reinicia solo al cerrar y volver
        a abrir la aplicación (self._catalogo_verificado_sesion vive en
        la instancia, no se persiste en tarifa_config.json).

        Devuelve True si se debe continuar (con o sin descarga previa,
        o si el catálogo todavía no existe -- nada que avisar), False
        si no se debe continuar.
        """
        if getattr(self, "_catalogo_verificado_sesion", False):
            return True   # ya se preguntó y se resolvió en esta misma sesión de trabajo

        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            return True   # sin proyecto seleccionado: cada acción ya valida esto por su cuenta

        ruta_catalogo = os.path.join(carpeta, FOLDER_RECURSOS, "catalogo-dali-completo.xlsx")
        if not os.path.isfile(ruta_catalogo):
            return True   # nunca se ha generado: nada que advertir todavía

        try:
            from datetime import datetime as _dt
            _mtime = os.path.getmtime(ruta_catalogo)
            _fecha_txt = _dt.fromtimestamp(_mtime).strftime("%d/%m/%Y a las %H:%M")
        except Exception:
            _fecha_txt = "(fecha no disponible)"

        respuesta = messagebox.askyesnocancel(
            "Catálogo DALI completo — comprobar actualización",
            "catalogo-dali-completo.xlsx se actualizó por última vez:\n\n"
            f"   {_fecha_txt}\n\n"
            "Esta acción va a modificarlo ahora. El catálogo también se "
            "mantiene desde la aplicación «Catálogo Asignaciones» (otras "
            "personas pueden haber asignado códigos proveedor después de "
            "esta fecha), así que puede convenir refrescarlo antes de "
            "seguir.\n\n"
            "  SÍ       → descargar ahora una copia actualizada desde "
            "Catálogo Asignaciones y continuar con ella\n"
            "  NO       → continuar sin descargar, con la copia actual\n"
            "  CANCELAR → no continuar con esta acción")

        if respuesta is None:
            return False   # cancelar: no marcar como resuelto, para que vuelva a preguntar
        if respuesta is False:
            self._catalogo_verificado_sesion = True
            return True
        email, password = self._obtener_credenciales_dali_app()
        if not email or not password:
            self.log_consola("Descarga del catálogo cancelada: sin credenciales.", "WARN")
            return False   # sin credenciales: no marcar, para que se pueda reintentar
        exito = self._descargar_catalogo_dali_web_bloqueante(email, password)
        if exito:
            self._catalogo_verificado_sesion = True
        return exito

    def _regenerar_catalogo_dali_completo(self):
        """
        Wrapper con bloqueo global (ver _adquirir_lock_catalogo_con_aviso):
        adquiere self._catalogo_lock antes de tocar catalogo-dali-completo.xlsx
        y lo libera siempre al terminar, delegando en _regenerar_catalogo_
        dali_completo_impl() para la lógica real (sin tocarla).
        """
        self._adquirir_lock_catalogo_con_aviso()
        try:
            return self._regenerar_catalogo_dali_completo_impl()
        finally:
            self._catalogo_lock.release()

    def _regenerar_catalogo_dali_completo_impl(self):
        """
        Regenera RECURSOS/catalogo-dali-completo.xlsx combinando:
          - RECURSOS/ARTICULOS_DALI.xlsx (Acción 1): código, descripción,
            unidad, familia, subfamilia, naturaleza, IGIC, proveedor
            asignado y activo (lado DALI).
          - <raíz>/LISTADO CODIGOS DALI - SAP.xlsx (Acciones 2/3): código
            SAP, descripciones/grupos SAP, categoría valoración, tipo de
            producto, grupo de compras (lado SAP).

        Llamado automáticamente al final de _actualizar_listado_dali()
        (Acción 2) y de _cruzar_y_escribir_listado_sap() (Acción 3) --
        cada vez que cualquiera de las dos actualiza el LISTADO, el
        catálogo se regenera para mantenerse siempre sincronizado. No es
        una acción independiente con botón propio: se ejecuta de forma
        síncrona dentro del hilo de fondo que ya está corriendo, así que
        no usa su propio threading.Thread ni diálogos de confirmación, y
        cualquier error aquí se registra en el log sin interrumpir el
        éxito de la acción que lo llamó (ver el try/except del llamador).

        La columna B (Código Proveedor) se conserva de la versión
        anterior del catálogo SOLO si la columna I (Proveedor Asignado -
        DALI) de ese código no ha cambiado; si cambió, se vacía B y se
        deja constancia en un .txt en RECURSOS/ (ver más abajo). Acción 4
        (_actualizar_codigo_proveedor_catalogo) escribe B por su cuenta,
        con datos más específicos (la tarifa concreta), cuando corresponda.
        """
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            return

        recursos_path        = os.path.join(carpeta, FOLDER_RECURSOS)
        ruta_articulos_dali  = os.path.join(recursos_path, "ARTICULOS_DALI.xlsx")
        ruta_listado         = os.path.join(carpeta, "LISTADO CODIGOS DALI - SAP.xlsx")
        ruta_catalogo        = os.path.join(recursos_path, "catalogo-dali-completo.xlsx")

        if not os.path.isfile(ruta_articulos_dali) or not os.path.isfile(ruta_listado):
            self.log_consola(
                "  ℹ catalogo-dali-completo.xlsx no regenerado: falta "
                "ARTICULOS_DALI.xlsx o LISTADO CODIGOS DALI - SAP.xlsx.", "INFO")
            return

        import openpyxl
        from copy import copy
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from datetime import datetime as _dt

        self.consola.separador("REGENERAR CATÁLOGO DALI COMPLETO")

        # ── 1. Leer Código Proveedor, Proveedor Asignado y Precio € del catálogo anterior ──
        #    Se localizan las columnas por NOMBRE DE CABECERA (fila 4), no por
        #    posición: el catálogo anterior puede tener el formato viejo (20
        #    columnas) o el nuevo (22, con Texto UM y Precio €) y en ambos cada
        #    valor debe salir de SU columna. A, B e I coinciden en los dos
        #    formatos, así que son el valor por defecto si falta la cabecera.
        self.after(0, lambda: self._barra.indeterminate("Actualizando catálogo DALI..."))
        _datos_previos = {}   # clave DALI -> (cod_proveedor_previo, proveedor_asignado_previo)
        _precios_previos = {}   # clave DALI -> precio previo (columna "Precio €")
        if os.path.isfile(ruta_catalogo):
            try:
                wb_prev = openpyxl.load_workbook(ruta_catalogo, read_only=True, data_only=True)
                ws_prev = wb_prev.active
                _mapa_prev = {}
                for _n_fila, row in enumerate(
                        ws_prev.iter_rows(min_row=CATALOGO_FILA_CABECERA, values_only=True),
                        start=CATALOGO_FILA_CABECERA):
                    if _n_fila == CATALOGO_FILA_CABECERA:
                        _mapa_prev = mapa_columnas_catalogo(row or ())
                        _i_clave = _mapa_prev.get("clave", 1)
                        _i_cprov = _mapa_prev.get("cod_prov", 2)
                        _i_prov  = _mapa_prev.get("proveedor", 9)
                        _i_prec  = _mapa_prev.get("precio")     # None si el catálogo viejo no la tiene
                        continue
                    if not row or len(row) < _i_clave or row[_i_clave - 1] is None:
                        continue
                    clave_prev = str(row[_i_clave - 1]).strip()
                    cod_prov_prev  = row[_i_cprov - 1] if len(row) >= _i_cprov else None
                    prov_asig_prev = row[_i_prov - 1]  if len(row) >= _i_prov  else None
                    _datos_previos[clave_prev] = (cod_prov_prev, prov_asig_prev)
                    if _i_prec and len(row) >= _i_prec and row[_i_prec - 1] not in (None, ""):
                        _precios_previos[clave_prev] = row[_i_prec - 1]
                wb_prev.close()
                _con_cod_prov_previo = sum(
                    1 for cp, _ in _datos_previos.values() if cp not in (None, ""))
                self.log_consola(
                    f"  └ {len(_datos_previos)} código(s) leídos del catálogo anterior "
                    f"({_con_cod_prov_previo} con código proveedor, "
                    f"{len(_precios_previos)} con precio)", "INFO")
            except Exception as _e_prev:
                self.log_consola(
                    f"  ⚠  No se pudo leer el catálogo anterior "
                    f"(se generará sin conservar códigos proveedor): {_e_prev}", "WARN")

        # ── 2. Leer ARTICULOS_DALI.xlsx ─────────────────────────────
        def _lim(v):
            """Quita espacios sobrantes de inicio/fin si v es texto; deja
            números/None tal cual. Evita que espacios arrastrados desde el
            origen (p.ej. Productos.xlsx) lleguen al catálogo."""
            return v.strip() if isinstance(v, str) else v

        self.log_consola("Leyendo ARTICULOS_DALI.xlsx...", "INFO")
        wb_dali = openpyxl.load_workbook(ruta_articulos_dali, read_only=True, data_only=True)
        ws_dali = wb_dali.active
        tabla_dali = {}
        for row in ws_dali.iter_rows(min_row=8, values_only=True):
            if not row or row[0] is None:
                continue
            clave = str(row[0]).strip()
            tabla_dali[clave] = {
                "descripcion": _lim(row[1]) if len(row) > 1 else None,
                "unidad":      _lim(row[2]) if len(row) > 2 else None,
                "familia":     _lim(row[3]) if len(row) > 3 else None,
                "subfamilia":  _lim(row[4]) if len(row) > 4 else None,
                "naturaleza":  _lim(row[5]) if len(row) > 5 else None,
                "igic":        _lim(row[6]) if len(row) > 6 else None,
                "proveedor":   _lim(row[7]) if len(row) > 7 else None,
                "activo":      _lim(row[8]) if len(row) > 8 else None,
            }
        wb_dali.close()
        self.log_consola(f"  └ {len(tabla_dali)} artículos DALI leídos", "OK")

        # ── 3. Leer LISTADO CODIGOS DALI - SAP.xlsx ─────────────────
        self.log_consola("Leyendo LISTADO CODIGOS DALI - SAP.xlsx...", "INFO")
        wb_lst = openpyxl.load_workbook(ruta_listado, read_only=True, data_only=True)
        ws_lst = wb_lst.active
        tabla_sap = {}
        for row in ws_lst.iter_rows(min_row=8, values_only=True):
            if not row or row[0] is None:
                continue
            clave = str(row[0]).strip()
            tabla_sap[clave] = datos_sap_desde_fila_listado(row)
        wb_lst.close()
        self.log_consola(f"  └ {len(tabla_sap)} entrada(s) del LISTADO leídas", "OK")

        # ── 4. Construir el catálogo ─────────────────────────────────
        wb_out = openpyxl.Workbook()
        ws_out = wb_out.active
        ws_out.title = "Catálogo completo"

        _codigos_ordenados = sorted(
            tabla_dali.keys(),
            key=lambda c: (0, int(c)) if c.isdigit() else (1, c))

        _c_titulo = ws_out.cell(row=1, column=1,
                    value="Listado completo de artículos DALI — importación DALI + SAP + código proveedor")
        _c_titulo.font = Font(size=14, bold=True)
        _c_subtitulo = ws_out.cell(row=2, column=1,
                    value=f"Generado el {_dt.now().strftime('%d/%m/%Y, %H:%M:%S')} · "
                          f"{len(_codigos_ordenados)} artículo(s)")
        _c_subtitulo.font = Font(size=9, italic=True, color="FF666666")

        # Layout desde la definición canónica (sap_helpers.CATALOGO_COLUMNAS):
        # cabecera, orden de columnas y anchos salen de ahí, así que este
        # método no puede volver a descuadrarse respecto a los demás.
        _claves_col = [k for k, _h, _a in CATALOGO_COLUMNAS]
        _fill_cabecera = PatternFill(fill_type="solid", fgColor="FFE5DECB")
        _lado = Side(style="thin", color="FF000000")
        _borde = Border(left=_lado, right=_lado, top=_lado, bottom=_lado)
        _centrado = Alignment(horizontal="center")
        for i, (_k, titulo, _a) in enumerate(CATALOGO_COLUMNAS, start=1):
            c = ws_out.cell(row=CATALOGO_FILA_CABECERA, column=i, value=titulo)
            c.font = Font(bold=True)
            c.fill = _fill_cabecera
            c.alignment = _centrado
            c.border = _borde

        # Estilo de las celdas de datos (centrado + borde fino, como el
        # catálogo que exporta la web). Se construye UNA vez y se copia por
        # celda: asignar font/alignment/border en cada una de las ~660.000
        # celdas sería mucho más lento.
        _proto = ws_out.cell(row=CATALOGO_PRIMERA_FILA_DATOS, column=1)
        _proto.alignment = _centrado
        _proto.border = _borde
        _estilo_datos = copy(_proto._style)

        _fila = CATALOGO_PRIMERA_FILA_DATOS
        _con_proveedor_conservado = 0
        _cambios_proveedor_asignado = []  # [(clave, prov_asig_anterior, prov_asig_nuevo, cod_prov_vaciado, precio_vaciado)]

        def _normalizar(v):
            return str(v).strip().upper() if v not in (None, "") else ""

        for clave in _codigos_ordenados:
            d = tabla_dali[clave]
            s = tabla_sap.get(clave, {})
            prov_asig_nuevo = d["proveedor"]

            cod_prov = None
            precio = None
            if clave in _datos_previos:
                cod_prov_prev, prov_asig_prev = _datos_previos[clave]
                precio_prev = _precios_previos.get(clave)
                if _normalizar(prov_asig_prev) == _normalizar(prov_asig_nuevo):
                    # Proveedor Asignado sin cambios -> conservar Código
                    # Proveedor y Precio € (ambos dependen del proveedor)
                    cod_prov = cod_prov_prev
                    precio = precio_prev
                    if cod_prov not in (None, ""):
                        _con_proveedor_conservado += 1
                elif cod_prov_prev not in (None, "") or precio_prev not in (None, ""):
                    # Proveedor Asignado cambió Y había un código proveedor y/o
                    # un precio que ya no son fiables -> vaciar ambos y dejar
                    # constancia en el .txt
                    _cambios_proveedor_asignado.append(
                        (clave, prov_asig_prev, prov_asig_nuevo, cod_prov_prev, precio_prev))

            # Cada valor va asociado a su CLAVE de columna (no a una posición):
            # el orden real lo decide CATALOGO_COLUMNAS al escribir.
            valores = {
                "clave":       int(clave) if clave.isdigit() else clave,
                "cod_prov":    cod_prov,
                "descripcion": d["descripcion"], "unidad": d["unidad"],
                "igic":        d["igic"],        "naturaleza": d["naturaleza"],
                "familia":     d["familia"],     "subfamilia": d["subfamilia"],
                "proveedor":   d["proveedor"],   "activo":     d["activo"],
                # "Precio €": igual que el código proveedor, se conserva del
                # catálogo anterior solo si el Proveedor Asignado no cambió.
                "precio":      precio,
            }
            for _k_sap in CATALOGO_CLAVES_SAP:
                valores[_k_sap] = s.get(_k_sap)
            for i, _k in enumerate(_claves_col, start=1):
                c = ws_out.cell(row=_fila, column=i, value=valores.get(_k))
                c._style = copy(_estilo_datos)
            _fila += 1

        for i, (_k, _h, ancho) in enumerate(CATALOGO_COLUMNAS, start=1):
            ws_out.column_dimensions[openpyxl.utils.get_column_letter(i)].width = ancho
        ws_out.auto_filter.ref = (
            f"A{CATALOGO_FILA_CABECERA}:"
            f"{openpyxl.utils.get_column_letter(len(CATALOGO_COLUMNAS))}"
            f"{CATALOGO_FILA_CABECERA}")

        os.makedirs(recursos_path, exist_ok=True)
        wb_out.save(ruta_catalogo)

        # ── 4b. Si hubo cambios de Proveedor Asignado, dejar constancia en .txt ──
        if _cambios_proveedor_asignado:
            _sello = _dt.now().strftime('%Y%m%d_%H%M%S')
            _ruta_txt_cambios = os.path.join(
                recursos_path, f"catalogo_cambios_proveedor_{_sello}.txt")
            with open(_ruta_txt_cambios, "w", encoding="utf-8") as _f_txt:
                _f_txt.write(
                    "CASOS IRREGULARES — CAMBIOS DE PROVEEDOR ASIGNADO DETECTADOS "
                    "AL REGENERAR EL CATÁLOGO\n"
                    f"Generado el {_dt.now().strftime('%d/%m/%Y, %H:%M:%S')}\n\n"
                    "Los siguientes códigos DALI han cambiado de \"Proveedor Asignado - DALI\"\n"
                    "respecto al catálogo anterior. Normalmente ambas bases de datos de\n"
                    "proveedores deberían coincidir exactamente en ambos listados -- esto\n"
                    "suele indicar un desfase puntual (p.ej. el nombre se actualizó en una\n"
                    "base de datos antes que en la otra). Por precaución, se han vaciado su\n"
                    "Código Proveedor (columna B) y su Precio € -- revíselos y complételos de\n"
                    "nuevo si procede.\n\n"
                )
                _f_txt.write(
                    f"{'Código DALI':<12} | {'Proveedor anterior':<35} | "
                    f"{'Proveedor nuevo':<35} | {'Código Proveedor vaciado':<25} | Precio € vaciado\n")
                _f_txt.write("-" * 130 + "\n")
                for _clave_c, _prov_ant, _prov_new, _cod_vaciado, _precio_vaciado in _cambios_proveedor_asignado:
                    _f_txt.write(
                        f"{_clave_c:<12} | {str(_prov_ant or ''):<35} | "
                        f"{str(_prov_new or ''):<35} | {str(_cod_vaciado or ''):<25} | "
                        f"{'' if _precio_vaciado in (None, '') else _precio_vaciado}\n")
            self.log_consola(
                f"  ⚠  CASO IRREGULAR: {len(_cambios_proveedor_asignado)} código(s) DALI "
                f"cambiaron de Proveedor Asignado — código proveedor y precio vaciados, "
                f"detalle en {os.path.basename(_ruta_txt_cambios)}", "WARN")
            for _clave_c, _prov_ant, _prov_new, _cod_vaciado, _precio_vaciado in _cambios_proveedor_asignado[:30]:
                _vaciado = []
                if _cod_vaciado not in (None, ""):
                    _vaciado.append(f"código {_cod_vaciado}")
                if _precio_vaciado not in (None, ""):
                    _vaciado.append(f"precio {_precio_vaciado}")
                self.log_consola(
                    f"      · DALI {_clave_c}: «{_prov_ant or '(vacío)'}» → "
                    f"«{_prov_new or '(vacío)'}» ({' y '.join(_vaciado)} vaciado)", "WARN")
            if len(_cambios_proveedor_asignado) > 30:
                self.log_consola(
                    f"      … y {len(_cambios_proveedor_asignado) - 30} más "
                    f"(omitidos del log, ver {os.path.basename(_ruta_txt_cambios)})", "WARN")

            def _avisar_cambios_proveedor_asignado(cambios=_cambios_proveedor_asignado,
                                                    nombre_txt=os.path.basename(_ruta_txt_cambios)):
                lineas = [
                    "⚠  CASO IRREGULAR — Proveedor Asignado cambió",
                    "",
                    f"Los siguientes {len(cambios)} código(s) DALI han cambiado de",
                    "\"Proveedor Asignado - DALI\" respecto al catálogo anterior.",
                    "Normalmente ambas bases de datos de proveedores deberían",
                    "coincidir exactamente -- revise estos códigos, puede tratarse",
                    "de un desfase puntual entre listados:",
                    "",
                ]
                for _clave_c, _prov_ant, _prov_new, _cod_vaciado, _precio_vaciado in cambios[:20]:
                    lineas.append(
                        f"   • DALI {_clave_c}: «{_prov_ant or '(vacío)'}» → "
                        f"«{_prov_new or '(vacío)'}»")
                if len(cambios) > 20:
                    lineas.append(f"   … y {len(cambios) - 20} más (ver log de consola).")
                lineas += [
                    "",
                    "Su Código Proveedor y su Precio € se han vaciado en catalogo-dali-completo.xlsx.",
                    f"Detalle completo en RECURSOS/{nombre_txt}",
                ]
                messagebox.showwarning(
                    "Caso irregular — Proveedor Asignado cambió", "\n".join(lineas))
            self.after(0, _avisar_cambios_proveedor_asignado)

        _n_filas = _fila - 5
        self.log_consola(f"  └ catalogo-dali-completo.xlsx regenerado: {_n_filas} fila(s), "
                          f"{_con_proveedor_conservado} código(s) proveedor conservados", "OK")
        self.after(0, lambda: self._barra.completado("Catálogo DALI actualizado"))
        self.consola.separador("FIN REGENERAR CATÁLOGO DALI COMPLETO")

    def _refrescar_sap_catalogo_para_dali(self, codigos_dali):
        """
        Actualiza SOLO las columnas J-S (datos de origen SAP: código SAP,
        descripción, unidad, grupo de producto, IGIC, categoría, tipo,
        grupo de compras) de las filas de catalogo-dali-completo.xlsx
        correspondientes a los códigos DALI indicados, leyendo su estado
        actual desde LISTADO CODIGOS DALI - SAP.xlsx. No toca el resto
        de columnas (proveedor, datos DALI) ni añade filas nuevas.

        Fix (2026-09-25, hallazgo real del usuario): cuando Ventana
        SAPFaltantes resuelve un código SAP duplicado y borra el SAP de
        un DALI que NO pertenece a la tarifa en curso (ver
        VentanaResolverSAPDuplicados), _actualizar_codigo_proveedor_
        catalogo() nunca lo refresca -- solo toca los códigos DALI de
        la propia tarifa. El catálogo se queda con el SAP viejo hasta
        la siguiente Acción 2/3. Este método cubre ese hueco con una
        actualización puntual, ligera, sin regenerar el catálogo entero.

        codigos_dali: iterable de códigos DALI (int o str) a refrescar.
        No hace nada si el catálogo o el LISTADO no existen todavía.
        """
        if not codigos_dali:
            return
        self._adquirir_lock_catalogo_con_aviso()
        try:
            import openpyxl

            carpeta       = self._carpeta_proyecto.get().strip()
            recursos_path = os.path.join(carpeta, FOLDER_RECURSOS)
            ruta_catalogo = os.path.join(recursos_path, "catalogo-dali-completo.xlsx")
            ruta_listado  = os.path.join(carpeta, "LISTADO CODIGOS DALI - SAP.xlsx")
            if not os.path.isfile(ruta_catalogo) or not os.path.isfile(ruta_listado):
                return

            claves = set()
            for c in codigos_dali:
                try:
                    claves.add(str(int(float(str(c).strip()))))
                except (ValueError, TypeError):
                    continue
            if not claves:
                return

            wb_l = openpyxl.load_workbook(ruta_listado, read_only=True, data_only=True)
            ws_l = wb_l.active
            _datos_sap = {}
            for row in ws_l.iter_rows(min_row=8, values_only=True):
                if not row or row[0] is None:
                    continue
                clave = str(row[0]).strip()
                if clave not in claves:
                    continue
                _datos_sap[clave] = datos_sap_desde_fila_listado(row)
            wb_l.close()

            wb_cat = openpyxl.load_workbook(ruta_catalogo)
            ws_cat = wb_cat.active
            # Columnas localizadas por cabecera: valen tanto el formato viejo
            # (20 columnas) como el nuevo (con Texto UM y Precio €).
            _mapa = mapa_columnas_catalogo(
                [c.value for c in ws_cat[CATALOGO_FILA_CABECERA]])
            _col_clave = _mapa.get("clave", 1)
            _fila_por_clave = {}
            for row in ws_cat.iter_rows(min_row=CATALOGO_PRIMERA_FILA_DATOS):
                if not row or len(row) < _col_clave or row[_col_clave - 1].value is None:
                    continue
                _fila_por_clave[str(row[_col_clave - 1].value).strip()] = row[0].row

            actualizadas = 0
            for clave in claves:
                fila = _fila_por_clave.get(clave)
                if fila is None:
                    continue
                s = _datos_sap.get(clave, {})
                # Solo se tocan las columnas de origen SAP; "Precio €", el
                # proveedor y los datos DALI no se modifican aquí.
                for _k in CATALOGO_CLAVES_SAP:
                    _col = _mapa.get(_k)
                    if _col:
                        # ws_cat.cell(..., value=v) NO borra si v es None (bug
                        # conocido de openpyxl, ya visto antes en este mismo
                        # proyecto) -- hace falta asignar .value en dos pasos.
                        ws_cat.cell(row=fila, column=_col).value = s.get(_k)
                actualizadas += 1

            wb_cat.save(ruta_catalogo)
            wb_cat.close()
            if actualizadas:
                self.log_consola(
                    f"  └ Catálogo: datos SAP refrescados para {actualizadas} "
                    f"código(s) DALI tras resolución de duplicados.", "OK")
        except Exception as _e_ref:
            self.log_consola(
                f"  ⚠  No se pudo refrescar el catálogo tras resolver "
                f"duplicados: {_e_ref}", "WARN")
        finally:
            self._catalogo_lock.release()

    def _actualizar_codigo_proveedor_catalogo(self, ruta_tarifa):
        """
        Wrapper con bloqueo global (ver _adquirir_lock_catalogo_con_aviso):
        adquiere self._catalogo_lock antes de tocar catalogo-dali-completo.xlsx
        y lo libera siempre al terminar, delegando en _actualizar_codigo_
        proveedor_catalogo_impl() para la lógica real (sin tocarla).
        """
        self._adquirir_lock_catalogo_con_aviso()
        try:
            return self._actualizar_codigo_proveedor_catalogo_impl(ruta_tarifa)
        finally:
            self._catalogo_lock.release()

    def _actualizar_codigo_proveedor_catalogo_impl(self, ruta_tarifa):
        """
        Llamada desde _ejecutar_reinscripcion_sap_desde_inc() (Acción 4)
        justo después de que una tarifa queda lista para crear/modificar
        el Registro Info (renombrada, sin [INCOMPLETO] ni prefijo INC).

        Actualiza en RECURSOS/catalogo-dali-completo.xlsx, ÚNICAMENTE las
        filas de los códigos DALI presentes en esta tarifa (col B, desde
        fila 8) CUYO Proveedor Asignado (ARTICULOS_DALI.xlsx, el mismo
        valor que alimenta la col I del catálogo) coincide con el
        proveedor de esta tarifa (celda C2). Cuando hay match, refresca
        las 20 columnas con los valores actuales de:
          - RECURSOS/ARTICULOS_DALI.xlsx (Acción 1)         → cols A, C-I
          - <raíz>/LISTADO CODIGOS DALI - SAP.xlsx (Acc. 2/3) → cols J-T
          - la propia tarifa (Acción 4, col C)               → col B (Código Proveedor)

        Si el Proveedor Asignado NO coincide con el proveedor de la
        tarifa, esa fila NO se toca en absoluto (ninguna de las 20
        columnas) -- se avisa en el log y con una ventana de advertencia
        al terminar, para que se pueda revisar la asignación DALI.

        Si un código DALI de la tarifa no existe todavía como fila en el
        catálogo, se añade al final (solo si el proveedor coincide). No
        se toca ninguna fila de códigos DALI ajenos a esta tarifa.

        Si el catálogo todavía no existe (no se ha ejecutado la Acción 8
        nunca), no hace nada salvo avisar en el log -- no es un error.
        """
        try:
            import openpyxl

            carpeta       = self._carpeta_proyecto.get().strip()
            recursos_path = os.path.join(carpeta, FOLDER_RECURSOS)
            ruta_catalogo = os.path.join(recursos_path, "catalogo-dali-completo.xlsx")
            if not os.path.isfile(ruta_catalogo):
                self.log_consola(
                    "  ℹ catalogo-dali-completo.xlsx no existe todavía "
                    "(genérelo con la Acción 8) — catálogo no actualizado.",
                    "INFO")
                return

            # ── 1. Proveedor (C2), códigos DALI + código proveedor de la tarifa ──
            wb_tar = openpyxl.load_workbook(ruta_tarifa, read_only=True, data_only=True)
            ws_tar = wb_tar.active
            _nombre_prov_tarifa = ws_tar["C2"].value
            _nombre_prov_tarifa = str(_nombre_prov_tarifa).strip() if _nombre_prov_tarifa else ""
            _cod_prov_tarifa = {}     # clave DALI (str) -> código proveedor (tarifa col C)
            _precio_tarifa = {}       # clave DALI (str) -> PRECIO de la tarifa final (tarifa col F)
            _fila_tarifa_por_clave = {}   # clave DALI (str) -> nº de fila en la tarifa (para completar col C luego)
            _fila_actual_tar = 7
            for row in ws_tar.iter_rows(min_row=8, values_only=True):
                _fila_actual_tar += 1
                if not row or len(row) < 2:
                    continue
                v_dali = row[1]   # col B
                if v_dali is None:
                    continue
                try:
                    clave = str(int(float(str(v_dali).strip())))
                except (ValueError, TypeError):
                    continue
                v_prov = row[2] if len(row) > 2 else None   # col C
                _cod_prov_tarifa[clave] = v_prov
                _precio_tarifa[clave] = row[5] if len(row) > 5 else None   # col F: PRECIO
                _fila_tarifa_por_clave[clave] = _fila_actual_tar
            wb_tar.close()

            if not _cod_prov_tarifa:
                self.log_consola(
                    "  ⚠  No se pudieron leer códigos DALI de la tarifa "
                    "— catálogo no actualizado.", "WARN")
                return

            if not _nombre_prov_tarifa:
                self.log_consola(
                    "  ⚠  No se pudo leer el proveedor de la tarifa (celda C2) "
                    "— ninguna fila se actualizará en el catálogo hasta corregirlo.",
                    "WARN")

            # ── 2. Datos DALI (Acción 1), solo los códigos que necesitamos ──
            def _lim(v):
                """Quita espacios sobrantes de inicio/fin si v es texto."""
                return v.strip() if isinstance(v, str) else v

            ruta_articulos_dali = os.path.join(recursos_path, "ARTICULOS_DALI.xlsx")
            _datos_dali = {}
            if os.path.isfile(ruta_articulos_dali):
                wb_a = openpyxl.load_workbook(ruta_articulos_dali, read_only=True, data_only=True)
                ws_a = wb_a.active
                for row in ws_a.iter_rows(min_row=8, values_only=True):
                    if not row or row[0] is None:
                        continue
                    clave = str(row[0]).strip()
                    if clave not in _cod_prov_tarifa:
                        continue
                    _datos_dali[clave] = {
                        "descripcion": _lim(row[1]) if len(row) > 1 else None,
                        "unidad":      _lim(row[2]) if len(row) > 2 else None,
                        "familia":     _lim(row[3]) if len(row) > 3 else None,
                        "subfamilia":  _lim(row[4]) if len(row) > 4 else None,
                        "naturaleza":  _lim(row[5]) if len(row) > 5 else None,
                        "igic":        _lim(row[6]) if len(row) > 6 else None,
                        "proveedor":   _lim(row[7]) if len(row) > 7 else None,
                        "activo":      _lim(row[8]) if len(row) > 8 else None,
                    }
                wb_a.close()
            else:
                self.log_consola(
                    "  ⚠  ARTICULOS_DALI.xlsx no encontrado — "
                    "catálogo actualizado solo con los datos de la tarifa.", "WARN")

            # ── 3. Datos SAP (Acciones 2/3), solo los códigos que necesitamos ──
            ruta_listado = os.path.join(carpeta, "LISTADO CODIGOS DALI - SAP.xlsx")
            _datos_sap = {}
            if os.path.isfile(ruta_listado):
                wb_l = openpyxl.load_workbook(ruta_listado, read_only=True, data_only=True)
                ws_l = wb_l.active
                for row in ws_l.iter_rows(min_row=8, values_only=True):
                    if not row or row[0] is None:
                        continue
                    clave = str(row[0]).strip()
                    if clave not in _cod_prov_tarifa:
                        continue
                    _datos_sap[clave] = datos_sap_desde_fila_listado(row)
                wb_l.close()
            else:
                self.log_consola(
                    "  ⚠  LISTADO CODIGOS DALI-SAP no encontrado — "
                    "catálogo actualizado solo parcialmente.", "WARN")

            # ── 4. Abrir catálogo, localizar filas existentes por código DALI ──
            wb_cat = openpyxl.load_workbook(ruta_catalogo)
            ws_cat = wb_cat.active
            # Columnas por CABECERA (no por posición): cada valor cae en su
            # columna tanto en el formato viejo (20 cols) como en el nuevo
            # (22, con Texto UM y Precio €). Las columnas que el fichero no
            # tenga se saltan en lugar de escribir en otra.
            _mapa = mapa_columnas_catalogo(
                [c.value for c in ws_cat[CATALOGO_FILA_CABECERA]])
            _col_clave = _mapa.get("clave", 1)
            _col_cprov = _mapa.get("cod_prov", 2)
            _fila_por_clave = {}
            for row in ws_cat.iter_rows(min_row=CATALOGO_PRIMERA_FILA_DATOS):
                if not row or len(row) < _col_clave or row[_col_clave - 1].value is None:
                    continue
                _fila_por_clave[str(row[_col_clave - 1].value).strip()] = row[0].row

            def _normalizar(v):
                return str(v).strip().upper() if v not in (None, "") else ""

            _nombre_norm = _normalizar(_nombre_prov_tarifa)
            _actualizadas   = 0
            _anadidas       = 0
            _omitidas_proveedor = []   # [(clave, proveedor_asignado)]
            _completados_desde_catalogo = []   # [(clave, cod_prov_recuperado)]
            _fila_siguiente = ws_cat.max_row + 1

            # "Precio €" funciona como "Código Proveedor": se pega el valor de
            # la tarifa (col F, precio final ya preparado para el paso RINFO)
            # y, si el catálogo ya tenía uno, se SUSTITUYE. Solo se acepta un
            # número mayor que 0: una celda vacía, un 0 o un texto no deben
            # borrar un precio bueno del catálogo (en ese caso se conserva el
            # que ya hubiera y se avisa).
            _precios_escritos = 0
            _precios_sustituidos = 0    # sustituyen un valor anterior DISTINTO
            _precios_invalidos = []     # DALI sin precio válido en la tarifa
            _col_precio = _mapa.get("precio")

            for clave, cod_prov in _cod_prov_tarifa.items():
                d = _datos_dali.get(clave, {})
                s = _datos_sap.get(clave, {})

                _proveedor_asignado = d.get("proveedor")
                if _normalizar(_proveedor_asignado) != _nombre_norm:
                    _omitidas_proveedor.append((clave, _proveedor_asignado))
                    continue

                if clave in _fila_por_clave:
                    fila = _fila_por_clave[clave]
                    _actualizadas += 1
                    # La tarifa trae el código proveedor vacío para este
                    # material -- si el catálogo ya tenía uno guardado de
                    # antes, se conserva (no se borra) y además se recupera
                    # para completar la propia tarifa (ver más abajo).
                    if cod_prov in (None, ""):
                        _cod_prov_existente = ws_cat.cell(row=fila, column=_col_cprov).value
                        if _cod_prov_existente not in (None, ""):
                            cod_prov = _cod_prov_existente
                            _completados_desde_catalogo.append((clave, cod_prov))
                else:
                    fila = _fila_siguiente
                    _fila_siguiente += 1
                    _anadidas += 1

                valores = {
                    "clave":       int(clave) if clave.isdigit() else clave,
                    "cod_prov":    cod_prov,
                    "descripcion": d.get("descripcion"), "unidad": d.get("unidad"),
                    "igic":        d.get("igic"),        "naturaleza": d.get("naturaleza"),
                    "familia":     d.get("familia"),     "subfamilia": d.get("subfamilia"),
                    "proveedor":   d.get("proveedor"),   "activo":     d.get("activo"),
                }
                for _k_sap in CATALOGO_CLAVES_SAP:
                    valores[_k_sap] = s.get(_k_sap)
                _precio = _precio_tarifa.get(clave)
                if es_precio_valido(_precio):
                    valores["precio"] = _precio
                    _precios_escritos += 1
                    if _col_precio:
                        _prev = ws_cat.cell(row=fila, column=_col_precio).value
                        if _prev not in (None, "") and _prev != _precio:
                            _precios_sustituidos += 1
                else:
                    _precios_invalidos.append(clave)
                # (cell(..., value=None) no borra: si un dato no viene, el que
                # ya hubiera en el catálogo se conserva -- comportamiento
                # previo que se mantiene, y es lo que protege el precio cuando
                # la tarifa no trae uno válido.)
                for _k, _v in valores.items():
                    _col = _mapa.get(_k)
                    if _col:
                        ws_cat.cell(row=fila, column=_col, value=_v)

            wb_cat.save(ruta_catalogo)
            wb_cat.close()

            self.log_consola(
                f"  └ Catálogo DALI actualizado desde la tarifa: "
                f"{_actualizadas} fila(s) modificadas, {_anadidas} fila(s) nuevas añadidas",
                "OK")
            if not _col_precio:
                self.log_consola(
                    "  ⚠  El catálogo no tiene la columna «Precio €» (formato antiguo): "
                    "los precios de la tarifa no se han guardado; se incorporarán "
                    "cuando se regenere el catálogo (Acción 2/3).", "WARN")
            else:
                self.log_consola(
                    f"  └ Precio €: {_precios_escritos} precio(s) pegado(s) desde la "
                    f"tarifa ({_precios_sustituidos} sustituyen un valor anterior "
                    f"distinto)", "OK")
                if _precios_invalidos:
                    _lista_inv = ", ".join(_precios_invalidos[:10]) + (
                        f" … y {len(_precios_invalidos) - 10} más"
                        if len(_precios_invalidos) > 10 else "")
                    self.log_consola(
                        f"  ⚠  {len(_precios_invalidos)} código(s) DALI sin precio "
                        f"válido en la tarifa (columna F vacía, 0 o no numérica) — se "
                        f"conserva el que ya tuviera el catálogo: DALI {_lista_inv}",
                        "WARN")

            # ── 5. Completar en la propia tarifa los códigos proveedor que ──
            #      faltaban y se han recuperado del catálogo, para que el
            #      Registro Info que se cree a partir de ella salga más
            #      completo.
            if _completados_desde_catalogo:
                try:
                    wb_tar2 = openpyxl.load_workbook(ruta_tarifa)
                    ws_tar2 = wb_tar2.active
                    for _clave_c, _cod_c in _completados_desde_catalogo:
                        _fila_c = _fila_tarifa_por_clave.get(_clave_c)
                        if _fila_c is not None:
                            ws_tar2.cell(row=_fila_c, column=3, value=_cod_c)
                    wb_tar2.save(ruta_tarifa)
                    wb_tar2.close()
                    self.log_consola(
                        f"  └ {len(_completados_desde_catalogo)} código(s) proveedor "
                        f"completados en la tarifa desde el catálogo DALI:", "OK")
                    for _clave_c, _cod_c in _completados_desde_catalogo[:30]:
                        self.log_consola(
                            f"      · DALI {_clave_c} → «{_cod_c}»", "OK")
                    if len(_completados_desde_catalogo) > 30:
                        self.log_consola(
                            f"      … y {len(_completados_desde_catalogo) - 30} más "
                            f"(omitidos del log)", "OK")

                    def _avisar_completados_desde_catalogo(
                            completados=_completados_desde_catalogo):
                        lineas = [
                            f"✔  {len(completados)} código(s) proveedor completados",
                            "",
                            "Estos artículos venían sin Código Proveedor en la tarifa,",
                            "pero el catálogo DALI completo ya tenía uno guardado de",
                            "antes -- se ha recuperado y escrito también en la tarifa,",
                            "para que el Registro Info quede más completo:",
                            "",
                        ]
                        for _clave_c, _cod_c in completados[:20]:
                            lineas.append(f"   • DALI {_clave_c} → «{_cod_c}»")
                        if len(completados) > 20:
                            lineas.append(f"   … y {len(completados) - 20} más (ver log de consola).")
                        messagebox.showinfo(
                            "Códigos proveedor completados desde el catálogo",
                            "\n".join(lineas))
                    self.after(0, _avisar_completados_desde_catalogo)
                except Exception as _e_completar:
                    self.log_consola(
                        f"  ⚠  No se pudieron escribir los códigos proveedor "
                        f"recuperados en la tarifa: {_e_completar}", "WARN")

            if _omitidas_proveedor:
                self.log_consola(
                    f"  ⚠  CASO IRREGULAR: {len(_omitidas_proveedor)} código(s) DALI de la "
                    f"tarifa NO tienen «{_nombre_prov_tarifa}» como Proveedor Asignado en "
                    f"ARTICULOS_DALI.xlsx — fila(s) NO modificada(s):",
                    "WARN")
                for _clave_om, _prov_om in _omitidas_proveedor[:30]:
                    self.log_consola(
                        f"      · DALI {_clave_om}: Proveedor Asignado real = "
                        f"«{_prov_om or '(vacío)'}»", "WARN")
                if len(_omitidas_proveedor) > 30:
                    self.log_consola(
                        f"      … y {len(_omitidas_proveedor) - 30} más (omitidos del log)", "WARN")

                def _avisar_proveedor_no_coincide(nombre=_nombre_prov_tarifa,
                                                   omitidas=_omitidas_proveedor):
                    lineas = [
                        f"⚠  CASO IRREGULAR — Proveedor Asignado no coincide",
                        "",
                        f"Esta tarifa es de «{nombre}», pero en ARTICULOS_DALI.xlsx los",
                        f"siguientes {len(omitidas)} código(s) DALI tienen asignado un",
                        "proveedor distinto. Normalmente ambas bases de datos de",
                        "proveedores deberían coincidir exactamente -- revise estos",
                        "códigos, puede tratarse de un desfase puntual entre listados:",
                        "",
                    ]
                    for _clave_om, _prov_om in omitidas[:20]:
                        lineas.append(f"   • DALI {_clave_om}: proveedor real = «{_prov_om or '(vacío)'}»")
                    if len(omitidas) > 20:
                        lineas.append(f"   … y {len(omitidas) - 20} más (ver log de consola).")
                    lineas += ["", "No se ha modificado su fila en catalogo-dali-completo.xlsx."]
                    messagebox.showwarning(
                        "Caso irregular — Proveedor Asignado no coincide", "\n".join(lineas))
                self.after(0, _avisar_proveedor_no_coincide)

        except Exception as e_cat:
            self.log_consola(
                f"  ⚠  No se pudo actualizar catalogo-dali-completo.xlsx: {e_cat}", "WARN")

    def _actualizar_listado_sap(self):
        """
        PUNTO 3 — Actualizar listado materiales SAP.

        Origen 1 : RECURSOS/Productos.xlsx           (fecha = última modificación del archivo)
        Origen 2 : RECURSOS/ARTICULOS_DALI.xlsx       (fecha = celda B3, texto "ACTUALIZADO …")
        Origen 3 : RECURSOS/Listas de productos.xlsx
        Destino  : <raíz proyecto>/LISTADO CODIGOS DALI - SAP.xlsx

        Cruce 1 — col D destino ↔ col B Productos.xlsx:
          C (Descripción producto)              → E  (DESCRIPCION SAP)
          J (Texto UM expl.)                    → F  (nueva)
          I (Unidad medida base)                → G  (UNIDAD SAP)
          F (Grupo de productos)                → H  (GRUPO PRODUCTOS SAP)
          G (Descripción de grupo de productos) → I  (Descripción de grupo de productos SAP)
          D (Tipo de producto)                  → L  (nueva)

        Cruce 1b — col D destino ↔ col A Listas de productos.xlsx:
          F (Categoría valoración) → K (nueva). Si un código SAP trae
          valores distintos según el centro, se usa el más frecuente y
          se avisa al usuario (mensaje visual + log detallado).

        Cruce 2 — col A destino ↔ col A ARTICULOS_DALI.xlsx:
          E (SUBFAMILIA) → diccionario SUBFAMILIA → M (GRUPO COMPRAS SAP) + N (NOMBRE GRUPO COMPRAS SAP)
          G (IGIC)       → diccionario IGIC       → J (Indicador de impuestos SAP)
        """
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            messagebox.showwarning("Sin proyecto",
                                   "Seleccione primero la carpeta del Proyecto.")
            return

        if not self._confirmar_catalogo_actualizado():
            self.log_consola(
                "Actualización SAP cancelada por el usuario (catálogo desactualizado).", "WARN")
            return

        recursos_path  = os.path.join(carpeta, FOLDER_RECURSOS)
        ruta_productos = os.path.join(recursos_path, "Productos.xlsx")
        ruta_dali      = os.path.join(recursos_path, "ARTICULOS_DALI.xlsx")
        ruta_listas    = os.path.join(recursos_path, "Listas de productos.xlsx")
        ruta_destino   = os.path.join(carpeta, "LISTADO CODIGOS DALI - SAP.xlsx")

        # ── 1. Verificar archivos de origen ──────────────────────────
        faltantes = []
        if not os.path.isfile(ruta_productos):
            faltantes.append("Productos.xlsx")
        if not os.path.isfile(ruta_dali):
            faltantes.append("ARTICULOS_DALI.xlsx")
        if not os.path.isfile(ruta_listas):
            faltantes.append("Listas de productos.xlsx")

        if faltantes:
            # ── Aviso específico si falta Productos.xlsx ──────────────
            if "Productos.xlsx" in faltantes:
                messagebox.showwarning(
                    "Archivo Productos.xlsx no encontrado",
                    "❌  No se encontró el archivo  Productos.xlsx\n"
                    f"    en la carpeta RECURSOS.\n\n"
                    "📥  Descárguelo desde SAP siguiendo estos pasos:\n\n"
                    "  1.  En SAP, abra la carpeta  CREAR ARTÍCULOS\n"
                    "      (Aplicaciones → Favoritos)\n\n"
                    "  2.  Entre en  Gestionar datos maestros de producto\n\n"
                    "  3.  Haga clic en el icono  Exportar  (parte superior)\n"
                    "      y elija  Exportar como…\n\n"
                    "  4.  En el diálogo:\n"
                    "        • Nombre de fichero  →  Productos\n"
                    "        • Formato            →  Microsoft Excel (*.xlsx)\n"
                    "        • ☑  Dividir celdas con varios valores\n\n"
                    "  5.  Pulse  Exportar  y guarde el archivo descargado\n"
                    f"      dentro de la carpeta:\n"
                    f"      📁  {recursos_path}\n\n"
                    "      A continuación vuelva a ejecutar este proceso.")
                return
            # ── Resto de archivos faltantes (aviso genérico) ─────────
            otros = [f for f in faltantes if f != "Productos.xlsx"]
            if otros:
                messagebox.showwarning(
                    "Archivos no encontrados",
                    "❌  No se encontraron los siguientes archivos\n"
                    f"    en la carpeta RECURSOS:\n\n"
                    + "".join(f"   • {f}\n" for f in otros) +
                    f"\n📁  Ruta buscada:\n   {recursos_path}")
            return

        # ── 2. Obtener fechas de los archivos de origen ───────────────
        # Productos.xlsx → última modificación del archivo
        try:
            ts_prod  = os.path.getmtime(ruta_productos)
            dt_prod  = datetime.fromtimestamp(ts_prod)
            fecha_productos      = dt_prod.strftime("%d/%m/%Y %H:%M")
            fecha_productos_solo = dt_prod.strftime("%d/%m/%Y")   # para D4
            hora_productos_solo  = dt_prod.strftime("%H:%M")      # para D5
        except Exception:
            fecha_productos      = "(no disponible)"
            fecha_productos_solo = "(no disponible)"
            hora_productos_solo  = "(no disponible)"

        # ARTICULOS_DALI.xlsx → texto de celdas B3 y B5
        try:
            import openpyxl
            wb_tmp = openpyxl.load_workbook(ruta_dali, read_only=True, data_only=True)
            ws_tmp = wb_tmp.active
            celda_b3 = ws_tmp["B3"].value or ""
            wb_tmp.close()
            fecha_dali = str(celda_b3).strip() if celda_b3 else "(fecha no disponible)"
        except Exception as e:
            messagebox.showerror("Error al leer ARTICULOS_DALI",
                                 f"No se pudo leer la celda B3:\n{e}")
            return

        # ── 3. Confirmación al usuario ────────────────────────────────
        if not messagebox.askyesno(
                "Confirmar actualización SAP",
                "Se han localizado los archivos de origen:\n\n"
                f"   📄 Productos.xlsx\n"
                f"      Última modificación: {fecha_productos}\n\n"
                f"   📄 ARTICULOS_DALI.xlsx\n"
                f"      {fecha_dali}\n\n"
                f"   📄 Listas de productos.xlsx\n\n"
                "Destino:\n"
                f"   LISTADO CODIGOS DALI - SAP.xlsx\n\n"
                "¿Desea continuar?"):
            self.log_consola("Actualización SAP cancelada por el usuario.", "WARN")
            return

        # ── 4. Verificar archivo destino ──────────────────────────────
        if not os.path.isfile(ruta_destino):
            messagebox.showerror(
                "Destino no encontrado",
                "❌  No se encontró el archivo destino:\n\n"
                f"   {ruta_destino}\n\n"
                "Verifique que 'LISTADO CODIGOS DALI - SAP.xlsx'\n"
                "esté en la raíz de la carpeta del proyecto.")
            return

        # ── 5. Diccionarios de referencia (equivalentes al VBA) ───────
        # TABLA_SUBFAMILIA_GRUPO_COMPRAS / TABLA_IGIC_GRUPO_COMPRAS
        # (movidas a tarifa_constants.py — eran datos estáticos sin lógica,
        #  mismo patrón que DALI_CLAVES_SUBFAMILIA/DALI_CLAVES_IGIC)

        # ── 6. Proceso en hilo ────────────────────────────────────────
        def proceso():
            self.consola.separador("INICIO ACTUALIZACIÓN SAP")
            self.log_consola("Iniciando actualización listado materiales SAP...", "INFO")
            self.after(0, lambda: self._barra.indeterminate("Leyendo Productos.xlsx..."))

            try:
                # ── LECTURA 1: Productos.xlsx ─────────────────────────────────
                #    Extraido a _leer_productos_sap() -- ver ese metodo para el
                #    mapeo exacto de columnas (identico al que habia aqui antes).
                tabla_productos, filas_prod = self._leer_productos_sap(ruta_productos)
                if tabla_productos is None:
                    return

                # ── LECTURA 2: ARTICULOS_DALI.xlsx ───────────────────────────
                #    Extraido a _leer_articulos_dali_sap().
                tabla_dali = self._leer_articulos_dali_sap(ruta_dali)

                # ── LECTURA 3: Listas de productos.xlsx ──────────────────────
                #    Extraido a _leer_categoria_valoracion().
                tabla_categoria, conflictos_categoria = \
                    self._leer_categoria_valoracion(ruta_listas)

                # ── CRUCE Y ESCRITURA EN DESTINO ─────────────────────────────
                #    Extraido a _cruzar_y_escribir_listado_sap() -- ver ese
                #    metodo para el detalle (identico al que habia aqui antes).
                self._cruzar_y_escribir_listado_sap(
                    ruta_destino, tabla_productos, tabla_dali,
                    tabla_categoria, conflictos_categoria, filas_prod,
                    fecha_productos_solo, hora_productos_solo)

            except Exception as e:
                self.log_consola(f"ERROR: {e}", "ERROR")
                self.consola.separador("FIN ACTUALIZACIÓN SAP (con error)")
                self.after(0, lambda: self._barra.reset("Error en actualización SAP"))
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error en actualización SAP",
                    "✘  El proceso no se completó correctamente.\n\n"
                    f"Detalle:\n{err}\n\n"
                    "Revise la ventana de log para más información."))

        threading.Thread(target=proceso, daemon=True).start()

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
                from transformar_tarifa import transformar
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
                        "          para incorporar los nuevos datos al LISTADO.\n\n"
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

                    # Buscar ruta del listado DALI-SAP
                    # Orden: 1) raíz proyecto  2) RECURSOS/  3) carpeta del .xls  4) padre carpeta .xls
                    _carpeta_raiz = self._carpeta_proyecto.get().strip()
                    _carpeta_rec  = os.path.join(_carpeta_raiz, "RECURSOS")
                    _carpeta_xls  = os.path.dirname(os.path.abspath(self._ruta_resultado)) if self._ruta_resultado else ""
                    _carpeta_xls_padre = os.path.dirname(_carpeta_xls) if _carpeta_xls else ""
                    self.log_consola(f"  [DIAG] Buscando listado DALI-SAP en:", "INFO")
                    self.log_consola(f"    1) {_carpeta_raiz}", "INFO")
                    self.log_consola(f"    2) {_carpeta_rec}", "INFO")
                    self.log_consola(f"    3) {_carpeta_xls}", "INFO")
                    self.log_consola(f"    4) {_carpeta_xls_padre}", "INFO")
                    _ruta_listado = None
                    _dirs_busq = [d for d in [
                        _carpeta_raiz, _carpeta_rec,
                        _carpeta_xls, _carpeta_xls_padre
                    ] if d and os.path.isdir(d)]
                    for _dir_busq in _dirs_busq:
                        for _fn in os.listdir(_dir_busq):
                            if (_fn.lower().endswith(".xlsx") and
                                    "LISTADO" in _fn.upper() and
                                    "DALI"    in _fn.upper() and
                                    "SAP"     in _fn.upper()):
                                _ruta_listado = os.path.join(_dir_busq, _fn)
                                self.log_consola(f"  [DIAG] ✔ Encontrado: {_ruta_listado}", "OK")
                                break
                        if _ruta_listado:
                            break
                    if not _ruta_listado:
                        self.log_consola("  [DIAG] ✘ Listado no encontrado en ninguna ruta.", "WARN")

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
                                    catalogo_refresh_fn=self._refrescar_sap_catalogo_para_dali)
                                self.wait_window(ven)

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
                                "No se encontr\xf3 el archivo LISTADO CODIGOS DALI-SAP en RECURSOS.\n"
                                "A\xf1\xe1delo y vuelve a ejecutar las acciones.")
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
                      "pero sin DESCRIPCIÓN ni UNIDAD (columnas E y F vacías del LISTADO):",
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
        from openpyxl import load_workbook as _lw

        # ── 1. Localizar LISTADO CODIGOS DALI-SAP ────────────────
        _carpeta_raiz = self._carpeta_proyecto.get().strip()
        _carpeta_rec  = os.path.join(_carpeta_raiz, FOLDER_RECURSOS)
        _ruta_listado = None
        for _dir_busq in [_carpeta_raiz, _carpeta_rec]:
            if not os.path.isdir(_dir_busq):
                continue
            for _fn in os.listdir(_dir_busq):
                if (_fn.lower().endswith(".xlsx") and
                        "LISTADO" in _fn.upper() and
                        "DALI"    in _fn.upper() and
                        "SAP"     in _fn.upper()):
                    _ruta_listado = os.path.join(_dir_busq, _fn)
                    break
            if _ruta_listado:
                break

        if not _ruta_listado:
            self.log_consola("LISTADO CODIGOS DALI-SAP no encontrado.", "ERROR")
            self.after(0, lambda: self._barra.reset("Listado no encontrado"))
            self.after(0, lambda: messagebox.showerror(
                "Listado no encontrado",
                "❌  No se encontró el archivo\n"
                "   LISTADO CODIGOS DALI - SAP.xlsx\n\n"
                "Verifique que el archivo está en la raíz del proyecto\n"
                "o en la carpeta RECURSOS."))
            return None, None

        self.log_consola(f"  └ Listado: {os.path.basename(_ruta_listado)}", "OK")

        # ── 2. Leer tabla DALI → (SAP, desc_sap, unidad_sap) ─────
        self.log_consola("Construyendo tabla de cruce desde LISTADO...", "INFO")
        _tabla_listado = {}   # {cod_dali_int: (cod_sap, desc_sap, unidad_sap)}
        try:
            wb_lst = _lw(_ruta_listado, read_only=True, data_only=True)
            ws_lst = wb_lst.active
            _n_lst = 0
            for _row in ws_lst.iter_rows(min_row=8, values_only=True):
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
            wb_lst.close()
            self.log_consola(f"  └ {_n_lst} entradas leídas del LISTADO.", "OK")
        except Exception as _e_lst:
            self.log_consola(f"Error leyendo LISTADO: {_e_lst}", "ERROR")
            self.after(0, lambda: self._barra.reset("Error leyendo listado"))
            self.after(0, lambda err=str(_e_lst): messagebox.showerror(
                "Error al leer LISTADO", err))
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
                f"  ⚠  {len(_sin_en_lista)} DALI no encontrados en LISTADO: "
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
                f"  ⚠  {len(_sin_sap)} código(s) DALI sin SAP asignado en el LISTADO.",
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
                    f"⚠  {len(falt)} código(s) DALI aún sin SAP en el LISTADO:\n\n"
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
                        catalogo_refresh_fn=self._refrescar_sap_catalogo_para_dali)
                    self.wait_window(ven)

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
                f"✔  Descripciones SAP actualizadas desde el LISTADO.\n"
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
                    "Actualice el LISTADO con descripción y unidad SAP y ejecute de nuevo Acción 4.\n"
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
                self.consola.separador("REINSCRIPCIÓN SAP DESDE LISTADO (archivo INC)")
                self.log_consola(f"Archivo INC: {os.path.basename(ruta_inc)}", "INFO")
                self.after(0, lambda: self._barra.indeterminate("Leyendo LISTADO DALI-SAP..."))

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

    def _ejecutar_transformacion_inc(self, ruta_inc):
        """
        Caso 2 — el archivo INC *.xlsx ya está procesado (es un .xlsx generado
        por Acción 3). No se vuelve a transformar: se leen directamente los
        códigos DALI que tienen col A vacía (SAP no asignado) y se abre
        VentanaSAPFaltantes para completarlos.
        Al guardar ≥1 código el archivo se renombra con prefijo INC.
        Si se completan todos y no hay duplicados se quita también [INCOMPLETO].
        El archivo NO se elimina — es el resultado final.
        """
        self._habilitar_btn_procesar(False)
        self._btn_buscar.config(state="disabled")
        self._habilitar_btn_resultado(False)

        def proceso():
            try:
                from openpyxl import load_workbook as _lw

                self.consola.separador("REVISIÓN SAP FALTANTES (archivo INC)")
                self.log_consola(f"Leyendo: {os.path.basename(ruta_inc)}", "INFO")
                self.after(0, lambda: self._barra.indeterminate("Leyendo INC..."))

                # ── Leer SAP faltantes: filas donde col A está vacía y col B tiene DALI
                # Estructura xlsx generado: A=SAP, B=COD_DALI, C=COD_PROV, D=DESC, E=FORMATO
                # Los datos empiezan en fila 8 (filas 1-7 son cabecera)
                _falt_sap    = []   # [(fila_xlsx, cod_dali), ...]
                _dup_sap     = []
                _dup_dali    = []
                _hay_dups    = False

                try:
                    wb_inc = _lw(ruta_inc, read_only=True, data_only=True)
                    ws_inc = wb_inc.active
                    _vistos_sap  = {}
                    _vistos_dali = {}
                    for _row in ws_inc.iter_rows(min_row=8, values_only=False):
                        _val_a = _row[0].value if _row[0] else None  # SAP
                        _val_b = _row[1].value if len(_row) > 1 else None  # DALI
                        if _val_b is None:
                            continue
                        try:
                            _cod_dali = int(float(str(_val_b)))
                        except (ValueError, TypeError):
                            continue
                        _fila = _row[0].row
                        # SAP faltante: col A vacía o cero
                        if not _val_a or str(_val_a).strip() in ("", "0", "None"):
                            _falt_sap.append((_fila, _cod_dali))
                        else:
                            try:
                                _cod_sap = int(float(str(_val_a)))
                                if _cod_sap in _vistos_sap:
                                    if _cod_sap not in _dup_sap:
                                        _dup_sap.append(_cod_sap)
                                else:
                                    _vistos_sap[_cod_sap] = _fila
                            except (ValueError, TypeError):
                                pass
                        if _cod_dali in _vistos_dali:
                            if _cod_dali not in _dup_dali:
                                _dup_dali.append(_cod_dali)
                        else:
                            _vistos_dali[_cod_dali] = _fila
                    wb_inc.close()
                    _hay_dups = bool(_dup_sap or _dup_dali)
                except Exception as _e_lec:
                    self.log_consola(f"Error leyendo INC: {_e_lec}", "ERROR")
                    self.after(0, lambda: self._barra.reset("Error leyendo INC"))
                    self.after(0, lambda err=str(_e_lec): messagebox.showerror(
                        "Error al leer archivo INC",
                        f"No se pudo leer el archivo:\n{os.path.basename(ruta_inc)}\n\n{err}"))
                    return

                self.log_consola(
                    f"  └ {len(_falt_sap)} código(s) DALI sin SAP  |  "
                    f"Duplicados SAP: {len(_dup_sap)}  DALI: {len(_dup_dali)}", "INFO")
                self.after(0, lambda: self._barra.completado("Lectura INC completada"))
                self.after(0, lambda: self._habilitar_btn_resultado(True))

                if not _falt_sap:
                    # No hay faltantes — solo informar
                    self.consola.separador("FIN REVISIÓN SAP (sin faltantes)")
                    self.after(0, lambda: messagebox.showinfo(
                        "Sin SAP faltantes",
                        f"El archivo INC no tiene códigos DALI sin SAP asignado.\n\n"
                        f"El archivo no tiene códigos DALI sin SAP.\n\n✔  TARIFA PREPARADA PARA CREAR O MODIFICAR REGISTRO INFO SAP."))
                    return

                # ── Buscar listado DALI-SAP ───────────────────
                _carpeta_raiz = self._carpeta_proyecto.get().strip()
                _carpeta_rec  = os.path.join(_carpeta_raiz, "RECURSOS")
                _ruta_listado = None
                for _dir_busq in [_carpeta_raiz, _carpeta_rec]:
                    if not os.path.isdir(_dir_busq):
                        continue
                    for _fn in os.listdir(_dir_busq):
                        if (_fn.lower().endswith(".xlsx") and
                                "LISTADO" in _fn.upper() and
                                "DALI"    in _fn.upper() and
                                "SAP"     in _fn.upper()):
                            _ruta_listado = os.path.join(_dir_busq, _fn)
                            break
                    if _ruta_listado:
                        break

                # ── Leer descripciones del propio INC ─────────
                _descripciones = {}
                try:
                    wb_d = _lw(ruta_inc, read_only=True, data_only=True)
                    ws_d = wb_d.active
                    for _frow, _dcod in _falt_sap:
                        _descripciones[_dcod] = (
                            str(ws_d.cell(_frow, 4).value or "").strip(),
                            str(ws_d.cell(_frow, 5).value or "").strip(),
                        )
                    wb_d.close()
                except Exception:
                    pass

                self.consola.separador("FIN REVISIÓN SAP (INC)")

                if not _ruta_listado:
                    self.log_consola("Listado DALI-SAP no encontrado en RECURSOS.", "WARN")
                    self.after(0, lambda: messagebox.showwarning(
                        "Listado no encontrado",
                        f"{len(_falt_sap)} código(s) DALI sin SAP.\n\n"
                        "No se encontró LISTADO CODIGOS DALI-SAP en RECURSOS.\n"
                        "Añádelo y vuelve a ejecutar Acción 4."))
                    return

                # ── Mostrar diálogo SAP faltantes ─────────────
                def _mostrar_dialogo():
                    resp = messagebox.askyesnocancel(
                        "Códigos SAP faltantes",
                        f"Se encontraron {len(_falt_sap)} código(s) DALI sin código SAP.\n\n"
                        "[SÍ]       →  Introducir ahora desde la aplicación\n"
                        "[NO]       →  Abrir el archivo en Excel\n"
                        "[CANCELAR] →  Ignorar por ahora",
                        icon="warning")

                    if resp is True:
                        ven = VentanaSAPFaltantes(
                            self, _falt_sap, _ruta_listado,
                            consola_fn=self.log_consola,
                            descripciones=_descripciones,
                            ruta_tarifa=ruta_inc,
                            catalogo_refresh_fn=self._refrescar_sap_catalogo_para_dali)
                        self.wait_window(ven)

                        if ven.guardado and ven.escritos_en_tarifa >= 1:
                            _todos = (ven.escritos_en_tarifa == len(_falt_sap))
                            try:
                                import re as _re
                                _dir  = os.path.dirname(ruta_inc)
                                _base = os.path.basename(ruta_inc)
                                if _todos and not _hay_dups:
                                    # Completado: quitar [INCOMPLETO], añadir INC
                                    _nuevo = _re.sub(
                                        r'\[INCOMPLETO\]\s*', '', _base,
                                        flags=_re.IGNORECASE).strip()
                                    if not _nuevo.upper().startswith("INC "):
                                        _nuevo = "INC " + _nuevo
                                else:
                                    # Parcial: mantener [INCOMPLETO], añadir INC
                                    _nuevo = _base
                                    if not _nuevo.upper().startswith("INC "):
                                        _nuevo = "INC " + _nuevo
                                _ruta_nueva = os.path.join(_dir, _nuevo)
                                if _nuevo != _base:
                                    if os.path.exists(_ruta_nueva):
                                        os.remove(_ruta_nueva)
                                        self.log_consola(
                                            f"  🗑 Sustituido archivo anterior: {_nuevo}", "INFO")
                                    os.rename(ruta_inc, _ruta_nueva)
                                    self._ruta_resultado = _ruta_nueva
                                    self.log_consola(
                                        f"  ✔ Archivo renombrado → {_nuevo}", "OK")
                                    if _todos and not _hay_dups:
                                        # ── Enganche del catálogo DALI completo (2026-09-23) ──
                                        # Caso 2 (archivo INC ya procesado por Acción 3):
                                        # al quedar la tarifa completa aquí también hay que
                                        # actualizar el código proveedor del catálogo -- antes
                                        # este camino no lo hacía.
                                        if self._catalogo_reciente_o_confirmar():
                                            self._actualizar_codigo_proveedor_catalogo(_ruta_nueva)
                                        messagebox.showinfo(
                                            "Tarifa completada",
                                            f"✔  Todos los códigos SAP completados "
                                            f"y sin duplicados.\n\n"
                                            f"Archivo renombrado a:\n{_nuevo}\n\n"
                                            f"\n✔  TARIFA PREPARADA PARA CREAR O MODIFICAR REGISTRO INFO SAP.")
                                    else:
                                        _pend = len(_falt_sap) - ven.escritos_en_tarifa
                                        self.log_consola(
                                            f"  ⚠ {_pend} código(s) SAP aún pendiente(s)", "WARN")
                            except Exception as _e_ren:
                                self.log_consola(
                                    f"  ✘ Error al renombrar: {_e_ren}", "WARN")

                    elif resp is False:
                        try:
                            os.startfile(ruta_inc)
                        except Exception:
                            import subprocess
                            subprocess.Popen(["start", "", ruta_inc], shell=True)

                self.after(200, _mostrar_dialogo)

            except Exception as e:
                self.log_consola(f"ERROR en revisión INC: {e}", "ERROR")
                self.after(0, lambda: self._barra.reset("Error en revisión INC"))
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error", f"Se produjo un error:\n{err}"))
            finally:
                self.after(0, lambda: self._habilitar_btn_procesar(True))
                self.after(0, lambda: self._btn_buscar.config(state="normal"))

        threading.Thread(target=proceso, daemon=True).start()

    def _organizar_dali(self):
        """
        Acción 1 — Organizar artículos DALI (versión optimizada)
        Lee ARTICULOS1.xls y ARTICULOS2.xls desde RECURSOS/, cruza por CODIGO
        con pandas.merge en memoria y escribe el resultado de golpe con ws.append().
        Los estilos se aplican por rango, no celda a celda.
        """
        import threading

        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta or not os.path.isdir(carpeta):
            from tkinter import messagebox
            messagebox.showwarning("Sin proyecto",
                                   "Seleccione primero la carpeta del Proyecto.")
            return

        recursos_path = os.path.join(carpeta, "RECURSOS")
        file1  = os.path.join(recursos_path, "ARTICULOS1.xls")
        file2  = os.path.join(recursos_path, "ARTICULOS2.xls")
        output = os.path.join(recursos_path, "ARTICULOS_DALI.xlsx")

        def proceso():
            try:
                self.after(0, lambda: self._barra.indeterminate("Leyendo archivos..."))
                self.consola.separador("INICIO ORGANIZACIÓN DALI")

                # ── Validaciones ──────────────────────────────────────────────
                if not os.path.isdir(recursos_path):
                    raise Exception(f"Carpeta RECURSOS no encontrada en:\n{carpeta}")

                faltan_xls = not os.path.exists(file1) or not os.path.exists(file2)

                if faltan_xls:
                    # ── Comprobar si existe ARTICULOS_DALI.xlsx como alternativa ──
                    if os.path.isfile(output):
                        try:
                            import openpyxl as _oxl
                            _wb = _oxl.load_workbook(output, read_only=True, data_only=True)
                            _ws = _wb.active
                            _b3 = str(_ws["B3"].value or "").strip()
                            _b5 = str(_ws["B5"].value or "").strip()
                            _wb.close()
                            _info_fecha = f"{_b3}  {_b5}".strip() or "(fecha no disponible)"
                        except Exception:
                            _info_fecha = "(no se pudo leer la fecha)"

                        _faltantes = []
                        if not os.path.exists(file1):
                            _faltantes.append("ARTICULOS1.xls")
                        if not os.path.exists(file2):
                            _faltantes.append("ARTICULOS2.xls")
                        _msg_falta = "\n".join(f"   \u2022 {f}" for f in _faltantes)

                        _respuesta = [None]

                        def _preguntar(r=_respuesta):
                            from tkinter import messagebox
                            r[0] = messagebox.askyesno(
                                "Archivos ARTICULOS no encontrados",
                                f"\u274c  No se encontraron los siguientes archivos en RECURSOS/:\n\n"
                                f"{_msg_falta}\n\n"
                                f"Para obtenerlos, acceda a DALI \u2192 opci\u00f3n 331,\n"
                                f"descargue el listado con los filtros:\n\n"
                                f"   ARTICULOS1:  C\u00f3digo, Descripci\u00f3n, Unidad,\n"
                                f"               Naturaleza, Familia, Subfamilia\n\n"
                                f"   ARTICULOS2:  C\u00f3digo, Descripci\u00f3n, Unidad,\n"
                                f"               Impuesto Compra, Proveedor por Defecto,\n"
                                f"               No permitir compra\n\n"
                                f"\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\n"
                                f"Se ha encontrado un ARTICULOS_DALI.xlsx existente:\n\n"
                                f"   {_info_fecha}\n\n"
                                f"\u00bfDesea actualizar el listado art\u00edculos DALI ahora?",
                                icon="warning"
                            )

                        self.after(0, _preguntar)
                        # Esperar sin límite de tiempo a que el usuario responda el diálogo
                        import time as _time
                        while _respuesta[0] is None:
                            _time.sleep(0.1)

                        if _respuesta[0] is True:
                            # Usuario quiere actualizar → lanzar acción 2 directamente
                            self.log_consola("Lanzando actualizaci\u00f3n DALI (acci\u00f3n 2)...", "INFO")
                            self.consola.separador("FIN ORGANIZACIÓN DALI")
                            self.after(0, lambda: self._barra.reset(""))
                            self.after(100, self._actualizar_listado_dali)
                            return
                        else:
                            # Usuario cierra → volver a la interfaz sin hacer nada
                            self.log_consola("Operaci\u00f3n cancelada por el usuario.", "WARN")
                            self.consola.separador("FIN ORGANIZACIÓN DALI")
                            self.after(0, lambda: self._barra.reset("Cancelado"))
                            return
                    else:
                        # No hay .xls ni ARTICULOS_DALI.xlsx → informar y explicar descarga
                        _faltantes = []
                        if not os.path.exists(file1):
                            _faltantes.append("ARTICULOS1.xls")
                        if not os.path.exists(file2):
                            _faltantes.append("ARTICULOS2.xls")
                        _msg_falta = "\n".join(f"   \u2022 {f}" for f in _faltantes)
                        self.after(0, lambda mf=_msg_falta: messagebox.showerror(
                            "Archivos ARTICULOS no encontrados",
                            f"\u274c  No se encontraron en RECURSOS/:\n\n"
                            f"{mf}\n\n"
                            f"Tambi\u00e9n se verific\u00f3 que no existe un\n"
                            f"ARTICULOS_DALI.xlsx previo en RECURSOS/.\n\n"
                            f"\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\n"
                            f"\ud83d\udce5  C\u00f3mo descargar desde DALI (opci\u00f3n 131):\n\n"
                            f"  ARTICULOS1  \u2192  seleccione en este orden:\n"
                            f"     1. C\u00f3digo\n"
                            f"     2. Descripci\u00f3n\n"
                            f"     3. Unidad\n"
                            f"     4. Familia\n"
                            f"     5. Subfamilia\n"
                            f"     6. Naturaleza\n\n"
                            f"  ARTICULOS2  \u2192  seleccione en este orden:\n"
                            f"     1. C\u00f3digo\n"
                            f"     2. Descripci\u00f3n\n"
                            f"     3. Unidad\n"
                            f"     4. Impuesto Compra\n"
                            f"     5. Proveedor por Defecto\n"
                            f"     6. No permitir compra\n\n"
                            f"Guarde los archivos descargados en la carpeta:\n"
                            f"   \ud83d\udcc1  RECURSOS/\n"
                            f"con los nombres exactos  ARTICULOS1.xls  y  ARTICULOS2.xls.\n\n"
                            f"A continuaci\u00f3n vuelva a ejecutar la Acci\u00f3n 1."
                        ))
                        self.log_consola("Proceso cancelado: faltan ARTICULOS1/2.xls y no hay fichero previo.", "ERROR")
                        self.consola.separador("FIN ORGANIZACIÓN DALI")
                        self.after(0, lambda: self._barra.reset("Faltan archivos"))
                        return

                # ── Núcleo de datos (Tanda 1 de la reorganización) ─────────────
                #    Lectura, dedup de duplicados en origen, cruce por CODIGO
                #    y saneamiento — extraído a acciones/organizar_dali.py.
                #    Ver ese módulo para el detalle línea a línea (idéntico
                #    al que había aquí, solo con self.log_consola -> log).
                _resultado = organizar_dali_datos.procesar_datos_dali(
                    file1, file2, self.log_consola)

                df_final          = _resultado["df_final"]
                df1               = _resultado["df1"]
                _ultimo_cod_art1  = _resultado["ultimo_cod_art1"]
                _incidencias_dup  = _resultado["incidencias_dup"]
                output_cols       = _resultado["output_cols"]
                errores           = _resultado["errores_sin_igic"]

                # ── Volcado del log de incidencias + aviso al usuario ─────────
                #    Solo si se han detectado duplicados con datos DISTINTOS
                #    entre repeticiones (errores reales de origen, no el caso
                #    inofensivo de duplicado con datos idénticos).
                _ruta_log_incidencias = organizar_dali_datos.guardar_log_incidencias(
                    _incidencias_dup, recursos_path, self.log_consola)

                if _ruta_log_incidencias:
                    self.after(0, lambda ruta=_ruta_log_incidencias, n=len(_incidencias_dup): messagebox.showwarning(
                        "Incidencias detectadas en datos de origen",
                        f"⚠  Se han detectado {n} código(s) duplicado(s) con datos "
                        f"DISTINTOS entre repeticiones en ARTICULOS1/ARTICULOS2.\n\n"
                        f"Se ha conservado la 1ª aparición de cada uno para generar "
                        f"ARTICULOS_DALI.xlsx, pero debería revisar y corregir el "
                        f"dato correcto en DALI.\n\n"
                        f"Detalle guardado en:\n{ruta}"
                    ))

                # ── Generar Excel + limpieza + validaciones (Tanda 2) ──────────
                #    Extraído a acciones/organizar_dali.py: construcción del
                #    libro (cabecera, logo, tabla, estilos), guardado,
                #    limpieza de ARTICULOS1/2 y validación de diccionarios
                #    SAP (SUBFAMILIA/IGIC). La DECISIÓN de qué hacer si la
                #    validación falla (borrar de nuevo + avisar) se queda
                #    aquí, no en el módulo, por ser un efecto secundario serio.
                organizar_dali_datos.generar_excel_dali(
                    df_final, output_cols, output,
                    self.log_consola,
                    progreso=lambda t: self.after(0, lambda: self._barra.indeterminate(t)))

                n_filas = len(df_final)
                self.log_consola(f"Archivo generado: {os.path.basename(output)}", "OK")
                self.log_consola(f"  └ {n_filas} artículos transferidos", "OK")

                _eliminados_ok = organizar_dali_datos.eliminar_archivos_origen(
                    file1, file2, self.log_consola, motivo="limpieza")

                # ═══════════════════════════════════════════════════════════════
                #  VALIDACIÓN DE COLUMNAS E (SUBFAMILIA) y G (IGIC)
                #  Se verifica el % de coincidencia con los diccionarios SAP.
                #  Si alguna columna tiene coincidencia 0 % se lanza error y se
                #  advierte al usuario que revise los filtros de la opción 131.
                # ═══════════════════════════════════════════════════════════════
                try:
                    _val = organizar_dali_datos.validar_diccionarios_sap(
                        df_final, self.log_consola)

                    if _val["columnas_sin_match"]:
                        _cols_str = "\n".join(f"   \u2022 {c}" for c in _val["columnas_sin_match"])

                        # Eliminar ARTICULOS1/2 (proceso incorrecto)
                        organizar_dali_datos.eliminar_archivos_origen(
                            file1, file2, self.log_consola, motivo="match fallido")

                        self.after(0, lambda cs=_cols_str: messagebox.showerror(
                            "Acción 1 — Proceso NO correcto: match con diccionarios fallido",
                            "✘  El proceso de la Acción 1 NO se ha completado correctamente.\n\n"
                            "Motivo: Las siguientes columnas del archivo generado\n"
                            "no contienen valores reconocidos por los diccionarios SAP:\n\n"
                            f"{cs}\n\n"
                            "Esto indica que ARTICULOS1.xls fue descargado con un\n"
                            "orden o selección de campos incorrectos desde DALI.\n\n"
                            "────────────────────────────────────────────────\n"
                            "📥  Vuelva a descargar desde DALI (opción 131)\n"
                            "    con los campos en este orden exacto:\n\n"
                            "  ARTICULOS1  →  seleccione en este orden:\n"
                            "     1. Código\n"
                            "     2. Descripción\n"
                            "     3. Unidad\n"
                            "     4. Familia\n"
                            "     5. Subfamilia\n"
                            "     6. Naturaleza\n\n"
                            "  ARTICULOS2  →  seleccione en este orden:\n"
                            "     1. Código\n"
                            "     2. Descripción\n"
                            "     3. Unidad\n"
                            "     4. Impuesto Compra\n"
                            "     5. Proveedor por Defecto\n"
                            "     6. No permitir compra\n\n"
                            "Guarde los archivos en RECURSOS/ con los nombres\n"
                            "exactos  ARTICULOS1.xls  y  ARTICULOS2.xls\n"
                            "y vuelva a ejecutar la Acción 1.\n\n"
                            "⚠   Los archivos ARTICULOS1 y ARTICULOS2 han sido\n"
                            "   eliminados de RECURSOS/ para forzar una nueva\n"
                            "   descarga correcta desde DALI."
                        ))
                except Exception as _e_val:
                    self.log_consola(
                        f"  ⚠  No se pudo completar la validación de columnas: {_e_val}", "WARN")

                # ── Detectar códigos faltantes entre 1 y el último de ARTICULOS1 ──
                saltos, faltantes_lista = organizar_dali_datos.detectar_codigos_faltantes(
                    df_final, _ultimo_cod_art1, self.log_consola)

                if errores:
                    self.log_consola(
                        f"  └ {errores} códigos sin coincidencia en ARTICULOS2", "WARN")

                self.log_consola(f"Ruta: {output}", "INFO")
                self.consola.separador("FIN ORGANIZACIÓN DALI")
                self.after(0, lambda: self._barra.completado(
                    f"DALI — {n_filas} artículos ✔"))

                # ── Construir mensaje para el usuario ─────────────────────────
                def _mostrar_resultado(n=n_filas, s=saltos, fl=faltantes_lista,
                                       e=errores, ruta=recursos_path,
                                       techo=_ultimo_cod_art1,
                                       elim=_eliminados_ok):
                    from tkinter import messagebox

                    # Construir texto del mensaje
                    lineas = [
                        "✔  Proceso completado correctamente",
                        "",
                        f"   Artículos transferidos:  {n}",
                    ]
                    if e:
                        lineas.append(f"   Sin coincidencia ARTICULOS2:  {e}")

                    if fl:
                        lineas.append("")
                        _techo_str = f"{techo:06d}" if techo else "??????"
                        lineas.append(f"✘  {len(fl)} código(s) faltante(s) entre 000001 y {_techo_str}:")
                        lineas.append("   Estos códigos de ARTICULOS1 no están en el resultado.")
                        for desde, hasta, num in s[:5]:
                            if desde == hasta:
                                lineas.append(f"      {desde:06d}  ({num} ausente)")
                            else:
                                lineas.append(f"      {desde:06d} – {hasta:06d}  ({num} ausente(s))")
                        if len(s) > 5:
                            lineas.append(f"      … y {len(s)-5} grupo(s) más")

                    lineas += [
                        "",
                        f"📁  Puede acceder al listado en:",
                        f"   {ruta}",
                    ]

                    if elim:
                        lineas += [
                            "",
                            "🗑  Archivos eliminados automáticamente de RECURSOS/:",
                        ] + [f"   • {f}" for f in elim]

                    tipo  = "error" if fl else "info"
                    titulo = "DALI — Proceso completado" + (" con errores" if fl else "")
                    msg = "\n".join(lineas)

                    if tipo == "error":
                        messagebox.showerror(titulo, msg)
                    else:
                        messagebox.showinfo(titulo, msg)

                self.after(0, _mostrar_resultado)

            except PermissionError as e:
                self.log_consola(f"ERROR: {e}", "ERROR")
                self.consola.separador("FIN ORGANIZACIÓN DALI (con error)")
                self.after(0, lambda: self._barra.reset("Error en DALI"))
                from tkinter import messagebox
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error en DALI — archivo en uso",
                    "\u2718  No se pudo guardar el archivo.\n\n"
                    "El archivo de destino está abierto en Excel (por ti o por otra "
                    "persona, ya que está en una carpeta compartida).\n\n"
                    "Cierra el archivo y vuelve a ejecutar la acción.\n\n"
                    f"Detalle técnico:\n{err}"))

            except Exception as e:
                self.log_consola(f"ERROR: {e}", "ERROR")
                self.consola.separador("FIN ORGANIZACIÓN DALI (con error)")
                self.after(0, lambda: self._barra.reset("Error en DALI"))
                from tkinter import messagebox
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error en DALI",
                    "\u2718  El proceso no se completó correctamente.\n\n"
                    f"Detalle del error:\n{err}\n\n"
                    "Revise la ventana de log para más información."))

        threading.Thread(target=proceso, daemon=True).start()


    def _abrir_resultado(self):
        if self._ruta_resultado and os.path.isfile(self._ruta_resultado):
            self.log_consola(f"Abriendo resultado: {self._ruta_resultado}", "INFO")
            try:
                os.startfile(self._ruta_resultado)
            except AttributeError:
                import subprocess
                subprocess.call(["xdg-open", self._ruta_resultado])
            except Exception as e:
                self.log_consola(f"Error al abrir: {e}", "ERROR")
                messagebox.showerror("Error", f"No se pudo abrir:\n{e}")
        else:
            messagebox.showwarning("Sin resultado",
                                   "A\xfan no se ha generado ning\xfan archivo procesado.")


# ---------------------------------------------------------
#  ENTRY POINT
# ---------------------------------------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()