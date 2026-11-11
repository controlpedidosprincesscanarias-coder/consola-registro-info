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

import tkinter as tk
import threading

from app.comun.tarifa_constants import (
    C_DARK, C_ACCENT, C_ACCENT2, C_ACCENT_DIM, C_BG, C_MUTED,
    SAP_LOGO_PNG,
)
from app.comun.tarifa_core import (
    resource_path,
    cargar_config, guardar_config,
)
from app.comun.credenciales import (
    migrar_credenciales_dali,
)

# ─── Versión de la aplicación ────────────────────────────────────────────────
APP_VERSION = "v 10.2.36"   # ← actualiza aquí cada vez que termines una fase

# DALI_APP_API_BASE (URL del Worker de Catálogo Asignaciones) vive en app/principal/catalogo_dali.py.
# ─────────────────────────────────────────────────────────────────────────────
from app.principal.interfaz import InterfazApp
from app.principal.accion_organizar_dali import AccionOrganizarDali
from app.principal.accion_listado_dali import AccionListadoDali
from app.principal.accion_listado_sap import AccionListadoSap
from app.principal.catalogo_dali import CatalogoDali
from app.principal.accion_tarifa import AccionesTarifa
from app.principal.catalogo_dali import DALI_APP_API_BASE  # noqa: F401


# (nov 2026, etapa 4) Los métodos de App están repartidos por temas en app/principal/ (mixins):
# interfaz, Acción 1 (organizar DALI), Acción 2 (listado DALI), Acción 3 (listado SAP), catálogo
# DALI y Acciones 4/6 (tarifa). Aquí quedan __init__, la cabecera y el pie (usan APP_VERSION).
class App(InterfazApp, AccionOrganizarDali, AccionListadoDali, AccionListadoSap, CatalogoDali, AccionesTarifa, tk.Tk):
    # SAP_LOGO_PNG (ruta de assets/logo_sap.png) centralizado en tarifa_constants.py

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
            from PIL import Image, ImageTk
            img = Image.open(resource_path("assets/logo.png"))   # va siempre dentro del .exe
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
            from PIL import Image as _Img, ImageTk as _ImgTk
            _sap_pil = _Img.open(SAP_LOGO_PNG).resize((72, 26), _Img.LANCZOS)
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


# ---------------------------------------------------------
#  ENTRY POINT
# ---------------------------------------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()