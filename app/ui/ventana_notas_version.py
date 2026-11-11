# -*- coding: utf-8 -*-
"""
ventana_notas_version.py
Novedades de la versión instalada (nov 2026), al pulsar el número de versión de la cabecera o del pie.
Petición de Víctor: en todas las aplicaciones se pueden ver las notas de la versión pulsando sobre ella.

Las notas son las de la Release de GitHub de esa versión (las genera el workflow a partir de los PR):
se piden en segundo plano, sin iniciar sesión, con actualizaciones.notas_de_version(), y se muestran con
el mismo formato que el aviso de versión nueva (_notas_legibles). Sin conexión, o desde el código (sin
Release), se dice y se ofrece la página de versiones.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import threading
import tkinter as tk
import webbrowser

from app.comun import actualizaciones as act
from app.comun.tarifa_constants import C_BG, C_MID, C_TEXT, FONT_LABEL, FONT_SUB, FONT_TITLE
from app.comun.tarifa_core import cargar_icono_ventana, centrar_sobre_padre
from app.ui.ventana_actualizacion import _notas_legibles

_ABIERTA = [None]


class VentanaNotasVersion(tk.Toplevel):
    def __init__(self, app, version):
        super().__init__(app)
        version = str(version).strip()
        self.title("Novedades de la versión")
        self.configure(bg=C_BG)
        self.resizable(False, False)
        cargar_icono_ventana(self)

        hdr = tk.Frame(self, bg=C_MID, pady=14, padx=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text="\U0001f4dc  Novedades de la versión", bg=C_MID, fg="white",
                 font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr, text=f"Consola Registro Info SAP  ·  {version}", bg=C_MID, fg="#BDD5F0",
                 font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        cuerpo = tk.Frame(self, bg=C_BG, padx=16, pady=12)
        cuerpo.pack(fill="both", expand=True)
        self._txt = tk.Text(cuerpo, width=72, height=14, wrap="word", font=FONT_LABEL, relief="solid", bd=1,
                            fg=C_TEXT)
        self._txt.pack(fill="both", expand=True)
        self._poner("Cargando las novedades…")

        botones = tk.Frame(self, bg=C_BG, padx=16, pady=12)
        botones.pack(fill="x")
        tk.Button(botones, text="Cerrar", command=self.destroy, font=FONT_LABEL, relief="flat", padx=14,
                  pady=6, cursor="hand2").pack(side="right")
        tk.Button(botones, text="\U0001f310  Todas las versiones", font=FONT_LABEL, relief="flat", padx=14,
                  pady=6, cursor="hand2", command=lambda: webbrowser.open(act.URL_VERSIONES)
                  ).pack(side="left")
        self.transient(app)
        centrar_sobre_padre(self, app)
        threading.Thread(target=self._cargar, args=(version,), daemon=True).start()

    def _poner(self, texto):
        self._txt.configure(state="normal")
        self._txt.delete("1.0", "end")
        self._txt.insert("1.0", texto)
        self._txt.configure(state="disabled")

    def _cargar(self, version):
        try:
            notas = act.notas_de_version(version)
            texto = (_notas_legibles(notas) if notas is not None else
                     f"La {version} no tiene notas publicadas (por ejemplo, si se abre desde el código).\n\n"
                     "Pulsa «Todas las versiones» para verlas en GitHub.")
        except Exception as e:  # noqa: BLE001
            texto = (f"No se han podido cargar las novedades ({e}).\n\n"
                     "Comprueba la conexión o pulsa «Todas las versiones».")
        try:
            self.after(0, lambda: self.winfo_exists() and self._poner(texto))
        except RuntimeError:
            pass                        # la ventana se cerró mientras se cargaba


def abrir(app, version):
    """Abre la ventana (o la trae al frente si ya está abierta)."""
    if _ABIERTA[0] is not None and _ABIERTA[0].winfo_exists():
        _ABIERTA[0].lift()
        return _ABIERTA[0]
    _ABIERTA[0] = VentanaNotasVersion(app, version)
    return _ABIERTA[0]


def hacer_pulsable(etiqueta, app, version):
    """Convierte la etiqueta con el número de versión en un enlace a sus novedades."""
    etiqueta.configure(cursor="hand2")
    etiqueta.bind("<Button-1>", lambda e: abrir(app, version))
