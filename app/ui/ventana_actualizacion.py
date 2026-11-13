# -*- coding: utf-8 -*-
"""
ventana_actualizacion.py
Aviso «Nueva versión disponible» de la actualización automática (nov 2026). La lógica (consulta,
descarga, SHA-256 y sustitución del .exe) está en app/comun/actualizaciones.py.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import threading
import tkinter as tk

from app.comun import actualizaciones as act
from app.comun.tarifa_constants import (
    C_BG, C_MID, C_MUTED, C_SUCCESS, C_TEXT, FONT_BOLD, FONT_LABEL, FONT_SMALL, FONT_SUB, FONT_TITLE,
)
from app.comun.tarifa_core import cargar_icono_ventana, centrar_sobre_padre


def _notas_legibles(texto):
    """Las novedades que genera GitHub vienen en Markdown; se dejan como texto llano legible."""
    import re
    lineas = []
    saltar = False
    for l in (texto or "").splitlines():
        l = l.rstrip()
        if l.startswith("**Full Changelog**"):
            continue
        if l.startswith("#"):
            l = l.lstrip("#").strip()
            saltar = l.lower().startswith("new contributors")
            if saltar:
                continue
            l = {"what's changed": "Cambios"}.get(l.lower(), l)
        elif saltar:
            continue
        if l.startswith(("* ", "- ")):
            # «* Título by @usuario in https://…/pull/8» → «•  Título»
            l = "•  " + re.sub(r"\s+by\s+@\S+(\s+in\s+\S+)?$", "", l[2:])
        lineas.append(l.replace("**", ""))
    return "\n".join(lineas).strip() or "Sin notas para esta versión."


class VentanaActualizacion(tk.Toplevel):
    """Avisa de la versión nueva. «Actualizar ahora» descarga, comprueba, sustituye el .exe, abre la
    versión nueva y cierra esta. `log(mensaje, nivel)` escribe en la consola de la app."""

    def __init__(self, app, info, version_local, log):
        super().__init__(app)
        self._app, self._info, self._log = app, info, log
        self.title("Nueva versión disponible")
        self.configure(bg=C_BG)
        self.resizable(False, False)
        cargar_icono_ventana(self)

        hdr = tk.Frame(self, bg=C_MID, pady=14, padx=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text="\U0001f195  Nueva versión de la Consola Registro Info",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr, text=f"Tienes la {version_local.strip()}  →  nueva: v {info['version']}",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        cuerpo = tk.Frame(self, bg=C_BG, padx=16, pady=12)
        cuerpo.pack(fill="both", expand=True)
        tk.Label(cuerpo, text="Novedades", bg=C_BG, fg=C_TEXT, font=FONT_BOLD).pack(anchor="w")
        txt = tk.Text(cuerpo, width=72, height=12, wrap="word", font=FONT_LABEL, relief="solid", bd=1)
        txt.insert("1.0", _notas_legibles(info.get("notas")))
        txt.configure(state="disabled")
        txt.pack(fill="both", expand=True, pady=(4, 8))
        self._estado = tk.Label(cuerpo, text="Tus datos y la carpeta del proyecto no se tocan.",
                                bg=C_BG, fg=C_MUTED, font=FONT_SMALL, justify="left", wraplength=520)
        self._estado.pack(anchor="w")

        botones = tk.Frame(self, bg=C_BG, padx=16, pady=12)
        botones.pack(fill="x")
        self._btn_ok = tk.Button(botones, text="⬇  Actualizar ahora", command=self._actualizar,
                                 bg=C_SUCCESS, fg="white", font=FONT_BOLD, relief="flat", padx=14, pady=6,
                                 cursor="hand2")
        self._btn_ok.pack(side="right")
        self._btn_no = tk.Button(botones, text="Más tarde", command=self.destroy, font=FONT_LABEL,
                                 relief="flat", padx=14, pady=6, cursor="hand2")
        self._btn_no.pack(side="right", padx=(0, 8))
        if not act.puede_actualizarse():
            self._btn_ok.configure(state="disabled")
            self._estado.configure(text="Esta copia no se puede actualizar sola (no es el .exe o no se puede "
                                        "escribir en su carpeta). Descárgala de la página de Releases.")
        if app.winfo_viewable():
            self.transient(app)
            centrar_sobre_padre(self, app)
        else:
            # (v10.2.50) Puede llegar mientras la ventana principal está oculta, con el inicio de sesión abierto:
            # así se ve igual y se puede actualizar aunque no se haya podido entrar.
            self.update_idletasks()
            w, h = self.winfo_reqwidth(), self.winfo_reqheight()
            self.geometry(f"+{(self.winfo_screenwidth() - w) // 2}+{max(40, (self.winfo_screenheight() - h) // 4)}")
            self.lift()
            self.attributes("-topmost", True)
            self.after(500, lambda: self.winfo_exists() and self.attributes("-topmost", False))

    def _poner_estado(self, texto):
        if self.winfo_exists():
            self._estado.configure(text=texto)

    def _actualizar(self):
        self._btn_ok.configure(state="disabled")
        self._btn_no.configure(state="disabled")
        ruta_exe = os.path.abspath(sys.executable)
        self._log(f"Actualización: descargando v {self._info['version']}...", "INFO")

        def progreso(hecho, total):
            txt = (f"Descargando... {hecho / 1048576:.1f} / {total / 1048576:.1f} MB" if total
                   else f"Descargando... {hecho / 1048576:.1f} MB")
            self.after(0, lambda: self._poner_estado(txt))

        def hilo():
            try:
                nuevo = act.descargar_y_verificar(self._info, ruta_exe, progreso)
            except Exception as e:  # noqa: BLE001
                msg = str(e)
                self.after(0, lambda: self._fallo(msg))
                return
            self.after(0, lambda: self._instalar(nuevo, ruta_exe))
        threading.Thread(target=hilo, daemon=True, name="Actualizacion").start()

    def _fallo(self, mensaje):
        self._log(f"Actualización no completada: {mensaje}", "ERROR")
        self._poner_estado(f"⚠  {mensaje}")
        self._btn_no.configure(state="normal", text="Cerrar")

    def _instalar(self, nuevo, ruta_exe):
        self._poner_estado("\U0001f512  Integridad verificada (SHA-256). Instalando y reiniciando...")
        try:
            act.sustituir_exe(nuevo, ruta_exe)
            act.lanzar(ruta_exe)
        except act.ExeEnUso:
            # Unidad de red: no deja renombrar el .exe abierto ([WinError 32]). Se termina al cerrar.
            try:
                act.sustituir_al_cerrar(nuevo, ruta_exe)
            except Exception as e:  # noqa: BLE001
                self._fallo(f"No se pudo preparar la actualización al cerrar: {e}")
                return
            self._log(f"Actualización a v {self._info['version']}: la carpeta del programa no deja cambiarlo "
                      "mientras está abierto (unidad de red). La Consola se cierra, se actualiza y se vuelve "
                      "a abrir sola en unos segundos.", "OK")
            self._poner_estado("Cerrando para terminar de actualizar: la Consola se volverá a abrir sola "
                               "en unos segundos.")
            self.after(2500, lambda: (self._app.destroy(), os._exit(0)))
            return
        except Exception as e:  # noqa: BLE001
            self._fallo(str(e))
            return
        self._log(f"Actualizada a v {self._info['version']}: se abre la versión nueva.", "OK")
        self.after(800, lambda: (self._app.destroy(), os._exit(0)))
