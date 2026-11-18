# -*- coding: utf-8 -*-
"""
ventana_sin_conexion.py — aviso «No se ha podido acceder a Catálogo Asignaciones» (18 nov 2026, petición de Víctor).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Sale la primera vez que falla la conexión con la web al traer o enviar cambios (ver app/comun/estado_web.py, que
decide cuándo y explica la causa). Dice qué no se ha podido hacer y que la acción sigue con el catálogo Excel.
Botones: «Reintentar» → "reintentar"; «Seguir sin la web» (o cerrar) → "seguir". Para los fallos que no son de
conexión (credenciales, permiso) solo hay «Entendido» → "seguir".
Se llama siempre desde el hilo principal (PublicarCatalogoWeb._ejecutar_en_hilo_principal_y_esperar).
"""
import tkinter as tk

from app.comun.estado_web import TIPOS_SIN_REINTENTO
from app.comun.tarifa_constants import (
    C_ACCENT, C_ACCENT2, C_BORDER, C_DARK, C_HOVER, C_MID, C_MUTED, C_TEXT_LIGHT, C_WARNING,
    FONT_BOLD, FONT_LABEL, FONT_SMALL, FONT_TITLE,
)
from app.comun.tarifa_core import cargar_icono_ventana


class VentanaSinConexion(tk.Toplevel):
    def __init__(self, parent, que, info, detalle=""):
        super().__init__(parent)
        self.title("Catálogo Asignaciones")
        self.configure(bg=C_MID)
        self.transient(parent)
        self.resizable(False, False)
        cargar_icono_ventana(self)
        self.resultado = "seguir"
        reintentar = info.get("tipo") not in TIPOS_SIN_REINTENTO

        hdr = tk.Frame(self, bg=C_DARK, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"⚠  {info.get('titulo', 'Sin conexión con Catálogo Asignaciones')}", font=FONT_TITLE,
                 fg=C_WARNING, bg=C_DARK).pack(side="left", padx=14)

        cuerpo = tk.Frame(self, bg=C_MID, padx=16, pady=10)
        cuerpo.pack(fill="both", expand=True)
        tk.Label(cuerpo, text=f"No se ha podido {que}.", font=FONT_BOLD, fg=C_TEXT_LIGHT, bg=C_MID,
                 justify="left", wraplength=520).pack(anchor="w")
        tk.Label(cuerpo, text=info.get("explicacion", ""), font=FONT_LABEL, fg=C_TEXT_LIGHT, bg=C_MID,
                 justify="left", wraplength=520).pack(anchor="w", pady=(6, 0))
        tk.Label(cuerpo, text=("La Consola sigue trabajando con el catálogo Excel de RECURSOS. Lo que haya cambiado en "
                               "la web no está en ese Excel, y lo que cambie la Consola se enviará a la web la próxima "
                               "vez que conecte."),
                 font=FONT_LABEL, fg=C_ACCENT2, bg=C_MID, justify="left", wraplength=520).pack(anchor="w", pady=(8, 0))
        if detalle:
            tk.Label(cuerpo, text=f"Detalle: {detalle}", font=FONT_SMALL, fg=C_MUTED, bg=C_MID, justify="left",
                     wraplength=520).pack(anchor="w", pady=(8, 0))

        abajo = tk.Frame(self, bg=C_MID, padx=16, pady=12)
        abajo.pack(fill="x")
        if reintentar:
            tk.Button(abajo, text="↻  Reintentar", font=FONT_BOLD, fg=C_DARK, bg=C_ACCENT,
                      activebackground=C_ACCENT2, relief="flat", cursor="hand2",
                      command=self._reintentar).pack(side="left", ipadx=10, ipady=4)
        tk.Button(abajo, text="Seguir sin la web" if reintentar else "Entendido", font=FONT_BOLD, fg=C_TEXT_LIGHT,
                  bg=C_BORDER, activebackground=C_HOVER, relief="flat", cursor="hand2",
                  command=self._seguir).pack(side="left", padx=8, ipadx=10, ipady=4)

        self.protocol("WM_DELETE_WINDOW", self._seguir)
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f"+{max(0, (sw - w) // 2)}+{max(5, (sh - h) // 3)}")
        self.lift()
        self.grab_set()

    def _reintentar(self):
        self.resultado = "reintentar"
        self.destroy()

    def _seguir(self):
        self.resultado = "seguir"
        self.destroy()


def preguntar(parent, que, info, detalle=""):
    """Enseña el aviso y espera. Devuelve "reintentar" o "seguir"."""
    v = VentanaSinConexion(parent, que, info, detalle)
    parent.wait_window(v)
    return v.resultado
