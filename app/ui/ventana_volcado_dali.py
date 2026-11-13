# -*- coding: utf-8 -*-
"""
ventana_volcado_dali.py — final de la Acción 1 (13 nov 2026): resumen, revisar ARTICULOS_DALI.xlsx y actualizar
Catálogo Asignaciones.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

1. Muestra el resumen de la Acción 1 y la pregunta «¿Desea revisar ARTICULOS_DALI.xlsx en Excel…?» con Sí / No.
2. El botón «Actualizar Catálogo Asignaciones» está oculto hasta contestar. Con «Sí» se abre el Excel
   (abrir_con_su_programa) y el botón aparece en esta misma ventana, para pulsarlo cuando se haya guardado y cerrado
   Excel; con «No» aparece directamente.
3. Al pulsarlo se comprueba el archivo (volcado_dali.estado_para_actualizar): si sigue abierto en Excel, no se sigue.
   resultado = "actualizar" y se cierra; la app pasa ARTICULOS_DALI.xlsx al catálogo y lo envía a la web.
4. «Cerrar sin actualizar» (o la X) → resultado "despues", con el aviso de que no se ha actualizado nada.
La lógica y los textos están en app/comun/volcado_dali.py.
"""
import tkinter as tk
from tkinter import messagebox

from app.comun import volcado_dali as V
from app.comun.tarifa_constants import (
    C_ACCENT, C_ACCENT2, C_BORDER, C_CARD, C_DARK, C_ERROR, C_HOVER, C_MID, C_MUTED, C_SUCCESS, C_TEXT_LIGHT,
    C_WARNING, FONT_BOLD, FONT_LABEL, FONT_TITLE,
)
from app.comun.tarifa_core import abrir_con_su_programa, cargar_icono_ventana

_COLOR_TIPO = {"ok": C_SUCCESS, "error": C_ERROR, "aviso": C_WARNING, "info": C_TEXT_LIGHT, "nota": C_MUTED}


class VentanaVolcadoDali(tk.Toplevel):
    """lineas: lista de (texto, tipo) con el resumen; tipo en ok / error / aviso / info / nota."""

    def __init__(self, parent, titulo, lineas, ruta_articulos, abrir=None):
        super().__init__(parent)
        self.title(titulo)
        self.configure(bg=C_MID)
        self.transient(parent)
        cargar_icono_ventana(self)
        self.resultado = "despues"
        self.revisado = None                 # None hasta contestar; True = «Sí»; False = «No»
        self._ruta = ruta_articulos
        self._mtime_al_abrir = None
        self._abrir = abrir or abrir_con_su_programa   # sustituible en las pruebas
        self._construir(titulo, lineas)
        self.protocol("WM_DELETE_WINDOW", self._cerrar_sin_actualizar)
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = min(max(self.winfo_reqwidth(), 640), sw - 40)
        h = min(self.winfo_reqheight() + 60, sh - 60)
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(5, (sh - h) // 4)}")
        self.grab_set()

    # ── Interfaz ─────────────────────────────────────────────────────────────────────────────────────
    def _boton(self, padre, texto, comando, principal=False):
        return tk.Button(padre, text=texto, font=FONT_BOLD, relief="flat", cursor="hand2", command=comando,
                         fg=C_DARK if principal else C_TEXT_LIGHT, bg=C_ACCENT if principal else C_BORDER,
                         activebackground=C_ACCENT2 if principal else C_HOVER)

    def _construir(self, titulo, lineas):
        hdr = tk.Frame(self, bg=C_DARK, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text=titulo, font=FONT_TITLE, fg=C_ACCENT, bg=C_DARK).pack(side="left", padx=14)

        resumen = tk.Frame(self, bg=C_CARD, padx=14, pady=10)
        resumen.pack(fill="x", padx=12, pady=(10, 6))
        for texto, tipo in lineas:
            tk.Label(resumen, text=texto, font=FONT_LABEL, fg=_COLOR_TIPO.get(tipo, C_TEXT_LIGHT), bg=C_CARD,
                     justify="left", anchor="w", wraplength=600).pack(fill="x", anchor="w")

        self._frm_pregunta = tk.Frame(self, bg=C_MID, padx=14, pady=6)
        self._frm_pregunta.pack(fill="x")
        tk.Label(self._frm_pregunta, text=V.PREGUNTA, font=FONT_BOLD, fg=C_TEXT_LIGHT, bg=C_MID, justify="left",
                 wraplength=600).pack(anchor="w", pady=(0, 6))
        fila = tk.Frame(self._frm_pregunta, bg=C_MID)
        fila.pack(anchor="w")
        self._btn_si = self._boton(fila, "Sí, abrir el Excel", self._si)
        self._btn_si.pack(side="left", ipadx=12, ipady=4)
        self._btn_no = self._boton(fila, "No", self._no)
        self._btn_no.pack(side="left", padx=8, ipadx=18, ipady=4)

        self._lbl_estado = tk.Label(self, text="", font=FONT_LABEL, fg=C_MUTED, bg=C_MID, justify="left",
                                    anchor="w", wraplength=600)
        self._lbl_estado.pack(fill="x", padx=14, pady=(4, 0))

        abajo = tk.Frame(self, bg=C_MID, pady=10)
        abajo.pack(fill="x", padx=14, side="bottom")
        # Oculto (sin pack) hasta contestar a la pregunta
        self._btn_actualizar = self._boton(abajo, "⬆  Actualizar Catálogo Asignaciones", self._actualizar,
                                           principal=True)
        self._abajo = abajo
        self._btn_cerrar = self._boton(abajo, "Cerrar sin actualizar", self._cerrar_sin_actualizar)
        self._btn_cerrar.pack(side="right", ipadx=10, ipady=4)

    def _mostrar_boton_actualizar(self, texto_estado):
        for b in (self._btn_si, self._btn_no):
            b.configure(state="disabled")
        self._lbl_estado.configure(text=texto_estado, fg=C_TEXT_LIGHT)
        if not self._btn_actualizar.winfo_manager():          # aún sin colocar (oculto)
            self._btn_actualizar.pack(side="left", ipadx=12, ipady=4)
        self._btn_actualizar.focus_set()

    # ── Acciones ─────────────────────────────────────────────────────────────────────────────────────
    def _si(self):
        self.revisado = True
        self._mtime_al_abrir = V.fecha_modificacion(self._ruta)
        try:
            self._abrir(self._ruta)
        except OSError as e:
            messagebox.showwarning("Abrir ARTICULOS_DALI.xlsx", f"No se pudo abrir el archivo en Excel:\n{e}\n\n"
                                   f"Ábrelo a mano desde:\n{self._ruta}", parent=self)
        self._mostrar_boton_actualizar(V.TEXTO_REVISANDO)

    def _no(self):
        self.revisado = False
        self._mostrar_boton_actualizar(V.TEXTO_SIN_REVISAR)

    def _actualizar(self):
        estado = V.estado_para_actualizar(self._ruta, bool(self.revisado), self._mtime_al_abrir)
        if estado in (V.ABIERTO, V.NO_EXISTE):
            messagebox.showwarning("ARTICULOS_DALI.xlsx", V.MENSAJES_ESTADO[estado], parent=self)
            self._lbl_estado.configure(text=V.MENSAJES_ESTADO[estado].split("\n")[0], fg=C_WARNING)
            return
        if estado == V.SIN_GUARDAR and not messagebox.askyesno(
                "ARTICULOS_DALI.xlsx", V.MENSAJES_ESTADO[estado], parent=self):
            return
        self.resultado = "actualizar"
        self.destroy()

    def _cerrar_sin_actualizar(self):
        messagebox.showwarning("Catálogo Asignaciones sin actualizar", V.AVISO_MAS_TARDE, parent=self)
        self.resultado = "despues"
        self.destroy()


def preguntar(parent, titulo, lineas, ruta_articulos):
    """Abre la ventana, espera a que se cierre y devuelve "actualizar" o "despues"."""
    ven = VentanaVolcadoDali(parent, titulo, lineas, ruta_articulos)
    parent.wait_window(ven)
    return ven.resultado
