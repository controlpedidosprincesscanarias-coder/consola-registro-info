# -*- coding: utf-8 -*-
"""Pruebas de «Abrir en Excel» a la primera (13 nov 2026): app/comun/tarifa_core.abrir_con_su_programa y la ventana
de la Acción 6 (app/principal/accion_tarifa.py), que antes había que pulsar dos veces.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import time
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import tkinter as tk
    HAY_TK = True
except ImportError:          # pragma: no cover
    HAY_TK = False


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class AbrirConSuPrograma(unittest.TestCase):
    def test_permite_ponerse_delante_antes_de_abrir(self):
        from app.comun.tarifa_core import abrir_con_su_programa
        orden = []
        abrir_con_su_programa("T.xlsx", startfile=lambda r: orden.append(("abrir", r)),
                              permitir_delante=lambda: orden.append("permitir"))
        self.assertEqual(orden, ["permitir", ("abrir", "T.xlsx")])

    def test_fuera_de_windows_no_falla_el_permiso(self):
        from app.comun.tarifa_core import abrir_con_su_programa
        abiertos = []
        abrir_con_su_programa("T.xlsx", startfile=abiertos.append)     # sin ctypes.windll: se ignora
        self.assertEqual(abiertos, ["T.xlsx"])


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class VentanaAccion6(unittest.TestCase):
    def setUp(self):
        from app.principal.accion_tarifa import AccionesTarifa

        class App(AccionesTarifa, tk.Tk):
            pass
        try:
            self.app = App()
        except tk.TclError:
            self.skipTest("sin pantalla")
        self.app.withdraw()
        self.app.log_consola = mock.Mock()

    def tearDown(self):
        try:
            self.app.destroy()
        except tk.TclError:
            pass

    def test_un_solo_clic_abre_y_cierra_la_ventana(self):
        with mock.patch("app.principal.accion_tarifa.abrir_con_su_programa") as abrir:
            self.app._ventana_accion6_incompletos([{"nombre": "DEMO [INCOMPLETO] [13-11-2026].xlsx",
                                                     "ruta": "/tmp/DEMO.xlsx", "fecha": "13-11-2026 11:54"}])
            top = [w for w in self.app.winfo_children() if isinstance(w, tk.Toplevel)][0]
            boton = next(w for w in _todos(top) if isinstance(w, tk.Button) and "Abrir en Excel" in w.cget("text"))
            boton.invoke()
            fin = time.time() + 2
            while top.winfo_exists() and time.time() < fin:
                self.app.update()
                time.sleep(0.02)
        abrir.assert_called_once_with("/tmp/DEMO.xlsx")
        self.assertFalse(top.winfo_exists())


def _todos(w):
    for h in w.winfo_children():
        yield h
        yield from _todos(h)


if __name__ == "__main__":
    unittest.main()
