# -*- coding: utf-8 -*-
"""Pruebas del inicio de sesión al abrir la Consola (v10.2.50): ventana app/ui/ventana_login.py y
_iniciar_sesion_consola (app/principal/catalogo_dali.py). La web se simula; sin red. Necesitan tkinter y
pantalla (en Linux, xvfb-run).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    import tkinter as tk
    from app.ui import ventana_login as VL
    from app.principal import catalogo_dali as CD
    _r = tk.Tk()
    _r.destroy()
    HAY_TK = True
except Exception:  # sin tkinter o sin pantalla
    HAY_TK = False


def _hijos(w):
    for h in w.winfo_children():
        yield h
        yield from _hijos(h)


@unittest.skipUnless(HAY_TK, "necesita tkinter y pantalla")
class TestVentanaLogin(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()

    def tearDown(self):
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def _ventana(self):
        return [w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel)][0]

    def _rellenar_y_entrar(self, email, clave, despues):
        def paso():
            win = self._ventana()
            entradas = [h for h in _hijos(win) if isinstance(h, tk.Entry)]
            entradas[0].delete(0, "end")
            entradas[0].insert(0, email)
            entradas[1].insert(0, clave)
            boton = [h for h in _hijos(win) if isinstance(h, tk.Button) and h["text"] == "Entrar"][0]
            boton.invoke()
            self.root.after(600, lambda: despues(win))
        self.root.after(200, paso)

    def _mensaje(self, win):
        return " ".join(h["text"] for h in _hijos(win) if isinstance(h, tk.Label))

    def test_credenciales_guardadas_entran_solas(self):
        usuario = {"nombre": "Ana", "rol": "admin"}
        r = VL.pedir_inicio_sesion(self.root, "a@x", "clave", lambda e, p: ("ok", "ok", usuario))
        self.assertEqual(r, ("a@x", "clave", usuario))

    def test_no_autorizado_muestra_mensaje_y_olvida_las_guardadas(self):
        vistos, olvidadas = [], []

        def comprobar(e, p):
            return ("no_autorizado", "Usuario no autorizado.", None) if e == "luis@x" else ("ok", "", {"rol": "admin"})

        def despues(win):
            vistos.append(self._mensaje(win))
            [h for h in _hijos(win) if isinstance(h, tk.Button) and h["text"] == "Salir"][0].invoke()
        self._rellenar_y_entrar("luis@x", "clave", despues)
        r = VL.pedir_inicio_sesion(self.root, None, None, comprobar, al_fallar_credenciales=lambda: olvidadas.append(1))
        self.assertIsNone(r)
        self.assertIn("Usuario no autorizado.", vistos[0])
        self.assertEqual(olvidadas, [1])

    def test_sin_conexion_no_olvida_y_se_puede_reintentar(self):
        llamadas, olvidadas = [], []

        def comprobar(e, p):
            llamadas.append(e)
            return ("sin_conexion", "No hay conexión", None) if len(llamadas) == 1 else ("ok", "", {"rol": "admin"})

        def despues(win):
            self.assertIn("No hay conexión", self._mensaje(win))
            entradas = [h for h in _hijos(win) if isinstance(h, tk.Entry)]
            entradas[1].insert(0, "clave")
            [h for h in _hijos(win) if isinstance(h, tk.Button) and h["text"] == "Entrar"][0].invoke()
        self._rellenar_y_entrar("a@x", "clave", despues)
        r = VL.pedir_inicio_sesion(self.root, None, None, comprobar, al_fallar_credenciales=lambda: olvidadas.append(1))
        self.assertEqual(r[0], "a@x")
        self.assertEqual(olvidadas, [])


@unittest.skipUnless(HAY_TK, "necesita tkinter y pantalla")
class TestIniciarSesionConsola(unittest.TestCase):
    def _app(self):
        app = object.__new__(CD.CatalogoDali)
        app.config_data = {}
        app.logs = []
        app.log_consola = lambda m, n=">>": app.logs.append((n, m))
        app.destroy = mock.Mock()
        app.deiconify = mock.Mock()
        app.lift = mock.Mock()
        app._tras_iniciar_sesion = mock.Mock()
        return app

    def test_entra_guarda_y_sigue(self):
        app = self._app()
        with mock.patch.object(VL, "pedir_inicio_sesion", return_value=("a@x", "c", {"nombre": "Ana"})), \
                mock.patch.object(CD, "guardar_credenciales_dali", return_value=True) as guardar, \
                mock.patch.object(CD, "guardar_config"):
            app._iniciar_sesion_consola()
        guardar.assert_called_once_with({}, "a@x", "c")
        app.deiconify.assert_called_once()
        app._tras_iniciar_sesion.assert_called_once()
        app.destroy.assert_not_called()
        self.assertEqual(app._obtener_credenciales_dali_app(), ("a@x", "c"))

    def test_salir_cierra_la_consola(self):
        app = self._app()
        with mock.patch.object(VL, "pedir_inicio_sesion", return_value=None):
            app._iniciar_sesion_consola()
        app.destroy.assert_called_once()
        app.deiconify.assert_not_called()
        app._tras_iniciar_sesion.assert_not_called()


if __name__ == "__main__":
    unittest.main()
