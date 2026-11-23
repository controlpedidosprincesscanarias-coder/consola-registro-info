# -*- coding: utf-8 -*-
"""(23 nov 2026, petición de Víctor) Al pulsar la Acción 3 la barra de progreso se pone en marcha (amarilla) en el
acto, antes de traer lo cambiado en la web y leer el catálogo: antes seguía en verde y parecía bloqueada.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from app.principal.accion_tarifa import AccionesTarifa
    HAY = True
except Exception:  # sin tkinter
    HAY = False


@unittest.skipUnless(HAY, "necesita tkinter")
class BarraEnMarchaDesdeElPrincipio(unittest.TestCase):
    def _app(self):
        app = object.__new__(AccionesTarifa)
        app.orden = []
        app._barra = mock.Mock()
        app._barra.indeterminate.side_effect = lambda t: app.orden.append(("barra", t))
        app._habilitar_btn_procesar = mock.Mock()
        app._habilitar_btn_resultado = mock.Mock()
        app._btn_buscar = mock.Mock()
        return app

    def _hilo(self, app):
        return mock.patch("threading.Thread", side_effect=lambda *a, **k: mock.Mock(
            start=lambda: app.orden.append(("hilo", None))))

    def test_transformacion(self):
        app = self._app()
        with self._hilo(app):
            app._ejecutar_transformacion()
        self.assertEqual(app.orden[0][0], "barra")
        self.assertIn("Catálogo Asignaciones", app.orden[0][1])
        self.assertEqual(app.orden[-1][0], "hilo")

    def test_reinscripcion_incompleto(self):
        import inspect
        nombre = next(n for n, f in inspect.getmembers(AccionesTarifa, inspect.isfunction)
                      if 'self._traer_cambios_web("la reinscripción SAP")' in inspect.getsource(f))
        app = self._app()
        firma = inspect.signature(getattr(AccionesTarifa, nombre))
        args = ["X"] * (len(firma.parameters) - 1)
        with self._hilo(app):
            getattr(app, nombre)(*args)
        self.assertEqual(app.orden[0][0], "barra")
        self.assertEqual(app.orden[-1][0], "hilo")


if __name__ == "__main__":
    unittest.main()
