# -*- coding: utf-8 -*-
"""Pruebas del aviso «Faltan los listados de SAP actualizados» (13 nov 2026): app/comun/listados_sap.py y
app/ui/ventana_listados_sap.py, y su enganche en la app (App._avisar_listados_sap).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import tempfile
import time
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import listados_sap as L  # noqa: E402
from app.comun.tarifa_constants import resource_path  # noqa: E402

try:
    import tkinter as tk
    HAY_TK = True
except ImportError:          # pragma: no cover
    HAY_TK = False


class Logica(unittest.TestCase):
    def test_codigos_sin_datos(self):
        tabla = {"10104081": ("LUBINA", "", "KG", "", "", "")}
        self.assertEqual(L.codigos_sin_datos([10104081, "10104082", " 10104082 ", ""], tabla), ["10104082"])
        self.assertEqual(L.codigos_sin_datos([10104081], tabla), [])
        self.assertEqual(L.codigos_sin_datos([10104081, 5], None), ["10104081", "5"])     # sin Productos.xlsx

    def test_estado_de_los_archivos(self):
        with tempfile.TemporaryDirectory() as rec:
            antes = time.time() - 100
            ruta = os.path.join(rec, L.PRODUCTOS)
            open(ruta, "w").close()
            os.utime(ruta, (antes - 50, antes - 50))                  # el de siempre, sin descargar de nuevo
            e = L.estado_archivos(rec, antes)
            self.assertEqual(e, {L.PRODUCTOS: (True, False), L.LISTAS: (False, False)})
            self.assertEqual(L.sin_actualizar(e), [L.PRODUCTOS, L.LISTAS])
            os.utime(ruta, None)                                       # descargado de nuevo
            open(os.path.join(rec, L.LISTAS), "w").close()
            self.assertEqual(L.sin_actualizar(L.estado_archivos(rec, antes)), [])

    def test_guia_completa(self):
        self.assertEqual([x["app_sap"] for x in L.LISTADOS], ["Gestionar datos maestros de producto",
                                                               "Lista de productos"])
        for img in ["assets/guia_sap_apps.png", "assets/guia_sap_barra.png"] + [x["imagen"] for x in L.LISTADOS]:
            self.assertTrue(os.path.isfile(resource_path(img)), img)
        self.assertIn("PARCIALMENTE", L.AVISO_PARCIAL)


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class Ventana(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("sin pantalla")
        self.root.withdraw()
        self.tmp = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmp.cleanup()
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def _ventana(self):
        from app.ui.ventana_listados_sap import VentanaListadosSap
        return VentanaListadosSap(self.root, ["10104082"], self.tmp.name, ahora=time.time() - 5)

    def test_mas_tarde_avisa_de_la_web_parcial(self):
        v = self._ventana()
        self.assertEqual(len(v._imagenes), 4)                       # las cuatro imágenes de la guía
        with mock.patch("app.ui.ventana_listados_sap.messagebox.showwarning") as aviso:
            v._mas_tarde()
        self.assertIn("PARCIALMENTE", aviso.call_args[0][1])
        self.assertEqual(v.resultado, "despues")

    def test_ya_estan_sin_descargar_pregunta(self):
        v = self._ventana()
        with mock.patch("app.ui.ventana_listados_sap.messagebox.askyesno", return_value=False) as pregunta:
            v._ya_estan()
        pregunta.assert_called_once()
        self.assertTrue(v.winfo_exists())                          # dijo que no: sigue la ventana
        for n in (L.PRODUCTOS, L.LISTAS):
            open(os.path.join(self.tmp.name, n), "w").close()
        with mock.patch("app.ui.ventana_listados_sap.messagebox.askyesno") as pregunta:
            v._ya_estan()
        pregunta.assert_not_called()                               # los dos descargados de nuevo: no pregunta
        self.assertEqual(v.resultado, "actualizar")


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class EngancheApp(unittest.TestCase):
    """App._avisar_listados_sap: solo avisa si algún código no está en Productos.xlsx, y con «actualizar» lanza
    la Acción 3."""

    def _app(self, carpeta):
        from app.principal.accion_listado_sap import AccionListadoSap
        app = object.__new__(AccionListadoSap)
        app._carpeta_proyecto = mock.Mock(get=lambda: carpeta)
        app.log_consola = mock.Mock()
        app.wait_window = lambda v: None
        app.after = mock.Mock()
        app._actualizar_listado_sap = mock.Mock()
        return app

    def test_no_avisa_si_estan_en_productos_y_lanza_la_accion_3(self):
        with tempfile.TemporaryDirectory() as carpeta:
            os.makedirs(os.path.join(carpeta, "RECURSOS"))
            app = self._app(carpeta)
            tabla = {"10104081": ("LUBINA", "", "KG", "", "", "")}
            open(os.path.join(carpeta, "RECURSOS", L.PRODUCTOS), "w").close()
            with mock.patch("app.principal.accion_listado_sap.leer_productos_xlsx", return_value=(tabla, [])), \
                    mock.patch("app.ui.ventana_listados_sap.VentanaListadosSap") as V:
                self.assertIsNone(app._avisar_listados_sap([10104081]))
                V.assert_not_called()
                V.return_value.resultado = "actualizar"
                self.assertEqual(app._avisar_listados_sap([10104081, 10104082]), "actualizar")
            self.assertEqual(V.call_args[0][1], ["10104082"])
            app.after.assert_called_once_with(300, app._actualizar_listado_sap)


if __name__ == "__main__":
    unittest.main()
