# -*- coding: utf-8 -*-
"""Pruebas de la lectura de «Listas de productos.xlsx» (13 nov 2026): columnas por cabecera, código entre
paréntesis y la protección que para la Acción 3 si el listado no casa con Productos.xlsx (ese día una exportación
con otras columnas dejó sin Categoría valoración a 7.159 materiales y lo envió a la web).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import openpyxl  # noqa: E402
from app.comun import sap_helpers as H  # noqa: E402

PRODUCTOS = {"10100050": ("QUESO BRIE",), "10100051": ("HELADO VAINILLA",), "10100052": ("HOJALDRES",)}


def _xlsx(carpeta, filas):
    ruta = os.path.join(carpeta, "Listas de productos.xlsx")
    wb = openpyxl.Workbook()
    for f in filas:
        wb.active.append(list(f))
    wb.save(ruta)
    return ruta


class LeerListas(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmp.cleanup()

    def test_formato_de_siempre(self):
        ruta = _xlsx(self.tmp.name, [("Producto", "B", "C", "D", "E", "Categoría valoración"),
                                     (10100050, "", "", "", "", "3101"), (10100050, "", "", "", "", "3101"),
                                     (10100051, "", "", "", "", "3106"), (10100051, "", "", "", "", "3101"),
                                     (10100051, "", "", "", "", "3106")])
        tabla, conflictos, n, por_cab = H.leer_listas_productos(ruta)
        self.assertEqual(tabla, {"10100050": "3101", "10100051": "3106"})
        self.assertEqual(conflictos, [("10100051", {"3106": 2, "3101": 1})])
        self.assertEqual(n, 5)
        self.assertTrue(por_cab)
        self.assertIsNone(H.comprobar_listas_productos(tabla, PRODUCTOS))

    def test_exportacion_con_otras_columnas_y_descripcion_clave(self):
        """La exportación del 13 nov 2026: «DESCRIPCIÓN (código)» y el centro donde antes iba la categoría."""
        cab = ("Producto", "Grupo de productos", "Tipo de producto", "Moneda", "Precio", "Centro",
               "Categoría valoración")
        filas = [cab, ("QUESO BRIE (10100050)", "", "", "EUR", 1, "GUAYARMINA PRINCESS (0110)", "Alimentación (3101)"),
                 ("QUESO BRIE (10100050)", "", "", "EUR", 1, "TAURITO PRINCESS (0310)", "Alimentación (3101)"),
                 ("HELADO VAINILLA (10100051)", "", "", "EUR", 1, "SUITE PRINCESS (0410)", "3106")]
        tabla, conflictos, _n, por_cab = H.leer_listas_productos(_xlsx(self.tmp.name, filas))
        self.assertTrue(por_cab)
        self.assertEqual(tabla, {"10100050": "3101", "10100051": "3106"})
        self.assertEqual(conflictos, [])
        self.assertIsNone(H.comprobar_listas_productos(tabla, PRODUCTOS))

    def test_sin_cabeceras_conocidas_no_adivina_y_la_proteccion_para(self):
        """Sin «Producto» / «Categoría valoración» se usan A y F tal cual: el archivo malo no casa y se para."""
        filas = [("Descripción", "x", "x", "x", "x", "Planta"),
                 ("QUESO BRIE (10100050)", "", "", "", "", "GUAYARMINA PRINCESS (0110)"),
                 ("HELADO VAINILLA (10100051)", "", "", "", "", "SUITE PRINCESS (0410)")]
        tabla, _c, _n, por_cab = H.leer_listas_productos(_xlsx(self.tmp.name, filas))
        self.assertFalse(por_cab)
        motivo = H.comprobar_listas_productos(tabla, PRODUCTOS)
        self.assertIn("no casa con Productos.xlsx", motivo)
        self.assertIn("No se ha cambiado nada", motivo)

    def test_vacio_tambien_para(self):
        self.assertIn("ninguna Categoría", H.comprobar_listas_productos({}, PRODUCTOS))

    def test_clave_entre_parentesis(self):
        self.assertEqual(H.clave_entre_parentesis("QUESO BRIE (10100050)"), "10100050")
        self.assertEqual(H.clave_entre_parentesis("PAN (PANADERIA) (10100052)"), "10100052")
        self.assertEqual(H.clave_entre_parentesis(" 3101 "), "3101")


class AccionTresParaAntesDeEscribir(unittest.TestCase):
    """_leer_categoria_valoracion lanza ValueError si el listado no casa: la Acción 3 no llega a escribir."""

    def test_lanza_y_lo_dice_en_la_consola(self):
        from app.principal.accion_listado_sap import AccionListadoSap
        app = object.__new__(AccionListadoSap)
        app.log_consola = mock.Mock()
        with tempfile.TemporaryDirectory() as d:
            ruta = _xlsx(d, [("Descripción", "", "", "", "", "Planta"),
                             ("QUESO BRIE (10100050)", "", "", "", "", "GUAYARMINA PRINCESS (0110)")])
            with self.assertRaises(ValueError):
                app._leer_categoria_valoracion(ruta, PRODUCTOS)
            self.assertTrue(any(c.args[1] == "ERROR" for c in app.log_consola.call_args_list))
            # sin tabla_productos (uso antiguo) no comprueba
            tabla, _c = app._leer_categoria_valoracion(ruta)
            self.assertEqual(tabla, {"QUESO BRIE (10100050)": "GUAYARMINA PRINCESS (0110)"})


class CategoriaDistintaSoloEnUso(unittest.TestCase):
    """(23 nov 2026, petición de Víctor) La Categoría valoración distinta según el centro solo se avisa si el código
    está en uso en el catálogo (y la web); si no, no se tiene en cuenta."""

    CONFLICTOS = [("10103148", {"3101": 8, "3106": 2}), ("10300003", {"3104": 9, "7901": 1}),
                  ("10800950", {"3104": 9, "3100": 1})]

    def test_filtra_por_codigos_en_uso(self):
        en_uso, ignorados = H.conflictos_en_uso(self.CONFLICTOS, [10103148.0, None, "", " 10800950 "])
        self.assertEqual([c for c, _v in en_uso], ["10103148", "10800950"])
        self.assertEqual(ignorados, 1)

    def test_ninguno_en_uso_no_avisa(self):
        from app.principal.accion_listado_sap import AccionListadoSap
        app = object.__new__(AccionListadoSap)
        app.log_consola = mock.Mock()
        en_uso, ignorados = app._avisar_conflictos_categoria(self.CONFLICTOS, ["99999999"])
        self.assertEqual((en_uso, ignorados), ([], 3))
        app.log_consola.assert_not_called()

    def test_avisa_solo_los_en_uso(self):
        from app.principal.accion_listado_sap import AccionListadoSap
        app = object.__new__(AccionListadoSap)
        app.log_consola = mock.Mock()
        app._avisar_conflictos_categoria(self.CONFLICTOS, ["10300003"])
        texto = " ".join(c.args[0] for c in app.log_consola.call_args_list)
        self.assertIn("1 código(s) en uso", texto)
        self.assertIn("10300003", texto)
        self.assertNotIn("10103148", texto)

    def test_la_lectura_ya_no_avisa_de_todos(self):
        from app.principal.accion_listado_sap import AccionListadoSap
        app = object.__new__(AccionListadoSap)
        app.log_consola = mock.Mock()
        with tempfile.TemporaryDirectory() as d:
            ruta = _xlsx(d, [("Producto", "Categoría valoración"), ("QUESO (10100050)", "A (3101)"),
                             ("QUESO (10100050)", "B (3106)")])
            _t, conflictos = app._leer_categoria_valoracion(ruta)
        self.assertEqual(len(conflictos), 1)
        self.assertFalse(any(c.args[1] == "WARN" for c in app.log_consola.call_args_list))


if __name__ == "__main__":
    unittest.main()
