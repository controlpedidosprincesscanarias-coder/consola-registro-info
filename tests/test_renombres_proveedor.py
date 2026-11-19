# -*- coding: utf-8 -*-
"""Pruebas de app/comun/renombres_proveedor.py: cambio de nombre del proveedor en DALI (19 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import renombres_proveedor as R  # noqa: E402

LARGO = "B REFRESCOS DEMO S.L. ( BEBIDAS DEMO CANARIAS SL )"
CORTO = "B REFRESCOS DEMO S.L."


class TestDetectar(unittest.TestCase):
    def test_todos_los_articulos_a_un_nombre_nuevo_es_cambio_de_nombre(self):
        antes = {78: LARGO, 79: LARGO, 80: LARGO, 1: "OTRO SL"}
        despues = {78: CORTO, 79: CORTO, 80: CORTO, 1: "OTRO SL"}
        self.assertEqual(R.detectar(antes, despues), [{"antes": LARGO, "despues": CORTO, "articulos": 3}])
        self.assertTrue(R.es_renombre(R.detectar(antes, despues), LARGO.lower(), "b  refrescos demo s.l. "))

    def test_si_queda_algun_articulo_con_el_nombre_antiguo_es_cambio_de_proveedor(self):
        antes = {1: "AHUMADOS DEMO SA", 2: "AHUMADOS DEMO SA"}
        despues = {1: "PESCADOS DEMO SL", 2: "AHUMADOS DEMO SA"}
        self.assertEqual(R.detectar(antes, despues), [])

    def test_si_el_nombre_nuevo_ya_existia_es_cambio_de_proveedor(self):
        antes = {1: "AHUMADOS DEMO SA", 2: "PESCADOS DEMO SL"}
        despues = {1: "PESCADOS DEMO SL", 2: "PESCADOS DEMO SL"}
        self.assertEqual(R.detectar(antes, despues), [])

    def test_si_se_reparten_entre_varios_proveedores_no_es_cambio_de_nombre(self):
        antes = {1: LARGO, 2: LARGO}
        despues = {1: CORTO, 2: "OTRO NUEVO SL"}
        self.assertEqual(R.detectar(antes, despues), [])

    def test_mayusculas_y_espacios_no_cuentan_como_cambio(self):
        antes = {1: "Huerta Demo SL"}
        despues = {1: "HUERTA  DEMO SL"}
        self.assertEqual(R.detectar(antes, despues), [])

    def test_articulos_dados_de_baja_o_sin_proveedor_no_cuentan(self):
        antes = {1: LARGO, 2: LARGO, 3: None}
        despues = {1: CORTO, 3: "NUEVO SL"}            # el 2 ya no está en DALI; el 3 no tenía proveedor
        self.assertEqual(R.detectar(antes, despues), [{"antes": LARGO, "despues": CORTO, "articulos": 1}])


if __name__ == "__main__":
    unittest.main()
