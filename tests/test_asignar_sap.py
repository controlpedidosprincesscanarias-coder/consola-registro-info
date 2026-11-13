# -*- coding: utf-8 -*-
"""Pruebas de app/comun/asignar_sap.py (acción «Asignar SAP pendientes», v10.2.50), con un catálogo inventado.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import asignar_sap as A  # noqa: E402
from app.comun import catalogo_local as C  # noqa: E402
from tests.test_catalogo_local import crear_catalogo  # noqa: E402


class TestAsignarSap(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.cat = C.ruta_catalogo(self.dir.name)
        C._CACHE.clear()
        crear_catalogo(self.cat, [
            {"clave": 1, "descripcion": "TOMATE", "proveedor": "FRUTAS NORTE", "familia": "VERDURAS", "activo": "SI",
             "cod_sap": "1000"},
            {"clave": 2, "descripcion": "Pimiento rojo", "proveedor": "FRUTAS NORTE", "familia": "VERDURAS",
             "subfamilia": "PIMIENTOS", "activo": "SI"},
            {"clave": 3, "descripcion": "LECHE ENTERA", "proveedor": "LÁCTEOS SA", "familia": "LÁCTEOS", "activo": "SI"},
            {"clave": 4, "descripcion": "YOGUR", "proveedor": "LÁCTEOS SA", "familia": "LÁCTEOS", "activo": "NO"},
            {"clave": 5, "descripcion": "AZÚCAR", "proveedor": "ALMACÉN", "familia": "SECOS", "activo": "SI"},
        ])

    def tearDown(self):
        self.dir.cleanup()

    def test_pendientes_activos_sin_sap(self):
        lista = A.pendientes(self.cat)
        self.assertEqual([p["dali"] for p in lista], [2, 3, 5])
        self.assertEqual(lista[0]["subfamilia"], "PIMIENTOS")
        self.assertEqual(A.valores_filtro(lista, "proveedor"), ["ALMACÉN", "FRUTAS NORTE", "LÁCTEOS SA"])

    def test_filtrar_sin_acentos_ni_mayusculas(self):
        lista = A.pendientes(self.cat)
        self.assertEqual([p["dali"] for p in A.filtrar(lista, "azucar")], [5])
        self.assertEqual([p["dali"] for p in A.filtrar(lista, "PIMIENTO")], [2])
        self.assertEqual([p["dali"] for p in A.filtrar(lista, "3")], [3])
        self.assertEqual([p["dali"] for p in A.filtrar(lista, proveedor="LÁCTEOS SA")], [3])
        self.assertEqual([p["dali"] for p in A.filtrar(lista, "leche", familia="SECOS")], [])

    def test_validar_codigo(self):
        self.assertEqual(A.validar_codigo(" 10100416 "), 10100416)
        self.assertEqual(A.validar_codigo("10100416.0"), 10100416)
        for malo in ("", "abc", "12-3", "0"):
            with self.assertRaises(ValueError):
                A.validar_codigo(malo)

    def test_info_productos(self):
        tabla = {"2000": ("PIMIENTO ROJO KG", "Kilogramo", "KG", "G", "D", "T")}
        self.assertEqual(A.info_productos(2000, tabla), ("PIMIENTO ROJO KG", "KG"))
        self.assertIsNone(A.info_productos(9999, tabla))
        self.assertIsNone(A.info_productos(2000, None))
        self.assertEqual(A.fuera_de_productos({2: 2000, 3: 9999}, tabla), [3])
        self.assertEqual(A.fuera_de_productos({2: 2000}, None), [2])

    def test_conflictos_y_aplicar(self):
        wb, ws, mapa = C.abrir_para_editar(self.cat)
        fila_de, sap_de = C.mapas_dali_sap(ws, mapa)
        nuevos = {2: 1000, 3: 3000, 5: 3000}         # 1000 ya es del DALI 1; 3000 repetido en el lote
        conf = A.calcular_conflictos(nuevos, sap_de, {1: "TOMATE", 2: "PIMIENTO"})
        self.assertEqual(sorted(conf), [1000, 3000])
        self.assertEqual(conf[1000][0][:2], (1, True))
        # el usuario decide: el 1000 es del 2 (se le quita al 1) y el 3000 es del 3 (el 5 no se graba)
        escritos, limpiados, no = A.aplicar(ws, mapa, fila_de, nuevos, conf, {1, 5})
        wb.save(self.cat)
        wb.close()
        self.assertEqual((escritos, limpiados, no), ({2: 1000, 3: 3000}, [1], []))
        tabla = C.leer_tabla_sap(self.cat)
        self.assertEqual((tabla["1"]["cod_sap"], tabla["2"]["cod_sap"], tabla["3"]["cod_sap"], tabla["5"]["cod_sap"]),
                         (None, "1000", "3000", None))

    def test_sin_conflictos(self):
        wb, ws, mapa = C.abrir_para_editar(self.cat)
        _fila_de, sap_de = C.mapas_dali_sap(ws, mapa)
        wb.close()
        self.assertEqual(A.calcular_conflictos({2: 2000, 3: 3000}, sap_de, {}), {})


if __name__ == "__main__":
    unittest.main()
