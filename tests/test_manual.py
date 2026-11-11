# -*- coding: utf-8 -*-
"""Pruebas de app/comun/manual.py: «Consultar el manual» encuentra el apartado que toca.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import json
import os
import sys
import tempfile
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
from app.comun import manual as M  # noqa: E402

INDICE = [
    {"capitulo": "4. Generar la tarifa", "titulo": "Códigos SAP faltantes", "pagina": 12,
     "texto": "Si algún código DALI no tiene SAP, se abre la ventana de códigos SAP faltantes. Escribe el SAP de cada "
              "artículo y pulsa Guardar. Se escriben en el LISTADO y en la tarifa."},
    {"capitulo": "6. Registro Info", "titulo": "Crear el Registro Info (RINFO)", "pagina": 18,
     "texto": "Pulsa Crear / Actualizar Registro Info SAP, elige la tarifa y el proveedor. Si hay Purchasing, se usa."},
    {"capitulo": "3. Listados", "titulo": "Actualizar el listado de materiales SAP", "pagina": 9,
     "texto": "La Acción 3 lee Productos.xlsx y Listas de productos.xlsx de RECURSOS."},
    {"capitulo": "8. Ajustes", "titulo": "Contraseña de Catálogo Asignaciones", "pagina": 25,
     "texto": "La primera vez se pide el usuario y la contraseña. Se guarda cifrada en el equipo."},
]


class Busqueda(unittest.TestCase):
    def test_normaliza_acentos_y_mayusculas(self):
        self.assertEqual(M.normalizar("Código SAP Ñ"), "codigo sap n")
        self.assertEqual(M.palabras("¿Cómo creo los códigos?"), [M.raiz("creo"), M.raiz("codigos")])

    def test_encuentra_el_apartado_aunque_cambien_las_palabras(self):
        casos = {"códigos sap faltantes": "Códigos SAP faltantes",
                 "me faltan codigos SAP en la tarifa": "Códigos SAP faltantes",
                 "cómo creo un registro info": "Crear el Registro Info (RINFO)",
                 "rinfo": "Crear el Registro Info (RINFO)",
                 "actualizar productos sap": "Actualizar el listado de materiales SAP",
                 "contraseña catalogo": "Contraseña de Catálogo Asignaciones"}
        for consulta, esperado in casos.items():
            r = M.buscar(INDICE, consulta)
            self.assertTrue(r, consulta)
            self.assertEqual(r[0]["titulo"], esperado, consulta)

    def test_resultado_con_pagina_y_fragmento(self):
        r = M.buscar(INDICE, "contraseña cifrada")[0]
        self.assertEqual(r["pagina"], 25)
        self.assertIn("cifrada", r["fragmento"])

    def test_sin_resultados(self):
        self.assertEqual(M.buscar(INDICE, "zzzz qqqq"), [])
        self.assertEqual(M.buscar(INDICE, ""), [])
        self.assertEqual(M.buscar(INDICE, "de la el"), [])      # solo palabras vacías
        self.assertEqual(M.buscar([], "rinfo"), [])

    def test_cargar_indice(self):
        with tempfile.TemporaryDirectory() as d:
            ruta = os.path.join(d, "i.json")
            with open(ruta, "w", encoding="utf-8") as f:
                json.dump({"version": "v 1.0.0", "secciones": INDICE}, f)
            self.assertEqual(len(M.cargar_indice(ruta)), 4)
            self.assertEqual(M.cargar_indice(os.path.join(d, "no.json")), [])

    def test_indice_real_de_la_app(self):
        indice = M.cargar_indice()
        if not indice:
            self.skipTest("aún no hay assets/manual_indice.json")
        self.assertTrue(all(s.get("titulo") and s.get("pagina") for s in indice))
        for consulta in ("crear registro info", "códigos SAP faltantes", "actualizar listado SAP"):
            self.assertTrue(M.buscar(indice, consulta), consulta)


if __name__ == "__main__":
    unittest.main()
