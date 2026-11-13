# -*- coding: utf-8 -*-
"""
test_accion1_mensajes.py — los mensajes de la Acción 1 salen de verdad (13 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Hasta el 13 nov 2026, `from tkinter import messagebox` estaba dentro de proceso() solo en los bloques except, así que
messagebox era variable local de proceso() y los mensajes que se lanzaban antes con self.after (faltan ARTICULOS1/2,
incidencias de duplicados, validación fallida) daban NameError y no se veían. Esta prueba vigila que ninguna función
de los módulos de la Acción 1 vuelva a importar messagebox por su cuenta.
"""
import ast
import os
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODULOS = ("app/principal/accion_organizar_dali.py", "app/principal/accion_listado_dali.py")


class TestMessageboxAccion1(unittest.TestCase):
    def test_messagebox_solo_arriba_del_todo(self):
        for rel in MODULOS:
            arbol = ast.parse(open(os.path.join(RAIZ, rel), encoding="utf-8").read())
            arriba = [n for n in arbol.body if isinstance(n, ast.ImportFrom) and n.module == "tkinter"
                      and any(a.name == "messagebox" for a in n.names)]
            self.assertTrue(arriba, f"{rel}: falta `from tkinter import messagebox` arriba del todo")
            for func in ast.walk(arbol):
                if not isinstance(func, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    continue
                for n in ast.walk(func):
                    if isinstance(n, ast.ImportFrom) and n.module == "tkinter":
                        self.assertNotIn("messagebox", [a.name for a in n.names],
                                         f"{rel}:{n.lineno}: messagebox importado dentro de {func.name}()")


if __name__ == "__main__":
    unittest.main()
