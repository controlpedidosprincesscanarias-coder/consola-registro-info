# -*- coding: utf-8 -*-
"""Pruebas de compilacion/version.py: la versión de cada compilación sale sola y nunca se repite.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(RAIZ, "compilacion"))
import version as V  # noqa: E402


class VersionAutomatica(unittest.TestCase):
    def test_sin_releases_usa_la_del_codigo(self):
        self.assertEqual(V.siguiente("10.2.36", []), "10.2.36")

    def test_salto_a_10_4_00(self):
        # Petición de Víctor (23 nov 2026): tras la v10.2.100 (en borrador), la v 10.4.00.
        self.assertEqual(V.siguiente("10.4.00", ["10.2.98", "10.2.100"]), "10.4.00")
        self.assertEqual(V.siguiente("10.4.00", ["10.2.100", "10.4.00"]), "10.4.02")
        self.assertEqual(V.siguiente("10.4.00", ["10.4.98"]), "10.4.100")

    def test_el_codigo_por_debajo_no_cuenta(self):
        self.assertEqual(V.siguiente("10.2.36", ["10.4.02"]), "10.4.04")

    def test_ultimo_numero_con_dos_cifras(self):
        self.assertEqual(V.siguiente("10.4.00", ["10.4.06"]), "10.4.08")
        self.assertEqual(V.siguiente("10.4.00", ["10.4.08"]), "10.4.10")

    def test_parte_de_la_ultima_release(self):
        self.assertEqual(V.siguiente("10.2.36", ["10.2.38", "10.2.40"]), "10.2.42")

    def test_compara_como_numeros(self):
        self.assertEqual(V.siguiente("10.2.36", ["10.2.100", "10.2.98"]), "10.2.102")

    def test_si_la_ultima_es_impar_salta_al_siguiente_par(self):
        self.assertEqual(V.siguiente("10.2.37", []), "10.2.38")

    def test_forzar(self):
        self.assertEqual(V.siguiente("10.2.36", ["10.2.38"], "10.3.0"), "10.3.00")
        self.assertEqual(V.siguiente("10.4.00", ["10.4.02"], "10.6.00"), "10.6.00")
        for mala in ("10.2.38", "10.2.20", "10.3.1", "10.3"):
            with self.assertRaises(SystemExit):
                V.siguiente("10.2.36", ["10.2.38"], mala)

    def test_etiquetas_de_releases(self):
        texto = "v10.2.38\nv10.2.40\nprueba\n﻿v10.2.42\n"
        self.assertEqual(V.versiones_releases(texto), ["10.2.38", "10.2.40", "10.2.42"])

    def test_escribir_conserva_formato_y_comentario(self):
        src = 'x = 1\nAPP_VERSION = "v 10.2.36"   # ← actualiza aquí\ny = 2\n'
        nuevo = V.escribir_app_version(src, "10.2.38")
        self.assertEqual(nuevo, 'x = 1\nAPP_VERSION = "v 10.2.38"   # ← actualiza aquí\ny = 2\n')
        self.assertEqual(V.leer_app_version(nuevo), "10.2.38")

    def test_lee_la_version_real_de_tarifa_app(self):
        with open(os.path.join(RAIZ, "tarifa_app.py"), encoding="utf-8") as f:
            self.assertRegex(V.leer_app_version(f.read()), r"^\d+\.\d+\.\d+$")


if __name__ == "__main__":
    unittest.main()
