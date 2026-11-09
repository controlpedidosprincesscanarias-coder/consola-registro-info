# -*- coding: utf-8 -*-
"""Pruebas de credenciales.py: la contraseña de Catálogo Asignaciones nunca se guarda en claro.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
Uso: python -m unittest test_credenciales -v   (en Windows prueba también el cifrado real con DPAPI)
"""
import sys
import unittest
from unittest import mock

import credenciales as cr


def _falso_proteger(texto):
    return cr.PREFIJO + texto[::-1] if texto else None


def _falso_desproteger(valor):
    return valor[len(cr.PREFIJO):][::-1] if valor and valor.startswith(cr.PREFIJO) else None


class ConCifradoSimulado(unittest.TestCase):
    def setUp(self):
        self._p = mock.patch.object(cr, "proteger", _falso_proteger)
        self._d = mock.patch.object(cr, "desproteger", _falso_desproteger)
        self._p.start(); self._d.start()

    def tearDown(self):
        self._p.stop(); self._d.stop()

    def test_guardar_y_leer(self):
        cfg = {}
        self.assertTrue(cr.guardar_credenciales_dali(cfg, "a@b.c", "secreto"))
        self.assertNotIn("secreto", repr(cfg))
        self.assertNotIn(cr.CLAVE_ANTIGUA, cfg)
        self.assertEqual(cr.leer_credenciales_dali(cfg), ("a@b.c", "secreto"))

    def test_migra_la_antigua_en_claro(self):
        cfg = {"carpeta_proyecto": "X", cr.CLAVE_EMAIL: "a@b.c", cr.CLAVE_ANTIGUA: "secreto"}
        self.assertTrue(cr.migrar_credenciales_dali(cfg))
        self.assertNotIn(cr.CLAVE_ANTIGUA, cfg)
        self.assertNotIn("secreto", repr(cfg))
        self.assertEqual(cfg["carpeta_proyecto"], "X")
        self.assertEqual(cr.leer_credenciales_dali(cfg), ("a@b.c", "secreto"))

    def test_sin_antigua_no_cambia_nada(self):
        cfg = {"carpeta_proyecto": "X"}
        self.assertFalse(cr.migrar_credenciales_dali(cfg))
        self.assertEqual(cfg, {"carpeta_proyecto": "X"})

    def test_borrar(self):
        cfg = {"carpeta_proyecto": "X"}
        cr.guardar_credenciales_dali(cfg, "a@b.c", "secreto")
        cr.borrar_credenciales_dali(cfg)
        self.assertEqual(cfg, {"carpeta_proyecto": "X"})


class SinCifrado(unittest.TestCase):
    """Si no se puede cifrar (no Windows, DPAPI falla): la contraseña no se guarda nunca."""
    def setUp(self):
        self._p = mock.patch.object(cr, "proteger", lambda texto: None)
        self._p.start()

    def tearDown(self):
        self._p.stop()

    def test_no_guarda_en_claro(self):
        cfg = {}
        self.assertFalse(cr.guardar_credenciales_dali(cfg, "a@b.c", "secreto"))
        self.assertNotIn("secreto", repr(cfg))

    def test_migracion_borra_la_antigua(self):
        cfg = {cr.CLAVE_EMAIL: "a@b.c", cr.CLAVE_ANTIGUA: "secreto"}
        self.assertTrue(cr.migrar_credenciales_dali(cfg))
        self.assertNotIn("secreto", repr(cfg))
        self.assertEqual(cr.leer_credenciales_dali(cfg), (None, None))

    def test_otro_equipo_no_descifra(self):
        cfg = {cr.CLAVE_EMAIL: "a@b.c", cr.CLAVE_CIFRADA: "dpapi1:basura"}
        with mock.patch.object(cr, "desproteger", lambda v: None):
            self.assertEqual(cr.leer_credenciales_dali(cfg), (None, None))


@unittest.skipUnless(sys.platform.startswith("win"), "DPAPI real solo en Windows")
class DpapiReal(unittest.TestCase):
    def test_ida_y_vuelta(self):
        c = cr.proteger("contraseña de prueba ñ")
        self.assertTrue(c.startswith(cr.PREFIJO))
        self.assertNotIn("prueba", c)
        self.assertEqual(cr.desproteger(c), "contraseña de prueba ñ")


if __name__ == "__main__":
    unittest.main()
