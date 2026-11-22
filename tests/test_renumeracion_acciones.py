# -*- coding: utf-8 -*-
"""Botones de acciones numerados 1 a 6 y Acción 4 solo con tarifas [INCOMPLETO] (20 nov 2026, v10.2.86).

Petición de Víctor: la 3 pasa a 2, la 4 a 3, la 5 desaparece (no se usa), la 6 a 4 (sombreada y solo utilizable si hay
tarifas [INCOMPLETO] en TARIFAS), la 8 a 5 y la 7 a 6. Datos inventados.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import re
import sys
import tempfile
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
from app.comun import incompletos  # noqa: E402


def _botones():
    with open(os.path.join(RAIZ, "app/principal/interfaz.py"), encoding="utf-8") as f:
        texto = f.read()
    return re.findall(r'frm_acc, "(\d)", "([^"]+)", "[^"]+",\s*\n?\s*self\.(\w+)', texto)


class Botones(unittest.TestCase):
    def test_orden_y_funcion_de_cada_boton(self):
        self.assertEqual([(n, m) for n, _t, m in _botones()], [
            ("1", "_organizar_dali"), ("2", "_accion_3_guiada"), ("3", "_procesar"),
            ("4", "_crear_registro_sap"), ("5", "_asignar_sap_pendientes"), ("6", "_generar_plantilla_desde_rinfo")])

    def test_ya_no_hay_boton_para_abrir_la_ultima_tarifa(self):
        self.assertFalse(any("ultima TARIFA" in t for _n, t, _m in _botones()))

    def test_la_accion_4_empieza_sombreada(self):
        with open(os.path.join(RAIZ, "app/principal/interfaz.py"), encoding="utf-8") as f:
            texto = f.read()
        bloque = texto[texto.index('frm_acc, "4"'):texto.index('frm_acc, "5"')]
        self.assertIn('state="disabled"', bloque)


class TarifasIncompletas(unittest.TestCase):
    def test_solo_cuentan_las_tarifas_incompletas(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertFalse(incompletos.hay(d))
            for nombre in ("SUMINISTROS DEMO SL [19-11-2026].xlsx",            # completa
                           "RINFO SUMINISTROS DEMO SL [INCOMPLETO] [19-11-2026].xlsx",   # RINFO: no cuenta
                           "notas [INCOMPLETO].txt"):                          # otro archivo
                open(os.path.join(d, nombre), "w").close()
            os.mkdir(os.path.join(d, "PAPELERIA EJEMPLO [INCOMPLETO] [18-11-2026].xlsx"))  # carpeta: no cuenta
            self.assertFalse(incompletos.hay(d))
            self.assertEqual(incompletos.listar(d), [])
            open(os.path.join(d, "SUMINISTROS DEMO SL [INCOMPLETO] [19-11-2026].xlsx"), "w").close()
            open(os.path.join(d, "menaje modelo [incompleto] [17-11-2026].XLSX"), "w").close()
            self.assertTrue(incompletos.hay(d))
            self.assertEqual([a["nombre"] for a in incompletos.listar(d)],
                             ["menaje modelo [incompleto] [17-11-2026].XLSX",
                              "SUMINISTROS DEMO SL [INCOMPLETO] [19-11-2026].xlsx"])

    def test_sin_carpeta(self):
        self.assertFalse(incompletos.hay(os.path.join(tempfile.gettempdir(), "no-existe-demo-123")))
        self.assertEqual(incompletos.listar(""), [])


if __name__ == "__main__":
    unittest.main()
