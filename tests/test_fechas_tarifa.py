# -*- coding: utf-8 -*-
"""Pruebas de las fechas de validez propuestas según el periodo de la tarifa (13 nov 2026). Caso real: una tarifa del
«4º TRIMESTRE 2026» se pasó con 01.01.2026 – 31.12.2026 (lo que venía por defecto) y el RINFO recortó 530 vigencias.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import tempfile
import unittest
from datetime import date
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import fechas_tarifa as FT  # noqa: E402

try:
    import tkinter as tk
    HAY_TK = True
except ImportError:          # pragma: no cover
    HAY_TK = False

HOY = date(2026, 11, 13)


class Periodo(unittest.TestCase):
    def test_periodos(self):
        casos = {
            "4º TRIMESTRE 2026": (date(2026, 10, 1), date(2026, 12, 31)),
            "1er trimestre 2027": (date(2027, 1, 1), date(2027, 3, 31)),
            "T2 2026": (date(2026, 4, 1), date(2026, 6, 30)),
            "3T 2026": (date(2026, 7, 1), date(2026, 9, 30)),
            "IV TRIMESTRE 2026": (date(2026, 10, 1), date(2026, 12, 31)),
            "2º SEMESTRE 2026": (date(2026, 7, 1), date(2026, 12, 31)),
            "OCTUBRE 2026": (date(2026, 10, 1), date(2026, 10, 31)),
            "OCT-DIC 2026": (date(2026, 10, 1), date(2026, 12, 31)),
            "NOVIEMBRE-FEBRERO 2026": (date(2026, 11, 1), date(2027, 2, 28)),
            "Octubre 2026 - Marzo 2027": (date(2026, 10, 1), date(2027, 3, 31)),
            "01/10/2026 - 31/03/2027": (date(2026, 10, 1), date(2027, 3, 31)),
            "AÑO 2026": (date(2026, 1, 1), date(2026, 12, 31)),
            "2027": (date(2027, 1, 1), date(2027, 12, 31)),
        }
        for texto, esperado in casos.items():
            self.assertEqual(FT.periodo_a_fechas(texto), esperado, texto)
        for texto in ("", None, "TARIFA GENERAL", "SIN PERIODO"):
            self.assertIsNone(FT.periodo_a_fechas(texto), texto)

    def test_propuesta(self):
        self.assertEqual(FT.propuesta("4º TRIMESTRE 2026", HOY), (date(2026, 10, 1), date(2026, 12, 31), "periodo"))
        ult = (date(2026, 10, 1), date(2027, 3, 31))
        self.assertEqual(FT.propuesta("TARIFA GENERAL", HOY, ult), (*ult, "ultimas"))
        self.assertEqual(FT.propuesta("", HOY), (date(2026, 1, 1), date(2026, 12, 31), "anio"))

    def test_avisos_caso_del_registro(self):
        a = FT.avisos(date(2026, 1, 1), date(2026, 12, 31), "4º TRIMESTRE 2026", HOY)
        self.assertEqual(len(a), 2)
        self.assertIn("anterior a hoy", a[0])
        self.assertIn("01.10.2026 – 31.12.2026", a[1])
        # Las del periodo de la tarifa, aunque el inicio sea anterior a hoy: nada que avisar
        self.assertEqual(FT.avisos(date(2026, 10, 1), date(2026, 12, 31), "4º TRIMESTRE 2026", HOY), [])
        # Sin periodo, solo el aviso del pasado
        self.assertEqual(FT.avisos(date(2026, 12, 1), date(2027, 3, 31), "", HOY), [])
        self.assertEqual(len(FT.avisos(date(2026, 1, 1), date(2026, 12, 31), "", HOY)), 1)

    def test_ultimas_fechas(self):
        with tempfile.TemporaryDirectory() as d:
            ruta = os.path.join(d, "x", "ultimas.json")
            self.assertIsNone(FT.leer_ultimas(ruta))
            FT.guardar_ultimas(date(2026, 10, 1), date(2027, 3, 31), ruta)
            self.assertEqual(FT.leer_ultimas(ruta), (date(2026, 10, 1), date(2027, 3, 31)))

    def test_leer_periodo_de_la_tarifa(self):
        import openpyxl
        with tempfile.TemporaryDirectory() as d:
            ruta = os.path.join(d, "T.xlsx")
            wb = openpyxl.Workbook()
            wb.active.cell(3, 2).value = "Periodo:"
            wb.active.cell(3, 3).value = "4º TRIMESTRE 2026"
            wb.save(ruta)
            self.assertEqual(FT.leer_periodo_tarifa(ruta), "4º TRIMESTRE 2026")
        self.assertEqual(FT.leer_periodo_tarifa("/no/existe.xlsx"), "")


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class Ventana(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("sin pantalla")
        self.root.withdraw()
        self.tmp = tempfile.TemporaryDirectory()
        self.env = mock.patch.dict(os.environ, {"LOCALAPPDATA": self.tmp.name})
        self.env.start()

    def tearDown(self):
        self.env.stop()
        self.tmp.cleanup()
        self.root.destroy()

    def test_propone_el_periodo_y_avisa_si_se_cambia(self):
        from app.rinfo.ventanas_rinfo import VentanaFechasRINFO
        r = []
        v = VentanaFechasRINFO(self.root, r.append, periodo="4º TRIMESTRE 2026")
        self.assertEqual((v._var_inicio.get(), v._var_fin.get()), ("01.10.2026", "31.12.2026"))
        self.assertIn("periodo de la tarifa", v._lbl_origen.cget("text"))
        v._var_inicio.set("01.01.2026")                      # como en el registro
        with mock.patch("app.rinfo.ventanas_rinfo.messagebox.askyesno", return_value=False) as pregunta:
            v._confirmar()
        pregunta.assert_called_once()
        self.assertEqual(r, [])                                # «No»: sigue en la ventana para corregir
        v._var_inicio.set("01.10.2026")
        with mock.patch("app.rinfo.ventanas_rinfo.messagebox.askyesno") as pregunta:
            v._confirmar()
        pregunta.assert_not_called()                           # las del periodo: sin preguntas
        self.assertEqual(r, [("01.10.2026", "31.12.2026")])
        self.assertEqual(FT.leer_ultimas(), (date(2026, 10, 1), date(2026, 12, 31)))   # recordadas


if __name__ == "__main__":
    unittest.main()
