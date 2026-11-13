# -*- coding: utf-8 -*-
"""Pruebas del calendario 📅 de «Fechas de validez del Registro Info» (13 nov 2026): la fecha elegida se escribe
como «dd.mm.aaaa», el mismo formato de siempre para el RINFO.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import unittest
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import tkinter as tk
    HAY_TK = True
except ImportError:          # pragma: no cover
    HAY_TK = False


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class Logica(unittest.TestCase):
    def test_formato_y_lectura(self):
        from app.comun import selector_fecha as S
        self.assertEqual(S.formatear(date(2026, 10, 1)), "01.10.2026")
        for t in ("01.10.2026", "01-10-2026", "01102026", "1.10.2026", " 01,10,2026 "):
            self.assertEqual(S.leer(t), date(2026, 10, 1), t)
        self.assertIsNone(S.leer("31.02.2026"))
        self.assertEqual(S.leer("", date(2000, 1, 1)), date(2000, 1, 1))

    def test_semanas_empiezan_en_lunes(self):
        from app.comun import selector_fecha as S
        sem = S.semanas(2026, 10)               # el 1 de octubre de 2026 es jueves
        self.assertEqual(sem[0], [0, 0, 0, 1, 2, 3, 4])
        self.assertEqual(max(max(s) for s in sem), 31)


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class VentanaFechas(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("sin pantalla")
        self.root.withdraw()

    def tearDown(self):
        self.root.destroy()

    def test_elegir_en_el_calendario_escribe_el_formato_del_rinfo(self):
        from app.comun import selector_fecha as S
        from app.rinfo.ventanas_rinfo import VentanaFechasRINFO
        resultado = []
        v = VentanaFechasRINFO(self.root, resultado.append)
        v.update()
        v._btn_cal_inicio.invoke()                          # abre el calendario
        cal = [w for w in v.winfo_children() if isinstance(w, S.Calendario)]
        cal = cal[0] if cal else next(w for w in self.root.winfo_children() if isinstance(w, S.Calendario))
        cal._var_mes.set("octubre"); cal._var_anio.set("2026"); cal._pintar()
        cal._elegir(date(2026, 10, 1))
        self.assertEqual(v._var_inicio.get(), "01.10.2026")
        v._var_fin.set("31.03.2027")                        # a mano sigue valiendo
        v._confirmar()
        self.assertEqual(resultado, [("01.10.2026", "31.03.2027")])


if __name__ == "__main__":
    unittest.main()
