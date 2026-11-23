# -*- coding: utf-8 -*-
"""
Bajas «sin tarifa» con una tarifa de solo algunos hoteles (23 nov 2026, caso real AHUMADOS CANARIOS SA LP).

Una tarifa de un solo hotel (LP) no puede dar de baja el Registro Info entero (H1 AJ) de los materiales que no trae:
pueden ser de las tarifas de otros hoteles (GC-TF, FV). Con el método real de GeneradorDesdeExistente:
  · material también en hoteles de fuera de la tarifa → baja solo en los hoteles de la tarifa (H2 AN, H3 Y);
  · material solo en hoteles de fuera → no se toca;
  · tarifa de todos sus hoteles → como siempre (H1 AJ y todos sus centros).

Datos inventados. Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.test_crear_registro_info import _CRI, _parchear_rinfo  # noqa: E402

LP, GC, FV = ("0210", "LA PALMA&TENEGUIA PRINCESS"), ("0610", "GRAN CANARIA PRINCESS"), ("0910", "FUERTEVENTURA PRINCESS")


@unittest.skipIf(_CRI is None, "crear_registro_info no importable (¿sin tkinter?)")
class BajasPorHotel(unittest.TestCase):
    def _libro(self, materiales):
        """materiales: {material: [centros]}. Precio vigente hasta 2090 en cada centro."""
        import openpyxl
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Info general p.reg.info"
        ws2 = wb.create_sheet("Datos org.compr.p.reg.info")
        ws3 = wb.create_sheet("Det.precios p.registros info")
        for ws, n in ((ws1, 44), (ws2, 41), (ws3, 27)):
            for c in range(1, n + 1):
                ws.cell(row=1, column=c, value=f"c{c}")
        filas_h1, r2, r3 = {}, 2, 2
        for i, (mat, centros) in enumerate(materiales.items(), start=2):
            reg = f"53000000{i}"
            ws1.cell(row=i, column=1, value=reg)
            ws1.cell(row=i, column=5, value=mat)
            ws1.cell(row=i, column=6, value=f"MATERIAL {mat}")
            filas_h1[mat] = i
            for centro in centros:
                ws2.cell(row=r2, column=1, value=reg)
                ws2.cell(row=r2, column=6, value=centro)
                ws2.cell(row=r2, column=10, value=mat)
                r2 += 1
                ws3.cell(row=r3, column=1, value=reg)
                ws3.cell(row=r3, column=5, value=centro)
                ws3.cell(row=r3, column=6, value=mat)
                ws3.cell(row=r3, column=9, value="01.01.2026")
                ws3.cell(row=r3, column=10, value="31.12.2090")
                ws3.cell(row=r3, column=15, value=2.5)
                r3 += 1
        return wb, ws1, ws2, ws3, filas_h1

    def _generar(self, materiales, centros_activos, en_tarifa=()):
        class _P:
            def after(self, ms, fn): fn()
            def winfo_exists(self): return True
        wb, ws1, ws2, ws3, filas_h1 = self._libro(materiales)
        gen = object.__new__(_CRI.GeneradorDesdeExistente)
        gen._parent = gen._app_root = _P()
        gen._barra = lambda *a, **k: None
        gen._fecha_inicio = "01.07.2090"
        gen._audit = {"materiales_modificados": set(), "materiales_sin_cambios": set()}
        gen._fila_h1_por_codigo = filas_h1
        gen._centros_activos = list(centros_activos)
        self.ventana = []
        logs = []

        def ventana(parent, pend):
            self.ventana.extend(p["codigo"] for p in pend)
            return [{"codigo": p["codigo"], "fila_h1": p["fila_h1"], "baja": True} for p in pend]   # «DAR BAJA»
        with _parchear_rinfo("_log", lambda _p, m, n="INFO": logs.append(m)), \
                _parchear_rinfo("_ventana_sin_tarifa", ventana), \
                _parchear_rinfo("_safe_after", lambda w, fn, ms=0: fn()):
            gen._limpiar_marca_aj_heredados(wb, ws1, set(materiales), set(en_tarifa))
            gen._propagar_borrado_h1_a_h2_h3(wb, ws1)
        self.logs = logs
        aj = {mat: ws1.cell(row=f, column=36).value for mat, f in filas_h1.items()}
        an = {(str(ws2.cell(row=r, column=10).value), str(ws2.cell(row=r, column=6).value)):
              ws2.cell(row=r, column=40).value for r in range(2, ws2.max_row + 1)}
        y = {(str(ws3.cell(row=r, column=6).value), str(ws3.cell(row=r, column=5).value)):
             ws3.cell(row=r, column=25).value for r in range(2, ws3.max_row + 1)}
        return aj, an, y

    def test_tarifa_de_un_hotel_no_borra_el_registro_entero(self):
        aj, an, y = self._generar({"10101660": ["0210", "0610", "0910"]}, [LP])
        self.assertIsNone(aj["10101660"])                               # H1 AJ sin marcar
        self.assertEqual(an[("10101660", "0210")], "X")                 # baja solo en LP
        self.assertIsNone(an[("10101660", "0610")])
        self.assertIsNone(an[("10101660", "0910")])
        self.assertEqual(y[("10101660", "0210")], "X")
        self.assertIsNone(y[("10101660", "0610")])
        self.assertTrue(any("SOLO en los hoteles de esta tarifa (0210)" in m for m in self.logs))

    def test_material_solo_en_otros_hoteles_no_se_toca(self):
        aj, an, y = self._generar({"10103211": ["0610", "0910"]}, [LP])
        self.assertIsNone(aj["10103211"])
        self.assertEqual(set(an.values()) | set(y.values()), {None})
        self.assertEqual(self.ventana, [])                              # ni siquiera se pregunta
        self.assertTrue(any("solo en hoteles fuera de esta tarifa" in m for m in self.logs))

    def test_tarifa_de_todos_sus_hoteles_como_siempre(self):
        aj, an, y = self._generar({"10101416": ["0210", "0610"]}, [LP, GC, FV])
        self.assertEqual(aj["10101416"], "X")
        self.assertEqual(an[("10101416", "0210")], "X")
        self.assertEqual(an[("10101416", "0610")], "X")
        self.assertEqual(y[("10101416", "0610")], "X")

    def test_los_que_siguen_en_la_tarifa_no_se_tocan(self):
        aj, an, y = self._generar({"10101723": ["0210", "0610"], "10101660": ["0210", "0610"]}, [LP],
                                  en_tarifa={"10101723"})
        self.assertIsNone(an[("10101723", "0210")])
        self.assertEqual(an[("10101660", "0210")], "X")
        self.assertIsNone(an[("10101660", "0610")])


if __name__ == "__main__":
    unittest.main()
