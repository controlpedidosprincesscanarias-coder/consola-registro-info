# -*- coding: utf-8 -*-
"""Pruebas de _eliminar_filas_agrupadas (app/rinfo/comun_rinfo.py): mismo resultado que delete_rows() fila a
fila, pero en una sola pasada (19 nov 2026: 690 filas salteadas de la Hoja 3 tardaban 94 s). Datos inventados.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import random
import sys
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import openpyxl  # noqa: E402
from openpyxl.styles import Font  # noqa: E402

from app.rinfo.comun_rinfo import _eliminar_filas_agrupadas  # noqa: E402


def _hoja(filas=400, columnas=12):
    ws = openpyxl.Workbook().active
    negrita = Font(bold=True)
    for f in range(1, filas + 1):
        for c in range(1, columnas + 1):
            if (f * c) % 7 == 0:
                continue                                    # huecos, como en un Purchasing real
            celda = ws.cell(row=f, column=c, value=f"F{f}C{c}" if c % 3 else f * 1.5)
            if c == 2 and f % 5 == 0:
                celda.font = negrita
    return ws


def _foto(ws):
    return [tuple((c.value, c.font.bold) for c in fila) for fila in ws.iter_rows()]


def _como_antes(ws, filas):
    for f in sorted(set(filas), reverse=True):
        ws.delete_rows(f, 1)


class EliminarFilas(unittest.TestCase):
    def _comparar(self, filas):
        a, b = _hoja(), _hoja()
        _como_antes(a, filas)
        _eliminar_filas_agrupadas(b, filas)
        self.assertEqual(a.max_row, b.max_row)
        self.assertEqual(_foto(a), _foto(b))
        for (f, c), celda in b._cells.items():
            self.assertEqual((celda.row, celda.column), (f, c))
        return b

    def test_filas_salteadas_como_el_caso_real(self):
        self._comparar(list(range(30, 330, 3)))

    def test_al_azar_con_bloques_repetidas_y_la_ultima(self):
        random.seed(7)
        self._comparar(random.sample(range(2, 400), 80) + list(range(50, 70)) + [55, 55, 400])

    def test_sin_filas_no_cambia_nada(self):
        ws = _hoja()
        antes = _foto(ws)
        _eliminar_filas_agrupadas(ws, [])
        self.assertEqual(_foto(ws), antes)

    def test_las_filas_que_quedan_suben_en_orden(self):
        ws = self._comparar([2, 3, 10])
        self.assertEqual(ws.cell(row=2, column=1).value, "F4C1")
        self.assertEqual(ws.cell(row=8, column=1).value, "F11C1")

    def test_es_rapido_con_muchas_filas_salteadas(self):
        ws = _hoja(filas=6000, columnas=27)
        t = time.time()
        _eliminar_filas_agrupadas(ws, list(range(1000, 3100, 3)))
        self.assertLess(time.time() - t, 5)                 # antes, del orden de un minuto
        self.assertEqual(ws.max_row, 6000 - 700)


if __name__ == "__main__":
    unittest.main()
