# -*- coding: utf-8 -*-
"""Materiales con tarifa cuyo Purchasing trae marcados para borrado solo algunos centros (20 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Caso real OASISFISH (log del 20 nov 2026): «Marca de borrado limpiada … H1 col AJ = 0 | H2 col AN = 50
… ver aviso siguiente sobre gestión manual», y el aviso no salía. Lo que SAP tiene borrado es el segmento
por centro (Hoja 2, col AN): aunque la cabecera (Hoja 1, col AJ) no venga marcada, el material hay que
reactivarlo a mano en SAP. Ahora sale en el aviso de la consola y en la ventana «Materiales a reactivar
manualmente en SAP». Prueba el código real de GeneradorDesdeExistente (sin tkinter se salta).
"""
import unittest

from tests.test_crear_registro_info import _CRI, _parchear_rinfo


@unittest.skipIf(_CRI is None, "crear_registro_info no importable (¿sin tkinter?)")
class TestReactivarSoloCentros(unittest.TestCase):
    REG, MAT = 5300009999, "10100412"

    def _generador(self, aj_h1=None):
        import openpyxl
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Info general p.reg.info"
        ws2 = wb.create_sheet("Datos org.compr.p.reg.info")
        for ws, n in ((ws1, 44), (ws2, 41)):
            for c in range(1, n + 1):
                ws.cell(row=1, column=c, value=f"c{c}")
        ws1.cell(row=2, column=1, value=self.REG)
        ws1.cell(row=2, column=5, value=self.MAT)
        ws1.cell(row=2, column=6, value="MERLUZA")
        ws1.cell(row=2, column=36, value=aj_h1)
        for i, centro in enumerate(["0110", "0120", "0210"]):
            ws2.cell(row=2 + i, column=1, value=self.REG)
            ws2.cell(row=2 + i, column=6, value=centro)
            ws2.cell(row=2 + i, column=10, value=self.MAT)
            ws2.cell(row=2 + i, column=40, value="X" if centro != "0210" else None)

        class _P:
            def after(self, ms, fn): fn()
            def winfo_exists(self): return True

        gen = object.__new__(_CRI.GeneradorDesdeExistente)
        gen._parent = _P()
        gen._app_root = _P()
        gen._barra = lambda *a, **k: None
        gen._audit = {"materiales_modificados": set(), "materiales_sin_cambios": set()}
        gen._fila_h1_por_codigo = {self.MAT: 2}
        return gen, wb, ws1, ws2

    def _limpiar(self, gen, wb, ws1, logs):
        with _parchear_rinfo("_log", lambda parent, texto, nivel="INFO": logs.append((nivel, texto))), \
             _parchear_rinfo("_safe_after", lambda w, fn, ms=0: fn()):
            gen._limpiar_marca_aj_heredados(wb, ws1, {self.MAT}, {self.MAT})
            return gen._avisar_materiales_reactivados_bloqueados(wb)

    def test_solo_centros_marcados_tambien_se_avisa(self):
        gen, wb, ws1, ws2 = self._generador(aj_h1=None)
        logs = []
        n = self._limpiar(gen, wb, ws1, logs)
        self.assertEqual(n, 1)
        self.assertEqual(gen._materiales_registro_a_renumerar, {self.MAT: self.REG})
        self.assertEqual(gen._materiales_registro_a_renumerar_desc, {self.MAT: "MERLUZA"})
        self.assertEqual(gen._materiales_centros_borrado, {self.MAT: ["0110", "0120"]})
        self.assertEqual([ws2.cell(row=r, column=40).value for r in (2, 3, 4)], [None, None, None])
        avisos = [t for nivel, t in logs if nivel == "WARN"]
        self.assertTrue(any(self.MAT in t and "0110, 0120" in t for t in avisos), avisos)

    def test_cabecera_marcada_sigue_igual(self):
        gen, wb, ws1, _ = self._generador(aj_h1="X")
        n = self._limpiar(gen, wb, ws1, [])
        self.assertEqual(n, 1)
        self.assertEqual(gen._materiales_registro_a_renumerar, {self.MAT: self.REG})
        self.assertIsNone(ws1.cell(row=2, column=36).value)

    def test_sin_marcas_no_avisa(self):
        gen, wb, ws1, ws2 = self._generador(aj_h1=None)
        for r in (2, 3, 4):
            ws2.cell(row=r, column=40).value = None
        logs = []
        self.assertEqual(self._limpiar(gen, wb, ws1, logs), 0)
        self.assertFalse(any("Marca de borrado limpiada" in t for _, t in logs))


if __name__ == "__main__":
    unittest.main()
