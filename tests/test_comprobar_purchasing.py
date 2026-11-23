# -*- coding: utf-8 -*-
"""
¿Está al día el Purchasing? (23 nov 2026, petición de Víctor) — app/comun/comprobar_purchasing.py y el aviso en la
ventana «Purchasing encontrado» de «Crear / Actualizar Registro Info SAP».

Datos inventados. Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import shutil
import sys
import tempfile
import time
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl  # noqa: E402

from app.comun import comprobar_purchasing as CP  # noqa: E402
from app.comun import datos_tarifa as DT  # noqa: E402

try:
    import tkinter  # noqa: F401
    HAY_TK = True
except ImportError:
    HAY_TK = False


def _purchasing(ruta, filas):
    """filas: [(registro, proveedor, nombre, material)] en la hoja 1."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = CP.HOJA_1
    ws.append(["Registro info de compras", "Desc", "Proveedor", "Nombre del proveedor", "Material"])
    for reg, prov, nombre, mat in filas:
        ws.append([reg, None, prov, nombre, mat])
    wb.save(ruta)
    return ruta


def _tarifa(ruta, filas, con_columna=True):
    """filas: [(sap, registro)]."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TARIFA"
    ws["C2"] = "PROVEEDOR EJEMPLO SL"
    cab = ["CODIGO SAP", "CODIGO DALI", "CODIGO PROVEEDOR", "DESCRIPCION", "FORMATO", "PRECIO", "DESCUENTO"]
    if con_columna:
        cab.append(DT.CAB_REGISTRO_INFO)
    for c, v in enumerate(cab, start=1):
        ws.cell(7, c).value = v
    for n, (sap, reg) in enumerate(filas, start=8):
        ws.cell(n, 1).value = sap
        ws.cell(n, 6).value = 1.0
        if con_columna:
            ws.cell(n, 8).value = reg
    wb.save(ruta)
    return ruta


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.pur = os.path.join(self.tmp, "Purchasing Info Records.xlsx")
        self.tar = os.path.join(self.tmp, "PROVEEDOR EJEMPLO SL [23-11-2026].xlsx")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _fecha(self, ruta, segundos):
        t = time.time() + segundos
        os.utime(ruta, (t, t))


class TestRegistrosQueFaltan(Base):
    def test_al_dia(self):
        _purchasing(self.pur, [("5300000001", "1009999", "PROVEEDOR EJEMPLO SL", "10100001"),
                               ("5300000002", "1009999", "PROVEEDOR EJEMPLO SL", "10100002")])
        _tarifa(self.tar, [(10100001, 5300000001), (10100002, 5300000002), (10100003, None)])
        r = CP.comprobar(self.tar, self.pur, self.tmp, "1009999", "PROVEEDOR EJEMPLO SL")
        self.assertEqual(r, {"faltan": [], "rinfo_posteriores": []})
        self.assertEqual(CP.textos_aviso(r), [])

    def test_la_tarifa_tiene_registros_que_el_purchasing_no(self):
        _purchasing(self.pur, [("5300000001", "1009999", "PROVEEDOR EJEMPLO SL", "10100001")])
        _tarifa(self.tar, [(10100001, 5300000001), (10100002, 5300000002)])
        r = CP.comprobar(self.tar, self.pur, self.tmp)
        self.assertEqual(r["faltan"], [("10100002", "5300000002")])
        textos = CP.textos_aviso(r)
        self.assertIn("10100002 (5300000002)", textos[0])
        self.assertIn("Descarga de nuevo el Purchasing", textos[-1])

    def test_tarifa_antigua_sin_columna(self):
        _purchasing(self.pur, [("5300000001", "1009999", "PROVEEDOR EJEMPLO SL", "10100001")])
        _tarifa(self.tar, [(10100002, None)], con_columna=False)
        self.assertEqual(CP.comprobar(self.tar, self.pur, self.tmp)["faltan"], [])

    def test_muchos_codigos_se_resumen(self):
        r = {"faltan": [(str(i), str(5300000000 + i)) for i in range(12)], "rinfo_posteriores": []}
        self.assertIn("y 4 más", CP.textos_aviso(r)[0])


class TestRinfoPosteriores(Base):
    def setUp(self):
        super().setUp()
        _purchasing(self.pur, [("5300000001", "1009999", "PROVEEDOR EJEMPLO SL", "10100001")])
        _tarifa(self.tar, [(10100001, 5300000001)])
        self._fecha(self.pur, -3600)

    def _rinfo(self, nombre, prov="1009999", nombre_prov="PROVEEDOR EJEMPLO SL", segundos=0):
        ruta = _purchasing(os.path.join(self.tmp, nombre), [("1", prov, nombre_prov, "10100009")])
        self._fecha(ruta, segundos)
        return ruta

    def test_rinfo_del_proveedor_guardado_despues(self):
        self._rinfo("RINFO PROVEEDOR EJEMPLO SL [23-11-2026].xlsx")
        r = CP.comprobar(self.tar, self.pur, self.tmp, "1009999", "PROVEEDOR EJEMPLO SL")
        self.assertEqual([n for n, _f in r["rinfo_posteriores"]], ["RINFO PROVEEDOR EJEMPLO SL [23-11-2026].xlsx"])
        self.assertIn("Si ya lo importaste en SAP", " ".join(CP.textos_aviso(r)))

    def test_no_cuentan_los_anteriores_ni_los_de_otro_proveedor(self):
        self._rinfo("RINFO PROVEEDOR EJEMPLO SL [20-11-2026].xlsx", segundos=-7200)     # anterior al Purchasing
        self._rinfo("RINFO OTRO PROVEEDOR SA [23-11-2026].xlsx", prov="2000456", nombre_prov="OTRO PROVEEDOR SA")
        self._rinfo("~$RINFO PROVEEDOR EJEMPLO SL [23-11-2026].xlsx")                    # temporal de Excel
        r = CP.comprobar(self.tar, self.pur, self.tmp, "1009999", "PROVEEDOR EJEMPLO SL")
        self.assertEqual(r["rinfo_posteriores"], [])

    def test_por_nombre_si_no_hay_codigo(self):
        self._rinfo("RINFO PROVEEDOR EJEMPLO SL (GY-IT) [23-11-2026].xlsx", prov=None)
        r = CP.comprobar(self.tar, self.pur, self.tmp, None, "Proveedor  Ejemplo SL")
        self.assertEqual(len(r["rinfo_posteriores"]), 1)


@unittest.skipUnless(HAY_TK, "la ventana importa tkinter")
class TestVentanaPurchasingEncontrado(Base):
    def test_avisa_en_la_ventana_y_en_la_consola(self):
        import tkinter as tk
        from app.rinfo import crear_registro_info as CRI
        _purchasing(self.pur, [("5300000001", "1009999", "PROVEEDOR EJEMPLO SL", "10100001")])
        _tarifa(self.tar, [(10100001, 5300000001), (10100002, 5300000002)])
        logs = []
        try:
            raiz = tk.Tk()
        except tk.TclError:
            self.skipTest("sin pantalla")
        raiz.withdraw()
        self.addCleanup(raiz.destroy)
        with mock.patch.object(CRI, "_log", lambda _p, m, n="INFO": logs.append((n, m))):
            v = CRI.VentanaPurchasingEncontrado(
                raiz, self.tar, os.path.basename(self.tar), "PROVEEDOR EJEMPLO SL", "PROVEEDOR EJEMPLO SL",
                "1009999", {"ruta": self.pur, "archivo": "Purchasing Info Records.xlsx", "score": 100},
                self.tmp, raiz)
        textos = []

        def recorrer(w):
            for h in w.winfo_children():
                try:
                    textos.append(h.cget("text"))
                except tk.TclError:
                    pass
                recorrer(h)
        recorrer(v)
        self.assertTrue(any("puede estar desactualizado" in t for t in textos))
        self.assertTrue(any("Cómo descargar el Purchasing" in t for t in textos))
        self.assertTrue(any(n == "WARN" and "5300000002" in m for n, m in logs))
        v.destroy()

    def test_sin_avisos_la_ventana_de_siempre(self):
        import tkinter as tk
        from app.rinfo import crear_registro_info as CRI
        _purchasing(self.pur, [("5300000001", "1009999", "PROVEEDOR EJEMPLO SL", "10100001")])
        _tarifa(self.tar, [(10100001, 5300000001)])
        try:
            raiz = tk.Tk()
        except tk.TclError:
            self.skipTest("sin pantalla")
        raiz.withdraw()
        self.addCleanup(raiz.destroy)
        with mock.patch.object(CRI, "_log", lambda *a, **k: None):
            v = CRI.VentanaPurchasingEncontrado(
                raiz, self.tar, os.path.basename(self.tar), "PROVEEDOR EJEMPLO SL", "PROVEEDOR EJEMPLO SL",
                "1009999", {"ruta": self.pur, "archivo": "Purchasing Info Records.xlsx", "score": 100},
                self.tmp, raiz)
        self.assertEqual(v._avisos_purchasing, [])
        v.destroy()


if __name__ == "__main__":
    unittest.main()
