# -*- coding: utf-8 -*-
"""
test_rinfo_sin_unidad.py — aviso de materiales sin unidad de medida antes de cargar el RINFO en SAP (18 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Caso real PRO A PRO (18 nov 2026): el SAP 10104121 (DALI 30245) se puso en la ventana «Códigos SAP faltantes»
después de la última Acción 3; el catálogo no tenía su unidad, el RINFO salió sin ella y SAP rechazó el registro con
«Introduzca UM de pedido». El catálogo es la fuente de los datos SAP de la tarifa: si falta la unidad, se avisa.
"""
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import catalogo_local as C  # noqa: E402
from tests.test_catalogo_local import crear_catalogo  # noqa: E402

try:
    import openpyxl
    from app.rinfo import comun_rinfo as R
    from app.rinfo import creador_rinfo as CR
    HAY_RINFO = True
except Exception:  # noqa: BLE001 — sin tkinter
    HAY_RINFO = False


@unittest.skipUnless(HAY_RINFO, "necesita tkinter")
class TestMaterialesSinUnidad(unittest.TestCase):
    def test_lista_los_que_no_tienen_unidad(self):
        dic = {"10100199": ("TARTA", "CJ", "", ""), "10104121": ("", "", "", "")}
        self.assertEqual(R._materiales_sin_unidad(["10100199", "10104121", "10104121", "999"], dic),
                         ["10104121", "999"])
        self.assertEqual(R._materiales_sin_unidad(["10100199"], dic), [])

    def test_texto_del_aviso(self):
        texto = R.AVISO_SIN_UNIDAD.format(n=1, lista="10104121")
        self.assertIn("10104121", texto)
        self.assertIn("Acción 2", texto)   # (v10.2.86) antes Acción 3


@unittest.skipUnless(HAY_RINFO, "necesita tkinter")
class TestCreadorAvisa(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.proyecto = self.dir.name
        os.makedirs(os.path.join(self.proyecto, "TARIFAS"))
        C._CACHE.clear()
        crear_catalogo(C.ruta_catalogo(self.proyecto), [
            {"clave": 1, "descripcion": "TARTA", "activo": "SI", "cod_sap": "10100199", "desc_sap": "TARTA FLAN",
             "unidad_sap": "CJ"},
            {"clave": 30245, "descripcion": "PULLED COCHINITA PIBIL", "activo": "SI", "cod_sap": "10104121"},
        ])

    def tearDown(self):
        self.dir.cleanup()

    def test_crear_desde_cero_avisa_del_material_sin_unidad(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = CR.CreadorRINFO.HOJA_PRINCIPAL
        ws["E2"], ws["E3"] = "10100199", "10104121"
        cre = object.__new__(CR.CreadorRINFO)
        cre._parent = None
        cre._carpeta_tarifas = os.path.join(self.proyecto, "TARIFAS")
        cre._carpeta_proyecto = self.proyecto
        logs = []
        with mock.patch.object(CR, "_log", lambda _p, m, n="INFO": logs.append((n, m))):
            cre._cruzar_con_listado_dali(wb)
        self.assertEqual(cre.nuevos_sin_unidad, ["10104121"])
        self.assertEqual((ws["N2"].value, ws["P2"].value), ("CJ", "CJ"))
        self.assertIsNone(ws["N3"].value)
        avisos = [m for n, m in logs if n == "WARN" and "sin unidad de medida" in m]
        self.assertEqual(len(avisos), 1)
        self.assertIn("10104121", avisos[0])


if __name__ == "__main__":
    unittest.main()
