# -*- coding: utf-8 -*-
"""Con varios Purchasing Info Records del mismo proveedor en TARIFAS, gana el más reciente (19 nov 2026).

Caso real: se usó el Purchasing de antes de importar el RINFO anterior y SAP contestó «Ya existe el registro Info de
compras…» con los materiales creados el día antes. Datos inventados.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import tempfile
import unittest
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import openpyxl  # noqa: E402

from app.rinfo.comun_rinfo import _buscar_purchasing_compatible  # noqa: E402


def _purchasing(ruta, proveedor, creado):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Info general p.reg.info"
    ws.append(["Registro info de compras", "Descripción", "Proveedor", "Nombre del proveedor", "Material"])
    ws.append(["5300000001", None, "1000001", proveedor, "10500001"])
    wb.properties.created = creado
    wb.save(ruta)


class PurchasingMasReciente(unittest.TestCase):
    def test_gana_el_mas_reciente_aunque_salga_antes_el_antiguo(self):
        with tempfile.TemporaryDirectory() as d:
            # El nuevo se escribe primero: la carpeta lo devuelve antes o después según el sistema
            for nombre, dia in (("Purchasing Info Records (2).xlsx", 19), ("Purchasing Info Records.xlsx", 18),
                                ("Purchasing Info Records (1).xlsx", 17)):
                _purchasing(os.path.join(d, nombre), "SUMINISTROS DEMO SL", datetime(2026, 11, dia, 8, 0))
            r = _buscar_purchasing_compatible("SUMINISTROS DEMO SL", d)
        self.assertEqual(r["archivo"], "Purchasing Info Records (2).xlsx")
        self.assertEqual(r["fecha"], datetime(2026, 11, 19, 8, 0))

    def test_el_parecido_del_proveedor_manda_sobre_la_fecha(self):
        with tempfile.TemporaryDirectory() as d:
            _purchasing(os.path.join(d, "Purchasing Info Records.xlsx"), "SUMINISTROS DEMO SL", datetime(2026, 11, 1))
            _purchasing(os.path.join(d, "Purchasing Info Records (1).xlsx"), "OTRO PROVEEDOR EJEMPLO SA",
                        datetime(2026, 11, 19))
            r = _buscar_purchasing_compatible("SUMINISTROS DEMO SL", d)
        self.assertEqual(r["archivo"], "Purchasing Info Records.xlsx")

    def test_sin_purchasing_del_proveedor(self):
        with tempfile.TemporaryDirectory() as d:
            _purchasing(os.path.join(d, "Purchasing Info Records.xlsx"), "OTRO PROVEEDOR EJEMPLO SA",
                        datetime(2026, 11, 19))
            self.assertIsNone(_buscar_purchasing_compatible("SUMINISTROS DEMO SL", d))


if __name__ == "__main__":
    unittest.main()
