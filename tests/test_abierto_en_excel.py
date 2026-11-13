# -*- coding: utf-8 -*-
"""Pruebas de «¿abierto en Excel?» antes de quitar [INCOMPLETO] (13 nov 2026): si la tarifa sigue abierta en Excel,
no se renombra (al volver a guardar en Excel reaparecía el «[INCOMPLETO]»).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import sap_helpers as H  # noqa: E402


class AbiertoEnExcel(unittest.TestCase):
    def test_detecta_el_archivo_de_bloqueo(self):
        with tempfile.TemporaryDirectory() as d:
            ruta = os.path.join(d, "PRO A PRO HOSTELERIA ORGANIZADA SAU [INCOMPLETO] [13-11-2026].xlsx")
            open(ruta, "w").close()
            self.assertFalse(H.abierto_en_excel(ruta))
            H.comprobar_cerrado_en_excel(ruta)                       # cerrado: no lanza
            bloqueo = os.path.join(d, "~$" + os.path.basename(ruta))
            open(bloqueo, "w").close()
            self.assertTrue(H.abierto_en_excel(ruta))
            with self.assertRaises(H.ArchivoAbiertoEnExcel) as e:
                H.comprobar_cerrado_en_excel(ruta)
            self.assertIn("abierto en Excel", str(e.exception))
            os.remove(bloqueo)
            open(os.path.join(d, "~$" + os.path.basename(ruta)[2:]), "w").close()   # nombre largo recortado
            self.assertTrue(H.abierto_en_excel(ruta))


if __name__ == "__main__":
    unittest.main()
