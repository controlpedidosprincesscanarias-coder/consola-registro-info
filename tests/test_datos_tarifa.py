# -*- coding: utf-8 -*-
"""
Teléfono del proveedor y hoteles en la cabecera de la TARIFA final (22 nov 2026, petición de Víctor).

- app/comun/datos_tarifa.py: escribir los huecos, leerlos (también en solo lectura, con X, ✔, «sí»…), teléfono válido.
- La TARIFA final de la Acción 3 (transformar) sale con «Teléfono:» en B4 y el recuadro HOTELES en E1:H6 (encima de la tabla), en blanco,
  sin mover la tabla (cabecera en la fila 7).
- «Crear / Actualizar Registro Info SAP» no pregunta el teléfono ni los centros si la tarifa los trae, y pregunta
  como siempre si faltan o en una tarifa antigua.

Datos inventados. Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import shutil
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl  # noqa: E402

from app.comun import datos_tarifa as DT  # noqa: E402

try:
    import tkinter  # noqa: F401
    HAY_TK = True
except ImportError:
    HAY_TK = False


def _tarifa_antigua(ws):
    """Cabecera como la de las tarifas de antes del 22 nov 2026 (fila 4 vacía, sin hoteles)."""
    ws["B1"], ws["C1"] = "Código:", 4321
    ws["B2"], ws["C2"] = "Proveedor:", "PROVEEDOR EJEMPLO SL"
    ws["B3"], ws["C3"] = "Periodo:", "01/01/2027 - 31/12/2027"
    ws["B5"], ws["C5"], ws["D5"] = "Dto.(%):", "línea", "general"
    for c, cab in enumerate(["CODIGO SAP", "CODIGO DALI", "CODIGO PROVEEDOR",
                             "DESCRIPCION", "FORMATO", "PRECIO", "DESCUENTO"], start=1):
        ws.cell(7, c).value = cab
    ws.append([10100001, 5001, "REF-1", "ACEITE DE PRUEBA 1L", "UD", 3.5, None])


class TestTelefonoYMarcas(unittest.TestCase):
    def test_telefono_valido(self):
        self.assertEqual(DT.limpiar_telefono("928 000 000"), "928000000")
        self.assertEqual(DT.limpiar_telefono("928-00-00-00"), "928000000")
        self.assertEqual(DT.limpiar_telefono(928000000), "928000000")
        self.assertEqual(DT.limpiar_telefono(928000000.0), "928000000")
        self.assertEqual(DT.limpiar_telefono("6000000000"), "6000000000")

    def test_telefono_no_valido(self):
        for v in (None, "", "1234567", "12345678901", "+34 928 000 000", "NO TIENE", "928 00O 000"):
            self.assertIsNone(DT.limpiar_telefono(v), v)

    def test_marcas(self):
        for v in ("✔", "X", "x", "sí", "SI", 1, True, "ok"):
            self.assertTrue(DT.es_marca(v), v)
        for v in (None, "", "  ", "NO", "n", 0, "0", False, "-"):
            self.assertFalse(DT.es_marca(v), v)

    def test_siglas_y_posiciones(self):
        self.assertEqual([s for s, _c, _n in DT.CENTROS_TARIFA],
                         ["GY", "IT", "LP", "TA", "ST", "MG", "GC", "MT", "JN", "FV"])
        self.assertEqual(DT.posicion_centro(0), (2, 5, 6))     # GY en E2, casilla F2
        self.assertEqual(DT.posicion_centro(4), (6, 5, 6))     # ST en E6
        self.assertEqual(DT.posicion_centro(5), (2, 7, 8))     # MG en G2, casilla H2
        self.assertEqual(DT.posicion_centro(9), (6, 7, 8))     # FV en G6

    @unittest.skipUnless(HAY_TK, "comun_rinfo importa tkinter")
    def test_mismos_centros_que_el_rinfo(self):
        from app.rinfo.comun_rinfo import CENTROS_SAP_CANARIAS
        self.assertEqual([(c, n) for _s, c, n in DT.CENTROS_TARIFA], list(CENTROS_SAP_CANARIAS.items()))


class TestEscribirYLeer(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.ruta = os.path.join(self.tmp, "PROVEEDOR EJEMPLO SL [22-11-2026].xlsx")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _guardar(self, telefono=None, codigos=()):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TARIFA"
        _tarifa_antigua(ws)
        DT.escribir_huecos(ws, telefono=telefono, codigos_marcados=codigos)
        wb.save(self.ruta)
        return self.ruta

    def test_en_blanco(self):
        d = DT.leer_datos_tarifa(self._guardar())
        self.assertTrue(d["tiene_huecos"])
        self.assertIsNone(d["telefono"])
        self.assertEqual(d["telefono_texto"], "")
        self.assertEqual(d["centros"], [])

    def test_rellena(self):
        d = DT.leer_datos_tarifa(self._guardar("928 111 222", ["0910", "0110", "0610"]))
        self.assertEqual(d["telefono"], "928111222")
        self.assertEqual(d["centros"], [("0110", "GUAYARMINA PRINCESS"), ("0610", "GRAN CANARIA PRINCESS"),
                                        ("0910", "FUERTEVENTURA PRINCESS")])
        self.assertEqual(d["siglas"], ["GY", "GC", "FV"])

    def test_formato_de_los_huecos(self):
        wb = openpyxl.load_workbook(self._guardar())
        ws = wb.active
        self.assertEqual(ws["B4"].value, "Teléfono:")
        self.assertEqual(ws["E1"].value, "HOTELES")
        self.assertIn("E1:H1", [str(r) for r in ws.merged_cells.ranges])
        self.assertEqual([ws["E2"].value, ws["E6"].value, ws["G2"].value, ws["G6"].value], ["GY", "ST", "MG", "FV"])
        # El aviso de descuentos pasa a B6:D6 y nada queda a la derecha de la columna H
        self.assertEqual(ws["B6"].value, "DESCUENTOS YA APLICADOS EN PRECIO FINAL")
        self.assertTrue(all(ws.cell(r, c).value is None for r in range(1, 7) for c in range(9, 15)))
        # La lista desplegable «✔» cubre las 10 casillas y sobrevive a abrir y guardar con openpyxl
        # (la Acción 4 reescribe así los archivos INC).
        wb.save(self.ruta)
        ws = openpyxl.load_workbook(self.ruta).active
        dvs = ws.data_validations.dataValidation
        self.assertEqual(len(dvs), 1)
        self.assertEqual(dvs[0].formula1, '"✔"')
        rango = str(dvs[0].sqref)
        for celda in ("F2", "F6", "H2", "H6"):
            self.assertIn(celda, rango)
        # La tabla no se mueve
        self.assertEqual(ws["A7"].value, "CODIGO SAP")
        self.assertEqual(ws["A8"].value, 10100001)

    def test_marcado_a_mano_en_excel(self):
        """Lo que escribiría el usuario: teléfono con espacios, X o «sí» en vez de ✔, «NO» = sin marcar."""
        self._guardar()
        wb = openpyxl.load_workbook(self.ruta)
        ws = wb.active
        ws["C4"] = "928 33 44 55"
        ws["F5"] = "x"        # TA
        ws["H3"] = "sí"       # GC
        ws["F2"] = "NO"       # GY, no
        wb.save(self.ruta)
        d = DT.leer_datos_tarifa(self.ruta)
        self.assertEqual(d["telefono"], "928334455")
        self.assertEqual(d["siglas"], ["TA", "GC"])

    def test_telefono_mal_escrito(self):
        d = DT.leer_datos_tarifa(self._guardar("928 11"))
        self.assertIsNone(d["telefono"])
        self.assertEqual(d["telefono_texto"], "928 11")

    def test_tarifa_antigua_sin_huecos(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TARIFA"
        _tarifa_antigua(ws)
        wb.save(self.ruta)
        d = DT.leer_datos_tarifa(self.ruta)
        self.assertEqual(d, {"tiene_huecos": False, "telefono_texto": "", "telefono": None,
                             "centros": [], "siglas": []})

    def test_archivo_que_no_se_puede_leer(self):
        d = DT.leer_datos_tarifa(os.path.join(self.tmp, "no existe.xlsx"))
        self.assertFalse(d["tiene_huecos"])
        self.assertEqual(d["centros"], [])


class _HojaXls:
    """Hoja .xls inventada con la forma de TARIFA.xls de DALI (formato A), para transformar() sin xlwt."""

    class _C:
        def __init__(self, v):
            self.value = v

    def __init__(self, celdas):
        self._celdas = celdas   # {(fila_1, col_1): valor}
        self.nrows = max(r for r, _c in celdas)
        self.ncols = 10

    def cell(self, r, c):
        return self._C(self._celdas.get((r + 1, c + 1), ""))


class TestTarifaFinalDeLaAccion3(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        os.makedirs(os.path.join(self.tmp, "TARIFAS"))
        self.ruta_xls = os.path.join(self.tmp, "TARIFAS", "TARIFA.xls")
        open(self.ruta_xls, "wb").close()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_sale_con_los_huecos_en_blanco(self):
        from app.tarifas import transformar_tarifa as TT
        celdas = {(2, 4): 4321, (4, 4): "PROVEEDOR EJEMPLO SL", (8, 4): "", (10, 4): "",
                  (17, 3): "01/01/2027 - 31/12/2027",
                  (24, 1): 5001, (24, 2): "REF-1", (24, 4): "ACEITE DE PRUEBA 1L", (24, 6): "UD", (24, 7): 3.5,
                  (25, 1): 5002, (25, 2): "REF-2", (25, 4): "VINAGRE DE PRUEBA 1L", (25, 6): "UD", (25, 7): 1.25}
        with mock.patch.object(TT, "_leer_xls", return_value=_HojaXls(celdas)):
            ruta, _analisis = TT.transformar(self.ruta_xls)
        ws = openpyxl.load_workbook(ruta).active
        self.assertEqual(ws["C2"].value, "PROVEEDOR EJEMPLO SL")
        self.assertEqual(ws["C3"].value, "01/01/2027 - 31/12/2027")
        self.assertEqual(ws["B4"].value, "Teléfono:")
        self.assertIsNone(ws["C4"].value)
        self.assertEqual(ws["E1"].value, "HOTELES")
        self.assertEqual(ws["B6"].value, "DESCUENTOS YA APLICADOS EN PRECIO FINAL")
        self.assertLess(ws.column_dimensions["B"].width, 30)   # el aviso de B6:D6 no ensancha la columna B
        self.assertEqual(ws["A7"].value, "CODIGO SAP")
        self.assertEqual(ws["B8"].value, 5001)
        self.assertGreater(ws.row_dimensions[4].height, 6)   # ya no es un separador
        d = DT.leer_datos_tarifa(ruta)
        self.assertTrue(d["tiene_huecos"])
        self.assertIsNone(d["telefono"])
        self.assertEqual(d["centros"], [])


@unittest.skipUnless(HAY_TK, "crear_registro_info importa tkinter")
class TestCrearRinfoNoPregunta(unittest.TestCase):
    def setUp(self):
        from app.rinfo import crear_registro_info as CRI
        self.CRI = CRI
        self.tmp = tempfile.mkdtemp()
        self.ruta = os.path.join(self.tmp, "PROVEEDOR EJEMPLO SL [22-11-2026].xlsx")
        self.logs = []
        self.p_log = mock.patch.object(CRI, "_log", lambda _p, m, n="INFO": self.logs.append((n, m)))
        self.p_log.start()

    def tearDown(self):
        self.p_log.stop()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _tarifa(self, telefono=None, codigos=(), huecos=True):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TARIFA"
        _tarifa_antigua(ws)
        if huecos:
            DT.escribir_huecos(ws, telefono=telefono, codigos_marcados=codigos)
        wb.save(self.ruta)

    def test_con_datos_no_abre_ventanas(self):
        self._tarifa("928 111 222", ["0110", "0810"])
        recibido = []
        with mock.patch.object(self.CRI, "VentanaTelefonoRINFO") as v_tel, \
                mock.patch.object(self.CRI, "VentanaCentrosRINFO") as v_cen:
            self.CRI._telefono_de_tarifa_o_preguntar(None, self.ruta, recibido.append)
            self.CRI._centros_de_tarifa_o_preguntar(None, self.ruta, recibido.append)
        v_tel.assert_not_called()
        v_cen.assert_not_called()
        self.assertEqual(recibido, ["928111222", [("0110", "GUAYARMINA PRINCESS"), ("0810", "CLUB JANDIA PRINCESS")]])
        self.assertTrue(any("928111222" in m for _n, m in self.logs))
        self.assertTrue(any("GY, JN" in m for _n, m in self.logs))

    def test_sin_datos_pregunta_como_siempre(self):
        for huecos in (True, False):
            self._tarifa(huecos=huecos)
            cb = mock.Mock()
            with mock.patch.object(self.CRI, "VentanaTelefonoRINFO") as v_tel, \
                    mock.patch.object(self.CRI, "VentanaCentrosRINFO") as v_cen:
                self.CRI._telefono_de_tarifa_o_preguntar("padre", self.ruta, cb)
                self.CRI._centros_de_tarifa_o_preguntar("padre", self.ruta, cb)
            v_tel.assert_called_once_with("padre", cb)
            v_cen.assert_called_once_with("padre", cb)
            cb.assert_not_called()

    def test_telefono_mal_escrito_avisa_y_pregunta(self):
        self._tarifa("12 34", ["0110"])
        with mock.patch.object(self.CRI, "VentanaTelefonoRINFO") as v_tel:
            self.CRI._telefono_de_tarifa_o_preguntar(None, self.ruta, mock.Mock())
        v_tel.assert_called_once()
        self.assertTrue(any(n == "WARN" and "12 34" in m for n, m in self.logs))

    def test_los_dos_flujos_usan_la_tarifa(self):
        """Purchasing encontrado y «crear desde cero» pasan por las dos funciones (ninguna llamada directa queda)."""
        import inspect
        src = inspect.getsource(self.CRI)
        self.assertEqual(src.count("_telefono_de_tarifa_o_preguntar(self._parent_ventana"), 2)
        self.assertEqual(src.count("_centros_de_tarifa_o_preguntar(self._parent_ventana"), 2)
        self.assertEqual(src.count("VentanaTelefonoRINFO(self._parent_ventana"), 0)
        self.assertEqual(src.count("VentanaCentrosRINFO(self._parent_ventana"), 0)


if __name__ == "__main__":
    unittest.main()
