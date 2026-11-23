# -*- coding: utf-8 -*-
"""
Acción 6 «Generar plantilla Tarifa desde RINFO» (22 nov 2026, petición de Víctor).

- app/comun/plantilla_rinfo.py: leer un RINFO / Purchasing (hojas 1 a 4), precio vigente, borrados, caducados,
  escalados, una tarifa por proveedor y por grupo de hoteles con los mismos precios, nombres sin pisar nada.
- app/tarifas/plantilla_tarifa.py: la plantilla Excel (cabecera con teléfono y hoteles, columnas amarillas).
- La vuelta: la plantilla rellena por el proveedor en «Crear / Actualizar Registro Info SAP» (PRECIO NUEVO y
  FECHA INICIO, datos_tarifa.precios_tarifa / periodo_propuesto).
- Botón 6 y ventana (sin Windows no se puede arrastrar: se prueba el filtro de archivos).

Todo con datos inventados. Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import datetime as dt
import os
import shutil
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl  # noqa: E402

from app.comun import datos_tarifa as DT  # noqa: E402
from app.comun import plantilla_rinfo as PR  # noqa: E402

try:
    import tkinter  # noqa: F401
    HAY_TK = True
except ImportError:
    HAY_TK = False

HOY = dt.date(2026, 11, 22)


def _fila(n_cols, valores):
    fila = [None] * n_cols
    for col, v in valores.items():
        fila[col - 1] = v
    return fila


def crear_purchasing(ruta, h1=(), h2=(), h3=(), h4=()):
    """Purchasing inventado: cada elemento es {columna (1 = A): valor}."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for nombre, filas, n in ((PR.HOJA_1, h1, 37), (PR.HOJA_2, h2, 41), (PR.HOJA_3, h3, 27), (PR.HOJA_4, h4, 29)):
        ws = wb.create_sheet(nombre)
        ws.append([f"Columna {i}" for i in range(1, n + 1)])
        for f in filas:
            ws.append(_fila(n, f))
    wb.save(ruta)
    return ruta


def h1(reg, mat, prov="1000123", nombre="SUMINISTROS DEMO SL", ref=None, tel=None, borrado=None, texto=None):
    return {1: reg, 3: prov, 4: nombre, 5: mat, 6: texto or f"MATERIAL DEMO {mat}", 14: "UN", 28: ref, 30: tel,
            36: borrado}


def h3(reg, mat, centro, precio, ini="01.01.2026", fin="31.12.2026", por=1, borrado=None):
    return {1: reg, 5: centro, 6: mat, 9: ini, 10: fin, 15: precio, 17: por, 25: borrado}


TODOS = [c for _s, c, _n in DT.CENTROS_TARIFA]


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.ruta = os.path.join(self.tmp, "Purchasing Info Records.xlsx")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)


class TestLeer(Base):
    def test_precio_vigente_y_datos(self):
        crear_purchasing(self.ruta,
                         h1=[h1("5300001", "10100001", ref="REF-1", tel="928 000 000")],
                         h3=[h3("5300001", "10100001", "0110", 4.10, "01.01.2025", "31.12.2025"),
                             h3("5300001", "10100001", "0110", 4.35, "01.01.2026", "31.12.2026"),
                             h3("5300001", "10100001", "0110", 9.99, "01.01.2026", "31.12.2026", borrado="X")])
        d = PR.leer_purchasing(self.ruta, hoy=HOY)
        p = d["proveedores"]["1000123"]
        self.assertEqual((p["nombre"], p["telefono"]), ("SUMINISTROS DEMO SL", "928 000 000"))
        m = p["materiales"]["10100001"]
        self.assertEqual(m["cod_prov_mat"], "REF-1")
        self.assertEqual(m["precios"]["0110"], (4.35, dt.date(2026, 1, 1), dt.date(2026, 12, 31)))

    def test_sin_vigente_usa_el_proximo_y_los_caducados_no_salen(self):
        crear_purchasing(self.ruta,
                         h1=[h1("1", "A1"), h1("2", "A2")],
                         h3=[h3("1", "A1", "0110", 3.0, "01.01.2027", "31.12.2027"),
                             h3("1", "A1", "0110", 2.0, "01.06.2027", "31.12.2027"),
                             h3("2", "A2", "0110", 5.0, "01.01.2025", "31.12.2025")])
        d = PR.leer_purchasing(self.ruta, hoy=HOY)
        mats = d["proveedores"]["1000123"]["materiales"]
        self.assertEqual(mats["A1"]["precios"]["0110"][0], 3.0)
        self.assertEqual(mats["A2"]["precios"], {})
        self.assertIn(("A2", "0110"), d["avisos"]["caducados"])
        self.assertIn("A2", d["avisos"]["sin_precio"])

    def test_borrados_registro_y_centro(self):
        crear_purchasing(self.ruta,
                         h1=[h1("1", "B1", borrado="X"), h1("2", "B2")],
                         h2=[{1: "2", 6: "0810", 10: "B2", 40: "X"}],
                         h3=[h3("1", "B1", "0110", 1.0), h3("2", "B2", "0110", 2.0), h3("2", "B2", "0810", 2.0)])
        d = PR.leer_purchasing(self.ruta, hoy=HOY)
        mats = d["proveedores"]["1000123"]["materiales"]
        self.assertNotIn("B1", mats)
        self.assertEqual(set(mats["B2"]["precios"]), {"0110"})
        self.assertEqual(d["avisos"]["borrados"], {"B1"})
        self.assertEqual(d["avisos"]["centros_borrados"], 1)

    def test_rinfo_de_la_consola_sin_registro_info_y_precio_por_unidades(self):
        """En un RINFO recién generado los materiales nuevos aún no tienen nº de registro: se une por material.
        Fechas como datetime (descarga de SAP) y precio por 10 unidades."""
        crear_purchasing(self.ruta,
                         h1=[h1(None, "C1")],
                         h3=[h3(None, "C1", "0910", 25.0, dt.datetime(2026, 1, 1), dt.datetime(2026, 12, 31),
                                por=10)])
        d = PR.leer_purchasing(self.ruta, hoy=HOY)
        self.assertAlmostEqual(d["proveedores"]["1000123"]["materiales"]["C1"]["precios"]["0910"][0], 2.5)

    def test_escalas_y_centros_desconocidos(self):
        crear_purchasing(self.ruta, h1=[h1("1", "D1")],
                         h3=[h3("1", "D1", "0110", 1.0), h3("1", "D1", "9999", 1.0)],
                         h4=[{1: "1", 5: "0110", 6: "D1"}])
        d = PR.leer_purchasing(self.ruta, hoy=HOY)
        self.assertEqual(d["avisos"]["escalas"], {"D1"})
        self.assertEqual(d["avisos"]["centros_desconocidos"], {"9999"})

    def test_no_es_purchasing(self):
        wb = openpyxl.Workbook()
        wb.active.title = "TARIFA"
        wb.save(self.ruta)
        with self.assertRaises(PR.NoEsPurchasing):
            PR.leer_purchasing(self.ruta)


class TestAgrupar(Base):
    def _datos(self):
        filas_h1 = [h1("1", "E1", ref="R1", tel="928111222"), h1("2", "E2", ref="R2"),
                    h1("3", "F1", prov="2000456", nombre="BEBIDAS EJEMPLO SA", tel="922333444")]
        filas_h3 = []
        for c in TODOS:
            filas_h3.append(h3("1", "E1", c, 4.10 if c in ("0810", "0910") else 4.35))
            filas_h3.append(h3("2", "E2", c, 1.20))
        filas_h3 += [h3("3", "F1", "0110", 0.80), h3("3", "F1", "0120", 0.80)]
        crear_purchasing(self.ruta, h1=filas_h1, h3=filas_h3)
        return PR.leer_purchasing(self.ruta, hoy=HOY)

    def test_una_tarifa_por_proveedor_y_por_grupo_de_precios(self):
        catalogo = PR.mapa_catalogo([(5001, "ACEITE DEMO 1L", "UD", 10100001), ("5002", "x", "CJ", "E2"),
                                     ("5003", "VINAGRE DEMO 1L", "UD", "E2")])
        pls = PR.preparar_plantillas(self._datos(), catalogo, fecha=HOY)
        nombres = [p["nombre_archivo"] for p in pls]
        self.assertEqual(nombres, [
            "SUMINISTROS DEMO SL (GY-IT-LP-TA-ST-MG-GC-MT) [22-11-2026].xlsx",
            "SUMINISTROS DEMO SL (JN-FV) [22-11-2026].xlsx",
            "BEBIDAS EJEMPLO SA [22-11-2026].xlsx"])
        demo8, demo2, bebidas = pls
        self.assertEqual(demo8["centros"], TODOS[:8])
        self.assertEqual(demo2["siglas"], ["JN", "FV"])
        self.assertEqual(demo2["telefono"], "928111222")
        self.assertEqual({(f["sap"], f["precio"]) for f in demo2["filas"]}, {("E1", 4.10), ("E2", 1.20)})
        # E2 tiene dos DALI en el catálogo → dos filas; E1 no está → descripción y unidad de SAP
        e2 = [f for f in demo8["filas"] if f["sap"] == "E2"]
        self.assertEqual([f["dali"] for f in e2], ["5002", "5003"])
        e1 = [f for f in demo8["filas"] if f["sap"] == "E1"][0]
        self.assertEqual((e1["dali"], e1["descripcion"], e1["formato"], e1["cod_prov"]),
                         ("", "MATERIAL DEMO E1", "UN", "R1"))
        self.assertEqual(bebidas["siglas"], ["GY", "IT"])

    def test_todos_los_hoteles_con_el_mismo_precio_una_sola_tarifa(self):
        crear_purchasing(self.ruta, h1=[h1("1", "G1")], h3=[h3("1", "G1", c, 2.0) for c in TODOS])
        pls = PR.preparar_plantillas(PR.leer_purchasing(self.ruta, hoy=HOY), {}, fecha=HOY)
        self.assertEqual([p["nombre_archivo"] for p in pls], ["SUMINISTROS DEMO SL [22-11-2026].xlsx"])
        self.assertEqual(pls[0]["centros"], TODOS)

    def test_nombres_sin_pisar_y_sin_caracteres_raros(self):
        ya = ["SUMINISTROS DEMO SL [22-11-2026].xlsx", "suministros demo sl (2) [22-11-2026].xlsx"]
        self.assertEqual(PR.nombre_archivo("SUMINISTROS DEMO SL", ["GY"], True, HOY, ya),
                         "SUMINISTROS DEMO SL (3) [22-11-2026].xlsx")
        self.assertEqual(PR.nombre_archivo('A/B: "C" S.L.', ["GY"], True, HOY), "A B C S.L [22-11-2026].xlsx")
        # Y con el patrón de las tarifas válidas (las ofrece «Crear / Actualizar Registro Info SAP»)
        import re
        patron = re.compile(r'^.+\[\d{2}-\d{2}-\d{4}\]\.xlsx$', re.IGNORECASE)
        self.assertTrue(patron.match(PR.nombre_archivo("X", ["JN", "FV"], False, HOY)))


class TestPlantillaExcelYVuelta(Base):
    def _generar(self):
        crear_purchasing(self.ruta,
                         h1=[h1("5300001", "10100001", ref="REF-1", tel="928 111 222"),
                             h1("5300002", "10100002", ref="REF-2")],
                         h3=[h3("5300001", "10100001", "0110", 4.35), h3("5300001", "10100001", "0410", 4.35),
                             h3("5300002", "10100002", "0110", 1.2), h3("5300002", "10100002", "0410", 1.2)])
        from app.tarifas.plantilla_tarifa import escribir_plantilla
        cat = PR.mapa_catalogo([(90001, "ACEITE DEMO 1L", "UD", 10100001)])
        pl = PR.preparar_plantillas(PR.leer_purchasing(self.ruta, hoy=HOY), cat, fecha=HOY)[0]
        return escribir_plantilla(pl, self.tmp), pl

    def test_aspecto_de_la_plantilla(self):
        ruta, pl = self._generar()
        self.assertEqual(os.path.basename(ruta), "SUMINISTROS DEMO SL [22-11-2026].xlsx")
        ws = openpyxl.load_workbook(ruta).active
        self.assertEqual(ws.title, "TARIFA")
        self.assertEqual((ws["C1"].value, ws["C2"].value), (1000123, "SUMINISTROS DEMO SL"))
        self.assertIsNone(ws["C3"].value)                         # nuevo periodo, lo pone el proveedor
        self.assertEqual([ws.cell(7, c).value for c in range(1, 11)],
                         ["CODIGO SAP", "CODIGO DALI", "CODIGO PROVEEDOR", "DESCRIPCION", "FORMATO", "PRECIO",
                          "DESCUENTO", "REGISTRO INFO", "PRECIO NUEVO", "FECHA INICIO"])
        self.assertEqual([ws.cell(8, c).value for c in range(1, 9)],
                         [10100001, 90001, "REF-1", "ACEITE DEMO 1L", "UD", 4.35, None, 5300001])
        self.assertEqual(ws.cell(9, 8).value, 5300002)                   # Registro Info de cada material
        self.assertIsNone(ws.cell(8, 8).fill.fgColor.rgb if ws.cell(8, 8).fill.fill_type else None)
        self.assertEqual(ws.cell(8, 9).fill.fgColor.rgb[-6:], "FFF2A8")   # amarillo
        self.assertEqual(ws.freeze_panes, "A8")
        d = DT.leer_datos_tarifa(ruta)
        self.assertEqual(d["telefono"], "928111222")
        self.assertEqual(d["siglas"], ["GY", "ST"])

    def test_no_pisa_un_archivo_existente(self):
        from app.tarifas.plantilla_tarifa import escribir_plantilla
        ruta, pl = self._generar()
        with self.assertRaises(FileExistsError):
            escribir_plantilla(pl, self.tmp)

    def test_vuelta_precio_nuevo_y_fecha_inicio(self):
        ruta, _pl = self._generar()
        # Sin rellenar: los precios de siempre y el periodo vacío
        ws = openpyxl.load_workbook(ruta).active
        self.assertEqual(DT.precios_tarifa(ws), ({"10100001": 4.35, "10100002": 1.2}, 0))
        self.assertEqual(DT.periodo_propuesto(ruta), "")
        # Lo que hace el proveedor: precio nuevo en una fila (otra con texto «1,30 €»), fecha y periodo
        wb = openpyxl.load_workbook(ruta)
        ws = wb.active
        ws["I8"] = 4.60
        ws["J8"] = dt.datetime(2027, 1, 1)
        ws["I9"] = "1,30 €"
        ws["J9"] = "15/01/2027"
        ws["C3"] = "AÑO 2027"
        wb.save(ruta)
        ws = openpyxl.load_workbook(ruta, read_only=True, data_only=True).active
        self.assertEqual(DT.precios_tarifa(ws), ({"10100001": 4.60, "10100002": 1.30}, 2))
        self.assertEqual(DT.periodo_propuesto(ruta), "01/01/2027 - 31/12/2027")

    def test_periodo_sin_fin_reconocible_hasta_fin_de_anio(self):
        ruta, _pl = self._generar()
        wb = openpyxl.load_workbook(ruta)
        wb.active["J8"] = "01/03/2027"
        wb.save(ruta)
        self.assertEqual(DT.periodo_propuesto(ruta), "01/03/2027 - 31/12/2027")

    def test_tarifa_normal_sin_columnas_del_proveedor(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TARIFA"
        ws["C3"] = "4º TRIMESTRE 2026"
        ws.append([])
        for _ in range(5):
            ws.append([])
        ws["A7"], ws["F7"] = "CODIGO SAP", "PRECIO"
        ws["A8"], ws["F8"] = 10100001.0, 2.5
        ruta = os.path.join(self.tmp, "normal.xlsx")
        wb.save(ruta)
        self.assertEqual(DT.precios_tarifa(openpyxl.load_workbook(ruta).active), ({"10100001": 2.5}, 0))
        self.assertEqual(DT.periodo_propuesto(ruta), "4º TRIMESTRE 2026")


class TestRegistroInfo(Base):
    """(22 nov 2026, petición de Víctor) Columna REGISTRO INFO (H) en las dos tarifas."""

    def test_registros_del_purchasing_tambien_los_borrados(self):
        crear_purchasing(self.ruta,
                         h1=[h1("5300001", "10100001"), h1("5300009", "10100009", borrado="X"),
                             h1("5400001", "20200001", prov="2000456", nombre="BEBIDAS EJEMPLO SA")],
                         h3=[h3("5300001", "10100001", "0110", 1.0)])
        self.assertEqual(PR.registros_info_de(self.ruta, "SUMINISTROS DEMO SL"),
                         {"10100001": "5300001", "10100009": "5300009"})
        self.assertEqual(PR.registros_info_de(self.ruta, "otro nombre", codigo_proveedor="2000456"),
                         {"20200001": "5400001"})
        self.assertEqual(PR.registros_info_de(self.ruta, "NO ESTÁ"), {})        # dos proveedores y ninguno casa

    def test_rinfo_de_la_consola_sin_los_correlativos_provisionales(self):
        """(23 nov 2026, caso real) En un RINFO de la Consola los materiales nuevos llevan en la columna A un
        correlativo (1, 2…) y «Nueva» = X: no es un Registro Info de SAP y no sale en la plantilla."""
        nuevo1 = {**h1("1", "10104146"), 37: "X"}
        crear_purchasing(self.ruta, h1=[h1("5300008221", "10100210"), nuevo1],
                         h3=[h3("5300008221", "10100210", "0110", 1.19), h3("1", "10104146", "0110", 5.91)])
        d = PR.leer_purchasing(self.ruta, hoy=HOY)
        self.assertEqual(d["proveedores"]["1000123"]["registros"], {"10100210": "5300008221"})
        filas = PR.preparar_plantillas(d, {}, fecha=HOY)[0]["filas"]
        self.assertEqual({f["sap"]: f["registro_info"] for f in filas}, {"10100210": "5300008221", "10104146": ""})
        self.assertEqual(d["proveedores"]["1000123"]["materiales"]["10104146"]["precios"]["0110"][0], 5.91)

    def test_un_solo_proveedor_aunque_el_nombre_no_case(self):
        crear_purchasing(self.ruta, h1=[h1("5300001", "10100001")], h3=[])
        self.assertEqual(PR.registros_info_de(self.ruta, "Suministros Demo, S.L."), {"10100001": "5300001"})

    def test_rellenar_no_pisa_y_salta_sin_sap(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A8"], ws["A9"], ws["A10"], ws["H10"] = 10100001, 10100002, 10100003, 999
        ws["A11"] = None
        n = PR.rellenar_registros_info(ws, {"10100001": "5300001", "10100003": "5300003"})
        self.assertEqual(n, 1)
        self.assertEqual((ws["H8"].value, ws["H9"].value, ws["H10"].value), (5300001, None, 999))

    def test_tarifa_final_de_la_accion_3(self):
        """transformar() con la búsqueda del Purchasing que le pasa la Acción 3, y sin ella (columna vacía)."""
        from tests.test_datos_tarifa import _HojaXls
        from app.tarifas import transformar_tarifa as TT
        celdas = {(2, 4): 4321, (4, 4): "SUMINISTROS DEMO SL", (17, 3): "2027",
                  (24, 1): 5001, (24, 2): "R1", (24, 4): "ACEITE DEMO", (24, 6): "UD", (24, 7): 3.5,
                  (25, 1): 5002, (25, 2): "R2", (25, 4): "VINAGRE DEMO", (25, 6): "UD", (25, 7): 1.25}
        os.makedirs(os.path.join(self.tmp, "TARIFAS"))
        os.makedirs(os.path.join(self.tmp, "RECURSOS"))
        # Catálogo mínimo para que el paso 27b ponga el código SAP
        cat = openpyxl.Workbook()
        cws = cat.active
        from app.comun.sap_helpers import CATALOGO_COLUMNAS, CATALOGO_FILA_CABECERA
        for i, (_k, cab, _a) in enumerate(CATALOGO_COLUMNAS, start=1):
            cws.cell(CATALOGO_FILA_CABECERA, i).value = cab
        claves = [k for k, _c, _a in CATALOGO_COLUMNAS]
        for n, (dali, sap) in enumerate(((5001, 10100001), (5002, 10100002)), start=CATALOGO_FILA_CABECERA + 1):
            cws.cell(n, claves.index("clave") + 1).value = dali
            cws.cell(n, claves.index("cod_sap") + 1).value = sap
            cws.cell(n, claves.index("descripcion") + 1).value = "X"
        cat.save(os.path.join(self.tmp, "RECURSOS", "catalogo-dali-completo.xlsx"))
        pedidos = []

        def buscar(nombre):
            pedidos.append(nombre)
            return {"mapa": {"10100001": "5300001"}, "archivo": "Purchasing Info Records.xlsx", "fecha": None}

        for funcion, esperado in ((buscar, [5300001, None]), (None, [None, None])):
            xls = os.path.join(self.tmp, "TARIFAS", "TARIFA.xls")
            open(xls, "wb").close()
            with mock.patch.object(TT, "_leer_xls", return_value=_HojaXls(celdas)):
                ruta, _ = TT.transformar(xls, registros_info=funcion)
            ws = openpyxl.load_workbook(ruta).active
            self.assertEqual(ws["H7"].value, "REGISTRO INFO")
            self.assertEqual([ws["A8"].value, ws["A9"].value], [10100001, 10100002])
            self.assertEqual([ws["H8"].value, ws["H9"].value], esperado)
            self.assertEqual(ws["G7"].value, "DESCUENTO")              # A..G no se mueven
            os.remove(ruta)
        self.assertEqual(pedidos, ["SUMINISTROS DEMO SL"])


@unittest.skipUnless(HAY_TK, "las ventanas y el RINFO importan tkinter")
class TestBotonYVentana(unittest.TestCase):
    def test_filtro_de_archivos(self):
        from app.ui.ventana_plantilla_rinfo import solo_excel
        self.assertEqual(solo_excel(["/t/a.xlsx", "/t/~$a.xlsx", "/t/b.xls", "/t/c.XLSX", "/t/a.xlsx"]),
                         ["/t/a.xlsx", "/t/c.XLSX"])

    def test_arrastrar_no_rompe_fuera_de_windows(self):
        from app.ui import dnd_windows
        with mock.patch.object(dnd_windows.sys, "platform", "linux"):
            self.assertFalse(dnd_windows.hook_dropfiles(object(), lambda r: None))

    def test_accion_6_de_punta_a_punta(self):
        """El método del botón 6 con la clase real (sin ventana): genera en TARIFAS, dice todo en la consola y no
        abre ninguna ventana si ha ido bien; con un archivo que no es un Purchasing, ventana de error."""
        from app.principal.accion_tarifa import AccionesTarifa
        tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, tmp, True)
        tarifas = os.path.join(tmp, "TARIFAS")
        os.makedirs(tarifas)
        os.makedirs(os.path.join(tmp, "RECURSOS"))
        ruta = crear_purchasing(os.path.join(tarifas, "Purchasing Info Records.xlsx"),
                                h1=[h1("1", "H1", tel="928000000")],
                                h3=[h3("1", "H1", "0110", 3.0, "01.01.2026", "31.12.2099")],
                                h4=[{1: "1", 5: "0110", 6: "H1"}])
        app = object.__new__(AccionesTarifa)
        logs = []
        app.log_consola = lambda m, n=">>": logs.append((n, m))
        app.after = lambda _ms, fn: fn()
        app._barra = mock.Mock()
        with mock.patch("app.principal.accion_tarifa.messagebox") as mb:
            app._generar_plantillas_rinfo([ruta], tmp, tarifas)
            mb.showerror.assert_not_called()
            hechas = [n for n in os.listdir(tarifas) if n.startswith("SUMINISTROS")]
            self.assertEqual(len(hechas), 1)
            self.assertTrue(any(n == "WARN" and "escalado" in m for n, m in logs))
            self.assertTrue(any(n == "OK" and "Plantillas generadas" in m for n, m in logs))
            app._barra.completado.assert_called_once()
            # Otra vez: no pisa, sale «(2)»
            app._generar_plantillas_rinfo([ruta], tmp, tarifas)
            self.assertEqual(len([n for n in os.listdir(tarifas) if n.startswith("SUMINISTROS")]), 2)
            # Un archivo que no es un Purchasing
            otro = os.path.join(tarifas, "otra cosa.xlsx")
            openpyxl.Workbook().save(otro)
            app._generar_plantillas_rinfo([otro], tmp, tarifas)
            mb.showerror.assert_called_once()

    def test_el_rinfo_lee_el_precio_nuevo(self):
        """Creador y Generador leen los precios con datos_tarifa.precios_tarifa (ya no la columna F a mano)."""
        import inspect
        from app.rinfo import creador_rinfo, generador_existente, crear_registro_info
        self.assertIn("DT.precios_tarifa(hoja_t)", inspect.getsource(creador_rinfo))
        self.assertIn("DT.precios_tarifa(hoja_t)", inspect.getsource(generador_existente))
        self.assertEqual(inspect.getsource(crear_registro_info).count("periodo=DT.periodo_propuesto("), 2)


if __name__ == "__main__":
    unittest.main()
