# -*- coding: utf-8 -*-
"""Pruebas de app/comun/catalogo_local.py (v10.2.50): el catálogo completo como único archivo de artículos, con la
forma del antiguo LISTADO para quien lo leía, escritura de códigos SAP por código DALI y el paso único del LISTADO
al catálogo. Con archivos inventados en una carpeta temporal.

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
from app.comun import catalogo_local as C  # noqa: E402
from app.comun.sap_helpers import CATALOGO_COLUMNAS, CATALOGO_FILA_CABECERA  # noqa: E402

TITULO = {k: t for k, t, _a in CATALOGO_COLUMNAS}


def crear_catalogo(ruta, filas, desordenar=False):
    """filas: lista de dicts con claves locales (clave, descripcion, cod_sap…)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    claves = [k for k, _t, _a in CATALOGO_COLUMNAS]
    if desordenar:
        claves = list(reversed(claves))
    for i, k in enumerate(claves, start=1):
        ws.cell(row=CATALOGO_FILA_CABECERA, column=i, value=TITULO[k])
    for n, f in enumerate(filas, start=CATALOGO_FILA_CABECERA + 1):
        for i, k in enumerate(claves, start=1):
            if f.get(k) is not None:
                ws.cell(row=n, column=i, value=f[k])
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    wb.save(ruta)


def crear_listado(ruta, filas):
    """filas: tuplas A..O desde la fila 8 (como el LISTADO real)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A7"] = "CODIGO DALI"
    for n, f in enumerate(filas, start=8):
        for i, v in enumerate(f, start=1):
            ws.cell(row=n, column=i, value=v)
    wb.save(ruta)


class Base(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.proyecto = self.dir.name
        self.cat = C.ruta_catalogo(self.proyecto)
        C._CACHE.clear()

    def tearDown(self):
        self.dir.cleanup()


class TestLeer(Base):
    def test_forma_de_listado_y_tipos(self):
        crear_catalogo(self.cat, [
            {"clave": 101, "descripcion": "TOMATE", "unidad": "KG", "cod_sap": "10100416", "desc_sap": "TOMATE PERA",
             "texto_um": "Kilogramo", "unidad_sap": "KG", "grp_prod": "G1", "desc_grp_prod": "VERDURA",
             "ind_imp": "I7", "categoria": "C1", "tipo_prod": "T", "grp_compras": "P01", "desc_grp_compras": "ALIM",
             "activo": "SI", "proveedor": "PROV A", "precio": 1.5},
            {"clave": "102", "descripcion": "PERA", "activo": "NO"},
        ], desordenar=True)
        filas = C.filas_formato_listado(self.cat)
        self.assertEqual(filas[0], (101, "TOMATE", "KG", 10100416, "TOMATE PERA", "Kilogramo", "KG", "G1", "VERDURA",
                                    "I7", "C1", "T", "P01", "ALIM", "SI"))
        self.assertEqual(filas[1][0], 102)
        self.assertIsNone(filas[1][3])
        self.assertEqual(filas[1][14], "NO")

    def test_libro_en_memoria_se_lee_como_el_listado(self):
        crear_catalogo(self.cat, [{"clave": 5, "cod_sap": 7, "desc_sap": "X", "unidad_sap": "UN"}])
        wb = C.abrir_como_listado(self.cat)
        filas = list(wb.active.iter_rows(min_row=8, values_only=True))
        wb.close()
        self.assertEqual(filas[0][0], 5)
        self.assertEqual(filas[0][3], 7)
        self.assertEqual(filas[0][6], "UN")
        self.assertEqual(list(wb.active.iter_rows(min_row=8, min_col=1, max_col=4, values_only=True))[0], (5, None, None, 7))

    def test_tabla_sap_y_localizar(self):
        crear_catalogo(self.cat, [{"clave": 5, "cod_sap": 7, "desc_sap": "X"}, {"clave": 6}])
        tabla = C.leer_tabla_sap(self.cat)
        self.assertEqual(tabla["5"]["cod_sap"], "7")
        self.assertIsNone(tabla["6"]["cod_sap"])
        self.assertEqual(C.localizar_catalogo(self.proyecto), self.cat)
        tarifas = os.path.join(self.proyecto, "TARIFAS")
        os.makedirs(tarifas)
        self.assertEqual(C.localizar_catalogo(carpeta_tarifas=tarifas), self.cat)
        self.assertIsNone(C.localizar_catalogo(os.path.join(self.proyecto, "otro")))

    def test_sin_columna_codigo_dali(self):
        wb = openpyxl.Workbook()
        wb.active.cell(row=4, column=1, value="Otra cosa")
        os.makedirs(os.path.dirname(self.cat))
        wb.save(self.cat)
        with self.assertRaises(ValueError):
            C.leer_catalogo(self.cat)


class TestEscribir(Base):
    def test_escribe_por_codigo_dali(self):
        crear_catalogo(self.cat, [{"clave": 1, "cod_sap": 10}, {"clave": 2}, {"clave": 3, "cod_sap": "30"}])
        escritos, no = C.escribir_codigos_sap(self.cat, {2: 20, 3: None, 9: 90})
        self.assertEqual((escritos, no), (2, [9]))
        tabla = C.leer_tabla_sap(self.cat)
        self.assertEqual((tabla["1"]["cod_sap"], tabla["2"]["cod_sap"], tabla["3"]["cod_sap"]), ("10", "20", None))

    def test_mapas_para_duplicados(self):
        crear_catalogo(self.cat, [{"clave": 1, "cod_sap": "10"}, {"clave": 2}])
        wb, ws, mapa = C.abrir_para_editar(self.cat)
        fila_de, sap_de = C.mapas_dali_sap(ws, mapa)
        wb.close()
        self.assertEqual(fila_de, {1: 5, 2: 6})
        self.assertEqual(sap_de, {10: (1, 5)})


class TestDescarga(Base):
    def test_conserva_sap_local_si_la_web_no_tiene(self):
        crear_catalogo(self.cat, [{"clave": 1, "cod_sap": "10", "desc_sap": "LOCAL", "unidad_sap": "KG"},
                                  {"clave": 2, "cod_sap": "20"}, {"clave": 3}])
        local = C.leer_tabla_sap(self.cat)
        nuevo = os.path.join(self.proyecto, "descargado.xlsx")
        crear_catalogo(nuevo, [{"clave": 1}, {"clave": 2, "cod_sap": "22", "desc_sap": "WEB"}, {"clave": 3},
                               {"clave": 4}])
        self.assertEqual(C.conservar_sap_local(nuevo, local), ["1"])
        tabla = C.leer_tabla_sap(nuevo)
        self.assertEqual((tabla["1"]["cod_sap"], tabla["1"]["desc_sap"], tabla["1"]["unidad_sap"]), ("10", "LOCAL", "KG"))
        self.assertEqual(tabla["2"]["cod_sap"], "22")      # los dos tienen: manda la web
        self.assertIsNone(tabla["3"]["cod_sap"])


class TestMigrar(Base):
    def test_sin_listado_no_hace_nada(self):
        crear_catalogo(self.cat, [{"clave": 1}])
        self.assertIsNone(C.migrar_listado(self.proyecto))

    def test_el_listado_manda_informe_y_renombra(self):
        crear_catalogo(self.cat, [
            {"clave": 1, "cod_sap": "10", "desc_sap": "A", "proveedor": "P", "precio": 2.0},
            {"clave": 2, "cod_sap": "20", "desc_sap": "B"},
            {"clave": 3},
        ])
        lst = C.ruta_listado(self.proyecto)
        crear_listado(lst, [
            (1, "UNO", "KG", 10, "A", None, None, None, None, None, None, None, None, None, "SI"),
            (2, "DOS", "KG", 21, "B NUEVA", None, None, None, None, None, None, None, None, None, "SI"),
            (3, "TRES", "KG", 30, None, None, None, None, None, None, None, None, None, None, "SI"),
            (4, "CUATRO", "KG", 40, None, None, None, None, None, None, None, None, None, None, "SI"),
        ])
        r = C.migrar_listado(self.proyecto, ahora=datetime(2026, 11, 13, 10, 0, 0))
        self.assertEqual(r["articulos_cambiados"], 2)
        self.assertEqual(r["solo_listado"], ["4"])
        self.assertFalse(os.path.exists(lst))
        self.assertTrue(os.path.isfile(os.path.join(self.proyecto, C.NOMBRE_LISTADO_ANTIGUO)))
        with open(r["informe"], encoding="utf-8") as f:
            texto = f.read()
        self.assertIn("B NUEVA", texto)
        tabla = C.leer_tabla_sap(self.cat)
        self.assertEqual((tabla["1"]["cod_sap"], tabla["2"]["cod_sap"], tabla["3"]["cod_sap"]), ("10", "21", "30"))
        self.assertEqual(tabla["2"]["desc_sap"], "B NUEVA")
        filas = {f[0]: f for f in C.filas_formato_listado(self.cat)}
        self.assertEqual(len(filas), 3)                    # no se añaden artículos
        datos = dict(C.leer_catalogo(self.cat))
        self.assertEqual((datos[1]["proveedor"], datos[1]["precio"]), ("P", 2.0))   # lo demás no se toca
        # una segunda vez ya no hay LISTADO
        self.assertIsNone(C.migrar_listado(self.proyecto))

    def test_iguales_sin_informe(self):
        crear_catalogo(self.cat, [{"clave": 1, "cod_sap": "10", "desc_sap": "A"}])
        crear_listado(C.ruta_listado(self.proyecto),
                      [(1, "UNO", "KG", 10.0, "A", None, None, None, None, None, None, None, None, None, "SI")])
        r = C.migrar_listado(self.proyecto)
        self.assertEqual(r["diferencias"], [])
        self.assertIsNone(r["informe"])


if __name__ == "__main__":
    unittest.main()
