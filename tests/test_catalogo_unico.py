# -*- coding: utf-8 -*-
"""Pruebas de punta a punta del catálogo único (v10.2.50) con un proyecto inventado: regeneración del catálogo
sin LISTADO (por código DALI, también cuando DALI quita o intercala códigos), paso único del LISTADO al catálogo,
Acción 3 sobre el libro en memoria y la ventana de SAP faltantes escribiendo en el catálogo. Usan las clases
reales de la app (object.__new__), sin red. Necesitan tkinter (y pantalla para la ventana: xvfb-run).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import tempfile
import threading
import time
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    import tkinter as tk
    import openpyxl
    from app.comun import catalogo_local as C
    from app.principal.catalogo_dali import CatalogoDali
    from app.principal.accion_listado_sap import AccionListadoSap
    from tests.test_catalogo_local import crear_catalogo, crear_listado
    HAY_TK = True
except ImportError:
    HAY_TK = False


class _Var:
    def __init__(self, v):
        self.v = v

    def get(self):
        return self.v


def crear_articulos_dali(ruta, filas):
    """filas: (codigo, descripcion, unidad, familia, subfamilia, naturaleza, igic, proveedor, activo) desde la fila 8."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["B3"] = "Actualizado: 13/11/2026"
    for n, f in enumerate(filas, start=8):
        for i, v in enumerate(f, start=1):
            ws.cell(row=n, column=i, value=v)
    wb.save(ruta)


def _dali(codigo, desc, proveedor="PROV A", activo="SI", subfamilia="VERDURA", igic="7%"):
    return (codigo, desc, "KG", "FRUTAS", subfamilia, "ALIMENTACION", igic, proveedor, activo)


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class Base(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.proyecto = self.dir.name
        os.makedirs(os.path.join(self.proyecto, "RECURSOS"))
        self.cat = C.ruta_catalogo(self.proyecto)
        self.dali = os.path.join(self.proyecto, "RECURSOS", "ARTICULOS_DALI.xlsx")
        C._CACHE.clear()
        clase = type("AppPrueba", (AccionListadoSap, CatalogoDali), {})
        app = object.__new__(clase)
        app._carpeta_proyecto = _Var(self.proyecto)
        app.logs = []
        app.log_consola = lambda m, n=">>": app.logs.append((n, m))
        app.log_consola_batch = lambda lineas: app.logs.extend((n, m) for m, n in lineas)
        app.after = lambda *a, **k: None          # sin ventanas ni barra en las pruebas
        app._barra = mock.Mock()
        app.consola = mock.Mock()
        app._publicar_catalogo_web = mock.Mock()
        app._catalogo_lock = threading.Lock()
        self.app = app

    def tearDown(self):
        self.dir.cleanup()

    def tabla(self):
        return {str(k): d for k, d in C.leer_catalogo(self.cat)}


class TestRegenerar(Base):
    def test_sap_por_codigo_aunque_dali_quite_e_intercale_codigos(self):
        crear_catalogo(self.cat, [
            {"clave": 10, "descripcion": "MANZANA", "proveedor": "PROV A", "cod_prov": "M-1", "precio": 1.2,
             "cod_sap": "1000", "desc_sap": "MANZANA SAP", "activo": "SI"},
            {"clave": 20, "descripcion": "PERA", "proveedor": "PROV A", "cod_sap": "2000", "activo": "SI"},
            {"clave": 30, "descripcion": "UVA", "proveedor": "PROV A", "cod_sap": "3000", "activo": "SI"},
        ])
        # DALI quita el 20 e intercala el 15: con el LISTADO por posición, el 3000 habría acabado en otro artículo
        crear_articulos_dali(self.dali, [_dali(10, "MANZANA"), _dali(15, "KIWI"), _dali(30, "UVA")])
        self.app._regenerar_catalogo_dali_completo()
        t = self.tabla()
        self.assertEqual(sorted(t), ["10", "15", "30"])
        self.assertEqual((t["10"]["cod_sap"], t["10"]["desc_sap"], t["10"]["cod_prov"], t["10"]["precio"]),
                         ("1000", "MANZANA SAP", "M-1", 1.2))
        self.assertIsNone(t["15"]["cod_sap"])
        self.assertEqual(t["30"]["cod_sap"], "3000")
        self.app._publicar_catalogo_web.assert_called()

    def test_con_listado_lo_pasa_una_vez_y_lo_renombra(self):
        crear_catalogo(self.cat, [{"clave": 10, "descripcion": "MANZANA", "cod_sap": "1000"}])
        crear_listado(C.ruta_listado(self.proyecto),
                      [(10, "MANZANA", "KG", 1001, "MANZANA NUEVA", None, None, None, None, None, None, None, None,
                        None, "SI")])
        crear_articulos_dali(self.dali, [_dali(10, "MANZANA")])
        self.app._regenerar_catalogo_dali_completo()
        self.assertEqual(self.tabla()["10"]["cod_sap"], "1001")
        self.assertFalse(os.path.exists(C.ruta_listado(self.proyecto)))
        self.assertTrue(os.path.exists(os.path.join(self.proyecto, C.NOMBRE_LISTADO_ANTIGUO)))
        self.assertTrue(any("Detalle en RECURSOS/paso_listado_a_catalogo_" in m for _n, m in self.app.logs))

    def test_primera_vez_sin_catalogo_sale_del_listado(self):
        crear_listado(C.ruta_listado(self.proyecto),
                      [(10, "MANZANA", "KG", 1000, "MANZANA SAP", None, "KG", None, None, None, None, None, None,
                        None, "SI")])
        crear_articulos_dali(self.dali, [_dali(10, "MANZANA")])
        self.app._regenerar_catalogo_dali_completo()
        t = self.tabla()
        self.assertEqual((t["10"]["cod_sap"], t["10"]["unidad_sap"]), ("1000", "KG"))
        self.assertFalse(os.path.exists(C.ruta_listado(self.proyecto)))


class TestAccion3(Base):
    def test_cruza_productos_sobre_el_catalogo(self):
        crear_catalogo(self.cat, [
            {"clave": 10, "descripcion": "MANZANA", "proveedor": "PROV A", "cod_prov": "M-1", "precio": 1.2,
             "cod_sap": "1000", "desc_sap": "VIEJA", "activo": "SI"},
            {"clave": 20, "descripcion": "PERA", "proveedor": "PROV A", "activo": "SI"},
        ])
        crear_articulos_dali(self.dali, [_dali(10, "MANZANA"), _dali(20, "PERA")])
        tabla_productos = {"1000": ("MANZANA GOLDEN", "Kilogramo", "KG", "G01", "FRUTA", "ZMAT")}
        tabla_dali = {"10": ("VERDURA", "7%"), "20": ("VERDURA", "7%")}
        self.app._cruzar_y_escribir_listado_sap(self.cat, tabla_productos, tabla_dali, {"1000": "C1"}, [], 1,
                                                "13/11/2026", "10:00")
        t = self.tabla()
        self.assertEqual((t["10"]["desc_sap"], t["10"]["texto_um"], t["10"]["unidad_sap"], t["10"]["grp_prod"],
                          t["10"]["tipo_prod"], t["10"]["categoria"]),
                         ("MANZANA GOLDEN", "Kilogramo", "KG", "G01", "ZMAT", "C1"))
        self.assertEqual((t["10"]["cod_prov"], t["10"]["precio"]), ("M-1", 1.2))   # lo de la tarifa no se toca
        self.assertIsNone(t["20"]["cod_sap"])
        self.assertTrue(any("Cruce terminado" in m for _n, m in self.app.logs))


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class TestVentanaSapFaltantes(Base):
    def test_guarda_en_el_catalogo_y_avisa_para_enviar(self):
        try:
            root = tk.Tk()
        except tk.TclError:
            self.skipTest("sin pantalla")
        root.withdraw()
        root._catalogo_lock = threading.Lock()
        from app.ui.tarifa_dialogs import VentanaSAPFaltantes
        crear_catalogo(self.cat, [{"clave": 10, "cod_sap": "1000"}, {"clave": 20}])
        cambiados = []
        ven = VentanaSAPFaltantes(root, [(8, 20)], self.cat, consola_fn=lambda m, n="INFO": None,
                                  descripciones={20: "PERA"}, catalogo_refresh_fn=cambiados.extend)
        with mock.patch("tkinter.messagebox.showinfo"):
            ven._entradas[8][1].set("2000")
            ven._guardar()
            # Se espera a que la ventana se cierre sola (_fin_ok), como en la app: así el hilo de guardado ha
            # terminado de hablar con Tk antes de destruir la raíz.
            fin = time.time() + 10
            while ven.winfo_exists() and time.time() < fin:
                root.update()
                time.sleep(0.02)
            for _ in range(5):
                root.update()
                time.sleep(0.02)
        self.assertTrue(ven.guardado)
        self.assertEqual(self.tabla()["20"]["cod_sap"], 2000)
        self.assertEqual(cambiados, [20])
        self.assertFalse(root._catalogo_lock.locked())
        root.destroy()


if __name__ == "__main__":
    unittest.main()
