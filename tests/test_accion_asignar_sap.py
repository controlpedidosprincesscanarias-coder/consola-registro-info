# -*- coding: utf-8 -*-
"""Pruebas de la acción «Asignar SAP pendientes» (v10.2.56): guardado en el catálogo con duplicados resueltos,
completar los datos SAP (cruce de la Acción 3) y la ventana (app/ui/ventana_asignar_sap.py). Con un proyecto
inventado y las clases reales; la ventana necesita tkinter y pantalla (xvfb-run).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import time
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    import tkinter as tk
    from app.comun import asignar_sap as A
    from app.comun import catalogo_local as C
    from app.principal.accion_asignar_sap import AccionAsignarSap
    from app.principal.accion_listado_sap import AccionListadoSap
    from app.principal.catalogo_dali import CatalogoDali
    from tests.test_catalogo_local import crear_catalogo
    from tests.test_catalogo_unico import Base, crear_articulos_dali, _dali
    HAY_TK = True
except ImportError:
    HAY_TK = False
    Base = unittest.TestCase


def crear_productos(ruta, filas):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["x", "Producto", "Descripción", "Tipo", "e", "Grupo", "Desc grupo", "h", "UM", "Texto UM"])
    for sap, desc, um in filas:
        ws.append([None, str(sap), desc, "ZMAT", None, "G01", "FRUTA", None, um, "Kilogramo"])
    wb.save(ruta)


def crear_listas(ruta, filas):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Producto", "b", "c", "d", "e", "Categoría valoración"])
    for sap, cat in filas:
        ws.append([str(sap), None, None, None, None, cat])
    wb.save(ruta)


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class TestGuardar(Base):
    def setUp(self):
        super().setUp()
        clase = type("AppAsignar", (AccionAsignarSap, AccionListadoSap, CatalogoDali), {})
        viejo = self.app
        self.app = object.__new__(clase)
        self.app.__dict__.update(viejo.__dict__)
        self.app._tras_cambiar_sap_en_catalogo = mock.Mock()
        crear_catalogo(self.cat, [
            {"clave": 1, "descripcion": "TOMATE", "activo": "SI", "cod_sap": "1000", "proveedor": "P"},
            {"clave": 2, "descripcion": "PIMIENTO", "activo": "SI", "proveedor": "P"},
            {"clave": 3, "descripcion": "LECHE", "activo": "SI", "proveedor": "P"},
        ])
        self.lista = A.pendientes(self.cat)

    def test_guarda_resuelve_duplicados_y_sin_archivos_solo_envia(self):
        resolver = mock.Mock(return_value={1})          # el 1000 pasa del DALI 1 al 2
        r = self.app._guardar_sap_asignados(self.cat, self.lista, {"1000": ("TOMATE", "KG", "KG", "", "", "")},
                                            {2: 1000, 3: 9999}, resolver)
        resolver.assert_called_once()
        self.assertEqual(r["escritos"], {2: 1000, 3: 9999})
        self.assertEqual(r["limpiados"], [1])
        self.assertEqual(r["fuera_productos"], [3])
        self.assertFalse(r["completando"])               # faltan Productos/Listas/ARTICULOS_DALI en RECURSOS
        self.app._tras_cambiar_sap_en_catalogo.assert_called_once()
        t = C.leer_tabla_sap(self.cat)
        self.assertEqual((t["1"]["cod_sap"], t["2"]["cod_sap"], t["3"]["cod_sap"]), (None, "1000", "9999"))
        self.assertFalse(self.app._catalogo_lock.locked())

    def test_cancelar_duplicados_no_guarda_nada(self):
        r = self.app._guardar_sap_asignados(self.cat, self.lista, None, {2: 1000}, lambda c: None)
        self.assertIsNone(r)
        self.assertIsNone(C.leer_tabla_sap(self.cat)["2"]["cod_sap"])
        self.assertFalse(self.app._catalogo_lock.locked())

    def test_completa_los_datos_sap_con_el_cruce_de_la_accion_3(self):
        rec = os.path.join(self.proyecto, "RECURSOS")
        crear_productos(os.path.join(rec, "Productos.xlsx"), [(2000, "PIMIENTO ROJO", "KG")])
        crear_listas(os.path.join(rec, "Listas de productos.xlsx"), [(2000, "C9")])
        crear_articulos_dali(self.dali, [_dali(1, "TOMATE", "P"), _dali(2, "PIMIENTO", "P"), _dali(3, "LECHE", "P")])
        productos = __import__("app.comun.sap_helpers", fromlist=["x"]).leer_productos_xlsx(
            os.path.join(rec, "Productos.xlsx"))[0]
        r = self.app._guardar_sap_asignados(self.cat, self.lista, productos, {2: 2000}, mock.Mock())
        self.assertTrue(r["completando"])
        fin = time.time() + 20
        while time.time() < fin:
            t = C.leer_tabla_sap(self.cat)
            if t["2"].get("categoria"):
                break
            time.sleep(0.1)
        t = C.leer_tabla_sap(self.cat)
        self.assertEqual((t["2"]["cod_sap"], t["2"]["desc_sap"], t["2"]["unidad_sap"], t["2"]["categoria"]),
                         ("2000", "PIMIENTO ROJO", "KG", "C9"))
        self.assertEqual(t["1"]["cod_sap"], "1000")


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class TestVentana(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("sin pantalla")
        self.root.withdraw()
        from app.ui.ventana_asignar_sap import VentanaAsignarSap
        self.lista = [{"dali": 2, "descripcion": "PIMIENTO", "unidad": "KG", "proveedor": "P", "naturaleza": "",
                       "familia": "VERDURAS", "subfamilia": ""},
                      {"dali": 3, "descripcion": "LECHE", "unidad": "L", "proveedor": "Q", "naturaleza": "",
                       "familia": "LÁCTEOS", "subfamilia": ""}]
        self.guardados = []

        def guardar(nuevos, resolver):
            self.guardados.append(dict(nuevos))
            return {"escritos": dict(nuevos), "limpiados": [], "no_encontrados": [], "fuera_productos": [],
                    "completando": False}
        self.ven = VentanaAsignarSap(self.root, self.lista, {"2000": ("PIMIENTO ROJO", "", "KG", "", "", "")}, guardar)

    def tearDown(self):
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def _bombear(self, segundos=0.5):
        fin = time.time() + segundos
        while time.time() < fin:
            self.root.update()
            time.sleep(0.02)

    def test_asignar_con_intro_pasa_al_siguiente_y_guardar(self):
        v = self.ven
        v._tabla.selection_set("2")
        v._al_elegir()
        v._var_sap.set("2000")
        self.assertIn("PIMIENTO ROJO", v._lbl_productos["text"])
        v._asignar()
        self.assertEqual(v._asignados, {2: 2000})
        self.assertEqual(v._dali_elegido(), 3)           # pasa al siguiente
        v._var_sap.set("9999")
        self.assertIn("no está en Productos.xlsx", v._lbl_productos["text"])
        v._asignar()
        self.assertEqual(v._asignados, {2: 2000, 3: 9999})
        with mock.patch("app.ui.ventana_asignar_sap.messagebox.showinfo") as info:
            v._guardar()
            fin = time.time() + 5
            while v._asignados and time.time() < fin:
                self._bombear(0.05)
        self.assertEqual(self.guardados, [{2: 2000, 3: 9999}])
        self.assertEqual(v._todos, [])
        info.assert_called_once()

    def test_filtrar_y_codigo_no_valido(self):
        v = self.ven
        v._var_buscar.set("leche")
        self.assertEqual(v._tabla.get_children(), ("3",))
        v._quitar_filtros()
        self.assertEqual(len(v._tabla.get_children()), 2)
        v._tabla.selection_set("2")
        v._var_sap.set("abc")
        with mock.patch("app.ui.ventana_asignar_sap.messagebox.showwarning") as aviso:
            v._asignar()
        aviso.assert_called_once()
        self.assertEqual(v._asignados, {})


if __name__ == "__main__":
    unittest.main()
