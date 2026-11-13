# -*- coding: utf-8 -*-
"""Pruebas de la acción «Asignar SAP pendientes» (v10.2.50): guardado en el catálogo con duplicados resueltos,
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

    def test_escribir_en_la_tabla_intro_pasa_a_la_siguiente_y_guardar(self):
        """13 nov 2026: la tabla es como la de «Códigos SAP faltantes» de la Acción 4: se escribe en cada fila."""
        v = self.ven
        self.assertEqual(v._filas_visibles(), [2, 3])
        v._vars[2].set("2000")
        self.assertIn("PIMIENTO ROJO", v._estados[2]["text"])
        self.assertEqual(v._asignados, {2: 2000})
        with mock.patch.object(v._celdas_sap[3], "focus_set") as foco:
            self.assertEqual(v._siguiente_fila(2), "break")    # Intro: pasa a la fila siguiente
        foco.assert_called_once()
        v._vars[3].set("9999")
        self.assertIn("no está en Productos.xlsx", v._estados[3]["text"])
        self.assertEqual(v._asignados, {2: 2000, 3: 9999})
        self.assertIn("(2)", v._btn_guardar["text"])
        with mock.patch("app.ui.ventana_asignar_sap.messagebox.showinfo") as info:
            v._guardar()
            fin = time.time() + 5
            while v._asignados and time.time() < fin:
                self._bombear(0.05)
        self.assertEqual(self.guardados, [{2: 2000, 3: 9999}])
        self.assertEqual(v._todos, [])
        self.assertEqual(v._filas_visibles(), [])
        info.assert_called_once()

    def test_tras_guardar_codigos_fuera_de_productos_abre_la_guia(self):
        """13 nov 2026: los códigos que no están en Productos.xlsx llevan al aviso de los listados de SAP."""
        v = self.ven
        avisos = []
        v._listados_fn = avisos.append
        r = {"escritos": {2: 2000, 3: 9999}, "limpiados": [], "no_encontrados": [], "fuera_productos": [3],
             "completando": True}
        v._asignados = {2: 2000, 3: 9999}
        with mock.patch("app.ui.ventana_asignar_sap.messagebox.showinfo") as info:
            v._tras_guardar({2: 2000, 3: 9999}, r, None, False)
        self.assertIn("A continuación se explica cómo descargar", info.call_args[0][1])
        self.assertEqual(avisos, [["9999"]])

    def test_copiar_para_pegar_en_sap(self):
        v = self.ven
        v._copiar_campo("descripcion", "Descripción", dali=3)
        self.assertEqual(v.clipboard_get(), "LECHE")
        v._copiar_campo("dali", "Código DALI", dali=3)
        self.assertEqual(v.clipboard_get(), "000003")
        v._copiar_fila(3)
        self.assertEqual(v.clipboard_get(), "000003\tLECHE\tL\tQ\tLÁCTEOS")
        self.assertIn("Copiado", v._lbl_info["text"])
        # Ctrl+C en una celda sin seleccionar nada copia la celda entera
        celda = [w for w in v._rejilla.grid_slaves(row=2, column=1)][0]
        v._copiar_celda(celda)
        self.assertEqual(v.clipboard_get(), "LECHE")

    def test_filtrar_conserva_lo_escrito_y_codigo_no_valido(self):
        v = self.ven
        v._vars[2].set("2000")
        v._var_buscar.set("leche")
        self.assertEqual(v._filas_visibles(), [3])
        v._quitar_filtros()
        self.assertEqual(v._filas_visibles(), [2, 3])
        self.assertEqual(v._vars[2].get(), "2000")              # lo escrito sigue ahí tras filtrar
        self.assertEqual(v._celdas_sap[2].get(), "2000")
        v._vars[3].set("abc")
        self.assertTrue(v._estados[3]["text"].startswith("✘"))
        self.assertEqual(v._asignados, {2: 2000})               # el no válido no se guarda
        v._vars[3].set("2000")                                   # repetido con el DALI 2
        self.assertIn("repetido con el DALI 000002", v._estados[3]["text"])
        self.assertIn("repetido con el DALI 000003", v._estados[2]["text"])
        v._vars[3].set("")
        self.assertNotIn("repetido", v._estados[2]["text"])


if __name__ == "__main__":
    unittest.main()
