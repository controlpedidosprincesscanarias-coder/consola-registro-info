# -*- coding: utf-8 -*-
"""Pruebas del aviso «Faltan los listados de SAP actualizados» (13 nov 2026): app/comun/listados_sap.py y
app/ui/ventana_listados_sap.py, y su enganche en la app (App._avisar_listados_sap).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import tempfile
import time
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import listados_sap as L  # noqa: E402
from app.comun.tarifa_constants import resource_path  # noqa: E402

try:
    import tkinter as tk
    HAY_TK = True
except ImportError:          # pragma: no cover
    HAY_TK = False


class Logica(unittest.TestCase):
    def test_codigos_sin_datos(self):
        tabla = {"10104081": ("LUBINA", "", "KG", "", "", "")}
        self.assertEqual(L.codigos_sin_datos([10104081, "10104082", " 10104082 ", ""], tabla), ["10104082"])
        self.assertEqual(L.codigos_sin_datos([10104081], tabla), [])
        self.assertEqual(L.codigos_sin_datos([10104081, 5], None), ["10104081", "5"])     # sin Productos.xlsx

    def test_estado_de_los_archivos(self):
        with tempfile.TemporaryDirectory() as rec:
            ruta = os.path.join(rec, L.PRODUCTOS)
            open(ruta, "w").close()
            os.utime(ruta, (1_700_000_000, 1_700_000_000))             # el de siempre
            antes = L.fechas_listados(rec)                              # al salir el aviso
            e = L.estado_archivos(rec, antes)
            self.assertEqual(e, {L.PRODUCTOS: (True, False), L.LISTAS: (False, False)})
            self.assertEqual(L.sin_actualizar(e), [L.PRODUCTOS, L.LISTAS])
            # (v10.2.76) Descargado de nuevo aunque su fecha sea «anterior» a la hora del PC (reloj del servidor de
            # la unidad G: atrasado): cuenta que ha cambiado, no la hora del PC
            os.utime(ruta, (1_700_000_100, 1_700_000_100))
            open(os.path.join(rec, L.LISTAS), "w").close()
            self.assertEqual(L.sin_actualizar(L.estado_archivos(rec, antes)), [])

    def test_guia_completa(self):
        self.assertEqual([x["app_sap"] for x in L.LISTADOS], ["Gestionar datos maestros de producto",
                                                               "Lista de productos"])
        for img in ["assets/guia_sap_apps.png", "assets/guia_sap_barra.png"] + [x["imagen"] for x in L.LISTADOS]:
            self.assertTrue(os.path.isfile(resource_path(img)), img)
        self.assertIn("PARCIALMENTE", L.AVISO_PARCIAL)


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class Ventana(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("sin pantalla")
        self.root.withdraw()
        self.tmp = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmp.cleanup()
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def _ventana(self):
        from app.ui.ventana_listados_sap import VentanaListadosSap
        return VentanaListadosSap(self.root, ["10104082"], self.tmp.name)

    def test_mas_tarde_avisa_de_la_web_parcial(self):
        v = self._ventana()
        self.assertEqual(len(v._imagenes), 4)                       # las cuatro imágenes de la guía
        with mock.patch("app.ui.ventana_listados_sap.messagebox.showwarning") as aviso:
            v._mas_tarde()
        self.assertIn("PARCIALMENTE", aviso.call_args[0][1])
        self.assertEqual(v.resultado, "despues")

    def test_ya_estan_sin_descargar_pregunta(self):
        v = self._ventana()
        with mock.patch("app.ui.ventana_listados_sap.messagebox.askyesno", return_value=False) as pregunta:
            v._ya_estan()
        pregunta.assert_called_once()
        self.assertTrue(v.winfo_exists())                          # dijo que no: sigue la ventana
        for n in (L.PRODUCTOS, L.LISTAS):
            open(os.path.join(self.tmp.name, n), "w").close()
        with mock.patch("app.ui.ventana_listados_sap.messagebox.askyesno") as pregunta:
            v._ya_estan()
        pregunta.assert_not_called()                               # los dos descargados de nuevo: no pregunta
        self.assertEqual(v.resultado, "actualizar")


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class EngancheApp(unittest.TestCase):
    """App._avisar_listados_sap: solo avisa si algún código no está en Productos.xlsx, y con «actualizar» lanza
    la Acción 3."""

    def _app(self, carpeta):
        from app.principal.accion_listado_sap import AccionListadoSap
        app = object.__new__(AccionListadoSap)
        app._carpeta_proyecto = mock.Mock(get=lambda: carpeta)
        app.log_consola = mock.Mock()
        app.wait_window = lambda v: None
        app.after = mock.Mock()
        app._actualizar_listado_sap = mock.Mock()
        return app

    def test_no_avisa_si_estan_en_productos_y_lanza_la_accion_3(self):
        with tempfile.TemporaryDirectory() as carpeta:
            os.makedirs(os.path.join(carpeta, "RECURSOS"))
            app = self._app(carpeta)
            tabla = {"10104081": ("LUBINA", "", "KG", "", "", "")}
            open(os.path.join(carpeta, "RECURSOS", L.PRODUCTOS), "w").close()
            with mock.patch("app.principal.accion_listado_sap.leer_productos_xlsx", return_value=(tabla, [])), \
                    mock.patch("app.ui.ventana_listados_sap.VentanaListadosSap") as V:
                self.assertIsNone(app._avisar_listados_sap([10104081]))
                V.assert_not_called()
                V.return_value.resultado = "actualizar"
                self.assertEqual(app._avisar_listados_sap([10104081, 10104082]), "actualizar")
            self.assertEqual(V.call_args[0][1], ["10104082"])
            app.after.assert_called_once_with(300, app._actualizar_listado_sap)


# ── Antes de la Acción 3 (v10.2.74) ─────────────────────────────────────────────────────────────────────────────────
AHORA = 1_800_000_000.0
DIA = 86400


def _art(cod_sap, desc="X", unidad="KG", grupo="G1", categoria="3100", clave=1):
    return (clave, {"clave": clave, "cod_sap": cod_sap, "desc_sap": desc, "unidad_sap": unidad, "grp_prod": grupo,
                    "categoria": categoria})


class LogicaAccion3(unittest.TestCase):
    def test_articulos_sin_datos_sap(self):
        filas = [_art(10100001), _art(None, desc=None), _art(10100002.0, desc=None),
                 _art("10100003", categoria=""), _art(10100004, unidad=" ", categoria=None)]
        self.assertEqual(L.articulos_sin_datos_sap(filas), {
            "10100002": {"productos": True, "listas": False},
            "10100003": {"productos": False, "listas": True},
            "10100004": {"productos": True, "listas": True},
        })

    def test_fuera_por_listado(self):
        pend = {"1": {"productos": True, "listas": False}, "2": {"productos": False, "listas": True},
                "3": {"productos": True, "listas": True}}
        prod, cat = {"1": (), "3": ()}, {"3": "3100"}
        self.assertEqual(L.codigos_fuera_por_listado(pend, prod, cat), {L.PRODUCTOS: [], L.LISTAS: ["2"]})
        self.assertEqual(L.codigos_fuera_por_listado(pend, None, None), {L.PRODUCTOS: ["1", "3"],
                                                                        L.LISTAS: ["2", "3"]})

    def test_diagnostico_tres_casos(self):
        recientes = {L.PRODUCTOS: AHORA - DIA, L.LISTAS: AHORA - 2 * DIA}
        d = L.diagnostico([_art(10100001)], {}, {}, recientes, AHORA)
        self.assertEqual((d["estado"], d["preguntar"]), (L.NADA, False))           # nada: directo al botón
        filas = [_art(10100001, desc=None)]
        d = L.diagnostico(filas, {"10100001": ()}, {}, recientes, AHORA)
        self.assertEqual((d["estado"], d["pendientes"], d["preguntar"]), (L.EN_RECURSOS, 1, False))
        d = L.diagnostico(filas + [_art(10100009, categoria=None)], {"10100001": ()}, {}, recientes, AHORA)
        self.assertEqual((d["estado"], d["fuera"], d["preguntar"]), (L.FALTAN, ["10100009"], True))
        self.assertEqual(d["fuera_por_listado"], {L.PRODUCTOS: [], L.LISTAS: ["10100009"]})

    def test_antiguedad_7_dias_y_listado_que_no_esta(self):
        self.assertEqual(L.DIAS_ANTIGUEDAD, 7)                                       # decisión de Víctor
        fechas = {L.PRODUCTOS: AHORA - 8 * DIA, L.LISTAS: AHORA - 6 * DIA}
        d = L.diagnostico([_art(10100001)], {}, {}, fechas, AHORA)
        self.assertEqual((d["estado"], d["antiguos"], d["preguntar"]), (L.NADA, [L.PRODUCTOS], True))
        d = L.diagnostico([_art(10100001)], {}, {}, {L.PRODUCTOS: AHORA, L.LISTAS: None}, AHORA)
        self.assertEqual((d["no_estan"], d["antiguos"], d["preguntar"]), ([L.LISTAS], [], True))

    def test_lineas_diagnostico(self):
        fechas = {L.PRODUCTOS: AHORA - 8 * DIA, L.LISTAS: None}
        d = L.diagnostico([_art(10100009, desc=None)], {}, {}, fechas, AHORA)
        lineas = L.lineas_diagnostico(d, fechas, AHORA)
        texto = "\n".join(t for t, _ in lineas)
        self.assertIn("hace 8 días", texto)
        self.assertIn("10100009", texto)
        self.assertIn("más de 7 días", texto)
        self.assertIn("No está en RECURSOS: Listas de productos.xlsx", texto)
        self.assertEqual(lineas[0][1], "aviso")
        self.assertEqual(lineas[1][1], "error")
        d = L.diagnostico([_art(10100001, desc=None)], {"10100001": ()}, {}, {L.PRODUCTOS: AHORA, L.LISTAS: AHORA},
                          AHORA)
        self.assertIn("ya tienen los datos de los 1", "\n".join(t for t, _ in L.lineas_diagnostico(d, fechas, AHORA)))

    def test_estado_descarga_pide_que_el_archivo_tenga_los_codigos(self):
        with tempfile.TemporaryDirectory() as rec:
            ruta_prod = os.path.join(rec, L.PRODUCTOS)
            open(ruta_prod, "w").close()
            os.utime(ruta_prod, (AHORA, AHORA))
            sin_cambios = L.fechas_listados(rec)                       # al contestar «Sí»: Productos, sin Listas
            antes = {L.PRODUCTOS: AHORA - 100, L.LISTAS: None}           # el de antes, con otra fecha
            open(os.path.join(rec, L.LISTAS), "w").close()
            fuera = {L.PRODUCTOS: ["10100009"], L.LISTAS: []}
            e = {n: (listo, t) for n, listo, t in L.estado_descarga(rec, antes, fuera, {})}
            self.assertEqual(e[L.PRODUCTOS], (False, "descargado de nuevo, comprobando…"))
            self.assertEqual(e[L.LISTAS], (True, "descargado de nuevo"))
            e = dict((n, listo) for n, listo, _ in L.estado_descarga(rec, antes, fuera, {L.PRODUCTOS: ["10100009"]}))
            self.assertFalse(e[L.PRODUCTOS])                                         # cambió la fecha, pero no lo tiene
            e = dict((n, listo) for n, listo, _ in L.estado_descarga(rec, antes, fuera, {L.PRODUCTOS: []}))
            self.assertTrue(e[L.PRODUCTOS])
            e = dict((n, t) for n, _l, t in L.estado_descarga(rec, sin_cambios, fuera, {}))
            self.assertEqual(e[L.PRODUCTOS], "sin descargar de nuevo")             # misma fecha que al contestar
            self.assertEqual(e[L.LISTAS], "descargado de nuevo")                   # antes no estaba
            # La hora del PC no cuenta: un archivo con fecha «anterior» (reloj del servidor atrasado) también vale
            self.assertTrue(L.descargado_de_nuevo(AHORA, AHORA - 30))
            self.assertFalse(L.descargado_de_nuevo(AHORA, None))

    def test_codigos_que_faltan_en_lee_los_excel(self):
        import openpyxl
        with tempfile.TemporaryDirectory() as rec:
            wb = openpyxl.Workbook()
            wb.active.append(["A", "Producto"])
            wb.active.append(["x", "10100001"])
            wb.save(os.path.join(rec, L.PRODUCTOS))
            wb = openpyxl.Workbook()
            wb.active.append(["Producto", "Centro", "Categoría valoración"])
            wb.active.append(["QUESO (10100002)", "C1", "VAL (3100)"])
            wb.save(os.path.join(rec, L.LISTAS))
            self.assertEqual(L.codigos_que_faltan_en(L.PRODUCTOS, os.path.join(rec, L.PRODUCTOS),
                                                     ["10100001", "10100002"]), ["10100002"])
            self.assertEqual(L.codigos_que_faltan_en(L.LISTAS, os.path.join(rec, L.LISTAS),
                                                     ["10100001", "10100002"]), ["10100001"])


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class VentanaAccion3(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("sin pantalla")
        self.root.withdraw()
        self.tmp = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmp.cleanup()
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def _ventana(self, diag, leer=None):
        from app.ui.ventana_listados_sap import VentanaAntesAccion3
        return VentanaAntesAccion3(self.root, diag, [("texto", "info")], self.tmp.name, leer=leer)

    def _diag(self, preguntar, fuera_prod=()):
        return {"estado": L.FALTAN if fuera_prod else L.NADA, "pendientes": len(fuera_prod), "fuera": list(fuera_prod),
                "fuera_por_listado": {L.PRODUCTOS: list(fuera_prod), L.LISTAS: []}, "antiguos": [], "no_estan": [],
                "preguntar": preguntar}

    def test_sin_pregunta_el_boton_sale_directamente(self):
        v = self._ventana(self._diag(False))
        self.assertTrue(v._btn_actualizar.winfo_manager())
        self.assertEqual(v._imagenes, [])                                          # sin guía
        v._actualizar()
        self.assertEqual(v.resultado, "actualizar")

    def test_boton_oculto_hasta_contestar_no(self):
        v = self._ventana(self._diag(True, ["10100009"]))
        self.assertFalse(v._btn_actualizar.winfo_manager())
        self.assertEqual(len(v._imagenes), 4)                                      # la guía con sus imágenes
        v._no()
        self.assertTrue(v._btn_actualizar.winfo_manager())
        self.assertEqual(str(v._btn_si["state"]), "disabled")
        with mock.patch("app.ui.ventana_listados_sap.messagebox.askyesno") as pregunta:
            v._actualizar()
        pregunta.assert_not_called()                                               # «No»: sin comprobar descarga
        self.assertEqual((v.resultado, v.respuesta), ("actualizar", False))

    def test_si_abre_recursos_y_comprueba_que_los_traiga(self):
        leidos = []

        def leer(nombre, ruta, fuera):
            leidos.append(nombre)
            return []                                                              # el archivo nuevo ya los tiene

        v = self._ventana(self._diag(True, ["10100009"]), leer=leer)
        with mock.patch.object(v, "_abrir_recursos") as abrir:
            v._si()
        abrir.assert_called_once()
        self.assertTrue(v._btn_actualizar.winfo_manager())
        with mock.patch("app.ui.ventana_listados_sap.messagebox.askyesno", return_value=False) as pregunta:
            v._actualizar()                                                        # aún sin descargar: pregunta
        pregunta.assert_called_once()
        self.assertTrue(v.winfo_exists())
        for n in (L.PRODUCTOS, L.LISTAS):
            open(os.path.join(self.tmp.name, n), "w").close()
        v._comprobar_archivos_nuevos()
        for _ in range(50):                                                        # la lectura va en un hilo
            if L.PRODUCTOS in v._contiene:
                break
            time.sleep(0.02)
        self.assertEqual(leidos, [L.PRODUCTOS])                                    # solo el que tenía que traer
        with mock.patch("app.ui.ventana_listados_sap.messagebox.askyesno") as pregunta:
            v._actualizar()
        pregunta.assert_not_called()
        self.assertEqual((v.resultado, v.respuesta), ("actualizar", True))

    def test_cancelar(self):
        v = self._ventana(self._diag(True))
        v._cancelar()
        self.assertEqual(v.resultado, "cancelar")


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class EngancheAccion3(unittest.TestCase):
    """App._accion_3_guiada / _mostrar_antes_accion3, con archivos de verdad y la ventana sustituida."""

    def _app(self, carpeta):
        from app.principal.accion_listado_sap import AccionListadoSap
        app = object.__new__(AccionListadoSap)
        app._carpeta_proyecto = mock.Mock(get=lambda: carpeta)
        app.log_consola = mock.Mock()
        app.log_consola_batch = mock.Mock()
        app._barra = mock.Mock()
        app._traer_cambios_web = mock.Mock()
        app._actualizar_listado_sap = mock.Mock()
        app.wait_window = lambda v: None
        app.after = lambda _ms, f: f()
        return app

    def _proyecto(self, carpeta, articulos):
        import openpyxl
        from tests.test_catalogo_local import crear_catalogo
        rec = os.path.join(carpeta, "RECURSOS")
        crear_catalogo(os.path.join(rec, "catalogo-dali-completo.xlsx"), articulos)
        wb = openpyxl.Workbook()
        wb.active.append(["", "Producto", "Descripción"])
        wb.active.append(["", "10100001", "LUBINA"])
        wb.save(os.path.join(rec, L.PRODUCTOS))
        wb = openpyxl.Workbook()
        wb.active.append(["Producto", "Centro", "Categoría valoración"])
        wb.active.append(["LUBINA (10100001)", "C1", "VAL (3100)"])
        wb.save(os.path.join(rec, L.LISTAS))
        return rec

    def _lanzar(self, app):
        hilos = []
        real = __import__("threading").Thread

        def hilo(target, daemon):
            t = real(target=target, daemon=daemon)
            hilos.append(t)
            return t
        with mock.patch("app.principal.accion_listado_sap.threading.Thread", side_effect=hilo):
            app._accion_3_guiada()
        for t in hilos:
            t.join(30)

    def test_de_punta_a_punta_faltan_codigos_y_se_actualiza(self):
        with tempfile.TemporaryDirectory() as carpeta:
            self._proyecto(carpeta, [{"clave": 1, "descripcion": "LUBINA", "cod_sap": 10100001},
                                     {"clave": 2, "descripcion": "DORADA", "cod_sap": 10100009}])
            app = self._app(carpeta)
            with mock.patch("app.ui.ventana_listados_sap.VentanaAntesAccion3") as V:
                V.return_value.resultado = "actualizar"
                V.return_value.respuesta = True
                self._lanzar(app)
            app._traer_cambios_web.assert_called_once_with("la Acción 2")
            diag = V.call_args[0][1]
            self.assertEqual((diag["estado"], diag["fuera"], diag["pendientes"]), (L.FALTAN, ["10100009"], 2))
            self.assertTrue(diag["preguntar"])
            app._actualizar_listado_sap.assert_called_once_with(confirmar=False, traer_web=False)
            self.assertFalse(app._comprobando_accion3)

    def test_nada_que_completar_y_cancelar(self):
        with tempfile.TemporaryDirectory() as carpeta:
            self._proyecto(carpeta, [{"clave": 1, "descripcion": "LUBINA", "cod_sap": 10100001, "desc_sap": "L",
                                      "unidad_sap": "KG", "grp_prod": "G", "categoria": "3100"},
                                     {"clave": 2, "descripcion": "SIN SAP"}])
            app = self._app(carpeta)
            with mock.patch("app.ui.ventana_listados_sap.VentanaAntesAccion3") as V:
                V.return_value.resultado = "cancelar"
                self._lanzar(app)
            diag = V.call_args[0][1]
            self.assertEqual((diag["estado"], diag["preguntar"]), (L.NADA, False))
            app._actualizar_listado_sap.assert_not_called()

    def test_si_falla_la_comprobacion_sigue_la_accion_3_de_siempre(self):
        with tempfile.TemporaryDirectory() as carpeta:
            rec = os.path.join(carpeta, "RECURSOS")
            os.makedirs(rec)
            open(os.path.join(rec, "catalogo-dali-completo.xlsx"), "w").close()     # no es un Excel
            app = self._app(carpeta)
            self._lanzar(app)
            app._actualizar_listado_sap.assert_called_once_with(traer_web=False)

    def test_sin_catalogo_avisos_de_siempre(self):
        with tempfile.TemporaryDirectory() as carpeta:
            app = self._app(carpeta)
            app._accion_3_guiada()
            app._actualizar_listado_sap.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
