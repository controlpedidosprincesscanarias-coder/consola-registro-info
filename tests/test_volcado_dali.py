# -*- coding: utf-8 -*-
"""
test_volcado_dali.py — final de la Acción 1: revisar ARTICULOS_DALI.xlsx y actualizar el catálogo y Catálogo
Asignaciones (13 nov 2026), y que ya no haya botón 2.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

- Lógica sin tkinter (app/comun/volcado_dali.py): estado del Excel al pulsar «Actualizar», recuento, fecha y colores.
- La ventana (app/ui/ventana_volcado_dali.py) en pantalla virtual: el botón «Actualizar Catálogo Asignaciones» está
  oculto hasta contestar Sí / No, «Sí» abre el Excel, y con el Excel abierto no se sigue.
- De punta a punta con las clases reales: lo corregido a mano en ARTICULOS_DALI.xlsx llega al catálogo y se lanza el
  envío a la web.
Necesita tkinter (y pantalla para la ventana: xvfb-run).
"""
import os
import sys
import tempfile
import threading
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import volcado_dali as V  # noqa: E402

try:
    import tkinter as tk
    from app.comun import catalogo_local as C
    from app.principal.accion_listado_dali import AccionListadoDali
    from app.principal.accion_organizar_dali import AccionOrganizarDali
    from app.principal.catalogo_dali import CatalogoDali
    from tests.test_catalogo_local import crear_catalogo
    from tests.test_catalogo_unico import _dali, _Var, crear_articulos_dali
    HAY_TK = True
except ImportError:
    HAY_TK = False

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _escribir(ruta, texto="x"):
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(texto)


class TestLogica(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.ruta = os.path.join(self.dir.name, V.ARTICULOS_DALI)

    def tearDown(self):
        self.dir.cleanup()

    def test_estado_no_existe(self):
        self.assertEqual(V.estado_para_actualizar(self.ruta, False), V.NO_EXISTE)

    def test_estado_abierto_en_excel(self):
        _escribir(self.ruta)
        _escribir(os.path.join(self.dir.name, "~$" + V.ARTICULOS_DALI))   # lo que deja Excel al abrirlo
        self.assertEqual(V.estado_para_actualizar(self.ruta, True, V.fecha_modificacion(self.ruta)), V.ABIERTO)
        self.assertEqual(V.estado_para_actualizar(self.ruta, False), V.ABIERTO)

    def test_estado_sin_guardar_y_guardado(self):
        _escribir(self.ruta)
        antes = V.fecha_modificacion(self.ruta)
        self.assertEqual(V.estado_para_actualizar(self.ruta, True, antes), V.SIN_GUARDAR)
        os.utime(self.ruta, (antes + 5, antes + 5))                       # el usuario guarda en Excel
        self.assertEqual(V.estado_para_actualizar(self.ruta, True, antes), V.LISTO)

    def test_sin_revisar_no_mira_la_fecha(self):
        _escribir(self.ruta)
        self.assertEqual(V.estado_para_actualizar(self.ruta, False, V.fecha_modificacion(self.ruta)), V.LISTO)

    @unittest.skipUnless(HAY_TK, "necesita openpyxl")
    def test_contar_y_fecha(self):
        crear_articulos_dali(self.ruta, [_dali(10, "MANZANA"), _dali(20, "PERA")])
        self.assertEqual(V.contar_articulos(self.ruta), 2)
        self.assertEqual(V.fecha_encabezado(self.ruta), "Actualizado: 13/11/2026")

    def test_colores_del_resumen(self):
        lineas = ["✔  Proceso completado correctamente", "", "   Artículos transferidos:  5", "",
                  "✘  2 código(s) faltante(s) entre 000001 y 000009:", "      000003  (1 ausente)", "",
                  "📁  Puede acceder al listado en:", "   G:\\PROYECTO\\RECURSOS", "",
                  "🗑  Archivos eliminados automáticamente de RECURSOS/:", "   • ARTICULOS1.xls"]
        tipos = [t for _x, t in V.lineas_con_tipo(lineas)]
        self.assertEqual(tipos, ["ok", "info", "info", "info", "error", "error", "error", "nota", "nota", "info",
                                 "nota", "nota"])


class TestSinBoton2(unittest.TestCase):
    def test_no_queda_el_boton_ni_el_metodo_de_la_accion_2(self):
        with open(os.path.join(RAIZ, "app/principal/interfaz.py"), encoding="utf-8") as f:
            interfaz = f.read()
        self.assertNotIn('frm_acc, "2"', interfaz)
        for carpeta in ("app", "tests"):
            for raiz, _d, archivos in os.walk(os.path.join(RAIZ, carpeta)):
                for a in archivos:
                    if a.endswith(".py") and a != "test_volcado_dali.py":
                        with open(os.path.join(raiz, a), encoding="utf-8") as f:
                            texto = f.read()
                        self.assertNotIn("_actualizar_listado_dali", texto, a)


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class TestVentana(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("sin pantalla")
        self.root.withdraw()
        self.dir = tempfile.TemporaryDirectory()
        self.ruta = os.path.join(self.dir.name, V.ARTICULOS_DALI)
        _escribir(self.ruta)
        from app.ui.ventana_volcado_dali import VentanaVolcadoDali
        self.abiertos = []
        self.ven = VentanaVolcadoDali(self.root, "DALI — Proceso completado",
                                      [("✔  Proceso completado correctamente", "ok")], self.ruta,
                                      abrir=self.abiertos.append)
        self.ven.update()

    def tearDown(self):
        try:
            self.root.destroy()
        finally:
            self.dir.cleanup()

    def test_boton_oculto_hasta_contestar_y_si_abre_el_excel(self):
        self.assertEqual(self.ven._btn_actualizar.winfo_manager(), "")     # sin colocar: oculto
        self.ven._btn_si.invoke()
        self.ven.update()
        self.assertEqual(self.ven._btn_actualizar.winfo_manager(), "pack")
        self.assertEqual(self.abiertos, [self.ruta])
        self.assertEqual(str(self.ven._btn_si["state"]), "disabled")

    def test_no_muestra_el_boton_sin_abrir_nada(self):
        self.ven._btn_no.invoke()
        self.ven.update()
        self.assertEqual(self.ven._btn_actualizar.winfo_manager(), "pack")
        self.assertEqual(self.abiertos, [])
        self.ven._btn_actualizar.invoke()
        self.assertEqual(self.ven.resultado, "actualizar")

    def test_con_el_excel_abierto_no_sigue(self):
        self.ven._btn_si.invoke()
        _escribir(os.path.join(self.dir.name, "~$" + V.ARTICULOS_DALI))
        with mock.patch("app.ui.ventana_volcado_dali.messagebox.showwarning") as aviso:
            self.ven._btn_actualizar.invoke()
        aviso.assert_called_once()
        self.assertTrue(self.ven.winfo_exists())
        self.assertEqual(self.ven.resultado, "despues")

    def test_guardado_tras_revisar_actualiza_sin_preguntar(self):
        self.ven._btn_si.invoke()
        t = os.path.getmtime(self.ruta) + 5
        os.utime(self.ruta, (t, t))
        with mock.patch("app.ui.ventana_volcado_dali.messagebox.askyesno") as pregunta:
            self.ven._btn_actualizar.invoke()
        pregunta.assert_not_called()
        self.assertEqual(self.ven.resultado, "actualizar")

    def test_sin_guardar_pregunta_y_se_puede_seguir_revisando(self):
        self.ven._btn_si.invoke()
        with mock.patch("app.ui.ventana_volcado_dali.messagebox.askyesno", return_value=False):
            self.ven._btn_actualizar.invoke()
        self.assertTrue(self.ven.winfo_exists())
        with mock.patch("app.ui.ventana_volcado_dali.messagebox.askyesno", return_value=True):
            self.ven._btn_actualizar.invoke()
        self.assertEqual(self.ven.resultado, "actualizar")

    def test_cerrar_sin_actualizar(self):
        with mock.patch("app.ui.ventana_volcado_dali.messagebox.showwarning") as aviso:
            self.ven._btn_cerrar.invoke()
        aviso.assert_called_once()
        self.assertEqual(self.ven.resultado, "despues")


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class TestDePuntaAPunta(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.proyecto = self.dir.name
        os.makedirs(os.path.join(self.proyecto, "RECURSOS"))
        self.cat = C.ruta_catalogo(self.proyecto)
        self.dali = os.path.join(self.proyecto, "RECURSOS", V.ARTICULOS_DALI)
        C._CACHE.clear()
        clase = type("AppPrueba", (AccionOrganizarDali, AccionListadoDali, CatalogoDali), {})
        app = object.__new__(clase)
        app._carpeta_proyecto = _Var(self.proyecto)
        app.logs = []
        app.log_consola = lambda m, n=">>": app.logs.append((n, m))
        app.log_consola_batch = lambda lineas: app.logs.extend((n, m) for m, n in lineas)
        app.after = lambda *a, **k: None
        app._barra = mock.Mock()
        app.consola = mock.Mock()
        app._publicar_catalogo_web = mock.Mock()
        app._catalogo_lock = threading.Lock()
        self.app = app

    def tearDown(self):
        self.dir.cleanup()

    @staticmethod
    def _hilo_en_el_acto(target=None, daemon=None, **_k):
        """En lugar de un hilo, el proceso corre al llamar a start(): así no se recogen objetos de Tk de otras
        pruebas desde otro hilo (Tcl_AsyncDelete)."""
        return mock.Mock(start=target)

    def test_lo_corregido_en_excel_llega_al_catalogo_y_se_envia_a_la_web(self):
        crear_catalogo(self.cat, [{"clave": 10, "descripcion": "MANZANA", "proveedor": "PROV A", "cod_prov": "M-1",
                                   "precio": 1.2, "cod_sap": "1000", "activo": "SI"}])
        crear_articulos_dali(self.dali, [_dali(10, "MANZANA"), _dali(20, "PERA")])
        # El usuario contesta «Sí», corrige una descripción en Excel, guarda y pulsa «Actualizar»
        def _ventana(_parent, _titulo, _lineas, ruta):
            crear_articulos_dali(ruta, [_dali(10, "MANZANA ROJA"), _dali(20, "PERA")])
            return "actualizar"
        with mock.patch("app.ui.ventana_volcado_dali.preguntar", side_effect=_ventana), \
                mock.patch("app.principal.accion_listado_dali.threading.Thread", side_effect=self._hilo_en_el_acto):
            self.app._final_accion1("DALI — Proceso completado", [], self.dali)
        t = {str(k): d for k, d in C.leer_catalogo(self.cat)}
        self.assertEqual(sorted(t), ["10", "20"])
        self.assertEqual(t["10"]["descripcion"], "MANZANA ROJA")
        self.assertEqual((t["10"]["cod_sap"], t["10"]["cod_prov"], t["10"]["precio"]), ("1000", "M-1", 1.2))
        self.app._publicar_catalogo_web.assert_called()

    def test_cerrar_sin_actualizar_no_toca_nada(self):
        crear_catalogo(self.cat, [{"clave": 10, "descripcion": "MANZANA", "activo": "SI"}])
        crear_articulos_dali(self.dali, [_dali(10, "MANZANA"), _dali(20, "PERA")])
        antes = os.path.getmtime(self.cat)
        with mock.patch("app.ui.ventana_volcado_dali.preguntar", return_value="despues"), \
                mock.patch("app.principal.accion_listado_dali.threading.Thread", side_effect=self._hilo_en_el_acto):
            self.app._final_accion1("DALI — Proceso completado", [], self.dali)
        self.assertEqual(os.path.getmtime(self.cat), antes)
        self.app._publicar_catalogo_web.assert_not_called()
        self.assertTrue(any("sin actualizar" in m for _n, m in self.app.logs))


if __name__ == "__main__":
    unittest.main()
