# -*- coding: utf-8 -*-
"""
test_fusion_web.py — lo cambiado en Catálogo Asignaciones baja al catálogo Excel y la Consola no lo pisa; en cada
dato gana la última modificación (13 nov 2026, petición de Víctor).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

- catalogo_web.fusionar / acordar / Cache.aplicar_pagina (sin tkinter).
- catalogo_local.escribir_campos (solo las celdas que bajan).
- De punta a punta con PublicarCatalogoWeb y una web falsa: antes de enviar y en _traer_cambios_web.
"""
import os
import sys
import tempfile
import threading
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import catalogo_web as W  # noqa: E402
from tests.test_catalogo_web import WebFalsa, _art  # noqa: E402

try:
    import tkinter  # noqa: F401
    from app.comun import catalogo_local as C
    from app.principal import publicar_catalogo_web as P
    from app.ui import ventana_publicar_catalogo as V
    from tests.test_publicar_catalogo_web import _Var, _excel
    HAY_TK = True
except ImportError:
    HAY_TK = False

VIEJO = "2026-11-01T08:00:00.000Z"
MEDIO = "2026-11-10T08:00:00.000Z"
NUEVO = "2026-11-12T11:00:00.000Z"   # después de la «hasta» de WebFalsa (10:00)


def _cache(web, base=None):
    c = W.Cache(os.devnull)
    c.articulos = {cod: {k: W.normalizar(k, a.get(k)) for k in set().union(*W.GRUPOS.values())}
                   | {"modificado_en": a.get("modificado_en")} for cod, a in web.items()}
    c.base = base or {}
    return c


def _loc(**campos):
    return {k: None for k in W.CAMPOS_COMPARTIDOS} | campos


class TestFusionar(unittest.TestCase):
    def test_solo_cambio_la_web_baja(self):
        c = _cache({1: _art(1, codigo_sap="200", modificado_en=NUEVO)}, base={1: {"codigo_sap": "100"}})
        f = W.fusionar({1: _loc(codigo_sap="100")}, c, hora_excel=NUEVO)
        self.assertEqual(f["bajar"], {1: {"codigo_sap": "200"}})
        self.assertEqual(f["conflictos"], [])

    def test_solo_cambio_el_excel_sube(self):
        c = _cache({1: _art(1, codigo_sap="100", modificado_en=NUEVO)}, base={1: {"codigo_sap": "100"}})
        f = W.fusionar({1: _loc(codigo_sap="300")}, c, hora_excel=VIEJO, ahora=MEDIO)
        self.assertEqual(f["bajar"], {})
        self.assertEqual(c.local_cambiado_en[1]["codigo_sap"], ["300", MEDIO])

    def test_cambiaron_los_dos_gana_el_mas_reciente(self):
        c = _cache({1: _art(1, codigo_sap="200", modificado_en=NUEVO)}, base={1: {"codigo_sap": "100"}})
        c.local_cambiado_en = {1: {"codigo_sap": ["300", MEDIO]}}            # la Consola cambió antes
        f = W.fusionar({1: _loc(codigo_sap="300")}, c, hora_excel=MEDIO)
        self.assertEqual(f["bajar"], {1: {"codigo_sap": "200"}})
        self.assertEqual(f["conflictos"], [(1, "codigo_sap", "300", "200", "web")])
        c.local_cambiado_en = {1: {"codigo_sap": ["300", "2026-11-13T08:00:00.000Z"]}}   # y ahora después
        f = W.fusionar({1: _loc(codigo_sap="300")}, c, hora_excel=MEDIO)
        self.assertEqual(f["bajar"], {})
        self.assertEqual(f["conflictos"][0][4], "excel")

    def test_hora_de_la_web_por_campo_si_se_vio_cambiar(self):
        c = _cache({1: _art(1, codigo_proveedor="P-2", modificado_en=VIEJO)}, base={1: {"codigo_proveedor": "P-1"}})
        c.web_cambiado_en = {1: {"codigo_proveedor": NUEVO}}     # el artículo es viejo, pero su código se cambió hoy
        c.local_cambiado_en = {1: {"codigo_proveedor": ["P-3", MEDIO]}}
        f = W.fusionar({1: _loc(codigo_proveedor="P-3")}, c, hora_excel=MEDIO)
        self.assertEqual(f["bajar"], {1: {"codigo_proveedor": "P-2"}})

    def test_sin_base_gana_la_hora_del_archivo_o_la_web_si_el_excel_esta_vacio(self):
        c = _cache({1: _art(1, codigo_sap="200", precio=2.0, modificado_en=MEDIO)})
        f = W.fusionar({1: _loc(codigo_sap="100", precio=None)}, c, hora_excel=NUEVO)
        self.assertEqual(f["bajar"], {1: {"precio": 2.0}})               # el Excel es más reciente, salvo el vacío
        f = W.fusionar({1: _loc(codigo_sap="100", precio=None)}, c, hora_excel=VIEJO)
        self.assertEqual(f["bajar"], {1: {"codigo_sap": "200", "precio": 2.0}})

    def test_precio_vacio_en_excel_no_borra_la_web(self):
        c = _cache({1: _art(1, precio=4.5, modificado_en=VIEJO)}, base={1: {"precio": 4.5}})
        f = W.fusionar({1: _loc(precio=None)}, c, hora_excel=NUEVO)
        self.assertEqual(f["bajar"], {1: {"precio": 4.5}})

    def test_datos_dali_no_se_tocan_y_articulos_nuevos_no_bajan(self):
        c = _cache({1: _art(1, nombre_articulo="WEB", modificado_en=NUEVO)})
        f = W.fusionar({1: _loc() | {"nombre_articulo": "CONSOLA"}, 2: _loc(codigo_sap="9")}, c, hora_excel=VIEJO)
        self.assertEqual(f["bajar"], {})

    def test_acordar_guarda_la_base_y_olvida_las_horas(self):
        c = _cache({1: _art(1, codigo_sap="200", precio=1.0)})
        c.web_cambiado_en = {1: {"codigo_sap": NUEVO}}
        c.local_cambiado_en = {1: {"precio": [2.0, NUEVO]}}
        W.acordar(c, {1: _loc(codigo_sap="200", precio=2.0)})
        self.assertEqual(c.base[1]["codigo_sap"], "200")
        self.assertNotIn("precio", c.base[1])                            # siguen distintos: sin base nueva
        self.assertNotIn(1, c.web_cambiado_en)
        self.assertEqual(c.local_cambiado_en, {1: {"precio": [2.0, NUEVO]}})

    def test_aplicar_pagina_apunta_cuando_cambia_en_la_web(self):
        c = W.Cache(os.devnull)
        c.aplicar_pagina([_art(1, codigo_sap="100", codigo_proveedor="A")], hasta=MEDIO)
        self.assertEqual(c.web_cambiado_en, {})                          # primera vez: no hay cambio que apuntar
        c.aplicar_pagina([_art(1, codigo_sap="200", codigo_proveedor="B", modificado_en=NUEVO)], hasta=VIEJO)
        self.assertEqual(c.web_cambiado_en[1], {"codigo_sap": NUEVO, "codigo_proveedor": VIEJO})

    def test_hora_iso_compara_con_la_de_la_web(self):
        self.assertEqual(W.hora_iso(0), "1970-01-01T00:00:00.000Z")
        self.assertGreater(W.hora_iso(), NUEVO)

    def test_cache_guarda_y_carga_la_base(self):
        with tempfile.TemporaryDirectory() as d:
            c = W.Cache(os.path.join(d, "c.json"))
            c.base = {1: {"codigo_sap": "100"}}
            c.web_cambiado_en = {1: {"precio": NUEVO}}
            c.local_cambiado_en = {2: {"precio": [1.5, MEDIO]}}
            c.guardar()
            c2 = W.Cache(c.ruta).cargar()
            self.assertEqual((c2.base, c2.web_cambiado_en, c2.local_cambiado_en),
                             ({1: {"codigo_sap": "100"}}, {1: {"precio": NUEVO}}, {2: {"precio": [1.5, MEDIO]}}))


@unittest.skipUnless(HAY_TK, "necesita openpyxl y tkinter")
class TestEscribirCampos(unittest.TestCase):
    def test_solo_toca_las_celdas_que_bajan(self):
        with tempfile.TemporaryDirectory() as d:
            ruta = os.path.join(d, "catalogo-dali-completo.xlsx")
            _excel(ruta, [{"Código DALI": 1, "Descripción - DALI": "TOMATE", "Código SAP": 100, "Precio €": 1.0},
                          {"Código DALI": 2, "Descripción - DALI": "PERA", "Código SAP": 200}])
            n, faltan = C.escribir_campos(ruta, {1: {"cod_sap": "150", "cod_prov": "P-1"}, 9: {"precio": 3.0}})
            self.assertEqual((n, faltan), (2, [9]))
            t = {k: d for k, d in C.leer_catalogo(ruta)}
            self.assertEqual((t[1]["cod_sap"], t[1]["cod_prov"], t[1]["precio"], t[1]["descripcion"]),
                             (150, "P-1", 1.0, "TOMATE"))
            self.assertEqual(t[2]["cod_sap"], 200)


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class TestDePuntaAPunta(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        os.makedirs(os.path.join(self.dir.name, "RECURSOS"))
        self.ruta = os.path.join(self.dir.name, "RECURSOS", "catalogo-dali-completo.xlsx")
        self.cache = os.path.join(self.dir.name, "cache.json")
        app = object.__new__(P.PublicarCatalogoWeb)
        app.logs = []
        app.log_consola = lambda m, n=">>": app.logs.append((n, m))
        app._carpeta_proyecto = _Var(self.dir.name)
        app._catalogo_lock = threading.Lock()
        app._cache_web_lock = threading.RLock()
        app._ejecutar_en_hilo_principal_y_esperar = lambda f: f()
        app._obtener_credenciales_dali_app = lambda: ("u@x", "clave")
        app.config_data = {}
        self.app = app
        self.web = WebFalsa([])
        web = self.web
        self.parches = [
            mock.patch.object(W, "ruta_cache_por_defecto", lambda: self.cache),
            mock.patch.object(W.ClienteCatalogo, "__init__",
                              lambda s, base=W.DALI_APP_API_BASE, abrir=None: (setattr(s, "base", "https://x"),
                                                                               setattr(s, "_abrir", web))[0]),
        ]
        for p in self.parches:
            p.start()

    def tearDown(self):
        for p in self.parches:
            p.stop()
        self.dir.cleanup()

    def _fila(self, **extra):
        return {"Código DALI": 1, "Descripción - DALI": "TOMATE", "Activo": "SI",
                "Proveedor Asignado - DALI": "PROV A"} | extra

    def _lotes(self):
        return [p[3] for p in self.web.peticiones if p[1] == "/admin/catalogo/lote" and not p[3]["simular"]]

    def _tabla(self):
        return {k: d for k, d in C.leer_catalogo(self.ruta)}

    def test_lo_cambiado_en_la_web_baja_y_no_se_pisa(self):
        # Primer envío: Excel y web coinciden → queda la base
        _excel(self.ruta, [self._fila(**{"Código SAP": 100, "Código Proveedor": "P-1", "Precio €": 1.0})])
        self.web.articulos = [_art(1, nombre_articulo="TOMATE", proveedor="PROV A", codigo_sap="100",
                                   codigo_proveedor="P-1", precio=1.0, modificado_en=VIEJO)]
        with mock.patch.object(V, "pedir_confirmacion", side_effect=AssertionError("no hay nada que enviar")):
            self.app._publicar_catalogo_web_una_vez("Acción 1", None)
        self.assertEqual(W.Cache(self.cache).cargar().base[1]["codigo_sap"], "100")
        # En la web cambian el código SAP y el código de proveedor; en la Consola, el precio (Acción 4)
        self.web.articulos = [_art(1, nombre_articulo="TOMATE", proveedor="PROV A", codigo_sap="150",
                                   codigo_proveedor="P-9", precio=1.0, modificado_en=NUEVO)]
        _excel(self.ruta, [self._fila(**{"Código SAP": 100, "Código Proveedor": "P-1", "Precio €": 1.25})])
        with mock.patch.object(V, "pedir_confirmacion", return_value=(True, True)):
            self.app._publicar_catalogo_web_una_vez("completar la tarifa (Acción 4)", None)
        t = self._tabla()
        self.assertEqual((t[1]["cod_sap"], t[1]["cod_prov"], t[1]["precio"]), (150, "P-9", 1.25))
        lotes = self._lotes()
        self.assertEqual([(l["origen"], l["filas"]) for l in lotes], [("accion4", [{"codigo_dali": 1, "precio": 1.25}])])

    def test_traer_antes_de_una_accion_sin_enviar_nada(self):
        _excel(self.ruta, [self._fila(**{"Código SAP": 100})])
        self.web.articulos = [_art(1, nombre_articulo="TOMATE", proveedor="PROV A", codigo_sap="100",
                                   modificado_en=VIEJO)]
        self.app._traer_cambios_web("la Acción 3")                       # primera vez: base
        self.web.articulos[0] = _art(1, nombre_articulo="TOMATE", proveedor="PROV A", codigo_sap="777",
                                     descripcion_sap="PUESTO EN LA WEB", modificado_en=NUEVO)
        self.app._traer_cambios_web("la Acción 3")
        t = self._tabla()
        self.assertEqual((t[1]["cod_sap"], t[1]["desc_sap"]), (777, "PUESTO EN LA WEB"))
        self.assertEqual(self._lotes(), [])
        self.assertTrue(any("Traído de Catálogo Asignaciones" in m for _n, m in self.app.logs))

    def test_con_el_excel_abierto_al_enviar_no_sube_lo_que_gana_la_web(self):
        _excel(self.ruta, [self._fila(**{"Código SAP": 100, "Precio €": 1.0})])
        self.web.articulos = [_art(1, nombre_articulo="TOMATE", proveedor="PROV A", codigo_sap="100", precio=1.0,
                                   modificado_en=VIEJO)]
        self.app._traer_cambios_web("la Acción 3")                       # base
        self.web.articulos[0] = _art(1, nombre_articulo="TOMATE", proveedor="PROV A", codigo_sap="200", precio=1.0,
                                     modificado_en=NUEVO)
        _excel(self.ruta, [self._fila(**{"Código SAP": 100, "Precio €": 1.5})])
        with mock.patch.object(C, "escribir_campos", side_effect=PermissionError), \
                mock.patch.object(V, "pedir_confirmacion", return_value=(True, True)):
            self.app._publicar_catalogo_web_una_vez("Acción 4", None)
        filas = [f for l in self._lotes() for f in l["filas"]]
        self.assertEqual(filas, [{"codigo_dali": 1, "precio": 1.5}])     # el SAP 100 no pisa el 200 de la web

    def test_sin_conexion_la_accion_sigue(self):
        _excel(self.ruta, [self._fila(**{"Código SAP": 100})])
        self.app._obtener_credenciales_dali_app = lambda: (None, None)
        self.app._traer_cambios_web("la Acción 3")                       # no lanza nada
        self.assertEqual(self._tabla()[1]["cod_sap"], 100)

    def test_con_el_excel_abierto_no_baja_y_lo_avisa(self):
        _excel(self.ruta, [self._fila(**{"Código SAP": 100})])
        self.web.articulos = [_art(1, nombre_articulo="TOMATE", proveedor="PROV A", codigo_sap="100",
                                   modificado_en=VIEJO)]
        self.app._traer_cambios_web("la Acción 3")                       # base
        self.web.articulos[0] = _art(1, nombre_articulo="TOMATE", proveedor="PROV A", codigo_sap="200",
                                     modificado_en=NUEVO)
        with mock.patch.object(C, "escribir_campos", side_effect=PermissionError):
            self.app._traer_cambios_web("la Acción 3")
        self.assertTrue(any("abierto" in m for n, m in self.app.logs if n == "WARN"))
        self.assertEqual(self._tabla()[1]["cod_sap"], 100)
        self.assertNotEqual(W.Cache(self.cache).cargar().base[1]["codigo_sap"], "200")   # se reintentará


if __name__ == "__main__":
    unittest.main()
