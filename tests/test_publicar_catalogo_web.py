# -*- coding: utf-8 -*-
"""Pruebas del envío a Catálogo Asignaciones tras las Acciones 2, 3 y 4 (etapa 6, v10.2.50):
app/principal/publicar_catalogo_web.py con una web falsa y la ventana sustituida. Sin red.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import tempfile
import threading
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    import tkinter  # noqa: F401
    from app.principal import publicar_catalogo_web as P
    from app.ui import ventana_publicar_catalogo as V
    from app.comun import catalogo_web as W
    from app.comun.sap_helpers import CATALOGO_COLUMNAS, CATALOGO_FILA_CABECERA
    from tests.test_catalogo_web import WebFalsa, _art
    HAY_TK = True
except ImportError:  # sin tkinter (la ventana lo importa)
    HAY_TK = False


class _Var:
    def __init__(self, v):
        self.v = v

    def get(self):
        return self.v


def _excel(ruta, filas):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    titulos = [t for _k, t, _a in CATALOGO_COLUMNAS]
    for i, t in enumerate(titulos, start=1):
        ws.cell(row=CATALOGO_FILA_CABECERA, column=i, value=t)
    col = {t: i for i, t in enumerate(titulos, start=1)}
    for n, fila in enumerate(filas, start=CATALOGO_FILA_CABECERA + 1):
        for titulo, valor in fila.items():
            ws.cell(row=n, column=col[titulo], value=valor)
    wb.save(ruta)


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class TestPublicar(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        os.makedirs(os.path.join(self.dir.name, "RECURSOS"))
        self.ruta = os.path.join(self.dir.name, "RECURSOS", "catalogo-dali-completo.xlsx")
        _excel(self.ruta, [
            {"Código DALI": 1, "Descripción - DALI": "Tomate pera", "Activo": "SI", "Proveedor Asignado - DALI": "PROV A",
             "Precio €": 1.2, "Código SAP": "S1"},
            {"Código DALI": 2, "Descripción - DALI": "Nuevo", "Activo": "SI", "Proveedor Asignado - DALI": "PROV A",
             "Precio €": 3.0},
        ])
        self.web = WebFalsa([_art(1, nombre_articulo="Tomate", activo="SI", proveedor="PROV A", precio=1.0,
                                  codigo_sap="S1")])
        self.app = object.__new__(P.PublicarCatalogoWeb)
        self.app.logs = []
        self.app.log_consola = lambda m, n=">>": self.app.logs.append((n, m))
        self.app._carpeta_proyecto = _Var(self.dir.name)
        self.app._catalogo_lock = threading.Lock()
        self.app._ejecutar_en_hilo_principal_y_esperar = lambda f: f()
        self.app._obtener_credenciales_dali_app = lambda: ("u@x", "clave")
        self.app.config_data = {}
        self.app._publicacion_web_estado = threading.Lock()
        self.app._publicacion_web_activa = False
        self.app._publicacion_web_pendiente = None
        self.cache = os.path.join(self.dir.name, "cache.json")
        web = self.web
        # (18 nov 2026) El aviso «sin conexión» no abre ventanas en las pruebas: se apunta y responde self.respuesta
        self.avisos, self.respuesta = [], "seguir"
        self.parches = [
            mock.patch.object(P.ventana_sin_conexion, "preguntar",
                              lambda parent, que, info, detalle="": (self.avisos.append((que, info["tipo"])),
                                                                     self.respuesta)[1]),
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

    def _lotes(self, simular):
        return [p[3] for p in self.web.peticiones if p[1] == "/admin/catalogo/lote" and p[3]["simular"] is simular]

    def test_simula_pregunta_y_aplica_solo_lo_distinto(self):
        vistos = {}

        def ventana(parent, dif, sims, motivo):
            vistos["dif"], vistos["sims"] = dif, sims
            return True, False
        with mock.patch.object(V, "pedir_confirmacion", ventana):
            self.app._publicar_catalogo_web_una_vez("Acción 2", None)
        self.assertEqual(sorted(vistos["sims"]), ["accion2", "accion4"])
        aplicados = self._lotes(False)
        self.assertEqual([l["origen"] for l in aplicados], ["accion2", "accion4"])   # altas antes que precios
        self.assertEqual(aplicados[0]["filas"][0], {"codigo_dali": 1, "nombre_articulo": "Tomate pera"})
        self.assertEqual(aplicados[0]["filas"][1]["codigo_dali"], 2)
        self.assertEqual(aplicados[1]["proveedor_tarifa"], "PROV A")
        self.assertEqual(aplicados[1]["filas"], [{"codigo_dali": 1, "precio": 1.2}, {"codigo_dali": 2, "precio": 3.0}])
        cache = W.Cache(self.cache).cargar()
        self.assertIsNotNone(cache.publicado_hasta)

    def test_codigo_de_proveedor_que_la_web_no_acepta_se_queda_el_de_la_web(self):
        """19 nov 2026 (DALI 21905): la web tiene 243528 y el Excel «COMODIN». La web no lo cambia; antes el aviso
        salía en cada envío. Ahora se queda el de la web también en el Excel y la vez siguiente ya no se envía."""
        _excel(self.ruta, [{"Código DALI": 1, "Descripción - DALI": "Tomate", "Activo": "SI",
                            "Proveedor Asignado - DALI": "PROV A", "Precio €": 1.0, "Código SAP": "S1",
                            "Código Proveedor": "COMODIN"}])
        self.web.articulos = [_art(1, nombre_articulo="Tomate", activo="SI", proveedor="PROV A", precio=1.0,
                                   codigo_sap="S1", codigo_proveedor="243528")]
        self.web.irregulares = {"accion4": [{"codigo_dali": 1, "tipo": "codigo_proveedor_distinto",
                                             "antes": "243528", "despues": "COMODIN", "detalle": "…"}]}
        with mock.patch.object(V, "pedir_confirmacion", lambda *a: (True, False)):
            self.app._publicar_catalogo_web_una_vez("Acción 4", None)
        self.assertEqual(W.leer_catalogo_local(self.ruta)[1]["codigo_proveedor"], "243528")
        self.assertTrue(any("se queda el de la web" in m for _n, m in self.app.logs))
        self.assertFalse(any("Siguen distintos" in m for _n, m in self.app.logs))
        antes = len(self._lotes(True))
        with mock.patch.object(V, "pedir_confirmacion", lambda *a: self.fail("no debía preguntar")):
            self.app._publicar_catalogo_web_una_vez("Acción 4", None)
        self.assertEqual(len(self._lotes(True)), antes, "ya coinciden: no se envía nada")

    def test_cambio_de_nombre_del_proveedor_se_manda_a_la_web_y_no_se_repite(self):
        """19 nov 2026: todos los artículos de «PROV A» pasan a «PROV A S.L.» (nombre nuevo): el lote de datos DALI
        lleva renombres_proveedor; después la copia de la web ya tiene el nombre nuevo y no se vuelve a enviar."""
        _excel(self.ruta, [{"Código DALI": 1, "Descripción - DALI": "Tomate", "Activo": "SI",
                            "Proveedor Asignado - DALI": "PROV A S.L.", "Precio €": 1.0, "Código SAP": "S1"}])
        with mock.patch.object(V, "pedir_confirmacion", lambda *a: (True, False)):
            self.app._publicar_catalogo_web_una_vez("Acción 1", None)
        lotes = [l for l in self._lotes(False) if l["origen"] == "accion2"]
        self.assertEqual(lotes[0]["renombres_proveedor"], [{"antes": "PROV A", "despues": "PROV A S.L."}])
        self.assertTrue(any("Proveedor renombrado en DALI" in m for _n, m in self.app.logs))
        self.assertEqual(W.Cache(self.cache).cargar().articulos[1]["proveedor"], "PROV A S.L.")
        antes = len(self._lotes(True))
        with mock.patch.object(V, "pedir_confirmacion", lambda *a: self.fail("no debía preguntar")):
            self.app._publicar_catalogo_web_una_vez("Acción 1", None)
        self.assertEqual(len(self._lotes(True)), antes)

    def test_ahora_no_no_envia_nada(self):
        with mock.patch.object(V, "pedir_confirmacion", lambda *a: (False, False)):
            self.app._publicar_catalogo_web_una_vez("Acción 2", None)
        self.assertEqual(self._lotes(False), [])
        self.assertTrue(any("Ahora no" in m for _n, m in self.app.logs))

    def test_sin_diferencias_no_pregunta(self):
        _excel(self.ruta, [{"Código DALI": 1, "Descripción - DALI": "Tomate", "Activo": "SI",
                            "Proveedor Asignado - DALI": "PROV A", "Precio €": 1.0, "Código SAP": "S1"}])
        with mock.patch.object(V, "pedir_confirmacion", side_effect=AssertionError("no debía preguntar")):
            self.app._publicar_catalogo_web_una_vez("Acción 3", None)
        self.assertTrue(any("ya coincide" in m for _n, m in self.app.logs))

    def test_sin_permiso_lo_explica(self):
        self.web.estado_lote = 403
        self.app._publicar_catalogo_web_una_vez("Acción 4", "TARIFA.xls")
        self.assertTrue(any(n == "ERROR" and "Escribir desde la Consola Registro Info" in m for n, m in self.app.logs))

    def test_sin_permiso_avisa_sin_reintentar(self):
        self.web.estado_lote = 403
        self.app._publicar_catalogo_web_una_vez("Acción 4", "TARIFA.xls")
        self.assertEqual(self.avisos, [("enviar a la web los cambios de la Consola", "permiso")])

    def test_sin_conexion_avisa_una_vez_y_el_indicador_lo_dice(self):
        # (18 nov 2026, petición de Víctor) La primera vez sale el aviso; mientras siga sin conexión, no se repite
        self.web.fallos_login = 99
        self.app._publicar_catalogo_web_una_vez("Acción 2", None)
        self.app._traer_cambios_web("la Acción 3")
        self.assertEqual(self.avisos, [("enviar a la web los cambios de la Consola", "sin_red")])
        texto, clave = self.app._estado_web.indicador()
        self.assertEqual(clave, "fallo")
        self.assertIn("sin conexión desde las", texto)
        self.assertEqual(self._lotes(False), [])
        # al volver la conexión, el indicador se pone en verde y el siguiente fallo vuelve a avisar
        self.web.fallos_login = 0
        with mock.patch.object(V, "pedir_confirmacion", lambda *a: (False, False)):
            self.app._publicar_catalogo_web_una_vez("Acción 2", None)
        self.assertEqual(self.app._estado_web.indicador()[1], "ok")
        self.web.fallos_login = 1
        self.app._traer_cambios_web("la Acción 3")
        self.assertEqual(len(self.avisos), 2)
        self.assertEqual(self.avisos[1], ("traer lo cambiado en la web antes de la Acción 3", "sin_red"))

    def test_reintentar_vuelve_a_conectar_y_sigue(self):
        self.web.fallos_login = 1
        self.respuesta = "reintentar"
        with mock.patch.object(V, "pedir_confirmacion", lambda *a: (True, False)):
            self.app._publicar_catalogo_web_una_vez("Acción 2", None)
        self.assertEqual(len(self.avisos), 1)
        self.assertEqual([l["origen"] for l in self._lotes(False)], ["accion2", "accion4"])
        self.assertEqual(self.app._estado_web.indicador()[1], "ok")

    def test_sin_credenciales(self):
        self.app._obtener_credenciales_dali_app = lambda: (None, None)
        self.app._publicar_catalogo_web_una_vez("Acción 2", None)
        self.assertEqual(self.web.peticiones, [])

    def test_dos_seguidas_se_juntan(self):
        llamadas = []
        evento = threading.Event()

        def una_vez(motivo, archivo):
            llamadas.append(motivo)
            evento.wait(2)
        self.app._publicar_catalogo_web_una_vez = una_vez
        self.app._publicar_catalogo_web("A", None)
        self.app._publicar_catalogo_web("B", None)
        self.app._publicar_catalogo_web("C", None)       # sustituye a B: solo se repite una vez
        evento.set()
        for _ in range(100):
            if not self.app._publicacion_web_activa:
                break
            threading.Event().wait(0.02)
        self.assertEqual(llamadas, ["A", "C"])

    def test_detalle_de_la_ventana(self):
        dif = {"accion2": [{"codigo_dali": 2, "campos": {"nombre_articulo": (None, "Nuevo")}, "nuevo": True,
                            "editado_en_web": False, "proveedor": "P"}],
               "accion3": [{"codigo_dali": 1, "campos": {"codigo_sap": ("S1", "S2")}, "nuevo": False,
                            "editado_en_web": True, "proveedor": "P"}],
               "accion4": []}
        self.assertEqual(V.lineas_detalle(dif), ["DALI 2 · ALTA «Nuevo»",
                                                 "DALI 1 · Código SAP: S1 → S2 ✎ editado en la web"])
        sims = {"accion3": {"irregulares": [{"codigo_dali": 2, "tipo": "no_existe"},
                                            {"codigo_dali": 9, "tipo": "proveedor_distinto", "detalle": "x"}]}}
        self.assertEqual([i["codigo_dali"] for _g, i in V.irregulares_utiles(sims, dif)], [9])


if __name__ == "__main__":
    unittest.main()
