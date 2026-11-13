# -*- coding: utf-8 -*-
"""Pruebas de app/comun/catalogo_web.py (etapa 6: la Consola escribe en Catálogo Asignaciones).
Sin red: la web se simula con un servidor falso que responde como la v1.20.63.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import io
import json
import os
import sys
import tempfile
import unittest
import urllib.error
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import catalogo_web as W  # noqa: E402
from app.comun.sap_helpers import CATALOGO_COLUMNAS, CATALOGO_FILA_CABECERA  # noqa: E402


class _Resp(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *a):
        self.close()


def _art(cod, **campos):
    base = {c: None for c in set().union(*W.GRUPOS.values())}
    base.update(codigo_dali=cod, activo="SI", modificado_en="2026-11-10T08:00:00.000Z")
    base.update(campos)
    return base


class WebFalsa:
    """Responde a /auth/login, /admin/catalogo/cambios y /admin/catalogo/lote; apunta cada petición."""

    def __init__(self, articulos, por_pagina=2, estado_lote=200):
        self.articulos = articulos
        self.por_pagina = por_pagina
        self.estado_lote = estado_lote
        self.peticiones = []

    def __call__(self, req, timeout=None):
        url = urlparse(req.full_url)
        cuerpo = json.loads(req.data.decode("utf-8")) if req.data else None
        self.peticiones.append((req.get_method(), url.path, parse_qs(url.query), cuerpo))
        if url.path == "/auth/login":
            return _Resp(json.dumps({"ok": True}).encode())
        if url.path == "/admin/catalogo/cambios":
            q = parse_qs(url.query)
            pagina = int(q.get("pagina", ["1"])[0])
            lista = self.articulos
            if "desde" in q:
                lista = [a for a in lista if a["modificado_en"] > q["desde"][0]]
            trozo = lista[(pagina - 1) * self.por_pagina: pagina * self.por_pagina]
            return _Resp(json.dumps({"hasta": "2026-11-12T10:00:00.000Z", "articulos": trozo,
                                     "hay_mas": pagina * self.por_pagina < len(lista)}).encode())
        if url.path == "/admin/catalogo/lote":
            if self.estado_lote != 200:
                raise urllib.error.HTTPError(req.full_url, self.estado_lote, "Forbidden", {},
                                             io.BytesIO(b'{"error":"Sin permiso"}'))
            n = len(cuerpo["filas"])
            return _Resp(json.dumps({"ok": True, "simulado": cuerpo["simular"],
                                     "resumen": {"recibidas": n, "cambiados": n},
                                     "cambios": [{"codigo_dali": f["codigo_dali"]} for f in cuerpo["filas"]]}).encode())
        raise AssertionError(url.path)


class TestValores(unittest.TestCase):
    def test_normalizacion_como_la_web(self):
        self.assertEqual(W.texto(10100416.0), "10100416")
        self.assertIsNone(W.texto("   "))
        self.assertEqual(W.precio("12,345"), 12.35)
        self.assertEqual(W.precio(2.675), 2.68)          # igual que Math.round(n*100)/100 de la web
        self.assertIsNone(W.precio(""))
        self.assertTrue(W.activo("Sí"))
        self.assertFalse(W.activo("NO"))
        self.assertEqual(W.codigo_dali("1234.0"), 1234)
        self.assertIsNone(W.codigo_dali("abc"))


class TestCatalogoLocal(unittest.TestCase):
    def test_lee_por_cabecera(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        cabecera = [titulo for _k, titulo, _a in CATALOGO_COLUMNAS]
        cabecera.reverse()                              # en otro orden: se buscan por título
        for i, t in enumerate(cabecera, start=1):
            ws.cell(row=CATALOGO_FILA_CABECERA, column=i, value=t)
        col = {t: i for i, t in enumerate(cabecera, start=1)}
        ws.cell(row=5, column=col["Código DALI"], value=101)
        ws.cell(row=5, column=col["Descripción - DALI"], value=" Tomate  ")
        ws.cell(row=5, column=col["Precio €"], value=1.5)
        ws.cell(row=5, column=col["Activo"], value="SI")
        ws.cell(row=5, column=col["Código SAP"], value=10100416.0)
        ws.cell(row=6, column=col["Descripción - DALI"], value="sin código: se ignora")
        with tempfile.TemporaryDirectory() as d:
            ruta = os.path.join(d, "catalogo.xlsx")
            wb.save(ruta)
            local = W.leer_catalogo_local(ruta)
        self.assertEqual(list(local), [101])
        self.assertEqual(local[101]["nombre_articulo"], "Tomate")
        self.assertEqual(local[101]["precio"], 1.5)
        self.assertEqual(local[101]["codigo_sap"], "10100416")
        self.assertTrue(local[101]["activo"])


class TestCacheYDiferencias(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.ruta = os.path.join(self.dir.name, "sub", "catalogo_web.json")

    def tearDown(self):
        self.dir.cleanup()

    def test_sincroniza_por_paginas_y_luego_solo_lo_cambiado(self):
        web = WebFalsa([_art(1, nombre_articulo="A"), _art(2, nombre_articulo="B"), _art(3, nombre_articulo="C")])
        cli = W.ClienteCatalogo(abrir=web)
        cache = W.Cache(self.ruta)
        self.assertEqual(cache.sincronizar(cli), 3)
        self.assertEqual(len([p for p in web.peticiones if p[1] == "/admin/catalogo/cambios"]), 2)
        self.assertEqual(W.Cache(self.ruta).cargar().articulos[3]["nombre_articulo"], "C")
        # segunda vez: solo lo cambiado desde «hasta»
        web.articulos[1] = _art(2, nombre_articulo="B2", modificado_en="2026-11-12T11:00:00.000Z")
        web.peticiones.clear()
        cache2 = W.Cache(self.ruta).cargar()
        self.assertEqual(cache2.sincronizar(cli), 1)
        self.assertEqual(web.peticiones[0][2]["desde"], ["2026-11-12T10:00:00.000Z"])
        self.assertEqual(cache2.articulos[2]["nombre_articulo"], "B2")

    def test_cache_danada_empieza_de_cero(self):
        os.makedirs(os.path.dirname(self.ruta))
        with open(self.ruta, "w") as f:
            f.write("{no es json")
        c = W.Cache(self.ruta).cargar()
        self.assertIsNone(c.hasta)
        self.assertEqual(c.articulos, {})

    def test_diferencias_por_grupo(self):
        cache = W.Cache(self.ruta)
        cache.aplicar_pagina([
            _art(1, nombre_articulo="Tomate", proveedor="PROV A", precio=1.0, codigo_proveedor="X1", codigo_sap="S1"),
            _art(2, nombre_articulo="Pera", proveedor="PROV B", precio=2.0, modificado_en="2026-11-12T09:00:00.000Z"),
        ])
        base = {c: None for c in set().union(*W.GRUPOS.values())}
        local = {
            1: dict(base, nombre_articulo="Tomate", proveedor="PROV A", activo=True, precio=1.2,
                    codigo_proveedor="X1", codigo_sap="S9"),
            2: dict(base, nombre_articulo="Pera conferencia", proveedor="PROV B", activo=True, precio=None),
            3: dict(base, nombre_articulo="Nuevo", proveedor="PROV A", activo=True, precio=5.0),
        }
        dif = W.calcular_diferencias(local, cache.articulos, desde_sincronizacion="2026-11-11T00:00:00.000Z")
        a2 = {d["codigo_dali"]: d for d in dif["accion2"]}
        self.assertEqual(set(a2), {2, 3})                         # 1 no cambia en lo DALI
        self.assertTrue(a2[3]["nuevo"])
        self.assertTrue(a2[2]["editado_en_web"])                  # cambió en la web tras la última sincronización
        self.assertEqual([d["codigo_dali"] for d in dif["accion3"]], [1])
        self.assertEqual(dif["accion3"][0]["campos"], {"codigo_sap": ("S1", "S9")})
        # accion4: el precio vacío del 2 no borra el de la web; el 3 (nuevo) lleva su precio tras el alta
        self.assertEqual([d["codigo_dali"] for d in dif["accion4"]], [1, 3])
        self.assertEqual(dif["accion4"][0]["campos"], {"precio": (1.0, 1.2)})

        # el 2 se tocó en la web, pero su descripción DALI solo se carga desde la Consola: va igual
        a2_filas = W.filas_para_enviar(dif, "accion2")
        self.assertEqual([f["codigo_dali"] for f in a2_filas[None]], [2, 3])
        self.assertEqual(a2_filas[None][0], {"codigo_dali": 2, "nombre_articulo": "Pera conferencia"})
        a2_todas = W.filas_para_enviar(dif, "accion2", incluir_editados=True)
        self.assertEqual(sorted(f["codigo_dali"] for f in a2_todas[None]), [2, 3])
        nuevo = next(f for f in a2_todas[None] if f["codigo_dali"] == 3)
        self.assertEqual(nuevo["activo"], "SI")                   # «activo» viaja como SI/NO, igual que en el Excel
        self.assertEqual(nuevo["nombre_articulo"], "Nuevo")
        a4 = W.filas_para_enviar(dif, "accion4")
        self.assertEqual(a4, {"PROV A": [{"codigo_dali": 1, "precio": 1.2}, {"codigo_dali": 3, "precio": 5.0}]})

    def test_editado_en_web_no_pisa_lo_sap(self):
        cache = W.Cache(os.path.join(self.dir.name, "c.json"))
        cache.aplicar_pagina([_art(5, nombre_articulo="Leche", codigo_sap="S1",
                                   modificado_en="2026-11-12T09:00:00.000Z")])
        base = {c: None for c in set().union(*W.GRUPOS.values())}
        local = {5: dict(base, nombre_articulo="Leche", activo=True, codigo_sap="S2")}
        dif = W.calcular_diferencias(local, cache.articulos, desde_sincronizacion="2026-11-11T00:00:00.000Z")
        self.assertEqual(W.filas_para_enviar(dif, "accion3"), {})
        self.assertEqual(W.filas_para_enviar(dif, "accion3", incluir_editados=True),
                         {None: [{"codigo_dali": 5, "codigo_sap": "S2"}]})

    def test_protegidos_siguen_protegidos(self):
        cache = W.Cache(os.path.join(self.dir.name, "p.json"))
        cache.aplicar_pagina([_art(5, codigo_sap="S1", modificado_en="2026-11-12T09:00:00.000Z"),
                              _art(6, codigo_sap="T1")])
        base = {c: None for c in set().union(*W.GRUPOS.values())}
        local = {5: dict(base, activo=True, codigo_sap="S2"), 6: dict(base, activo=True, codigo_sap="T2")}
        dif = W.calcular_diferencias(local, cache.articulos, "2026-11-11T00:00:00.000Z")
        self.assertEqual(W.contar(dif, False)["accion3"], 1)
        W.actualizar_protegidos(cache, dif, incluir_editados=False)
        self.assertEqual(cache.protegidos, {5})
        cache.guardar()
        cache = W.Cache(cache.ruta).cargar()
        # la sincronización siguiente ya es posterior al cambio en la web: sigue protegido
        dif = W.calcular_diferencias(local, cache.articulos, "2026-11-13T00:00:00.000Z", cache.protegidos)
        self.assertEqual(W.filas_para_enviar(dif, "accion3"), {None: [{"codigo_dali": 6, "codigo_sap": "T2"}]})
        # si se incluye a propósito, deja de estarlo
        W.actualizar_protegidos(cache, dif, incluir_editados=True)
        self.assertEqual(cache.protegidos, set())

    def test_en_lotes(self):
        self.assertEqual([len(x) for x in W.en_lotes(list(range(901)), 400)], [400, 400, 101])


class TestEnvio(unittest.TestCase):
    def test_envia_por_lotes_y_proveedor(self):
        web = WebFalsa([])
        cli = W.ClienteCatalogo(abrir=web)
        cli.login("u@x", "clave")
        por_prov = {"PROV A": [{"codigo_dali": i, "precio": 1.0} for i in range(1, 451)],
                    "PROV B": [{"codigo_dali": 999, "precio": 2.0}]}
        r = W.enviar_grupo(cli, "accion4", por_prov, simular=True, archivo="TARIFA.xls",
                           sincronizado_desde="2026-11-12T10:00:00.000Z")
        lotes = [p[3] for p in web.peticiones if p[1] == "/admin/catalogo/lote"]
        self.assertEqual([len(l["filas"]) for l in lotes], [400, 50, 1])
        self.assertEqual([l["proveedor_tarifa"] for l in lotes], ["PROV A", "PROV A", "PROV B"])
        self.assertTrue(all(l["simular"] is True and l["archivo"] == "TARIFA.xls" for l in lotes))
        self.assertEqual(r["resumen"]["recibidas"], 451)
        self.assertEqual(len(r["cambios"]), 451)

    def test_accion2_sin_proveedor_tarifa(self):
        web = WebFalsa([])
        W.enviar_grupo(W.ClienteCatalogo(abrir=web), "accion2", {None: [{"codigo_dali": 1}]}, simular=False)
        cuerpo = web.peticiones[0][3]
        self.assertNotIn("proveedor_tarifa", cuerpo)
        self.assertIs(cuerpo["simular"], False)

    def test_403_explica_el_permiso(self):
        web = WebFalsa([], estado_lote=403)
        with self.assertRaises(W.ErrorWeb) as ctx:
            W.ClienteCatalogo(abrir=web).lote("accion2", [{"codigo_dali": 1}], True)
        self.assertEqual(ctx.exception.codigo, 403)
        self.assertIn("Escribir desde la Consola Registro Info", str(ctx.exception))

    def test_sin_conexion(self):
        def caido(req, timeout=None):
            raise OSError("red caída")
        with self.assertRaises(W.ErrorWeb) as ctx:
            W.ClienteCatalogo(abrir=caido).cambios()
        self.assertIsNone(ctx.exception.codigo)


class _ClienteLogin:
    def __init__(self, respuesta=None, error=None):
        self.respuesta, self.error = respuesta, error

    def login(self, email, password):
        if self.error:
            raise self.error
        return self.respuesta


class TestAccesoConsola(unittest.TestCase):
    def test_solo_admin_con_permiso_o_admin_principal(self):
        self.assertTrue(W.usuario_autorizado({"rol": "admin", "permisos": ["admin-catalogo-escribir"]}))
        self.assertTrue(W.usuario_autorizado({"rol": "admin", "esAdminPrincipal": True, "permisos": []}))
        self.assertFalse(W.usuario_autorizado({"rol": "admin", "permisos": ["admin-articulos"]}))
        self.assertFalse(W.usuario_autorizado({"rol": "hotel", "permisos": ["admin-catalogo-escribir"]}))
        self.assertFalse(W.usuario_autorizado({"rol": "hotel", "esAdminPrincipal": True}))
        self.assertFalse(W.usuario_autorizado(None))

    def test_comprobar_acceso(self):
        ok = {"ok": True, "usuario": {"nombre": "Ana", "rol": "admin", "permisos": ["admin-catalogo-escribir"]}}
        estado, mensaje, usuario = W.comprobar_acceso(_ClienteLogin(ok), "a@x", "c")
        self.assertEqual((estado, usuario["nombre"]), ("ok", "Ana"))
        self.assertIn("Ana", mensaje)
        sin_permiso = {"ok": True, "usuario": {"nombre": "Luis", "rol": "admin", "permisos": []}}
        self.assertEqual(W.comprobar_acceso(_ClienteLogin(sin_permiso), "l@x", "c")[:2],
                         ("no_autorizado", "Usuario no autorizado."))
        hotel = {"ok": True, "usuario": {"rol": "hotel", "permisos": []}}
        self.assertEqual(W.comprobar_acceso(_ClienteLogin(hotel), "h@x", "c")[0], "no_autorizado")
        casos = [(W.ErrorWeb("HTTP 401", 401), "credenciales"), (W.ErrorWeb("verif", 428), "verificacion"),
                 (W.ErrorWeb("sin red"), "sin_conexion"), (W.ErrorWeb("HTTP 429", 429), "error"),
                 (W.ErrorWeb("HTTP 403", 403), "no_autorizado")]
        for error, esperado in casos:
            self.assertEqual(W.comprobar_acceso(_ClienteLogin(error=error), "a@x", "c")[0], esperado)
        self.assertEqual(W.comprobar_acceso(_ClienteLogin(ok), " ", "c")[0], "credenciales")


if __name__ == "__main__":
    unittest.main()
