# -*- coding: utf-8 -*-
"""Pruebas de app/comun/actualizaciones.py: aviso de versión nueva, descarga con SHA-256 y sustitución
del .exe. Sin red: las descargas se simulan.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import hashlib
import io
import json
import urllib.error
from unittest import mock
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import actualizaciones as A  # noqa: E402


def _release(tag="v10.2.40", con_sha=True, **extra):
    assets = [{"name": "TARIFA_Finder.exe", "size": 4, "browser_download_url": "https://x/exe"}]
    if con_sha:
        assets.append({"name": "TARIFA_Finder.exe.sha256", "browser_download_url": "https://x/sha"})
    r = {"tag_name": tag, "name": "Consola v 10.2.40", "body": "## Novedades\n* Algo", "assets": assets}
    r.update(extra)
    return r


class _Resp(io.BytesIO):
    def __init__(self, datos):
        super().__init__(datos)
        self.headers = {"Content-Length": str(len(datos))}

    def __enter__(self):
        return self

    def __exit__(self, *a):
        self.close()


def _abridor(exe=b"NUEVO", sha=None):
    sha = hashlib.sha256(exe).hexdigest() if sha is None else sha
    return lambda url, timeout: _Resp(exe if url.endswith("exe") else f"{sha}  TARIFA_Finder.exe\n".encode())


class Versiones(unittest.TestCase):
    def test_numeros_y_comparacion(self):
        self.assertEqual(A.numeros("v 10.2.38"), (10, 2, 38))
        self.assertTrue(A.es_mas_nueva("10.2.40", "v 10.2.38"))
        self.assertTrue(A.es_mas_nueva("10.2.100", "v 10.2.98"))
        self.assertFalse(A.es_mas_nueva("10.2.38", "v 10.2.38"))
        self.assertFalse(A.es_mas_nueva("10.2.36", "v 10.2.38"))
        self.assertFalse(A.es_mas_nueva("", "v 10.2.38"))

    def test_datos_release(self):
        d = A.datos_release(_release())
        self.assertEqual((d["version"], d["etiqueta"], d["url_exe"], d["url_sha"]),
                         ("10.2.40", "v10.2.40", "https://x/exe", "https://x/sha"))
        self.assertIsNone(A.datos_release(_release(draft=True)))
        self.assertIsNone(A.datos_release(_release(prerelease=True)))
        self.assertIsNone(A.datos_release({"tag_name": "v1.0.0", "assets": []}))
        self.assertIsNone(A.datos_release(None))
        self.assertIsNone(A.datos_release(_release(con_sha=False))["url_sha"])

    def test_extraer_hash(self):
        h = "a" * 64
        self.assertEqual(A.extraer_hash(f"{h}  TARIFA_Finder.exe"), h)
        self.assertEqual(A.extraer_hash(f"{h.upper()} *TARIFA_Finder.exe"), h)
        self.assertEqual(A.extraer_hash(h), h)
        self.assertIsNone(A.extraer_hash(f"{h}  otro.exe"))
        self.assertIsNone(A.extraer_hash("basura"))


class SinReleasePublicada(unittest.TestCase):
    def test_404_es_que_no_hay_version_nueva(self):
        import urllib.error
        from unittest import mock
        err = urllib.error.HTTPError(A.API_ULTIMA, 404, "Not Found", {}, None)
        with mock.patch.object(A.urllib.request, "urlopen", side_effect=err):
            self.assertIsNone(A.consultar_ultima())
        err = urllib.error.HTTPError(A.API_ULTIMA, 403, "rate limit", {}, None)
        with mock.patch.object(A.urllib.request, "urlopen", side_effect=err):
            with self.assertRaises(urllib.error.HTTPError):
                A.consultar_ultima()


class NotasDeUnaVersion(unittest.TestCase):
    """(nov 2026) Al pulsar la versión: las notas de la Release de esa versión, por su etiqueta."""
    def test_pide_la_etiqueta_de_esa_version(self):
        cuerpo = json.dumps({"tag_name": "v10.2.44", "body": "## What's Changed\n* Icono nuevo by @x in #11"}).encode()
        with mock.patch.object(A.urllib.request, "urlopen", return_value=_Resp(cuerpo)) as u:
            notas = A.notas_de_version("v 10.2.44")
        self.assertIn("Icono nuevo", notas)
        self.assertTrue(u.call_args.args[0].full_url.endswith("/releases/tags/v10.2.44"))

    def test_sin_release_o_sin_numero(self):
        err = urllib.error.HTTPError(A.API_ETIQUETA, 404, "Not Found", {}, None)
        with mock.patch.object(A.urllib.request, "urlopen", side_effect=err):
            self.assertIsNone(A.notas_de_version("v 10.2.36"))
        self.assertIsNone(A.notas_de_version("sin versión"))


class NotasLegibles(unittest.TestCase):
    def test_notas_de_github_en_castellano(self):
        try:
            from app.ui.ventana_actualizacion import _notas_legibles
        except Exception:  # sin tkinter
            self.skipTest("sin tkinter")
        txt = ("## What's Changed\n* Actualizaciones automáticas by @usuario in https://github.com/x/pull/8\n"
               "## New Contributors\n* @otro made their first contribution\n\n**Full Changelog**: https://x")
        self.assertEqual(_notas_legibles(txt), "Cambios\n•  Actualizaciones automáticas")


class DescargaYSustitucion(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.exe = os.path.join(self.tmp.name, "TARIFA_Finder.exe")
        with open(self.exe, "wb") as f:
            f.write(b"VIEJO")
        self.info = A.datos_release(_release())

    def tearDown(self):
        self.tmp.cleanup()

    def test_descarga_correcta_junto_al_exe(self):
        avances = []
        nuevo = A.descargar_y_verificar(self.info, self.exe, lambda h, t: avances.append(h), abrir=_abridor())
        self.assertEqual(nuevo, os.path.join(self.tmp.name, "TARIFA_Finder.nuevo.exe"))
        with open(nuevo, "rb") as f:
            self.assertEqual(f.read(), b"NUEVO")
        self.assertTrue(avances)

    def test_hash_distinto_borra_y_no_toca_nada(self):
        with self.assertRaises(RuntimeError):
            A.descargar_y_verificar(self.info, self.exe, abrir=_abridor(sha="b" * 64))
        self.assertFalse(os.path.exists(A.rutas(self.exe)[0]))
        with open(self.exe, "rb") as f:
            self.assertEqual(f.read(), b"VIEJO")

    def test_sin_sha256_no_se_instala(self):
        info = A.datos_release(_release(con_sha=False))
        with self.assertRaises(RuntimeError):
            A.descargar_y_verificar(info, self.exe, abrir=_abridor())

    def test_sustituir_y_limpiar(self):
        nuevo = A.descargar_y_verificar(self.info, self.exe, abrir=_abridor())
        A.sustituir_exe(nuevo, self.exe)
        with open(self.exe, "rb") as f:
            self.assertEqual(f.read(), b"NUEVO")
        anterior = A.rutas(self.exe)[1]
        with open(anterior, "rb") as f:
            self.assertEqual(f.read(), b"VIEJO")
        self.assertEqual(A.limpiar_restos(self.exe), (["TARIFA_Finder.anterior.exe"], []))
        self.assertFalse(os.path.exists(anterior))
        self.assertEqual(A.limpiar_restos(self.exe), ([], []))     # ya no queda nada
        self.assertTrue(os.path.exists(self.exe))

    def test_si_falla_colocar_el_nuevo_se_deja_el_actual(self):
        with self.assertRaises(RuntimeError):
            A.sustituir_exe(os.path.join(self.tmp.name, "no_existe.exe"), self.exe)
        with open(self.exe, "rb") as f:
            self.assertEqual(f.read(), b"VIEJO")

    def test_si_no_se_puede_borrar_queda_pendiente(self):
        from unittest import mock
        anterior = A.rutas(self.exe)[1]
        open(anterior, "wb").close()
        with mock.patch.object(A.os, "remove", side_effect=PermissionError("en uso")):
            self.assertEqual(A.limpiar_restos(self.exe), ([], ["TARIFA_Finder.anterior.exe"]))
        self.assertTrue(os.path.exists(anterior))

    def test_entorno_limpio_para_el_exe_nuevo(self):
        env = A.entorno_limpio({"PATH": "x", "_PYI_ARCHIVE_FILE": "y", "_MEIPASS2": "z", "_PYI_PARENT_PROCESS_LEVEL": "1"})
        self.assertEqual(env, {"PATH": "x", "PYINSTALLER_RESET_ENVIRONMENT": "1"})

    def test_desde_el_codigo_no_se_actualiza(self):
        self.assertFalse(getattr(sys, "frozen", False))
        self.assertFalse(A.puede_actualizarse())


if __name__ == "__main__":
    unittest.main()
