# -*- coding: utf-8 -*-
"""
test_estado_web.py — estado de la conexión con Catálogo Asignaciones y cuándo avisar (18 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import unittest
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.comun import catalogo_web as W  # noqa: E402
from app.comun import estado_web as E  # noqa: E402


class TestClasificar(unittest.TestCase):
    def test_tipos(self):
        casos = [
            (W.ErrorWeb("sin conexión con Catálogo Asignaciones (<urlopen error [Errno 11001] getaddrinfo failed>)"),
             "sin_red"),
            (W.ErrorWeb("sin conexión con Catálogo Asignaciones (The read operation timed out)"), "arrancando"),
            (W.ErrorWeb("sin conexión con Catálogo Asignaciones ([WinError 10060] ...)"), "arrancando"),
            (W.ErrorWeb("HTTP 502: Bad Gateway", 502), "servidor"),
            (W.ErrorWeb("HTTP 401: Credenciales incorrectas", 401), "credenciales"),
            (W.ErrorWeb("Catálogo Asignaciones pide un código de verificación", 428), "credenciales"),
            (W.ErrorWeb("HTTP 403: sin permiso", 403), "permiso"),
            (W.ErrorWeb("algo raro"), "otro"),
        ]
        for error, tipo in casos:
            info = E.clasificar(error)
            self.assertEqual(info["tipo"], tipo, str(error))
            self.assertTrue(info["titulo"] and info["explicacion"])

    def test_arrancando_explica_que_espere(self):
        self.assertIn("arrancando", E.clasificar(W.ErrorWeb("timed out"))["explicacion"])


class TestEstadoConexion(unittest.TestCase):
    def setUp(self):
        self.t = datetime(2026, 11, 18, 11, 30).timestamp()
        self.e = E.EstadoConexion(reloj=lambda: self.t)

    def test_sin_comprobar_al_principio(self):
        self.assertEqual(self.e.indicador(), ("● Catálogo Asignaciones: sin comprobar", "pendiente"))

    def test_avisa_una_vez_mientras_siga_sin_conexion(self):
        self.assertTrue(self.e.marcar_fallo(W.ErrorWeb("timed out")))
        self.t += 600
        self.assertFalse(self.e.marcar_fallo(W.ErrorWeb("timed out")))
        self.assertEqual(self.e.indicador(), ("● Catálogo Asignaciones: sin conexión desde las 11:30", "fallo"))

    def test_reintentar_permite_otro_aviso(self):
        self.e.marcar_fallo(W.ErrorWeb("x"))
        self.e.permitir_nuevo_aviso()
        self.assertTrue(self.e.marcar_fallo(W.ErrorWeb("x")))

    def test_al_conectar_se_reinicia(self):
        self.e.marcar_fallo(W.ErrorWeb("x"))
        self.assertTrue(self.e.marcar_ok())
        self.assertFalse(self.e.marcar_ok())               # sin cambios: no hace falta repintar
        self.assertEqual(self.e.indicador(), ("● Catálogo Asignaciones: conectado", "ok"))
        self.t += 3600
        self.assertTrue(self.e.marcar_fallo(W.ErrorWeb("x")))
        self.assertIn("12:30", self.e.indicador()[0])


class TestVentana(unittest.TestCase):
    def setUp(self):
        try:
            import tkinter as tk
            self.root = tk.Tk()
        except Exception:  # noqa: BLE001 — sin tkinter o sin pantalla
            self.skipTest("sin pantalla")
        self.root.withdraw()

    def tearDown(self):
        self.root.destroy()

    def _botones(self, v):
        import tkinter as tk
        vistos = []

        def recorrer(w):
            for h in w.winfo_children():
                if isinstance(h, tk.Button):
                    vistos.append(h)
                recorrer(h)
        recorrer(v)
        return vistos

    def test_con_reintentar(self):
        from app.ui.ventana_sin_conexion import VentanaSinConexion
        v = VentanaSinConexion(self.root, "traer lo cambiado en la web antes de la Acción 3",
                               E.clasificar(W.ErrorWeb("timed out")), "timed out")
        botones = self._botones(v)
        self.assertEqual([b.cget("text") for b in botones], ["↻  Reintentar", "Seguir sin la web"])
        botones[0].invoke()
        self.assertEqual(v.resultado, "reintentar")

    def test_sin_permiso_solo_entendido(self):
        from app.ui.ventana_sin_conexion import VentanaSinConexion
        v = VentanaSinConexion(self.root, "enviar a la web los cambios de la Consola",
                               E.clasificar(W.ErrorWeb("HTTP 403", 403)))
        botones = self._botones(v)
        self.assertEqual([b.cget("text") for b in botones], ["Entendido"])
        botones[0].invoke()
        self.assertEqual(v.resultado, "seguir")


if __name__ == "__main__":
    unittest.main()
