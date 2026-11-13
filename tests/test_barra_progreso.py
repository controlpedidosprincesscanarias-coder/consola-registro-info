# -*- coding: utf-8 -*-
"""Pruebas de la barra de progreso (13 nov 2026): no se pone en verde mientras siga el envío a Catálogo
Asignaciones en segundo plano (Víctor: «termina en verde siempre sin realmente haber terminado todo el proceso»).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import tkinter as tk
    HAY_TK = True
except ImportError:          # pragma: no cover
    HAY_TK = False


@unittest.skipUnless(HAY_TK, "necesita tkinter")
class BarraHastaElFinal(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("sin pantalla")
        self.root.withdraw()
        from app.comun.tarifa_core import CanvasProgressBar
        self.b = CanvasProgressBar(self.root)
        self.b.pack()

    def tearDown(self):
        self.root.destroy()

    def _estado(self):
        return self.b._mode, self.b._lbl_estado.cget("text")

    def test_como_en_el_registro_acaba_lo_local_y_sigue_la_web(self):
        b = self.b
        b.indeterminate("Leyendo Productos.xlsx...")
        b.completado("Catálogo DALI actualizado")          # termina la parte local...
        b.ocupar("Enviando cambios a Catálogo Asignaciones…")  # ...y empieza el envío a la web
        self.assertEqual(self._estado(), ("indeterminate", "Enviando cambios a Catálogo Asignaciones…"))
        self.assertTrue(b.ocupada)
        b.liberar()                                          # acaba el envío: ahora sí, verde
        self.assertEqual(self._estado(), ("complete", "Catálogo DALI actualizado"))
        self.assertFalse(b.ocupada)

    def test_completado_mientras_se_envia_espera(self):
        b = self.b
        b.ocupar("Enviando cambios a Catálogo Asignaciones…")
        b.completado("Proceso finalizado")
        self.assertEqual(b._mode, "indeterminate")           # aún no
        b.ocupar("Enviando cambios a Catálogo Asignaciones…")   # otro envío encadenado
        b.liberar()
        self.assertEqual(b._mode, "indeterminate")           # queda uno
        b.liberar()
        self.assertEqual(self._estado(), ("complete", "Proceso finalizado"))

    def test_error_durante_el_envio_se_ve_al_acabar(self):
        b = self.b
        b.ocupar("Enviando…")
        b.reset("Error al guardar")
        self.assertEqual(b._mode, "indeterminate")
        b.liberar()
        self.assertEqual(self._estado(), ("determinate", "Error al guardar"))

    def test_sin_nada_en_segundo_plano_como_siempre(self):
        self.b.completado("Transformación completada")
        self.assertEqual(self._estado(), ("complete", "Transformación completada"))
        self.b.liberar()                                     # liberar de más no rompe nada
        self.assertEqual(self.b._ocupados, 0)


if __name__ == "__main__":
    unittest.main()
