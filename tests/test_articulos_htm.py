# -*- coding: utf-8 -*-
"""Pruebas de la Acción 1 con ARTICULOS.htm (v10.2.74): app/acciones/articulos_htm.py y su uso en
app/acciones/organizar_dali.py. Datos inventados con la forma del informe real de DALI.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).
"""
import os
import sys
import tempfile
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.acciones import articulos_htm as H  # noqa: E402

LEFTS = ("0.02083331", "0.9208333", "1.820833", "2.720833", "3.620833", "4.520833", "5.420834", "6.320834",
         "7.220833")


def _span(top, left, texto):
    return (f'<span style="position:absolute;top:{top}in;left:{left}in;width:0.8583333in;height:0.3353882in;'
            f'font-family:Arial;font-size:7pt;color:#000000;vertical-align:top;"><nobr>{texto}</nobr></span>\n')


def _pagina(filas, salto=True):
    """filas: listas de 9 textos ya en HTML (None = celda vacía, que el informe no escribe)."""
    estilo = "page-break-inside:avoid;page-break-after:always;" if salto else "page-break-inside:avoid;"
    trozos = [f'<div style="{estilo}"><div style="position:relative;width:7.47in;height:10.49in;">\n']
    for i, fila in enumerate(filas):
        top = f"{0.02083331 + i * 0.3770555:.7f}"
        trozos += [_span(top, left, t) for left, t in zip(LEFTS, fila) if t is not None]
    return "".join(trozos)


def informe(paginas, total=None):
    cuerpo = "".join(p + "</div></div>" for p in paginas)
    pie = ""
    if total is not None:
        pie = ('<div style="page-break-inside:avoid;"><div style="position:relative;width:7.47in;">'
               '<span style="position:absolute;top:6.137729in;left:6.638889in;width:0.74in;height:0.16in;'
               f'font-family:Arial;"><nobr>{total}</nobr></span>'
               '<span style="position:absolute;top:6.137729in;left:4.951389in;width:1.66in;height:0.16in;'
               'font-family:Arial;"><nobr>Total&nbsp;L&#237;neas:&nbsp;</nobr></span></div></div>')
    return ('<html><head><title>ActiveReports Document</title><meta HTTP-EQUIV="Content-Type" '
            'CONTENT="text/html; charset=utf-8" /></head><body>' + cuerpo + pie + "</body></html>")


AGUA = ["00000001", "AGUA&nbsp;GARRAFA&nbsp;<br/>5L", "UNIDAD", "BEBIDAS&nbsp;NO&nbsp;<br/>ALCOHOLICAS&nbsp;",
        "AGUAS", "BEBIDAS", "0&nbsp;Exento", "AGUAS&nbsp;DE&nbsp;<br/>TEROR&nbsp;SA", "Fa<br/>lse"]
VITTEL = ["00000002", "AGUA&nbsp;VITTEL&nbsp;75&nbsp;CL", "UNIDAD", "BEBIDAS&nbsp;NO&nbsp;<br/>ALCOHOLICAS&nbsp;",
          "AGUAS", "BEBIDAS", "0&nbsp;Exento", None, "True"]                     # sin proveedor: no hay celda
SMOOTHIE = ["00000049", "SMOOTHIES&nbsp;PI&#209;A-PLATANO-<br/>MANZANA&nbsp;110&nbsp;ML", "UNIDAD",
            "BEBIDAS&nbsp;NO&nbsp;<br/>ALCOHOLICAS", "ZUMOS", "BEBIDAS", "03&nbsp;&nbsp;reducido", "AHEMBO,SL",
            "Fa<br/>lse"]


class Lector(unittest.TestCase):
    def test_lee_filas_columnas_y_cortes_de_linea(self):
        filas, total = H.leer_contenido(informe([_pagina([AGUA, VITTEL]), _pagina([SMOOTHIE])], total=3))
        self.assertEqual(total, 3)
        self.assertEqual(filas[0], ["00000001", "AGUA GARRAFA 5L", "UNIDAD", "BEBIDAS NO ALCOHOLICAS ", "AGUAS",
                                    "BEBIDAS", "0 Exento", "AGUAS DE TEROR SA", "False"])
        self.assertEqual(filas[1][7], "")                                          # celda vacía
        self.assertEqual(filas[2][1], "SMOOTHIES PIÑA-PLATANO-MANZANA 110 ML")     # sin el espacio del .xls
        self.assertEqual(len(filas), 3)

    def test_pagina_sin_salto_no_pisa_la_primera_fila_de_la_siguiente(self):
        # Caso real: no todas las páginas llevan «page-break-after»; las dos empiezan con la misma posición
        contenido = informe([_pagina([AGUA], salto=False), _pagina([VITTEL]), _pagina([SMOOTHIE])], total=3)
        self.assertEqual([f[0] for f in H.leer_contenido(contenido)[0]], ["00000001", "00000002", "00000049"])

    def test_sin_total_y_formato_no_valido(self):
        self.assertIsNone(H.leer_contenido(informe([_pagina([AGUA, VITTEL])]))[1])
        ocho = [f[:8] for f in (AGUA, VITTEL)]
        with self.assertRaises(H.FormatoHtmNoValido) as e:
            H.leer_contenido(informe([_pagina(ocho)]))
        self.assertIn("8 columnas", str(e.exception))
        with self.assertRaises(H.FormatoHtmNoValido):
            H.leer_contenido("<html><body>nada</body></html>")

    def test_forma_de_los_xls(self):
        filas = H.leer_contenido(informe([_pagina([AGUA])]))[0]
        f1, f2 = H.tablas_como_xls(filas)
        self.assertEqual(f1[0], ["00000001", "AGUA GARRAFA 5L", "UNIDAD", "BEBIDAS NO ALCOHOLICAS ", "AGUAS",
                                 "BEBIDAS"])
        self.assertEqual([f2[0][i] for i in (0, 3, 5, 7)], ["00000001", "0 Exento", "AGUAS DE TEROR SA", "False"])


class Eleccion(unittest.TestCase):
    def test_el_mas_reciente(self):
        with tempfile.TemporaryDirectory() as d:
            x1, x2, h = (os.path.join(d, n) for n in ("ARTICULOS1.xls", "ARTICULOS2.xls", H.NOMBRE_HTM))
            self.assertEqual(H.elegir_origen(x1, x2, h)[0], None)
            open(h, "w").close()
            self.assertEqual(H.elegir_origen(x1, x2, h)[0], H.HTM)
            open(x1, "w").close()
            self.assertEqual(H.elegir_origen(x1, x2, h)[0], H.HTM)                 # falta ARTICULOS2
            open(x2, "w").close()
            ahora = time.time()
            os.utime(h, (ahora - 100, ahora - 100))
            os.utime(x1, (ahora - 200, ahora - 200))
            os.utime(x2, (ahora - 50, ahora - 50))                                 # el más nuevo de los .xls
            self.assertEqual(H.elegir_origen(x1, x2, h)[0], H.XLS)
            os.utime(h, (ahora, ahora))
            origen, texto = H.elegir_origen(x1, x2, h)
            self.assertEqual(origen, H.HTM)
            self.assertIn("más reciente", texto)
            os.utime(h, (ahora - 50, ahora - 50))                                  # misma fecha: como hasta ahora
            self.assertEqual(H.elegir_origen(x1, x2, h)[0], H.XLS)


class AccionUno(unittest.TestCase):
    """organizar_dali con ARTICULOS.htm: el mismo resultado que daba con los .xls."""

    def test_procesar_datos_dali_desde_htm(self):
        from app.acciones import organizar_dali
        mensajes = []
        with tempfile.TemporaryDirectory() as d:
            ruta = os.path.join(d, H.NOMBRE_HTM)
            with open(ruta, "w", encoding="utf-8") as fh:
                fh.write(informe([_pagina([SMOOTHIE, AGUA]), _pagina([VITTEL])], total=3))
            r = organizar_dali.procesar_datos_dali("no-existe-1.xls", "no-existe-2.xls",
                                                   lambda m, n="INFO": mensajes.append((m, n)), ruta_htm=ruta)
        df = r["df_final"]
        self.assertEqual(list(df["CODIGO"]), [1, 2, 49])                            # ordenado por código
        fila = df.iloc[0].to_dict()
        self.assertEqual(fila, {"CODIGO": 1, "DESCRIPCION": "AGUA GARRAFA 5L", "UNIDAD": "UNIDAD",
                                "FAMILIA": "BEBIDAS NO ALCOHOLICAS", "SUBFAMILIA": "AGUAS",
                                "NATURALEZA": "BEBIDAS", "IGIC": "0 Exento", "PROVEEDOR": "AGUAS DE TEROR SA",
                                "ACTIVO": "SI"})
        self.assertEqual(df.iloc[1]["ACTIVO"], "NO")                                # True → dado de baja
        self.assertEqual(df.iloc[1]["PROVEEDOR"], "")
        self.assertEqual(df.iloc[2]["IGIC"], "03 reducido")
        self.assertEqual(r["ultimo_cod_art1"], 49)
        self.assertEqual(r["errores_sin_igic"], 0)
        self.assertTrue(any("coincide con «Total Líneas»" in m for m, _ in mensajes))

    def test_avisa_si_no_cuadra_el_total(self):
        from app.acciones import organizar_dali
        mensajes = []
        with tempfile.TemporaryDirectory() as d:
            ruta = os.path.join(d, H.NOMBRE_HTM)
            with open(ruta, "w", encoding="utf-8") as fh:
                fh.write(informe([_pagina([AGUA])], total=5))
            organizar_dali.procesar_datos_dali("", "", lambda m, n="INFO": mensajes.append((m, n)), ruta_htm=ruta)
        self.assertIn(("  └ ⚠ 1 artículos leídos, pero el informe dice «Total Líneas: 5»", "WARN"), mensajes)

    def test_se_borran_las_tres_descargas(self):
        from app.acciones import organizar_dali
        with tempfile.TemporaryDirectory() as d:
            rutas = [os.path.join(d, n) for n in ("ARTICULOS1.xls", "ARTICULOS2.xls", H.NOMBRE_HTM)]
            for r in rutas:
                open(r, "w").close()
            borrados = organizar_dali.eliminar_archivos_origen(rutas[0], rutas[1], lambda *a: None,
                                                               ruta_htm=rutas[2])
            self.assertEqual(borrados, ["ARTICULOS1.xls", "ARTICULOS2.xls", H.NOMBRE_HTM])
            self.assertFalse(any(os.path.exists(r) for r in rutas))


if __name__ == "__main__":
    unittest.main()
