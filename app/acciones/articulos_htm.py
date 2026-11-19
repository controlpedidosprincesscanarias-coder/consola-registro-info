# -*- coding: utf-8 -*-
"""
articulos_htm.py — la Acción 1 también con ARTICULOS.htm (v10.2.74, petición de Víctor, 19 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Petición de Víctor: «cuando se pulsa la Acción 1, actualmente la aplicación busca los Excel ARTICULOS1 y 2 para
trabajar; necesito que también se pueda realizar la misma operación con ARTICULOS.htm».

ARTICULOS.htm es el listado de artículos de DALI exportado como informe (ActiveReports): una página tras otra
(`<div style="position:relative;…">`) con un `<span style="position:absolute;top:…in;left:…in;…">` por celda. Cada
fila es un artículo (mismo `top` dentro de su página) y cada columna, una posición `left`. Trae en UN solo archivo las
9 columnas que la Acción 1 saca de ARTICULOS1.xls y ARTICULOS2.xls, en este orden:
    código, descripción, unidad, familia, subfamilia, naturaleza, impuesto (IGIC), proveedor, «No permitir compra».
Al final, «Total Líneas: N», que se usa para comprobar que se han leído todos.

Los textos largos se parten en varias líneas con `<br/>`: si el corte es entre palabras, el espacio va antes
(«AGUA&nbsp;GARRAFA&nbsp;<br/>5L»); si no, se parte la palabra («Fa<br/>lse»). Por eso `<br/>` se quita sin poner nada.
Comprobado con la descarga real del 19 nov 2026 (30.246 artículos): sale lo mismo que en ARTICULOS1/2.xls, salvo los
espacios que DALI mete al pasar a Excel donde partía la línea («PLATANO- MANZANA», «Fals e»), que en el .htm no están.

Ojo: no todas las páginas llevan «page-break-after»; las páginas se separan por su `<div style="position:relative`.

Sin tkinter. Pruebas en tests/test_articulos_htm.py.
"""
import html
import os
import re

NOMBRE_HTM = "ARTICULOS.htm"
COLUMNAS = ("CODIGO", "DESCRIPCION", "UNIDAD", "FAMILIA", "SUBFAMILIA", "NATURALEZA", "IGIC", "PROVEEDOR",
            "ACTIVO")

_TOKEN = re.compile(
    r'(?P<pagina><div style="position:relative;)'
    r'|<span style="[^"]*?\btop:(?P<top>[\d.]+)in;left:(?P<left>[\d.]+)in;[^"]*">(?P<texto>.*?)</span>',
    re.S)


class FormatoHtmNoValido(ValueError):
    pass


def _texto(crudo):
    t = re.sub(r"</?nobr>", "", crudo)
    t = re.sub(r"<br\s*/?>", "", t)
    t = re.sub(r"<[^>]+>", "", t)
    return html.unescape(t).replace("\xa0", " ")


def _celdas_por_pagina(contenido):
    """[[(top, left, texto)] por página] en el orden del archivo."""
    paginas, actual = [], None
    for m in _TOKEN.finditer(contenido):
        if m.group("pagina"):
            actual = []
            paginas.append(actual)
        elif actual is not None:
            actual.append((round(float(m.group("top")), 3), round(float(m.group("left")), 2),
                           _texto(m.group("texto"))))
    return paginas


def _total_declarado(paginas):
    """El número de «Total Líneas: N» del final del informe (la celda numérica de la misma fila que el texto),
    o None si no está."""
    for celdas in reversed(paginas):
        for top, _left, texto in celdas:
            if re.match(r"\s*Total\s+L[ií]neas", texto):
                numeros = [t.strip() for tp, _l, t in celdas if tp == top and t.strip().isdigit()]
                return int(numeros[0]) if numeros else None
    return None


def leer_contenido(contenido):
    """Del texto de ARTICULOS.htm, (filas, total): filas = [[9 textos]] en el orden del informe (una por
    artículo); total = «Total Líneas» o None. Lanza FormatoHtmNoValido si no tiene la forma esperada."""
    paginas = _celdas_por_pagina(contenido)
    filas_crudas = []
    for celdas in paginas:
        por_top = {}
        for top, left, texto in celdas:
            por_top.setdefault(top, {})[left] = texto
        for top in sorted(por_top):
            filas_crudas.append(por_top[top])
    if not filas_crudas:
        raise FormatoHtmNoValido("ARTICULOS.htm no tiene el formato del informe de DALI (no se encontró ninguna "
                                 "celda).")
    # La columna del código es la de más a la izquierda; un artículo es una fila con un número en ella.
    izquierda = min(min(f) for f in filas_crudas)
    articulos = [f for f in filas_crudas if f.get(izquierda, "").strip().isdigit()]
    if not articulos:
        raise FormatoHtmNoValido("ARTICULOS.htm no tiene códigos de artículo en la primera columna.")
    # Columnas = las posiciones que aparecen en una parte importante de los artículos (las celdas vacías no salen)
    cuenta = {}
    for f in articulos:
        for left in f:
            cuenta[left] = cuenta.get(left, 0) + 1
    columnas = sorted(left for left, n in cuenta.items() if n >= max(1, len(articulos) // 20))
    if len(columnas) != len(COLUMNAS):
        raise FormatoHtmNoValido(
            f"ARTICULOS.htm tiene {len(columnas)} columnas y se esperan {len(COLUMNAS)} (código, descripción, "
            "unidad, familia, subfamilia, naturaleza, impuesto compra, proveedor por defecto y no permitir "
            "compra). Descárgalo de DALI con esas columnas y en ese orden.")
    filas = [[f.get(c, "") for c in columnas] for f in articulos]
    return filas, _total_declarado(paginas)


def leer_articulos_htm(ruta):
    """Lee ARTICULOS.htm del disco: (filas, total). Ver leer_contenido."""
    with open(ruta, encoding="utf-8", errors="replace") as fh:
        return leer_contenido(fh.read())


def tablas_como_xls(filas):
    """Las filas del .htm con la MISMA forma que tenían ARTICULOS1.xls y ARTICULOS2.xls al leerlos
    (procesar_datos_dali), para no tocar el resto de la Acción 1:
      ARTICULOS1: columnas 0..5 = código, descripción, unidad, familia, subfamilia, naturaleza;
      ARTICULOS2: columnas 0, 3, 5 y 7 = código, IGIC, proveedor, no permitir compra (las demás, vacías).
    Devuelve (filas1, filas2) como listas de listas de texto."""
    filas1 = [f[0:6] for f in filas]
    filas2 = [[f[0], f[1], f[2], f[6], "", f[7], "", f[8]] for f in filas]
    return filas1, filas2


# ── Qué descarga usar (decisión de Víctor, 19 nov 2026: la más reciente) ───────────────────────────────────────────
XLS = "xls"
HTM = "htm"


def _fecha(ruta):
    try:
        return os.path.getmtime(ruta)
    except OSError:
        return None


def elegir_origen(file1, file2, ruta_htm):
    """Qué usar en la Acción 1: XLS (ARTICULOS1 y 2), HTM (ARTICULOS.htm) o None (no hay ninguna descarga
    completa). Con las dos, la más reciente (la fecha de los .xls es la del más nuevo de los dos); a igual fecha,
    los .xls, como hasta ahora. Devuelve (origen, texto para la consola)."""
    f1, f2, fh = _fecha(file1), _fecha(file2), _fecha(ruta_htm)
    hay_xls = f1 is not None and f2 is not None
    hay_htm = fh is not None
    if hay_htm and not hay_xls:
        return HTM, f"Se usa {NOMBRE_HTM}."
    if hay_xls and not hay_htm:
        return XLS, "Se usan ARTICULOS1.xls y ARTICULOS2.xls."
    if not hay_xls and not hay_htm:
        return None, ""
    if fh > max(f1, f2):
        return HTM, f"Se usa {NOMBRE_HTM}, que es más reciente que ARTICULOS1/2.xls."
    return XLS, f"Se usan ARTICULOS1.xls y ARTICULOS2.xls, más recientes que {NOMBRE_HTM}."
