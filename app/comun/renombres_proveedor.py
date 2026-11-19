# -*- coding: utf-8 -*-
"""Cambios de nombre del proveedor en DALI (19 nov 2026, decisión de Víctor: opción A).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Caso real: DALI pasó «B REFRESHING COMPANY S.L. ( BEBIDAS REFRESCANTES CANARIAS SL )» a «B REFRESHING COMPANY S.L.».
Se comparaba solo el nombre, así que se tomó como un cambio de proveedor: se vaciaron el código de proveedor y el
precio de sus 12 artículos y la web dio de alta un proveedor nuevo.

Se considera **cambio de nombre** (el mismo proveedor) cuando:
  - todos los artículos que tenían el proveedor «antes» (y siguen en el catálogo) pasan a un mismo nombre «despues»;
  - ya no queda ningún artículo con «antes»;
  - «despues» no lo tenía antes ningún artículo (es un nombre nuevo).
Si solo cambian algunos artículos, o el nombre nuevo ya existía, es un cambio de proveedor de verdad, como siempre.

Sin tkinter. Pruebas en tests/test_renombres_proveedor.py.
"""


def normal(nombre):
    """Mismo criterio que el resto de la Consola: sin espacios de más y en mayúsculas ("" si está vacío)."""
    return " ".join(str(nombre).split()).upper() if nombre not in (None, "") else ""


def detectar(antes, despues):
    """`antes` y `despues`: {código DALI: nombre del proveedor asignado}.

    Devuelve [{"antes", "despues", "articulos"}] (nombres tal como vienen, sin espacios de los extremos), ordenado
    por el nombre de antes."""
    por_proveedor, nombre_antes = {}, {}
    for cod, prov in antes.items():
        n = normal(prov)
        if not n or cod not in despues:
            continue
        por_proveedor.setdefault(n, []).append(cod)
        nombre_antes.setdefault(n, str(prov).strip())
    quedan = {normal(p) for p in despues.values()}
    existian = {normal(p) for p in antes.values()}
    salida = []
    for n_antes, cods in por_proveedor.items():
        destinos = {normal(despues[c]) for c in cods}
        if len(destinos) != 1:
            continue
        n_despues = destinos.pop()
        if not n_despues or n_despues == n_antes or n_antes in quedan or n_despues in existian:
            continue
        nombre_despues = next(str(despues[c]).strip() for c in cods)
        salida.append({"antes": nombre_antes[n_antes], "despues": nombre_despues, "articulos": len(cods)})
    return sorted(salida, key=lambda r: r["antes"])


def aplicar_a(articulos, renombres):
    """Pone el nombre nuevo en una copia de artículos ({cod: {"proveedor": …}}), p. ej. la copia local de la web
    después de que la web renombre el proveedor (sus artículos no cambian, así que no vuelven en «cambios»)."""
    mapa = {normal(r["antes"]): r["despues"] for r in renombres}
    n = 0
    for art in articulos.values():
        nuevo = mapa.get(normal(art.get("proveedor")))
        if nuevo is not None:
            art["proveedor"] = nuevo
            n += 1
    return n


def es_renombre(renombres, antes, despues):
    """True si el paso de `antes` a `despues` es uno de los cambios de nombre detectados."""
    na, nd = normal(antes), normal(despues)
    return any(normal(r["antes"]) == na and normal(r["despues"]) == nd for r in renombres)
