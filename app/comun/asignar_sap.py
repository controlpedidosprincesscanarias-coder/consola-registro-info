# -*- coding: utf-8 -*-
"""
asignar_sap.py — lógica de la acción «Asignar SAP pendientes» (v10.2.56). Sin tkinter: se prueba en
tests/test_asignar_sap.py; la ventana está en app/ui/ventana_asignar_sap.py.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Petición de Víctor (13 nov 2026): una acción nueva, con su propio botón, parecida a la ventana de códigos SAP
faltantes de una tarifa, pero con TODOS los códigos DALI activos sin SAP del catálogo, para ir grabando los SAP en
el orden que se quiera. Un código que no esté en Productos.xlsx se guarda igualmente y se avisa de que, para
completar sus datos SAP, hay que descargar de nuevo Productos.xlsx y Listas de productos.xlsx en RECURSOS y volver
a pulsar la Acción 3.
"""
import unicodedata
from collections import defaultdict

from app.comun import catalogo_local

ACTIVOS = ("SI", "SÍ", "S", "TRUE", "1")


def _texto(v):
    return "" if v is None else str(v).strip()


def _plano(v):
    """Minúsculas y sin acentos, para buscar."""
    t = unicodedata.normalize("NFD", _texto(v).lower())
    return "".join(c for c in t if unicodedata.category(c) != "Mn")


def pendientes(ruta_catalogo):
    """Artículos ACTIVOS sin código SAP del catálogo, en el orden del archivo:
    [{dali, descripcion, unidad, proveedor, naturaleza, familia, subfamilia}]."""
    salida = []
    for clave, d in catalogo_local.leer_catalogo(ruta_catalogo):
        if _texto(d.get("activo")).upper() not in ACTIVOS:
            continue
        if _texto(d.get("cod_sap")):
            continue
        if not isinstance(clave, int):
            continue
        salida.append({
            "dali": clave,
            "descripcion": _texto(d.get("descripcion")),
            "unidad": _texto(d.get("unidad")),
            "proveedor": _texto(d.get("proveedor")),
            "naturaleza": _texto(d.get("naturaleza")),
            "familia": _texto(d.get("familia")),
            "subfamilia": _texto(d.get("subfamilia")),
        })
    return salida


def valores_filtro(lista, campo):
    """Valores distintos de un campo, ordenados (para los desplegables de filtro)."""
    return sorted({p[campo] for p in lista if p[campo]}, key=_plano)


def filtrar(lista, texto="", proveedor="", familia=""):
    """Los que contienen `texto` (en código DALI, descripción, proveedor, familia o subfamilia, sin distinguir
    mayúsculas ni acentos) y coinciden con el proveedor y la familia elegidos ("" = todos)."""
    t = _plano(texto)
    salida = []
    for p in lista:
        if proveedor and p["proveedor"] != proveedor:
            continue
        if familia and p["familia"] != familia:
            continue
        if t and t not in _plano(" ".join(str(p[k]) for k in ("dali", "descripcion", "proveedor", "familia",
                                                              "subfamilia"))):
            continue
        salida.append(p)
    return salida


def validar_codigo(texto):
    """Código SAP escrito por el usuario → int. Lanza ValueError con el motivo si no vale."""
    t = _texto(texto).replace(" ", "")
    if not t:
        raise ValueError("Escribe el código SAP.")
    if t.endswith(".0"):
        t = t[:-2]
    if not t.isdigit():
        raise ValueError(f"El código SAP tiene que ser un número entero (has escrito «{texto}»).")
    n = int(t)
    if n <= 0:
        raise ValueError("El código SAP tiene que ser mayor que cero.")
    return n


def info_productos(codigo, tabla_productos):
    """(descripción, unidad) del código en Productos.xlsx, o None si no está (o no hay Productos.xlsx)."""
    if not tabla_productos:
        return None
    datos = tabla_productos.get(str(codigo))
    if datos is None:
        return None
    return _texto(datos[0]), _texto(datos[2])


def calcular_conflictos(nuevos, sap_de, descripciones):
    """Códigos SAP que quedarían en más de un DALI: ya asignados a otro en el catálogo o repetidos en lo que se
    va a guardar. Misma forma que usa VentanaResolverSAPDuplicados:
    {sap: [(dali, es_existente, fila_o_None, descripcion)]}."""
    por_sap = defaultdict(list)
    for dali, sap in nuevos.items():
        por_sap[sap].append(dali)
    conflictos = {}
    for sap, dalis in por_sap.items():
        candidatos = []
        if sap in sap_de:
            dali_existente, fila = sap_de[sap]
            if dali_existente not in dalis:
                candidatos.append((dali_existente, True, fila, descripciones.get(dali_existente, "")))
        if len(dalis) > 1 or candidatos:
            for d in sorted(dalis):
                candidatos.append((d, False, None, descripciones.get(d, "")))
        if len(candidatos) > 1:
            conflictos[sap] = candidatos
    return conflictos


def aplicar(ws, mapa, fila_de, nuevos, conflictos, a_borrar):
    """Escribe en el catálogo abierto (catalogo_local.abrir_para_editar) los códigos de `nuevos`, tras resolver
    los conflictos: a los DALI de `a_borrar` que ya tenían ese SAP se les quita, y los que se estaban escribiendo
    no se graban. Devuelve (escritos {dali: sap}, limpiados [dali], no_encontrados [dali])."""
    col = mapa["cod_sap"]
    nuevos = dict(nuevos)
    limpiados = []
    for _sap, candidatos in conflictos.items():
        for dali, es_existente, fila, _desc in candidatos:
            if dali in a_borrar:
                if es_existente and fila:
                    ws.cell(row=fila, column=col).value = None
                    limpiados.append(dali)
                nuevos.pop(dali, None)
    escritos, no_encontrados = {}, []
    for dali, sap in nuevos.items():
        fila = fila_de.get(dali)
        if fila is None:
            no_encontrados.append(dali)
            continue
        ws.cell(row=fila, column=col).value = sap
        escritos[dali] = sap
    return escritos, limpiados, no_encontrados


def fuera_de_productos(escritos, tabla_productos):
    """Los DALI guardados cuyo código SAP no está en Productos.xlsx (todos, si no hay Productos.xlsx)."""
    return sorted(d for d, s in escritos.items() if info_productos(s, tabla_productos) is None)
