# -*- coding: utf-8 -*-
"""
listados_sap.py — aviso «faltan los listados de SAP actualizados» tras asignar códigos SAP (13 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Petición de Víctor: cuando se asignan códigos SAP (Acción 8 o la ventana de códigos SAP faltantes de la Acción 4),
la Consola envía los códigos a Catálogo Asignaciones, pero los datos SAP de un material nuevo (descripción, unidad,
grupos, categoría de valoración) solo llegan con Productos.xlsx y Listas de productos.xlsx descargados de SAP
después de crearlo. La Consola tiene que avisarlo, guiar la descarga y, cuando estén en RECURSOS, lanzar sola la
Acción 3; si no se hace ahora, la web queda parcialmente actualizada hasta entonces.

Sin tkinter: la ventana está en app/ui/ventana_listados_sap.py. Pruebas en tests/test_listados_sap.py.
"""
import os

PRODUCTOS = "Productos.xlsx"
LISTAS = "Listas de productos.xlsx"

# Cómo se descarga cada listado (aplicación de SAP y nombre con el que se exporta)
LISTADOS = (
    {"archivo": PRODUCTOS, "app_sap": "Gestionar datos maestros de producto", "nombre": "Productos",
     "imagen": "assets/guia_sap_exportar_productos.png"},
    {"archivo": LISTAS, "app_sap": "Lista de productos", "nombre": "Listas de productos",
     "imagen": "assets/guia_sap_exportar_listas.png"},
)

PASOS = (
    "En SAP, abre la aplicación del listado (imagen de arriba).",
    "Pulsa «Ir» para cargar la tabla y, a la derecha de su barra, el botón ⌄ de exportar (el último, «Abrir menú») "
    "→ «Exportar como…».",
    "Deja el nombre de fichero que se indica y el formato «Microsoft Excel (*.xlsx)», y pulsa «Exportar».",
    "Guarda el archivo descargado en la carpeta RECURSOS del proyecto, sustituyendo al anterior.",
)

AVISO_PARCIAL = ("La web (Catálogo Asignaciones) queda PARCIALMENTE actualizada: ya tiene los códigos SAP, pero no "
                 "sus datos SAP (descripción, unidad, grupos, categoría de valoración) hasta que descargues "
                 "Productos.xlsx y Listas de productos.xlsx en RECURSOS y pulses la Acción 3.")


def rutas(carpeta_recursos):
    """{nombre de archivo: ruta} de los dos listados en RECURSOS."""
    return {L["archivo"]: os.path.join(carpeta_recursos, L["archivo"]) for L in LISTADOS}


def codigos_sin_datos(codigos_sap, tabla_productos):
    """Los códigos SAP (como texto, sin repetir y en orden) que no están en Productos.xlsx: los que se quedarán sin
    sus datos SAP hasta descargar los listados actualizados. Sin Productos.xlsx (tabla None o vacía), todos."""
    vistos, fuera = set(), []
    for c in codigos_sap:
        txt = str(c).strip()
        if not txt or txt in vistos:
            continue
        vistos.add(txt)
        if not tabla_productos or txt not in tabla_productos:
            fuera.append(txt)
    return fuera


def estado_archivos(carpeta_recursos, desde):
    """Para cada listado: (existe, actualizado) — actualizado = modificado después de `desde` (timestamp), es decir,
    descargado de nuevo desde que salió el aviso."""
    estado = {}
    for nombre, ruta in rutas(carpeta_recursos).items():
        try:
            mtime = os.path.getmtime(ruta)
            estado[nombre] = (True, mtime >= desde)
        except OSError:
            estado[nombre] = (False, False)
    return estado


def sin_actualizar(estado):
    """Nombres de los listados que faltan o no se han descargado de nuevo."""
    return [n for n, (existe, actualizado) in estado.items() if not (existe and actualizado)]
