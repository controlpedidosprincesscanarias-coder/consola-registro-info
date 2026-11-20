# -*- coding: utf-8 -*-
"""Tarifas [INCOMPLETO] de la carpeta TARIFAS (20 nov 2026, v10.2.86).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Las usa la Acción 4 (antes 6, «Actualizar manualmente archivos [INCOMPLETOS]») para listarlas, y la ventana
principal para dejar ese botón sombreado mientras no haya ninguna (petición de Víctor, 20 nov 2026).
Mismo criterio que tenía la Acción 6: «PROVEEDOR [INCOMPLETO] [dd-mm-aaaa].xlsx», sin los RINFO.
Sin tkinter. Pruebas en tests/test_incompletos.py.
"""
import os
import re
from datetime import datetime

_RE_TARIFA = re.compile(r".+\[\d{2}-\d{2}-\d{4}\].*\.xlsx$", re.IGNORECASE)


def es_incompleto(nombre: str) -> bool:
    nu = str(nombre).upper()
    return bool(_RE_TARIFA.match(str(nombre))) and not nu.startswith("RINFO") and "[INCOMPLETO]" in nu


def listar(carpeta_tarifas: str) -> list:
    """[{nombre, ruta, fecha}] de las tarifas [INCOMPLETO], por nombre. Sin carpeta o sin permiso: []."""
    salida = []
    try:
        entradas = sorted(os.scandir(carpeta_tarifas), key=lambda e: e.name.lower())
    except (OSError, TypeError):
        return salida
    for entry in entradas:
        try:
            if not entry.is_file() or not es_incompleto(entry.name):
                continue
            fecha = datetime.fromtimestamp(entry.stat().st_mtime).strftime("%d-%m-%Y %H:%M")
        except OSError:
            continue
        salida.append({"nombre": entry.name, "ruta": entry.path, "fecha": fecha})
    return salida


def hay(carpeta_tarifas: str) -> bool:
    """True si en TARIFAS hay al menos una tarifa [INCOMPLETO]."""
    try:
        with os.scandir(carpeta_tarifas) as it:
            return any(e.is_file() and es_incompleto(e.name) for e in it)
    except (OSError, TypeError):
        return False
