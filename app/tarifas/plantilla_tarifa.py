# -*- coding: utf-8 -*-
"""
plantilla_tarifa.py — escribe la plantilla de revisión de precios que se envía al proveedor (Acción 6, «Generar
plantilla Tarifa desde RINFO», 22 nov 2026, petición de Víctor). Sin tkinter.

Es una TARIFA final como la de la Acción 3 (mismo aspecto, cabecera con logo, Código, Proveedor, Periodo, Teléfono,
Dto. y el recuadro HOTELES; tabla en la fila 7) con los datos que ya tiene SAP, y dos columnas más en amarillo para el
proveedor: PRECIO NUEVO y FECHA INICIO (detrás de REGISTRO INFO, columna H). El «Periodo:» también va en amarillo, vacío, para el nuevo periodo. Como
lleva el nombre de las tarifas («PROVEEDOR [dd-mm-aaaa].xlsx»), lo que devuelva el proveedor se usa directamente en
«Crear / Actualizar Registro Info SAP» (datos_tarifa.precios_tarifa y periodo_propuesto).

Los datos los prepara app/comun/plantilla_rinfo.py (preparar_plantillas).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import os

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.comun.datos_tarifa import CAB_FECHA_INICIO, CAB_PRECIO_NUEVO, CAB_REGISTRO_INFO, escribir_huecos
from app.tarifas.transformar_tarifa import (
    _BORDE_TODO, _GRIS, _alin, _font, _insertar_logo_cabecera, _texto_visible,
)

CABECERAS = ["CODIGO SAP", "CODIGO DALI", "CODIGO PROVEEDOR", "DESCRIPCION", "FORMATO", "PRECIO", "DESCUENTO",
             CAB_REGISTRO_INFO, CAB_PRECIO_NUEVO, CAB_FECHA_INICIO]   # (22 nov 2026) REGISTRO INFO en H
AMARILLO = PatternFill(fill_type="solid", fgColor="FFF2A8")
AMARILLO_CAB = PatternFill(fill_type="solid", fgColor="FFD966")
NOTA = ("Indique el PRECIO NUEVO (sin IGIC y con los descuentos ya aplicados) y la FECHA INICIO en las columnas "
        "amarillas. Deje en blanco los artículos que no cambian.")


def _como_numero(v):
    try:
        return int(v) if str(v).strip().isdigit() else v
    except (TypeError, ValueError):
        return v


def escribir_plantilla(plantilla, carpeta_tarifas):
    """Escribe una plantilla (un elemento de plantilla_rinfo.preparar_plantillas) en carpeta_tarifas y devuelve su
    ruta. No pisa nada: el nombre ya viene libre, y si aun así existe, lanza FileExistsError."""
    ruta = os.path.join(carpeta_tarifas, plantilla["nombre_archivo"])
    if os.path.exists(ruta):
        raise FileExistsError(ruta)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TARIFA"

    # ── Cabecera (como transformar_tarifa, pasos 01-14) ──
    for fila, etiqueta, valor in ((1, "Código:", _como_numero(plantilla.get("codigo") or "")),
                                  (2, "Proveedor:", plantilla.get("proveedor") or ""),
                                  (3, "Periodo:", None)):
        ws.cell(fila, 2).value = etiqueta
        ws.cell(fila, 3).value = valor if valor != "" else None
    ws.cell(5, 2).value = "Dto.(%):"
    ws.cell(5, 3).value = "línea"
    ws.cell(5, 4).value = "general"
    for r in (1, 2, 3, 5):
        ws.cell(r, 2).font = _font(bold=True, size=11)
        ws.cell(r, 2).alignment = _alin(h="center", v="center")
    for r, c in ((1, 3), (2, 3), (3, 3), (5, 3), (5, 4)):
        ws.cell(r, c).font = _font(size=11)
        ws.cell(r, c).alignment = _alin(h="left", v="center")
    ws.cell(3, 3).fill = AMARILLO                      # nuevo periodo, lo rellena el proveedor
    ws.cell(3, 4).value = "← nuevo periodo"
    ws.cell(3, 4).font = Font(name="Arial", size=9, italic=True, color="7F6000")
    ws.merge_cells("E5:G5")
    ws.cell(5, 5).value = "DESCUENTOS YA APLICADOS EN PRECIO FINAL"
    ws.cell(5, 5).font = Font(name="Arial", size=11, bold=True, color="FF0000")
    ws.cell(5, 5).alignment = _alin(h="center", v="center")
    linea = Side(style="thin", color="000000")
    for fila in (1, 2, 3, 5):                          # borde exterior de B:D en cada fila, como la TARIFA final
        for col in range(2, 5):
            ws.cell(fila, col).border = Border(top=linea, bottom=linea, left=linea if col == 2 else Side(),
                                               right=linea if col == 4 else Side())
    escribir_huecos(ws, telefono=plantilla.get("telefono") or None, codigos_marcados=plantilla.get("centros") or ())
    _insertar_logo_cabecera(ws)

    # ── Tabla ──
    for c, cab in enumerate(CABECERAS, start=1):
        celda = ws.cell(7, c)
        celda.value = cab
        celda.font = _font(bold=True)
        celda.fill = AMARILLO_CAB if cab in (CAB_PRECIO_NUEVO, CAB_FECHA_INICIO) else _GRIS
        celda.border = _BORDE_TODO
        celda.alignment = _alin(h="center")
    fila = 8
    for f in plantilla["filas"]:
        valores = [_como_numero(f["sap"]), _como_numero(f["dali"]) if f["dali"] else None, f["cod_prov"] or None,
                   f["descripcion"] or None, f["formato"] or None, round(f["precio"], 4), None,
                   _como_numero(f.get("registro_info")) if f.get("registro_info") else None, None, None]
        for c, v in enumerate(valores, start=1):
            celda = ws.cell(fila, c)
            celda.value = v
            celda.font = _font()
            celda.border = _BORDE_TODO
            celda.alignment = _alin(h="left" if c == 4 else "center")
        ws.cell(fila, 6).number_format = '#,##0.00 "€"'
        ws.cell(fila, 2).number_format = "000000"
        ws.cell(fila, 9).fill = AMARILLO
        ws.cell(fila, 9).number_format = '#,##0.00 "€"'
        ws.cell(fila, 10).fill = AMARILLO
        ws.cell(fila, 10).number_format = "dd/mm/yyyy"
        fila += 1
    ultima = fila - 1
    ws.cell(fila + 1, 1).value = NOTA
    ws.cell(fila + 1, 1).font = Font(name="Arial", size=10, italic=True, color="7F6000")

    # ── Anchos, altos, filtros y paneles (como transformar_tarifa, paso 31) ──
    anchos_min = {1: 14, 2: 14, 3: 18, 4: 24}
    for c in range(1, 5):
        largo = max([len(_texto_visible(ws.cell(r, c).value, c)) for r in list(range(1, 6)) + list(range(8, ultima + 1))]
                    or [0])
        ws.column_dimensions[get_column_letter(c)].width = max(round(largo * 1.60) + 4, anchos_min[c])
    for letra, ancho in (("E", 19), ("F", 16), ("G", 16), ("H", 18), ("I", 18), ("J", 18)):
        ws.column_dimensions[letra].width = ancho
    for r in range(1, 6):
        ws.row_dimensions[r].height = 18
    ws.row_dimensions[6].height = 6
    ws.row_dimensions[7].height = 20
    ws.freeze_panes = "A8"
    ws.auto_filter.ref = f"A7:J{max(ultima, 7)}"
    ws.cell(7, 9).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wb.save(ruta)
    return ruta
