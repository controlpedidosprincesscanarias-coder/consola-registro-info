# -*- coding: utf-8 -*-
"""
transformar_tarifa.py  —  V10.1
Lee TARIFA.xls y genera un .xlsx nuevo siguiendo las 31 pautas oficiales.
Elimina el .xls original tras generar el .xlsx correctamente.

CAMBIOS V10.1:
  - Detección automática de formato DALI (A: col D, B: col E).
  - Nueva función _detectar_formato() — sin cambios en la lógica de salida.

CAMBIOS V10.0:
  - Al finalizar: cuenta duplicados SAP/DALI (únicos) y detecta DALI sin SAP.
  - Devuelve (ruta_salida, _analisis) en lugar de solo ruta_salida.
  - _analisis = {dup_sap, dup_dali, faltantes_sap}.

CAMBIOS V9.7:
  - Precio 0 con descuentos activos también se marca en rojo (_cero_con_dto).
  - Logo centrado horizontalmente en col A: se pre-calcula el ancho de col A
    antes del autoajuste y se aplica colOff = (ancho_col_A - 95px) / 2 en el
    AnchorMarker _from, manteniendo el ancho de imagen en 95px.
CAMBIOS V9.6:
  - Nueva cabecera: logo en A1:A6, etiquetas en B, valores en C/D.
  - Eliminadas columnas A1:A6 de etiquetas (ya no se usan).
  - Dto. línea → fila 4 (C4="línea", D4=valor); Dto. general → fila 6 (C6="general", D6=valor).
  - Periodo → C5; Código → C1; Proveedor → C2.
  - Logo Princess insertado encuadrado en A1:A6.
CAMBIOS V9.1:
  - Normalización de espacios múltiples → un solo espacio en todos los textos.
  - Autoajuste de columnas mejorado: usa la cadena *formateada* (€, %, 00000000).
  - Fecha del nombre de archivo entre corchetes: [dd-mm-yyyy].
  - Eliminación del TARIFA.xls original al finalizar correctamente.

NUEVO MAPEO CABECERA (filas 1-6 del xlsx):
  A1:A5  → Logo Princess (mitad A1 → mitad A5)
  B1     → "Código:"      C1 → D2 xls  (código proveedor)
  B2     → "Proveedor:"   C2 → D4 xls  (nombre proveedor)
  B3     → "Periodo:"     C3 → C17 xls (período)
  Fila 4 → vacía (alto 6, separador fino)
  B5     → "Dto.(%):"     C5=D8&" línea"  D5=D10&" general"  E5/F5 vacíos
  Fila 6 → vacía (separador)
  Fila 7 → cabecera tabla
  Fila 8 → datos

  Datos desde fila 24 del XLS:
    Col A → CODIGO DALI   → B nuevo
    Col B → COD PROVEEDOR → C nuevo
    Col D → DESCRIPCION   → D nuevo
    Col F → FORMATO       → E nuevo
    Col G → PRECIO        → F nuevo
    Col H → DESCUENTO     → G nuevo
    Col A(CODIGO SAP) queda vacía

Autor: Víctor Alejandro Martín Andrade (VAMA) © 2026
"""

import os
import re
from datetime import datetime
import xlrd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from openpyxl.utils import get_column_letter


# ============================================================
#  LEER XLS  (xlrd para .xls legacy)
# ============================================================

def _leer_xls(ruta_xls: str):
    """Devuelve la hoja activa del .xls como objeto xlrd."""
    book = xlrd.open_workbook(ruta_xls, formatting_info=False)
    return book.sheet_by_index(0)


def _celda(ws_xls, fila_1based: int, col_1based: int):
    """Lee una celda del .xls (indexación 1-based como Excel)."""
    r = fila_1based - 1
    c = col_1based - 1
    if r >= ws_xls.nrows or c >= ws_xls.ncols:
        return ""
    v = ws_xls.cell(r, c).value
    if v == "" or v is None:
        return ""
    # Convertir floats enteros a int para evitar "64.0" en vez de "64"
    if isinstance(v, float) and v == int(v):
        return int(v)
    # Normalizar espacios múltiples en cadenas de texto
    if isinstance(v, str):
        v = _norm_espacios(v)
    return v


# ============================================================
#  UTILIDADES
# ============================================================

def _norm_espacios(texto) -> str:
    """Colapsa cualquier secuencia de espacios en blanco a un solo espacio."""
    if not isinstance(texto, str):
        return texto
    return re.sub(r"[ \t]+", " ", texto).strip()


def _log(log_fn, msg, nivel="INFO"):
    if log_fn:
        log_fn(msg, nivel)


def _to_float(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    txt = str(v).replace(" ", "").replace("€", "").replace("¤", "")
    if "," in txt and "." in txt:
        txt = txt.replace(".", "").replace(",", ".")
    elif "," in txt:
        txt = txt.replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return None


def _ultima_fila(ws, desde=9) -> int:
    for r in range(ws.max_row, desde - 1, -1):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, 8)):
            return r
    return desde


def _texto_visible(valor, col: int) -> str:
    """
    Devuelve la cadena que Excel mostraría para estimar el ancho real de columna.
    Tiene en cuenta los formatos aplicados en cada columna:
      Col 2 (B) → CODIGO DALI    formato 000000     → 6 dígitos mínimo
      Col 6 (F) → PRECIO         formato #,##0.00 € → ej. "1.234,56 €"
      Col 7 (G) → DESCUENTO      formato 0%         → ej. "5%"
    Para el resto, convierte a str directamente.
    """
    if valor is None or valor == "":
        return ""
    if col == 2:                        # CODIGO DALI: siempre 6 chars mínimo
        try:
            return str(int(float(valor))).zfill(6)
        except (ValueError, TypeError):
            return str(valor)
    if col == 6:                        # PRECIO: simular formato europeo con €
        try:
            f = float(valor)
            # Formato tipo "1.234,56 €"
            entero = int(abs(f))
            dec = round(abs(f) % 1 * 100)
            ent_fmt = f"{entero:,}".replace(",", ".")
            return f"{'-' if f < 0 else ''}{ent_fmt},{dec:02d} €"
        except (ValueError, TypeError):
            return str(valor)
    if col == 7:                        # DESCUENTO: simular formato 0%
        try:
            # El valor ya está almacenado como fracción (0.05) en el xlsx
            pct = round(float(valor) * 100)
            return f"{pct}%"
        except (ValueError, TypeError):
            return str(valor)
    return _norm_espacios(str(valor))


# Estilos reutilizables
_LINEA      = Side(style="thin", color="000000")
_BORDE_TODO = Border(left=_LINEA, right=_LINEA, top=_LINEA, bottom=_LINEA)
_SIN_FILL   = PatternFill(fill_type=None)
_GRIS       = PatternFill(fill_type="solid", fgColor="D9D9D9")
_AMARILLO   = PatternFill(fill_type="solid", fgColor="FFFF00")

def _font(bold=False, size=12):
    return Font(name="Arial", size=size, bold=bold)

def _alin(h="general", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


# ============================================================
#  TRANSFORMACIÓN PRINCIPAL  (31 pasos)
# ============================================================

def _detectar_formato(src) -> dict:
    """
    Detecta cuál de los dos formatos DALI tiene el archivo .xls y devuelve
    un diccionario con los números de columna (1-based) para cada campo.

    Formato A (ej. MAHOU):
      - Cabecera: código/nombre/dto en col D (col 4)
      - Datos: Descripción=col D(4), Formato=col F(6), Precio=col G(7), Dto=col H(8)

    Formato B (ej. SERV_ELECT):
      - Cabecera: código/nombre/dto en col E (col 5)
      - Datos: Descripción=col E(5), Formato=col G(7), Precio=col H(8), Dto=col I(9)

    Detección: si la celda F2 (fila 2, col 4) contiene un valor numérico → Formato A.
    Si está vacía y F2 col 5 tiene valor → Formato B.
    """
    val_col4 = _celda(src, 2, 4)
    val_col5 = _celda(src, 2, 5)

    if val_col4 not in ("", None):
        return {
            "nombre": "A",
            "col_cab":  4,   # col D: código prov, nombre, dtos cabecera
            "col_desc": 4,   # col D: Descripción en filas de datos
            "col_fmt":  6,   # col F: Formato
            "col_prec": 7,   # col G: Precio
            "col_dto":  8,   # col H: Descuento
        }
    else:
        return {
            "nombre": "B",
            "col_cab":  5,   # col E: código prov, nombre, dtos cabecera
            "col_desc": 5,   # col E: Descripción en filas de datos
            "col_fmt":  7,   # col G: Formato
            "col_prec": 8,   # col H: Precio
            "col_dto":  9,   # col I: Descuento
        }


def _insertar_logo_cabecera(ws, log_fn=None) -> None:
    """
    Paso 14b (extraido de transformar()): inserta el logo Princess en
    A1:A5 de la cabecera nueva, centrado en la columna A. El logo va
    embebido en base64 dentro de esta funcion (no depende de ficheros
    externos, funciona igual en .py y en el .exe compilado).

    No devuelve nada: escribe directamente sobre ws (mismo objeto que
    el resto de transformar()). Si Pillow no esta disponible o falla
    la insercion por cualquier motivo, deja constancia en el log y
    continua sin logo -- nunca interrumpe la generacion de la tarifa.
    """
    # ── PASO 14b: Logo en A1:A6 ──────────────────────────────
    #
    #  Prioridad de búsqueda del logo:
    #    1. assets/logo.png  junto al script (o al .exe)
    #    2. assets/logo.png  junto al TARIFA.xls (carpeta TARIFAS)
    #    3. Fallback: rectángulo dorado con texto "PRINCESS" generado con Pillow
    #
    _log(log_fn, "Paso 14b: Insertando logo en A1:A5...", "INFO")
    try:
        from openpyxl.drawing.image import Image as XLImage
        from io import BytesIO


        from app.comun.tarifa_constants import resource_path
        with open(resource_path("assets/logo_dali_articulos.png"), "rb") as _f:   # (nov 2026) antes, base64 incrustado
            img_buf = BytesIO(_f.read())
        _log(log_fn, "  └ Logo cargado (assets/logo_dali_articulos.png)", "OK")

        # Pre-calcular ancho real de col A (se usará para centrar el logo).
        # Misma lógica que el paso 31 pero solo para col A, aplicada aquí
        # porque el logo se inserta antes del autoajuste.
        _max_len_a = max(
            len(_texto_visible(ws.cell(r, 1).value, 1))
            for r in range(1, ws.max_row + 1 if ws.max_row else 2)
        )
        _col_a_units_pre = max(round(_max_len_a * 1.60) + 4, 14)

        # -- Insertar en Excel con TwoCellAnchor --
        #  El logo arranca en la MITAD de A1 y termina en la MITAD de A5.
        #  Fila 4 tiene alto 6pt; filas 1,2,3,5 tienen 18pt cada una.
        #  1 px = 9525 EMU.  Mitad de fila (18pt×1.333px/pt ≈ 24px → 12px offset).
        from openpyxl.drawing.spreadsheet_drawing import (
            AnchorMarker, TwoCellAnchor
        )

        EMU_PX   = 9525          # 1 px en EMU
        FILA_MID = 12 * EMU_PX   # offset mitad de fila estándar (18pt ≈ 24px → 12px)
        ANCHO_PX = 95            # ancho del logo en píxeles

        xl_img        = XLImage(img_buf)
        xl_img.width  = ANCHO_PX
        xl_img.height = 60

        # Centrar horizontalmente en col A.
        # Conversión ancho Excel → px: (ancho * 7 + 5) calibrado para Arial 10pt.
        _col_a_px    = int(_col_a_units_pre * 7 + 5)
        _margen_px   = max(0, (_col_a_px - ANCHO_PX) // 2) + 20
        _margen_emu  = _margen_px * EMU_PX
        ANCHO_EMU    = ANCHO_PX * EMU_PX

        # _from: arranca con margen de centrado desde el borde izq. de col A
        # _to:   termina margen + ancho del logo (todo dentro de col A)
        _from = AnchorMarker(col=0, colOff=_margen_emu,              row=0, rowOff=FILA_MID)
        _to   = AnchorMarker(col=0, colOff=_margen_emu + ANCHO_EMU,  row=4, rowOff=FILA_MID)

        anchor        = TwoCellAnchor(editAs="oneCell", _from=_from, to=_to)
        anchor.pic    = None
        xl_img.anchor = anchor
        ws.add_image(xl_img)
        _log(log_fn, f"  └ Logo centrado en col A ({ANCHO_PX}px), margen={_margen_px}px", "OK")

    except ImportError as e:
        _log(log_fn, f"  └ AVISO: Pillow no disponible ({e}). Logo omitido.", "WARN")
    except Exception as e:
        _log(log_fn, f"  └ AVISO: error al insertar logo ({e}). Logo omitido.", "WARN")


def _cruzar_con_listado_dali_sap(ws, ruta_xls, ultima, log_fn=None) -> list:
    """
    Paso 27b (extraido de transformar()): busca el LISTADO CODIGOS
    DALI-SAP en la carpeta del .xls (y su carpeta padre como fallback),
    construye el mapa DALI -> (cod_sap, desc_sap, unidad_sap), y rellena
    col A (CODIGO SAP) y col D (DESCRIPCION SAP) en las filas 8..ultima
    de ws.

    Devuelve sin_desc_und: lista de codigos DALI cuyo SAP no tenia
    descripcion o unidad en el listado -- usada mas adelante en
    transformar() tanto para el diccionario _analisis como para decidir
    si el archivo se marca [INCOMPLETO] en el nombre.
    """
    # ── PASO 27b: CRUZAR COL B (CODIGO DALI) CON LISTADO DALI-SAP ──────────
    #
    #  Col B ya contiene enteros limpios tras paso 27.
    #  Busca el listado en la misma carpeta que el .xls (y en la carpeta padre
    #  como fallback). Construye mapa { DALI(int) → (cod_sap, desc_sap) }
    #  y rellena col A (CODIGO SAP) y col D (DESCRIPCION SAP).
    #
    _log(log_fn, "Paso 27b: Cruzando CODIGO DALI (col B) con el catálogo DALI-SAP...", "INFO")
    _sin_desc_und = []   # DALI cuyo SAP no tiene descripcion/unidad en el listado

    _carpeta_xls  = os.path.dirname(os.path.abspath(ruta_xls))
    _carpeta_raiz = os.path.dirname(_carpeta_xls)

    _log(log_fn, f"  [DIAG] Carpeta .xls : {_carpeta_xls}", "INFO")
    _log(log_fn, f"  [DIAG] Carpeta raíz : {_carpeta_raiz}", "INFO")

    # v10.2.50: ya no hay LISTADO CODIGOS DALI - SAP; se lee el catálogo completo (RECURSOS de la carpeta raíz,
    # la padre de la del .xls), con la misma forma de filas que tenía el LISTADO.
    from app.comun import catalogo_local as _cl
    _listado_ruta = _cl.localizar_catalogo(_carpeta_raiz, _carpeta_xls)
    if _listado_ruta:
        _log(log_fn, f"  [DIAG] ✔ Catálogo : {_listado_ruta}", "OK")
    else:
        _log(log_fn, "  ⚠ catalogo-dali-completo.xlsx no encontrado. Se omite paso 27b.", "WARN")

    _mapa_dali_sap = {}
    if _listado_ruta:
        try:
            _n   = 0
            for _row in _cl.filas_formato_listado(_listado_ruta):
                _cod_dali   = _row[0]      # col A del listado = CODIGO DALI
                _cod_sap    = _row[3]      # col D del listado = CODIGO SAP-TEST
                _desc_sap   = _row[4]      # col E del listado = DESCRIPCION SAP-TEST
                # col G = UNIDAD SAP. Hasta v10.2.50 se leía la F, que desde la inserción de columnas de
                # 2026-08 es «Texto UM (expl.)»: la descripción salía «… - Kilogramo» en vez de «… - KG»
                # (mismo fallo que el de los INC corregido en la sección 496 del historial).
                _unidad_sap = _row[6] if len(_row) > 6 else None
                if _cod_dali is None:
                    continue
                try:
                    _key = int(float(_cod_dali))
                except (ValueError, TypeError):
                    continue
                _mapa_dali_sap[_key] = (
                    _cod_sap                  if _cod_sap    not in (None, "") else None,
                    str(_desc_sap).strip()    if _desc_sap   not in (None, "") else None,
                    str(_unidad_sap).strip()  if _unidad_sap not in (None, "") else None,
                )
                _n += 1
            # Mostrar las 3 primeras claves del mapa para confirmar carga
            _muestra = list(_mapa_dali_sap.keys())[:3]
            _log(log_fn, f"  [DIAG] Mapa cargado: {_n} entradas. Primeras claves: {_muestra}", "OK")
        except Exception as _e:
            _log(log_fn, f"  ⚠ Error leyendo el catálogo: {_e}", "WARN")

    if _mapa_dali_sap:
        _enc         = 0
        _no_enc      = 0
        for r in range(8, ultima + 1):
            raw = ws.cell(r, 2).value      # col B = CODIGO DALI (int tras paso 27)
            if raw is None or raw == "":
                _no_enc += 1
                continue
            try:
                _key = int(raw)
            except (ValueError, TypeError):
                _log(log_fn, f"  [DIAG] Fila {r}: no se pudo convertir col B={repr(raw)} a int", "WARN")
                _no_enc += 1
                continue

            if _key in _mapa_dali_sap:
                _cod_sap, _desc_sap, _unidad_sap = _mapa_dali_sap[_key]
                # ── col A: CODIGO SAP (sin cambio respecto al comportamiento anterior)
                if _cod_sap not in (None, ""):
                    ws.cell(r, 1).value = _cod_sap    # → col A CODIGO SAP
                # ── col D: DESCRIPCION SAP sobreescribe valor original de TARIFA.xls
                _desc_ok   = _desc_sap   not in (None, "")
                _unidad_ok = _unidad_sap not in (None, "")
                if _desc_ok and _unidad_ok:
                    ws.cell(r, 4).value = f"{_desc_sap} - {_unidad_sap}"
                elif _desc_ok:
                    ws.cell(r, 4).value = _desc_sap
                    _sin_desc_und.append(_key)        # tiene desc pero sin unidad
                elif _unidad_ok:
                    ws.cell(r, 4).value = _unidad_sap
                    _sin_desc_und.append(_key)        # tiene unidad pero sin desc
                else:
                    _sin_desc_und.append(_key)        # ninguno → col D queda con valor de TARIFA.xls
                _enc += 1
            else:
                # Log de las primeras 3 filas sin match para diagnóstico
                if _no_enc < 3:
                    _log(log_fn, f"  [DIAG] Fila {r}: DALI={_key} → sin match en listado", "WARN")
                _no_enc += 1

        _log(log_fn,
             f"  └ {_enc} coincidencias / {_no_enc} sin match",
             "OK" if _enc > 0 else "WARN")

        # ── Aviso cuando algún SAP no tiene descripción ni unidad en el listado
        if _sin_desc_und:
            _log(log_fn,
                 f"  ⚠  {len(_sin_desc_und)} código(s) DALI sin DESCRIPCION/UNIDAD SAP en el listado: "
                 + ", ".join(str(d) for d in _sin_desc_und[:10])
                 + (" ..." if len(_sin_desc_und) > 10 else ""),
                 "WARN")
            _log(log_fn,
                 "     La columna D conserva la descripción original de TARIFA.xls para esos artículos.",
                 "WARN")
            _log(log_fn,
                 "     → Actualiza el archivo Productos en la carpeta RECURSOS:", "WARN")
            _log(log_fn,
                 "       1. En SAP: transacción ME3N (o ME2M) → ejecutar → exportar lista de materiales.", "WARN")
            _log(log_fn,
                 "       2. Guardar el fichero descargado como  Productos.xlsx  dentro de  RECURSOS/", "WARN")
            _log(log_fn,
                 "       3. Volver a ejecutar Acción 4 (Procesar y generar TARIFA).", "WARN")
    else:
        _log(log_fn, "  └ Mapa vacío — revisa los mensajes [DIAG] anteriores.", "WARN")

    # ── Fondo amarillo en col D (y col A) para filas sin descripción/unidad SAP
    if _sin_desc_und:
        _set_sin_desc = set(_sin_desc_und)
        for r in range(8, ultima + 1):
            raw = ws.cell(r, 2).value
            if raw is None or raw == "":
                continue
            try:
                _k = int(raw)
            except (ValueError, TypeError):
                continue
            if _k in _set_sin_desc:
                ws.cell(r, 4).fill = _AMARILLO   # col D — descripción no sobreescrita
                ws.cell(r, 1).fill = _AMARILLO   # col A — SAP asociado (señal visual conjunta)


    return _sin_desc_und

def _finalizar_y_guardar(wb, ws, ultima, ruta_xls, nombre_base,
                          sin_desc_und, log_fn=None):
    """
    Paso 31 en adelante (extraido de transformar()): autoajusta
    columnas/filas, aplica formato condicional de duplicados,
    analiza duplicados SAP/DALI y filas sin CODIGO SAP, decide el
    nombre final (con o sin [INCOMPLETO]), guarda el .xlsx y elimina
    el .xls original.

    Recibe sin_desc_und (calculado en el paso 27b, fuera de esta
    funcion) para el diccionario de analisis y para decidir si el
    archivo se marca [INCOMPLETO].

    Devuelve (ruta_salida, _analisis) -- identico contrato al que ya
    devuelve transformar() (esta funcion es literalmente su ultimo
    tramo, incluido el return final).
    """
    _sin_desc_und = sin_desc_und
    # ── PASO 31: AUTOAJUSTAR COLUMNAS Y FILAS ───────────────
    #
    # openpyxl NO accede al motor de renderizado de Excel.
    # Se usa _texto_visible() para simular la cadena formateada visible y se
    # aplica un factor 1.35 (calibrado para Arial 12pt).
    # Col D (DESCRIPCION) no tiene tope máximo — se respeta la longitud real.
    #
    _log(log_fn, "Paso 31: Autoajustando columnas y filas...", "INFO")

    # Columnas A-D: autoajuste dinámico por contenido
    anchos_min = {1: 14, 2: 12, 3: 18, 4: 24}
    for c in range(1, 5):
        max_len = 0
        for r in range(1, ultima + 1):
            tv = _texto_visible(ws.cell(r, c).value, c)
            max_len = max(max_len, len(tv))
        ancho = round(max_len * 1.60) + 4
        ws.column_dimensions[get_column_letter(c)].width = max(ancho, anchos_min.get(c, 12))

    # Columnas E, F, G: anchos fijos
    ws.column_dimensions["E"].width = 19
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 19

    # Altos de fila: logo filas 1-5, separador fila 6, tabla fila 7, datos desde 8
    for r in range(1, 6):
        ws.row_dimensions[r].height = 18   # 5 filas × 18 = 90pt ≈ alto logo
    ws.row_dimensions[4].height = 6        # fila 4: separador delgado
    ws.row_dimensions[6].height = 6        # fila 6: separador delgado
    ws.row_dimensions[7].height = 20       # cabecera tabla
    for r in range(8, ultima + 1):
        ws.row_dimensions[r].height = 15

    # ── FORMATO CONDICIONAL: DUPLICADOS EN COL A Y COL B ────
    #  Relleno rojo claro (FFC7CE) + texto rojo oscuro (9C0006)
    #  sobre el rango A9:A(ultima) y B9:B(ultima)
    _log(log_fn, "Formato condicional: duplicados en cols A y B...", "INFO")
    _rojo_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE",
                             fill_type="solid")
    _rojo_font = Font(name="Arial", size=12, color="9C0006")
    _dxf_dup   = DifferentialStyle(fill=_rojo_fill, font=_rojo_font)
    _regla_dup = Rule(type="duplicateValues", dxf=_dxf_dup)

    rango_a = f"A8:A{ultima}"
    rango_b = f"B8:B{ultima}"
    ws.conditional_formatting.add(rango_a, _regla_dup)
    # La regla es un objeto inmutable reutilizable — crear una segunda instancia
    _regla_dup2 = Rule(type="duplicateValues", dxf=_dxf_dup)
    ws.conditional_formatting.add(rango_b, _regla_dup2)
    _log(log_fn, f"  └ Duplicados aplicados en {rango_a} y {rango_b}", "OK")

    # ========================================================
    #  ANÁLISIS POST-PROCESO: DUPLICADOS Y SAP FALTANTES
    # ========================================================

    # ── Contar duplicados SAP (col A) y DALI (col B) ────────
    #   Solo contamos una vez por código único duplicado
    _log(log_fn, "Análisis: contando duplicados SAP y DALI...", "INFO")

    _vals_sap  = []
    _vals_dali = []
    for r in range(8, ultima + 1):
        v_sap  = ws.cell(r, 1).value
        v_dali = ws.cell(r, 2).value
        if v_sap  not in (None, ""):
            try:
                _vals_sap.append(int(float(str(v_sap).strip())))
            except (ValueError, TypeError):
                _vals_sap.append(str(v_sap).strip())
        if v_dali not in (None, ""):
            try:
                _vals_dali.append(int(float(str(v_dali).strip())))
            except (ValueError, TypeError):
                _vals_dali.append(str(v_dali).strip())

    from collections import Counter
    _cnt_sap  = Counter(_vals_sap)
    _cnt_dali = Counter(_vals_dali)
    _dup_sap  = sorted(k for k, v in _cnt_sap.items()  if v > 1)
    _dup_dali = sorted(k for k, v in _cnt_dali.items() if v > 1)

    _log(log_fn,
         f"  └ SAP duplicados: {len(_dup_sap)} código(s) únicos  |  "
         f"DALI duplicados: {len(_dup_dali)} código(s) únicos",
         "WARN" if (_dup_sap or _dup_dali) else "OK")

    # ── Detectar filas sin CODIGO SAP (col A vacía) ──────────
    #   Para cada una anotamos el CODIGO DALI correspondiente
    _faltantes_sap = []   # lista de (fila, codigo_dali)
    for r in range(8, ultima + 1):
        v_sap  = ws.cell(r, 1).value
        v_dali = ws.cell(r, 2).value
        if v_sap in (None, "") and v_dali not in (None, ""):
            try:
                _dali_int = int(float(str(v_dali).strip()))
            except (ValueError, TypeError):
                _dali_int = v_dali
            _faltantes_sap.append((r, _dali_int))

    _log(log_fn,
         f"  └ Filas sin CODIGO SAP: {len(_faltantes_sap)}",
         "WARN" if _faltantes_sap else "OK")

    # Guardar resultados en diccionario para devolver junto con la ruta
    _analisis = {
        "dup_sap":       _dup_sap,
        "dup_dali":      _dup_dali,
        "faltantes_sap": _faltantes_sap,   # [(fila_xlsx, cod_dali), ...]
        "sin_desc_und":  _sin_desc_und,
    }

    # ========================================================
    #  GUARDAR ARCHIVO FINAL
    #  Nombre: [proveedor] [dd-mm-yyyy].xlsx
    # ========================================================
    fecha_hoy     = datetime.now().strftime("%d-%m-%Y")
    nombre_limpio = "".join(
        ch for ch in nombre_base if ch not in r'\/:*?"<>|').strip()
    nombre_limpio = _norm_espacios(nombre_limpio)   # por si lleva dobles espacios
    if not nombre_limpio:
        nombre_limpio = "TARIFA_PROCESADA"

    # El archivo se marca como INCOMPLETO cuando hay artículos sin
    # descripción/unidad SAP o cuando existen códigos duplicados.
    # Esto impide que aparezca en Acción 6 hasta que se corrija manualmente
    # y se elimine [INCOMPLETO] del nombre.
    _hay_incompleto = (
        bool(_sin_desc_und)
        or bool(_dup_sap)
        or bool(_dup_dali)
    )
    if _hay_incompleto:
        nombre_archivo = f"{nombre_limpio} [INCOMPLETO] [{fecha_hoy}].xlsx"
    else:
        nombre_archivo = f"{nombre_limpio} [{fecha_hoy}].xlsx"

    ruta_salida = os.path.join(
        os.path.dirname(os.path.abspath(ruta_xls)),
        nombre_archivo)

    _log(log_fn, f"Guardando: {os.path.basename(ruta_salida)}", "INFO")
    if _hay_incompleto:
        _log(log_fn,
             "  ⚠ Archivo marcado como [INCOMPLETO] — no aparecerá en Acción 6.",
             "WARN")
        _log(log_fn,
             "    Corrige los errores en el archivo, elimina [INCOMPLETO] del nombre",
             "WARN")
        _log(log_fn,
             "    y vuelve a ejecutar Acción 6 para crear el Registro Info.",
             "WARN")
    wb.save(ruta_salida)
    _log(log_fn, f"✔ Completado: {ruta_salida}", "OK")

    # ── ELIMINAR EL .XLS ORIGINAL ───────────────────────────
    try:
        os.remove(ruta_xls)
        _log(log_fn, f"🗑 Eliminado origen: {os.path.basename(ruta_xls)}", "OK")
    except OSError as e:
        _log(log_fn, f"⚠ No se pudo eliminar el .xls original: {e}", "WARN")

    return ruta_salida, _analisis


def transformar(ruta_xls: str, log_fn=None) -> str:

    _log(log_fn, f"Leyendo origen: {os.path.basename(ruta_xls)}", "INFO")
    src = _leer_xls(ruta_xls)

    # ── DETECCIÓN DE FORMATO ─────────────────────────────────
    _fmt = _detectar_formato(src)
    _log(log_fn,
         f"Formato detectado: {_fmt['nombre']}  "
         f"(cab=col{_fmt['col_cab']}, desc=col{_fmt['col_desc']}, "
         f"fmt=col{_fmt['col_fmt']}, precio=col{_fmt['col_prec']}, "
         f"dto=col{_fmt['col_dto']})",
         "OK")

    # Crear libro nuevo en blanco
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TARIFA"

    # ── PASOS 01–12: NUEVA CABECERA ─────────────────────────
    #
    #  Diseño de cabecera (filas 1-5, tabla en fila 7):
    #
    #    A1:A5  → logo Princess (mitad A1 → mitad A5)
    #    B1     → "Código:"      C1 → código proveedor
    #    B2     → "Proveedor:"   C2 → nombre proveedor
    #    B3     → "Periodo:"     C3 → período  (C17 xls — igual en ambos formatos)
    #    Fila 4 → vacía (alto 6, separador fino)
    #    B5     → "Dto.(%):"     C5 → dto_línea   D5 → dto_general
    #    Fila 6 → vacía (separador)
    #    Fila 7 → cabecera tabla
    #    Fila 8 → datos
    #
    _log(log_fn, "Pasos 01-12: Construyendo nueva cabecera...", "INFO")

    # Leer valores del XLS origen usando columna detectada
    _cc = _fmt["col_cab"]
    val_cod_prov  = _celda(src, 2,  _cc)   # código proveedor
    val_nombre    = _celda(src, 4,  _cc)   # nombre proveedor
    val_dto_linea = _celda(src, 8,  _cc)   # dto. por línea
    val_dto_gral  = _celda(src, 10, _cc)   # dto. general
    val_periodo   = _celda(src, 17, 3)     # C17 → período (igual en ambos formatos)

    # Nombre base para el fichero de salida
    nombre_base = str(val_nombre) if val_nombre != "" else str(val_cod_prov)

    # ── Fila 1: Código proveedor ──────────────────────────────
    ws.cell(1, 2).value = "Código:"
    ws.cell(1, 3).value = val_cod_prov

    # ── Fila 2: Nombre proveedor ──────────────────────────────
    ws.cell(2, 2).value = "Proveedor:"
    ws.cell(2, 3).value = val_nombre

    # ── Fila 3: Periodo ───────────────────────────────────────
    ws.cell(3, 2).value = "Periodo:"
    ws.cell(3, 3).value = val_periodo

    # ── Fila 4: vacía ─────────────────────────────────────────
    # (sin valor)

    # ── Fila 5: Dto. línea y Dto. general en la misma fila ───
    #  C5 = val_dto_linea (D8 origen) & " línea"
    #  D5 = val_dto_gral  (D10 origen) & " general"
    #  E5 y F5 → vacíos
    ws.cell(5, 2).value = "Dto.(%):"
    _c5 = f"{val_dto_linea} línea"   if val_dto_linea != "" else "línea"
    _d5 = f"{val_dto_gral} general"  if val_dto_gral  != "" else "general"
    ws.cell(5, 3).value = _c5
    ws.cell(5, 4).value = _d5
    # E5 y F5: sin valor

    # ── Fila 6: vacía (separador) ─────────────────────────────
    # (sin valor)

    _log(log_fn, f"  └ cod={val_cod_prov}  nombre={val_nombre}  "
                 f"periodo={val_periodo}  dto_línea={val_dto_linea}  dto_gral={val_dto_gral}", "OK")

    # ── PASO 13: Formato etiquetas B col — negrita, derecha ───
    _log(log_fn, "Paso 13: Formato etiquetas B col (negrita, derecha)...", "INFO")
    for r in (1, 2, 3, 5):
        cell = ws.cell(r, 2)
        cell.font      = _font(bold=True, size=11)
        cell.alignment = _alin(h="center", v="center")

    # ── PASO 14: Formato valores cabecera — normal, izqda ─────
    _log(log_fn, "Paso 14: Formato valores cabecera (izquierda)...", "INFO")
    for r in (1, 2, 3):
        for c in (3,):
            cell = ws.cell(r, c)
            cell.font      = _font(size=11)
            cell.alignment = _alin(h="left", v="center")
    # Fila 5: B5 etiqueta negrita; C5 y D5 valores normales (sin negrita); E5:G5 → bloque aparte
    for c in (3, 4):
        cell = ws.cell(5, c)
        cell.font      = _font(bold=False, size=11)
        cell.alignment = _alin(h="left", v="center")

    # ── E5:G5 — "DESCUENTOS YA APLICADOS EN PRECIO FINAL" — negrita, rojo ──
    ws.merge_cells("E5:G5")
    cell_e5 = ws.cell(5, 5)
    cell_e5.value     = "DESCUENTOS YA APLICADOS EN PRECIO FINAL"
    cell_e5.font      = Font(name="Arial", size=11, bold=True, color="FF0000")
    cell_e5.alignment = _alin(h="center", v="center")

    # ── BORDES EXTERNOS DE CABECERA ───────────────────────────
    #  Rangos: B1:D1 / C2:D2 / B3:C3 / B5:D5
    #  Cada rango recibe borde externo (top/bottom/left/right en los bordes perimetrales).
    def _borde_externo(rng_inicio_fila, rng_inicio_col, rng_fin_fila, rng_fin_col):
        """Aplica borde externo (perimetral) a un rango de celdas."""
        for r in range(rng_inicio_fila, rng_fin_fila + 1):
            for c in range(rng_inicio_col, rng_fin_col + 1):
                cell = ws.cell(r, c)
                top    = _LINEA if r == rng_inicio_fila else Side(style=None)
                bottom = _LINEA if r == rng_fin_fila    else Side(style=None)
                left   = _LINEA if c == rng_inicio_col  else Side(style=None)
                right  = _LINEA if c == rng_fin_col     else Side(style=None)
                # Preservar bordes existentes en los lados interiores
                eb = cell.border
                cell.border = Border(
                    top    = top    if r == rng_inicio_fila else eb.top,
                    bottom = bottom if r == rng_fin_fila    else eb.bottom,
                    left   = left   if c == rng_inicio_col  else eb.left,
                    right  = right  if c == rng_fin_col     else eb.right,
                )

    _borde_externo(1, 2, 1, 4)   # B1:D1
    _borde_externo(2, 2, 2, 4)   # B2:D2
    _borde_externo(3, 2, 3, 4)   # B3:D3
    _borde_externo(5, 2, 5, 4)   # B5:D5
    _log(log_fn, "  └ Bordes externos cabecera aplicados.", "INFO")

    # ── PASO 14b: Logo en A1:A6 ──────────────────────────────
    #    Extraido a _insertar_logo_cabecera() -- ver esa funcion para
    #    el detalle (incluye el logo embebido en base64).
    _insertar_logo_cabecera(ws, log_fn)


    # ── PASO 15: CABECERA FILA 8 ───────────────────────────
    _log(log_fn, "Paso 15: Cabecera fila 7...", "INFO")
    cabeceras = ["CODIGO SAP", "CODIGO DALI", "CODIGO PROVEEDOR",
                 "DESCRIPCION", "FORMATO", "PRECIO", "DESCUENTO"]
    for c, cab in enumerate(cabeceras, start=1):
        ws.cell(7, c).value = cab

    # ── PASO 16: A8:G8 — Arial12, Negrita, Gris, Bordes, Filtros, Centrado ─
    _log(log_fn, "Paso 16: Formato cabecera A7:G7...", "INFO")
    for c in range(1, 8):
        cell = ws.cell(7, c)
        cell.font      = _font(bold=True)
        cell.fill      = _GRIS
        cell.border    = _BORDE_TODO
        cell.alignment = _alin(h="center")
    # Filtros (rango completo se actualiza al final)
    ws.auto_filter.ref = "A7:G7"

    # ── PASOS 17–22: COPIAR DATOS DESDE FILA 24 DEL XLS ────
    _log(log_fn, "Pasos 17-22: Copiando datos de artículos...", "INFO")

    ultima_src = src.nrows   # última fila del XLS (1-based = nrows)
    _c_desc = _fmt["col_desc"]
    _c_fmt  = _fmt["col_fmt"]
    _c_prec = _fmt["col_prec"]
    _c_dto  = _fmt["col_dto"]

    fila_dest = 8
    copiados  = 0
    for r_src in range(24, ultima_src + 1):   # fila 24 en adelante (1-based)
        r0 = r_src - 1  # 0-based para xlrd
        if r0 >= src.nrows:
            break

        col_a = _celda(src, r_src, 1)        # CODIGO DALI      → B destino
        col_b = _celda(src, r_src, 2)        # COD.PROVEEDOR    → C destino
        col_d = _celda(src, r_src, _c_desc)  # DESCRIPCION      → D destino
        col_f = _celda(src, r_src, _c_fmt)   # FORMATO          → E destino
        col_g = _celda(src, r_src, _c_prec)  # PRECIO           → F destino
        col_h = _celda(src, r_src, _c_dto)   # DESCUENTO        → G destino

        # Solo copiar filas que tengan algo en col A o Descripción del origen
        if col_a == "" and col_d == "":
            continue

        # Normalizar espacios en campos de texto antes de escribir
        ws.cell(fila_dest, 1).value = ""                                          # A = CODIGO SAP (vacío)
        ws.cell(fila_dest, 2).value = col_a                                       # B = CODIGO DALI
        ws.cell(fila_dest, 3).value = col_b                                       # C = COD.PROVEEDOR
        ws.cell(fila_dest, 4).value = _norm_espacios(str(col_d)) if col_d != "" else col_d  # D = DESCRIPCION
        ws.cell(fila_dest, 5).value = _norm_espacios(str(col_f)) if col_f != "" else col_f  # E = FORMATO
        ws.cell(fila_dest, 6).value = _to_float(col_g) if col_g != "" else None   # F = PRECIO
        ws.cell(fila_dest, 7).value = col_h                                       # G = DESCUENTO (→ paso 29)
        fila_dest += 1
        copiados  += 1

    _log(log_fn, f"  └ {copiados} filas copiadas.", "OK")

    # ── PASO 23: ELIMINAR FILAS CON "Código", "Familia", "" EN COL B ──
    _log(log_fn, 'Paso 23: Eliminando filas basura en col B...', "INFO")
    fila = ws.max_row
    elim = 0
    while fila >= 8:
        val_b = str(ws.cell(fila, 2).value or "").strip().lower()
        if (val_b == ""
                or "familia"   in val_b
                or "código"    in val_b
                or "codigo"    in val_b
                or "descripci" in val_b):   # elimina también filas de sub-cabecera
            ws.delete_rows(fila)
            elim += 1
        fila -= 1
    _log(log_fn, f"  └ {elim} filas eliminadas.", "OK")

    # Última fila real después de limpieza
    ultima = _ultima_fila(ws, desde=8)
    _log(log_fn, f"  └ Última fila con datos: {ultima}", "INFO")

    # ── PASO 24: A9:G(ultima) — Arial 12, Todos los bordes ──
    _log(log_fn, f"Paso 24: Arial 12 + bordes A8:G{ultima}...", "INFO")
    for r in range(8, ultima + 1):
        for c in range(1, 8):
            cell = ws.cell(r, c)
            cell.font   = _font()
            cell.border = _BORDE_TODO

    # ── PASO 25: A9:C(ultima) — Centrar ─────────────────────
    _log(log_fn, "Paso 25: Centrar A8:C(ultima)...", "INFO")
    for r in range(8, ultima + 1):
        for c in range(1, 4):
            ws.cell(r, c).alignment = _alin(h="center")

    # ── PASO 26: E9:G(ultima) — Centrar ─────────────────────
    _log(log_fn, "Paso 26: Centrar E8:G(ultima)...", "INFO")
    for r in range(8, ultima + 1):
        for c in range(5, 8):
            ws.cell(r, c).alignment = _alin(h="center")

    # ── PASO 27: B9:B(ultima) — Formato "00000000" ───────────
    _log(log_fn, "Paso 27: Formato 000000 en col B...", "INFO")
    for r in range(8, ultima + 1):
        cell = ws.cell(r, 2)
        v = cell.value
        if v is not None and v != "":
            # Asegurar que es numérico para que el formato funcione
            f = _to_float(v)
            if f is not None:
                cell.value = int(f)
        cell.number_format = "000000"

    # ── PASO 27b: CRUZAR COL B (CODIGO DALI) CON LISTADO DALI-SAP ──────────
    #    Extraido a _cruzar_con_listado_dali_sap() -- ver esa funcion para
    #    el detalle linea a linea (identico al que habia aqui antes).
    _sin_desc_und = _cruzar_con_listado_dali_sap(ws, ruta_xls, ultima, log_fn)

    # ── PASO 28: F9:F(ultima) — Formato € español ───────────
    _log(log_fn, "Paso 28: Formato € en col F...", "INFO")
    for r in range(8, ultima + 1):
        cell = ws.cell(r, 6)
        v = _to_float(cell.value)
        if v is not None:
            cell.value = v
        cell.number_format = '#,##0.00\\ €'

    # ── PASO 29: G9:G(ultima) — Porcentaje (5 → 5%, 100 → 100%) ─
    _log(log_fn, "Paso 29: Formato % en col G (5→5%, 100→100%)...", "INFO")
    for r in range(8, ultima + 1):
        cell = ws.cell(r, 7)
        v = _to_float(cell.value)
        if v is not None:
            cell.value         = v / 100.0    # 5 → 0.05 → Excel muestra 5%
            cell.number_format = "0%"
        else:
            cell.number_format = "0%"

    ws.sheet_view.showGridLines = False   # Líneas de cuadrícula desactivadas

    # ── PASO 29b: APLICAR DESCUENTOS EN COL F (PRECIO FINAL) ────────────────
    #
    #  Fórmula: precio_final = precio_original * (1 - dto_linea/100) * (1 - dto_gral/100)
    #  donde:
    #    dto_linea = valor numérico extraído de C5 (ej. "5 línea" → 5.0)
    #    dto_gral  = valor numérico extraído de D5 (ej. "3 general" → 3.0)
    #
    #  Si el precio final difiere del original → fuente roja en esa celda.
    #  El valor almacenado en col F ya está como float (paso 28).
    #  El formato #,##0.00 € se mantiene.
    #
    _log(log_fn, "Paso 29b: Calculando precio final con descuentos en col F...", "INFO")

    def _extraer_num(texto) -> float:
        """Extrae el primer número de un texto como '5 línea' → 5.0. Devuelve 0 si no hay."""
        if texto is None or texto == "":
            return 0.0
        txt = str(texto).strip()
        # Tomar la parte antes del primer espacio
        parte = txt.split()[0] if " " in txt else txt
        try:
            return float(parte.replace(",", "."))
        except ValueError:
            return 0.0

    _dto_linea_pct = _extraer_num(ws.cell(5, 3).value)   # C5
    _dto_gral_pct  = _extraer_num(ws.cell(5, 4).value)   # D5
    _factor_linea  = 1.0 - _dto_linea_pct / 100.0
    _factor_gral   = 1.0 - _dto_gral_pct  / 100.0

    _log(log_fn, f"  └ Dto. línea={_dto_linea_pct}%  Dto. gral={_dto_gral_pct}%  "
                 f"Factores: {_factor_linea:.4f} × {_factor_gral:.4f}", "INFO")

    _hay_descuentos = (_dto_linea_pct != 0.0 or _dto_gral_pct != 0.0)

    _precios_modificados = 0
    for r in range(8, ultima + 1):
        cell_f = ws.cell(r, 6)
        cell_g = ws.cell(r, 7)
        precio_orig = _to_float(cell_f.value)
        if precio_orig is None:
            continue

        # Descuento individual de col G (ya en fracción tras paso 29: 5% → 0.05)
        _dto_g_val    = _to_float(cell_g.value) or 0.0   # 0.0 si vacío/None
        _factor_g_ind = 1.0 - _dto_g_val                 # ej. 0.05 → factor 0.95

        # Precio final = precio × factor_línea × factor_general × factor_individual
        precio_orig_r2 = round(precio_orig, 2)
        precio_final   = round(precio_orig * _factor_linea * _factor_gral * _factor_g_ind, 2)
        cell_f.value   = precio_final

        # Rojo si el precio cambió, o si precio es 0 y hay algún descuento activo
        _cambio       = abs(precio_final - precio_orig_r2) > 0.001
        _cero_con_dto = (precio_orig_r2 == 0.0 and (_hay_descuentos or _dto_g_val != 0.0))
        _cero_por_g   = (precio_orig_r2 == 0.0 and abs(_dto_g_val - 1.0) < 0.001)
        if _cambio or _cero_con_dto or _cero_por_g:
            cell_f.font = Font(name="Arial", size=12, color="FF0000")
            _precios_modificados += 1

    _log(log_fn, f"  └ {_precios_modificados} precios modificados (texto rojo).", "OK")

    # ── PASO 30: INMOVILIZAR PANELES EN A9 ──────────────────
    _log(log_fn, "Paso 30: Inmovilizando paneles en A9...", "INFO")
    ws.freeze_panes = "A8"

    # Actualizar rango de filtros con la fila real
    ws.auto_filter.ref = f"A7:{get_column_letter(7)}{ultima}"

    # ── PASO 31: AUTOAJUSTAR COLUMNAS Y FILAS ───────────────
    #    Extraido a _finalizar_y_guardar() -- ver esa funcion para el
    #    resto del proceso (autoajuste, duplicados, guardado, [INCOMPLETO]).
    return _finalizar_y_guardar(
        wb, ws, ultima, ruta_xls, nombre_base, _sin_desc_und, log_fn)
