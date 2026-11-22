# -*- coding: utf-8 -*-
"""
datos_tarifa.py — Teléfono del proveedor y hoteles de la TARIFA final (sin tkinter).

(22 nov 2026, petición de Víctor) La cabecera de la TARIFA final lleva dos datos más, para que «Crear / Actualizar
Registro Info SAP» no se pare a pedirlos:

    B4  «Teléfono:»   C4  el teléfono del proveedor (fila 4, antes una fila separadora vacía)
    H1  «HOTELES»     I1:J5 y K1:L5  las siglas de los 10 hoteles y, al lado, una casilla para marcar con ✔

        H          I    J      K    L
    1  HOTELES     GY  [✔]     MG  [ ]
    2              IT  [ ]     GC  [✔]
    3              LP  [ ]     MT  [ ]
    4              TA  [✔]     JN  [ ]
    5              ST  [ ]     FV  [ ]

Decisiones de Víctor (22 nov 2026): el recuadro a la derecha de la cabecera (no en otra hoja); siglas GY IT LP TA ST
MG GC MT JN FV; las casillas salen en blanco en una tarifa nueva (si nadie marca ninguna, se pregunta como siempre);
las fechas se siguen preguntando. Si el teléfono falta o no es válido, o no hay ningún hotel marcado, sale la ventana
de siempre. Las tarifas antiguas, sin estos huecos, funcionan igual que antes.

La tabla de la tarifa no se mueve (cabecera en la fila 7, datos desde la 8).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import re

# Siglas de los hoteles en el orden de los centros SAP (mismo orden que comun_rinfo.CENTROS_SAP_CANARIAS).
CENTROS_TARIFA = (
    ("GY", "0110", "GUAYARMINA PRINCESS"),
    ("IT", "0120", "PRINCESS INSPIRE TENERIFE"),
    ("LP", "0210", "LA PALMA&TENEGUIA PRINCESS"),
    ("TA", "0310", "TAURITO PRINCESS"),
    ("ST", "0410", "SUITE PRINCESS"),
    ("MG", "0510", "MOGAN PRINCESS"),
    ("GC", "0610", "GRAN CANARIA PRINCESS"),
    ("MT", "0710", "MASPALOMAS&TABAIBA PRINCESS"),
    ("JN", "0810", "CLUB JANDIA PRINCESS"),
    ("FV", "0910", "FUERTEVENTURA PRINCESS"),
)

FILA_TELEFONO = 4
COL_ETIQUETA_TELEFONO = 2   # B
COL_TELEFONO = 3            # C
ETIQUETA_TELEFONO = "Teléfono:"

FILA_TITULO_HOTELES = 1
COL_TITULO_HOTELES = 8      # H
TITULO_HOTELES = "HOTELES"
MARCA = "✔"

# Donde se buscan las siglas al leer: filas 1..6, columnas H..T (por si alguien mueve el recuadro un poco).
_FILAS_BUSQUEDA = range(1, 7)
_COLS_BUSQUEDA = range(8, 21)

_RE_TELEFONO = re.compile(r"^\d{8,10}$")   # la misma regla que VentanaTelefonoRINFO
_NO_MARCADO = {"", "NO", "N", "0", "FALSE", "FALSO", "-", "✘", "✗"}


def posicion_centro(indice):
    """(fila, columna de la sigla, columna de la casilla) del hotel `indice` (0..9)."""
    fila = 1 + indice % 5
    col_sigla = 9 + 2 * (indice // 5)       # I o K
    return fila, col_sigla, col_sigla + 1   # casilla en J o L


def limpiar_telefono(valor):
    """El teléfono sin espacios, guiones ni puntos si tiene de 8 a 10 dígitos; si no, None."""
    if valor is None:
        return None
    if isinstance(valor, float) and valor == int(valor):
        valor = int(valor)
    texto = re.sub(r"[\s\-./]", "", str(valor))
    return texto if _RE_TELEFONO.match(texto) else None


def es_marca(valor):
    """True si la casilla de un hotel está marcada (✔, X, SÍ, 1…)."""
    if valor is None or valor is False:
        return False
    if valor is True:
        return True
    return str(valor).strip().upper() not in _NO_MARCADO


def _texto(valor):
    return "" if valor is None else str(valor).strip()


def leer_datos_hoja(ws):
    """Lee el teléfono y los hoteles marcados de la hoja de una TARIFA final (openpyxl, también en read_only).

    Devuelve un dict:
      tiene_huecos   la tarifa ya trae la etiqueta «Teléfono:» o las siglas de los hoteles (formato nuevo);
      telefono_texto lo escrito en C4, tal cual ("" si nada);
      telefono       el teléfono válido sin espacios, o None;
      centros        [(código SAP, nombre)] de los hoteles marcados, en el orden de los centros SAP;
      siglas         las siglas de esos hoteles.
    """
    celdas = {}
    for fila in ws.iter_rows(min_row=1, max_row=max(_FILAS_BUSQUEDA), min_col=1, max_col=max(_COLS_BUSQUEDA)):
        for c in fila:
            if c.value is not None and getattr(c, "column", None) is not None:
                celdas[(c.row, c.column)] = c.value

    etiqueta = _texto(celdas.get((FILA_TELEFONO, COL_ETIQUETA_TELEFONO))).lower()
    tiene_telefono = etiqueta.startswith("tel")
    telefono_texto = _texto(celdas.get((FILA_TELEFONO, COL_TELEFONO))) if tiene_telefono else ""
    telefono = limpiar_telefono(celdas.get((FILA_TELEFONO, COL_TELEFONO))) if tiene_telefono else None

    por_sigla = {}
    for (r, col), v in celdas.items():
        if r in _FILAS_BUSQUEDA and col in _COLS_BUSQUEDA:
            s = _texto(v).upper()
            if s and s not in por_sigla:
                por_sigla[s] = (r, col)
    centros, siglas, hay_siglas = [], [], False
    for sigla, codigo, nombre in CENTROS_TARIFA:
        pos = por_sigla.get(sigla)
        if pos is None:
            continue
        hay_siglas = True
        if es_marca(celdas.get((pos[0], pos[1] + 1))):
            centros.append((codigo, nombre))
            siglas.append(sigla)

    return {
        "tiene_huecos": tiene_telefono or hay_siglas,
        "telefono_texto": telefono_texto,
        "telefono": telefono,
        "centros": centros,
        "siglas": siglas,
    }


def leer_datos_tarifa(ruta_tarifa):
    """Como leer_datos_hoja, abriendo el archivo (hoja «TARIFA» o la activa). Si no se puede leer, sin datos."""
    vacio = {"tiene_huecos": False, "telefono_texto": "", "telefono": None, "centros": [], "siglas": []}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ruta_tarifa, read_only=True, data_only=True)
    except Exception:  # noqa: BLE001 — sin datos se pregunta como siempre
        return vacio
    try:
        ws = wb["TARIFA"] if "TARIFA" in wb.sheetnames else wb.active
        return leer_datos_hoja(ws)
    except Exception:  # noqa: BLE001
        return vacio
    finally:
        wb.close()


# ── Escribir (TARIFA final de la Acción 3 y, más adelante, la plantilla desde RINFO) ─────────────────────
def escribir_huecos(ws, telefono=None, codigos_marcados=()):
    """Escribe en la cabecera de la TARIFA la fila «Teléfono:» y el recuadro de los hoteles, con su formato y una
    lista desplegable «✔» en cada casilla. `telefono` y `codigos_marcados` (códigos SAP de centro) rellenan los
    huecos; por defecto quedan en blanco."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    linea = Side(style="thin", color="000000")
    borde = Border(left=linea, right=linea, top=linea, bottom=linea)
    gris = PatternFill(fill_type="solid", fgColor="D9D9D9")
    centro = Alignment(horizontal="center", vertical="center")

    # Teléfono (fila 4): etiqueta como las de B1..B5 y valor como C1..C3
    et = ws.cell(FILA_TELEFONO, COL_ETIQUETA_TELEFONO)
    et.value = ETIQUETA_TELEFONO
    et.font = Font(name="Arial", size=11, bold=True)
    et.alignment = centro
    tel = ws.cell(FILA_TELEFONO, COL_TELEFONO)
    tel.value = telefono or None
    tel.number_format = "@"
    tel.font = Font(name="Arial", size=11)
    tel.alignment = Alignment(horizontal="left", vertical="center")
    for col in range(COL_ETIQUETA_TELEFONO, COL_TELEFONO + 2):    # B4:D4 con borde exterior
        ws.cell(FILA_TELEFONO, col).border = Border(
            top=linea, bottom=linea,
            left=linea if col == COL_ETIQUETA_TELEFONO else Side(style=None),
            right=linea if col == COL_TELEFONO + 1 else Side(style=None))

    # Hoteles
    tit = ws.cell(FILA_TITULO_HOTELES, COL_TITULO_HOTELES)
    tit.value = TITULO_HOTELES
    tit.font = Font(name="Arial", size=11, bold=True)
    tit.alignment = centro
    marcados = {str(c) for c in codigos_marcados}
    dv = DataValidation(type="list", formula1=f'"{MARCA}"', allow_blank=True, showErrorMessage=False)
    dv.prompt = "Marque con ✔ (o X) los hoteles a los que se aplica la tarifa."
    dv.promptTitle = "Hoteles"
    ws.add_data_validation(dv)
    for i, (sigla, codigo, _nombre) in enumerate(CENTROS_TARIFA):
        fila, col_s, col_m = posicion_centro(i)
        cs = ws.cell(fila, col_s)
        cs.value = sigla
        cs.font = Font(name="Arial", size=11, bold=True)
        cs.fill = gris
        cs.border = borde
        cs.alignment = centro
        cm = ws.cell(fila, col_m)
        cm.value = MARCA if codigo in marcados else None
        cm.font = Font(name="Arial", size=12, bold=True, color="1A7A1A")
        cm.border = borde
        cm.alignment = centro
        dv.add(cm.coordinate)
    ws.column_dimensions[get_column_letter(COL_TITULO_HOTELES)].width = 12
    for col in range(9, 13):
        ws.column_dimensions[get_column_letter(col)].width = 6
