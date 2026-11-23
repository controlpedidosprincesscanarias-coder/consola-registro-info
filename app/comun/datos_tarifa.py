# -*- coding: utf-8 -*-
"""
datos_tarifa.py — Teléfono del proveedor y hoteles de la TARIFA final (sin tkinter).

(22 nov 2026, petición de Víctor) La cabecera de la TARIFA final lleva dos datos más, para que «Crear / Actualizar
Registro Info SAP» no se pare a pedirlos:

    B4  «Teléfono:»   C4  el teléfono del proveedor (fila 4, antes una fila separadora vacía)
    E1:H1 «HOTELES»   E2:H6  las siglas de los 10 hoteles y, al lado, una casilla para marcar con ✔

        E (FORMATO) F (PRECIO)  G (DESCUENTO) H (REGISTRO INFO)
    1  ─────────────── HOTELES ───────────────
    2   GY          [✔]         MG            [ ]
    3   IT          [ ]         GC            [✔]
    4   LP          [ ]         MT            [ ]
    5   TA          [✔]         JN            [ ]
    6   SU          [ ]         FV            [ ]

    (22 nov 2026, petición de Víctor) El recuadro va ENCIMA de las columnas E..H de la tabla, para que la tarifa no
    se ensanche; al principio estaba en H1:L5, a la derecha. Por eso la frase roja «DESCUENTOS YA APLICADOS EN
    PRECIO FINAL», que estaba en E5:G5, pasa a B6:D6, debajo de «Dto.(%)», y la fila 6 deja de ser un separador.

Decisiones de Víctor (22 nov 2026): el recuadro en la cabecera, encima de la tabla (no en otra hoja); siglas GY IT LP TA SU
MG GC MT JN FV (SU, no ST, desde el 23 nov 2026); las casillas salen en blanco en una tarifa nueva (si nadie marca ninguna, se pregunta como siempre);
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
    ("SU", "0410", "SUITE PRINCESS"),   # (23 nov 2026, Víctor) antes «ST»: se sigue leyendo
    ("MG", "0510", "MOGAN PRINCESS"),
    ("GC", "0610", "GRAN CANARIA PRINCESS"),
    ("MT", "0710", "MASPALOMAS&TABAIBA PRINCESS"),
    ("JN", "0810", "CLUB JANDIA PRINCESS"),
    ("FV", "0910", "FUERTEVENTURA PRINCESS"),
)

# (23 nov 2026) Siglas que se usaron antes y se siguen leyendo en las tarifas ya generadas.
SIGLAS_ANTIGUAS = {"ST": "SU"}

FILA_TELEFONO = 4
COL_ETIQUETA_TELEFONO = 2   # B
COL_TELEFONO = 3            # C
ETIQUETA_TELEFONO = "Teléfono:"

FILA_TITULO_HOTELES = 1
COL_TITULO_HOTELES = 5      # E (combinada E1:H1)
COL_FIN_HOTELES = 8         # H
TITULO_HOTELES = "HOTELES"
MARCA = "✔"
AVISO_DESCUENTOS = "DESCUENTOS YA APLICADOS EN PRECIO FINAL"
RANGO_AVISO_DESCUENTOS = "B6:D6"   # antes E5:G5, donde ahora van los hoteles

# Donde se buscan las siglas al leer: filas 1..6, columnas E..T (también vale si alguien mueve el recuadro un poco).
_FILAS_BUSQUEDA = range(1, 7)
_COLS_BUSQUEDA = range(5, 21)

_RE_TELEFONO = re.compile(r"^\d{8,10}$")   # la misma regla que VentanaTelefonoRINFO
_NO_MARCADO = {"", "NO", "N", "0", "FALSE", "FALSO", "-", "✘", "✗"}


def posicion_centro(indice):
    """(fila, columna de la sigla, columna de la casilla) del hotel `indice` (0..9)."""
    fila = 2 + indice % 5
    col_sigla = 5 + 2 * (indice // 5)       # E o G
    return fila, col_sigla, col_sigla + 1   # casilla en F o H


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
            s = SIGLAS_ANTIGUAS.get(s, s)
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
    """Escribe en la cabecera de la TARIFA la fila «Teléfono:», el recuadro de los hoteles (con su formato y una
    lista desplegable «✔» en cada casilla) y el aviso «DESCUENTOS YA APLICADOS EN PRECIO FINAL» en B6:D6. `telefono` y `codigos_marcados` (códigos SAP de centro) rellenan los
    huecos; por defecto quedan en blanco."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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

    # Hoteles (E1:H6, encima de la tabla)
    ws.merge_cells(start_row=FILA_TITULO_HOTELES, start_column=COL_TITULO_HOTELES,
                   end_row=FILA_TITULO_HOTELES, end_column=COL_FIN_HOTELES)
    tit = ws.cell(FILA_TITULO_HOTELES, COL_TITULO_HOTELES)
    tit.value = TITULO_HOTELES
    tit.font = Font(name="Arial", size=11, bold=True)
    tit.fill = gris
    tit.alignment = centro
    for col in range(COL_TITULO_HOTELES, COL_FIN_HOTELES + 1):
        ws.cell(FILA_TITULO_HOTELES, col).border = borde
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

    # Aviso de descuentos, debajo de «Dto.(%)» (antes en E5:G5)
    ws.merge_cells(RANGO_AVISO_DESCUENTOS)
    av = ws[RANGO_AVISO_DESCUENTOS.split(":")[0]]
    av.value = AVISO_DESCUENTOS
    av.font = Font(name="Arial", size=11, bold=True, color="FF0000")
    av.alignment = centro


# ── (22 nov 2026) Columnas del proveedor en la plantilla de revisión de precios (Acción 6) ───────────────
# La plantilla que se genera desde un RINFO (app/comun/plantilla_rinfo.py) lleva, después de DESCUENTO, dos columnas
# amarillas que rellena el proveedor. Al volver, «Crear / Actualizar Registro Info SAP» usa el PRECIO NUEVO donde
# esté relleno (si no, el PRECIO de siempre) y propone la FECHA INICIO (decisión de Víctor, 22 nov 2026, opción A).
CAB_PRECIO_NUEVO = "PRECIO NUEVO"
# (22 nov 2026, petición de Víctor) Columna H de las dos tarifas, detrás de DESCUENTO: el Registro Info SAP de cada
# material. Al final de la tabla para no mover ninguna columna que la Consola lee por posición (A..G).
CAB_REGISTRO_INFO = "REGISTRO INFO"
COL_REGISTRO_INFO = 8   # H
CAB_FECHA_INICIO = "FECHA INICIO"
FILA_CABECERA_TABLA = 7


def _norm_cab(v):
    return re.sub(r"\s+", " ", _texto(v)).upper()


def columnas_proveedor(ws):
    """Índices (0 = columna A) de PRECIO NUEVO y FECHA INICIO en la cabecera de la tabla (fila 7), o None."""
    nuevo = inicio = None
    for fila in ws.iter_rows(min_row=FILA_CABECERA_TABLA, max_row=FILA_CABECERA_TABLA, values_only=True):
        for i, v in enumerate(fila or ()):
            cab = _norm_cab(v)
            if cab == CAB_PRECIO_NUEVO and nuevo is None:
                nuevo = i
            elif cab == CAB_FECHA_INICIO and inicio is None:
                inicio = i
    return nuevo, inicio


def numero(v):
    """Un precio escrito en la celda (número o texto «4,35», «4,35 €») como float, o None."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = _texto(v).replace("€", "").replace(" ", "")
    if not t:
        return None
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    else:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def precios_tarifa(ws):
    """{código SAP: precio} de la tabla de la TARIFA (A = código SAP, F = PRECIO) desde la fila 8. Si la tarifa trae
    la columna PRECIO NUEVO y el proveedor la ha rellenado (número mayor que 0), manda ese precio. Devuelve
    (precios, cuántos son PRECIO NUEVO)."""
    from app.comun.sap_helpers import normalizar_codigo_sap
    col_nuevo, _ = columnas_proveedor(ws)
    precios, nuevos = {}, 0
    for fila in ws.iter_rows(min_row=FILA_CABECERA_TABLA + 1, values_only=True):
        if not fila or fila[0] is None:
            continue
        clave = str(normalizar_codigo_sap(fila[0])).strip()
        if not clave:
            continue
        precio = fila[5] if len(fila) > 5 else None
        if col_nuevo is not None and len(fila) > col_nuevo:
            nuevo = numero(fila[col_nuevo])
            if nuevo is not None and nuevo > 0:
                precio = nuevo
                nuevos += 1
        if precio is not None:
            precios[clave] = precio
    return precios, nuevos


def periodo_propuesto(ruta_tarifa):
    """El texto que se pasa a la ventana de fechas del RINFO. Normalmente el «Periodo:» (C3). Si la tarifa trae la
    columna FECHA INICIO rellena, «inicio - fin» con la primera FECHA INICIO y el fin del periodo de C3 (o el 31/12
    de ese año)."""
    from app.comun import fechas_tarifa as FT
    from app.comun.sap_helpers import parsear_fecha_celda
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ruta_tarifa, read_only=True, data_only=True)
    except Exception:  # noqa: BLE001
        return ""
    try:
        ws = wb["TARIFA"] if "TARIFA" in wb.sheetnames else wb.active
        periodo = ""
        for fila in ws.iter_rows(min_row=3, max_row=3, min_col=3, max_col=3, values_only=True):
            periodo = _texto(fila[0]) if fila else ""
        _, col_inicio = columnas_proveedor(ws)
        inicios = []
        if col_inicio is not None:
            for fila in ws.iter_rows(min_row=FILA_CABECERA_TABLA + 1, values_only=True):
                if fila and len(fila) > col_inicio:
                    f = parsear_fecha_celda(fila[col_inicio])
                    if f is None and isinstance(fila[col_inicio], str):
                        f = FT._fecha(fila[col_inicio].strip())
                    if f is not None:
                        inicios.append(f)
    except Exception:  # noqa: BLE001
        return ""
    finally:
        wb.close()
    if not inicios:
        return periodo
    inicio = min(inicios)
    rango = FT.periodo_a_fechas(periodo)
    fin = rango[1] if rango and rango[1] >= inicio else inicio.replace(month=12, day=31)
    return f"{inicio:%d/%m/%Y} - {fin:%d/%m/%Y}"
