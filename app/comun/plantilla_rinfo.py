# -*- coding: utf-8 -*-
"""
plantilla_rinfo.py — «Generar plantilla Tarifa desde RINFO» (Acción 6): lee un RINFO o un Purchasing Info Records y
prepara las tarifas que se envían a los proveedores para la nueva revisión de precios (sin tkinter).

Petición de Víctor (21 nov 2026): «arrastrar un RINFO o Purchasing y la aplicación generará una tarifa a razón del
RINFO. Sería el mismo paso de Crear / Actualizar Registro Info SAP, pero a la inversa». Decisiones (22 nov 2026):
  1. la plantilla es para enviarla a los proveedores y que la rellenen con la nueva revisión de precios
     (columnas amarillas PRECIO NUEVO y FECHA INICIO);
  2. una tarifa por proveedor, cada una con sus artículos, teléfono y hoteles;
  3. una tarifa por hotel según el precio: los hoteles con los MISMOS precios van juntos en una tarifa;
  4. es el botón 6 (antes «Generar Listado Tarifa Proveedor Vacía», en construcción);
  5. lo que devuelve el proveedor entra tal cual en «Crear / Actualizar Registro Info SAP» (PRECIO NUEVO donde esté
     relleno; FECHA INICIO como fecha propuesta), ver datos_tarifa.precios_tarifa / periodo_propuesto.

Qué se lee del Purchasing (las mismas 4 hojas en un RINFO de la Consola y en uno descargado de SAP):
  Hoja 1 «Info general p.reg.info»: A registro info, C proveedor, D nombre, E material, F texto breve, N unidad de
          pedido, AB número de material de proveedor, AD teléfono, AJ registro marcado para borrado.
  Hoja 2 «Datos org.compr.p.reg.info»: A registro info, F centro, J material, AN centro marcado para borrado.
  Hoja 3 «Det.precios p.registros info»: A registro info, E centro, F material, I inicio, J fin, O importe,
          Q por cuántas unidades (determinación de precio), Y indicador de borrado.
  Hoja 4 «Escalas p.cond.precio PIR»: F material (solo para avisar de los escalados).

Precio de cada material y hotel: el vigente hoy (si hay varios, el que empezó más tarde); si no hay ninguno vigente,
el próximo que empieza; si solo hay precios caducados, el material no sale en ese hotel. Lo marcado para borrar
(el registro entero, el centro o la línea de precio) no sale.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import datetime as _dt
import re
from collections import OrderedDict

from app.comun.datos_tarifa import CENTROS_TARIFA
from app.comun.sap_helpers import normalizar_codigo_sap, parsear_fecha_celda

HOJA_1 = "Info general p.reg.info"
HOJA_2 = "Datos org.compr.p.reg.info"
HOJA_3 = "Det.precios p.registros info"
HOJA_4 = "Escalas p.cond.precio PIR"

_ORDEN_CENTRO = {c: i for i, (_s, c, _n) in enumerate(CENTROS_TARIFA)}
_SIGLA = {c: s for s, c, _n in CENTROS_TARIFA}


class NoEsPurchasing(ValueError):
    """El archivo no tiene las hojas de un Purchasing Info Records / RINFO."""


def _txt(v):
    if v is None:
        return ""
    v = normalizar_codigo_sap(v)
    return str(v).strip()


def _num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace("€", "").replace(" ", "")
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


def _marcado(v):
    return _txt(v).upper() == "X"


def _celda(fila, col_1):
    return fila[col_1 - 1] if len(fila) >= col_1 else None


def _filas(ws):
    for fila in ws.iter_rows(min_row=2, values_only=True):
        if fila and any(v not in (None, "") for v in fila):
            yield fila


def leer_purchasing(ruta, hoy=None):
    """Lee el archivo y devuelve los datos por proveedor. Lanza NoEsPurchasing si no tiene sus hojas.

    {"proveedores": OrderedDict{cod: {"codigo", "nombre", "telefono",
                                      "materiales": OrderedDict{mat: {"texto", "um", "cod_prov_mat",
                                                                      "precios": {centro: (precio, inicio, fin)}}}}},
     "avisos": {"borrados": set(mat), "centros_borrados": n, "caducados": set((mat, centro)),
                "escalas": set(mat), "centros_desconocidos": set(centro), "sin_precio": set(mat)}}
    """
    import openpyxl
    hoy = hoy or _dt.date.today()
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        if HOJA_1 not in wb.sheetnames or HOJA_3 not in wb.sheetnames:
            raise NoEsPurchasing(f"No tiene las hojas «{HOJA_1}» y «{HOJA_3}».")
        h1 = list(_filas(wb[HOJA_1]))
        h2 = list(_filas(wb[HOJA_2])) if HOJA_2 in wb.sheetnames else []
        h3 = list(_filas(wb[HOJA_3]))
        h4 = list(_filas(wb[HOJA_4])) if HOJA_4 in wb.sheetnames else []
    finally:
        wb.close()

    avisos = {"borrados": set(), "centros_borrados": 0, "caducados": set(), "escalas": set(),
              "centros_desconocidos": set(), "sin_precio": set()}
    proveedores = OrderedDict()
    por_reg = {}          # registro info → (proveedor, material)
    por_mat = {}          # material → [(proveedor, material)] (los RINFO nuevos aún no tienen registro info)
    borrados = set()      # (proveedor, material) con el registro entero marcado para borrar

    for f in h1:
        reg, prov, mat = _txt(_celda(f, 1)), _txt(_celda(f, 3)), _txt(_celda(f, 5))
        nombre = _txt(_celda(f, 4))
        if not mat or not (prov or nombre):
            continue
        clave_prov = prov or nombre
        p = proveedores.setdefault(clave_prov, {"codigo": prov, "nombre": nombre or prov, "telefono": "",
                                                "materiales": OrderedDict()})
        if not p["telefono"]:
            p["telefono"] = _txt(_celda(f, 30))
        if _marcado(_celda(f, 36)):
            borrados.add((clave_prov, mat))
            avisos["borrados"].add(mat)
            continue
        m = p["materiales"].setdefault(mat, {"texto": "", "um": "", "cod_prov_mat": "", "precios": {}})
        m["texto"] = m["texto"] or _txt(_celda(f, 6))
        m["um"] = m["um"] or _txt(_celda(f, 14))
        m["cod_prov_mat"] = m["cod_prov_mat"] or _txt(_celda(f, 28))
        if reg:
            por_reg[reg] = (clave_prov, mat)
        por_mat.setdefault(mat, [])
        if (clave_prov, mat) not in por_mat[mat]:
            por_mat[mat].append((clave_prov, mat))

    def _dueno(reg, mat):
        if reg and reg in por_reg:
            return por_reg[reg]
        lista = por_mat.get(mat) or []
        return lista[0] if lista else None

    centros_borrados = set()
    for f in h2:
        reg, centro, mat = _txt(_celda(f, 1)), _txt(_celda(f, 6)), _txt(_celda(f, 10))
        if centro and _marcado(_celda(f, 40)):
            d = _dueno(reg, mat)
            if d:
                centros_borrados.add((d, centro))
    avisos["centros_borrados"] = len(centros_borrados)

    candidatos = {}       # (dueño, centro) → [(inicio, fin, precio)]
    for f in h3:
        reg, centro, mat = _txt(_celda(f, 1)), _txt(_celda(f, 5)), _txt(_celda(f, 6))
        if not mat or not centro or _marcado(_celda(f, 25)):
            continue
        d = _dueno(reg, mat)
        if d is None or d in borrados or (d, centro) in centros_borrados:
            continue
        importe = _num(_celda(f, 15))
        if importe is None:
            continue
        por = _num(_celda(f, 17))
        precio = importe / por if por and por > 0 else importe
        candidatos.setdefault((d, centro), []).append(
            (parsear_fecha_celda(_celda(f, 9)), parsear_fecha_celda(_celda(f, 10)), precio))

    for ((prov, mat), centro), lista in candidatos.items():
        vigentes = [c for c in lista if (c[0] is None or c[0] <= hoy) and (c[1] is None or c[1] >= hoy)]
        futuros = [c for c in lista if c[0] is not None and c[0] > hoy]
        if vigentes:
            elegido = max(vigentes, key=lambda c: (c[0] or _dt.date.min))
        elif futuros:
            elegido = min(futuros, key=lambda c: c[0])
        else:
            avisos["caducados"].add((mat, centro))
            continue
        if centro not in _ORDEN_CENTRO:
            avisos["centros_desconocidos"].add(centro)
            continue
        proveedores[prov]["materiales"][mat]["precios"][centro] = (elegido[2], elegido[0], elegido[1])

    for f in h4:
        mat = _txt(_celda(f, 6))
        if mat and not _marcado(_celda(f, 28)):
            avisos["escalas"].add(mat)

    for p in proveedores.values():
        for mat, m in p["materiales"].items():
            if not m["precios"]:
                avisos["sin_precio"].add(mat)
    return {"proveedores": proveedores, "avisos": avisos}


def agrupar_por_precios(proveedor):
    """Las tarifas de un proveedor: los hoteles con los mismos materiales y precios van juntos (decisión de
    Víctor, opción A). Devuelve [{"centros": [...], "siglas": [...], "materiales": [(mat, datos, (precio, ini,
    fin))]}] en el orden de los centros SAP."""
    vectores = OrderedDict()
    for centro in sorted({c for m in proveedor["materiales"].values() for c in m["precios"]},
                         key=lambda c: _ORDEN_CENTRO.get(c, 99)):
        firma = tuple(sorted((mat, round(m["precios"][centro][0], 4))
                             for mat, m in proveedor["materiales"].items() if centro in m["precios"]))
        vectores.setdefault(firma, []).append(centro)
    grupos = []
    for _firma, centros in vectores.items():
        primero = centros[0]
        materiales = [(mat, m, m["precios"][primero]) for mat, m in proveedor["materiales"].items()
                      if primero in m["precios"]]
        grupos.append({"centros": centros, "siglas": [_SIGLA[c] for c in centros], "materiales": materiales})
    return grupos


def mapa_catalogo(filas_listado):
    """{código SAP: [(código DALI, descripción DALI, unidad DALI)]} desde catalogo_local.filas_formato_listado."""
    mapa = {}
    for fila in filas_listado or ():
        if not fila or len(fila) < 4 or fila[3] in (None, ""):
            continue
        sap = _txt(fila[3])
        mapa.setdefault(sap, []).append((_txt(fila[0]), _txt(fila[1]), _txt(fila[2])))
    return mapa


def filas_tarifa(grupo, catalogo):
    """Filas de la tabla de la tarifa (una por código DALI del catálogo; si el SAP no está, una con los datos de
    SAP): dicts sap, dali, cod_prov, descripcion, formato, precio."""
    filas = []
    for mat, m, (precio, _ini, _fin) in grupo["materiales"]:
        dalis = catalogo.get(mat) or [("", "", "")]
        for dali, desc, unidad in dalis:
            filas.append({"sap": mat, "dali": dali, "cod_prov": m["cod_prov_mat"],
                          "descripcion": desc or m["texto"], "formato": unidad or m["um"], "precio": precio})
    return filas


_NO_VALIDO = re.compile(r'[\\/:*?"<>|]+')


def nombre_archivo(nombre_proveedor, siglas, todas, fecha, existentes=()):
    """«PROVEEDOR [dd-mm-aaaa].xlsx» o, si el proveedor tiene precios distintos según el hotel,
    «PROVEEDOR (GY-IT-…) [dd-mm-aaaa].xlsx». Nunca uno que ya exista: «… (2) [dd-mm-aaaa].xlsx». Siempre con el
    patrón de las tarifas válidas, para que «Crear / Actualizar Registro Info SAP» la ofrezca."""
    base = _NO_VALIDO.sub(" ", nombre_proveedor or "PROVEEDOR").strip().rstrip(".") or "PROVEEDOR"
    base = re.sub(r"\s+", " ", base)
    if not todas:
        base += f" ({'-'.join(siglas)})"
    sello = f"[{fecha:%d-%m-%Y}]"
    existentes = {e.lower() for e in existentes}
    nombre, n = f"{base} {sello}.xlsx", 2
    while nombre.lower() in existentes:
        nombre, n = f"{base} ({n}) {sello}.xlsx", n + 1
    return nombre


def preparar_plantillas(datos, catalogo, fecha=None, existentes=()):
    """Todas las tarifas que hay que escribir: [{"nombre_archivo", "proveedor", "codigo", "telefono", "centros",
    "siglas", "filas"}]."""
    fecha = fecha or _dt.date.today()
    ocupados = set(existentes)
    salida = []
    for p in datos["proveedores"].values():
        grupos = agrupar_por_precios(p)
        for g in grupos:
            filas = filas_tarifa(g, catalogo)
            if not filas:
                continue
            nombre = nombre_archivo(p["nombre"], g["siglas"], len(grupos) == 1, fecha, ocupados)
            ocupados.add(nombre)
            salida.append({"nombre_archivo": nombre, "proveedor": p["nombre"], "codigo": p["codigo"],
                           "telefono": p["telefono"], "centros": g["centros"], "siglas": g["siglas"],
                           "filas": filas})
    return salida
