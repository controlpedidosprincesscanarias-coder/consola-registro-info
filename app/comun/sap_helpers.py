"""
sap_helpers.py — Helpers de negocio puros para RINFO / SAP.

Sin dependencias de tkinter ni de ningún otro módulo con UI: se puede
importar tanto desde crear_registro_info.py (la app real) como desde
test_crear_registro_info.py (los tests), sin arrastrar tkinter al
proceso de test.

Central Compras Princess Canarias — (c) VAMA 2026
"""

import datetime as _dt


def normalizar_codigo_sap(valor):
    """
    Normaliza un código de material/SAP a string, tal como lo espera SAP:
    sin decimales cuando el valor es un float entero (10800044.0 ->
    "10800044"), preservando decimales reales si los hubiera (1080.5 ->
    "1080.5"). Devuelve None si valor es None.

    Antes de consolidarse aquí, este mismo patrón de 3 líneas
    (isinstance/float/int/str.strip) estaba copiado y pegado 31 veces
    en distintos puntos de crear_registro_info.py, con nombres de
    variable distintos cada vez pero comportamiento idéntico.
    """
    if valor is None:
        return None
    if isinstance(valor, float) and valor == int(valor):
        valor = int(valor)
    return str(valor).strip()


def ultima_fila_con_datos(ws, columna=5, fila_inicio=2):
    """
    Devuelve el número de la última fila con un valor no vacío en
    `columna` (1=A, 2=B... por defecto 5=E), a partir de `fila_inicio`.
    Devuelve 1 si no hay ninguna fila con datos.

    Antes de consolidarse aquí, este mismo bucle de 3 líneas estaba
    copiado y pegado 5 veces en crear_registro_info.py, siempre para
    la columna E (Hoja2/Hoja3), con distinto nombre de hoja cada vez.
    """
    ultima_fila = 1
    for fila in ws.iter_rows(min_row=fila_inicio, min_col=columna, max_col=columna):
        if fila[0].value is not None:
            ultima_fila = fila[0].row
    return ultima_fila


def parsear_fecha_celda(val):
    """
    Normaliza el valor de una celda Excel (None, datetime, date, o str
    en formato "dd.mm.yyyy" / "dd/mm/yyyy" / "yyyy-mm-dd") a un objeto
    datetime.date, o None si no se puede interpretar.

    Antes de consolidarse aquí, este mismo patrón de ~12 líneas
    (isinstance datetime/date/str + bucle de 3 formatos con
    strptime/ValueError) estaba copiado y pegado 5 veces en
    crear_registro_info.py, con nombres distintos cada vez
    (_parse_fecha, _a_date_h3, _parse_date_bis, y dos variantes
    inline) pero comportamiento idéntico.
    """
    if val is None:
        return None
    if isinstance(val, _dt.datetime):
        return val.date()
    if isinstance(val, _dt.date):
        return val
    if isinstance(val, str):
        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return _dt.datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                continue
    return None


def formatear_fecha_dmy(d):
    """
    Formatea un datetime.date (o datetime.datetime) como "dd.mm.yyyy",
    el formato de fecha que usa SAP en Hoja 3.

    Antes de consolidarse aquí, este mismo one-liner (d.strftime("%d.%m.%Y"))
    estaba definido como función anidada 2 veces en crear_registro_info.py
    (_date_str en _ajustar_fechas_fin_h3, _date_a_str_h3 en
    _actualizar_h3_heredados), con nombre distinto cada vez.
    """
    return d.strftime("%d.%m.%Y")


def clave_pir_h3(fila_row):
    """
    Extrae la clave SAP PIR (Purchasing Info Record) completa de una fila
    de Hoja 3, a partir de la tupla de celdas que devuelve
    ws.iter_rows(...) (celdas, no valores).

    Clave PIR = (Registro_info, Org_compras, Material, Centro), que
    corresponde a las columnas (A=1, D=4, E=5, F=6) de Hoja 3, es decir
    índices 0-based (0, 3, 4, 5). Usar solo (material, centro) es
    insuficiente: el mismo material puede tener condiciones para
    distintos proveedores u organizaciones de compra, y mezclarlas
    produce solapamientos fantasma o actualizaciones cruzadas.

    Devuelve (reg_info_str, org_str, material_str, centro_str), o None
    si la fila no tiene suficientes columnas o no tiene Material (col F).

    Extraída de _actualizar_h3_heredados() (antes _clave_pir_h3, función
    anidada) — lógica pura, sin dependencias de self/wb/ws.
    """
    if len(fila_row) < 6:
        return None
    val_a = fila_row[0].value   # col A  Registro info
    val_d = fila_row[3].value   # col D  Org. compras
    val_e = fila_row[4].value   # col E  Centro
    val_f = fila_row[5].value   # col F  Material
    if val_f is None:
        return None
    raw_f = normalizar_codigo_sap(val_f)
    return (
        str(val_a).strip() if val_a is not None else "",
        str(val_d).strip() if val_d is not None else "",
        str(raw_f).strip(),
        str(val_e).strip() if val_e is not None else "",
    )


def material_de_clave_pir(clave_pir):
    """
    Extrae solo el campo Material (3er elemento) de una clave PIR de 4
    campos (reg_info, org, material, centro), tal como la devuelve
    clave_pir_h3(). Devuelve "" si clave_pir es None/vacía.

    Extraída de _actualizar_h3_heredados() (antes _mat_de_clave).
    """
    return clave_pir[2] if clave_pir else ""


def formatear_clave_pir(clave_pir):
    """
    Formatea una clave PIR de 4 campos (reg_info, org, material, centro)
    como "reg|org|mat|centro" para mostrarla en logs, omitiendo los
    campos vacíos. Si no queda ningún campo no vacío, devuelve str(clave_pir).

    Extraída de _actualizar_h3_heredados() (antes _fmt_clave).
    """
    partes = [p for p in (clave_pir[0], clave_pir[1], clave_pir[2], clave_pir[3]) if p]
    return "|".join(partes) if partes else str(clave_pir)

def eliminar_filas_agrupadas(ws: "openpyxl.worksheet.worksheet.Worksheet",
                               filas) -> None:
    """
    Elimina físicamente de `ws` todas las filas de `filas`, agrupando los
    índices consecutivos en rangos y llamando a ws.delete_rows(inicio,
    cantidad) UNA sola vez por rango contiguo (en orden descendente).

    Motivo: openpyxl.delete_rows() desplaza hacia arriba todas las filas
    que quedan por debajo del punto de borrado en cada llamada. Invocarlo
    fila a fila (como se hacía antes) sobre miles de filas dispersas repite
    ese desplazamiento una vez por fila, lo que da un coste ~O(n²) y es lo
    que provocaba que la generación de un RINFO desde un Purchasing grande
    tardase varios minutos en el paso "eliminación física (coherencia con
    Hoja2)". Agrupando en rangos contiguos el número de desplazamientos se
    reduce al mínimo imprescindible, sin cambiar el resultado final.

    Nota: se deduplica con set() antes de agrupar. Si `filas` contuviera
    el mismo número repetido (no debería ocurrir: cada llamador construye
    la lista en una única pasada por fila), el código antiguo fila-a-fila
    habría llamado a delete_rows() dos veces sobre esa posición, borrando
    de facto una fila distinta la segunda vez -- un bug latente. Aquí se
    deduplica explícitamente para evitarlo.
    """
    if not filas:
        return
    filas_ordenadas = sorted(set(filas))
    rangos = []
    inicio = anterior = filas_ordenadas[0]
    for f in filas_ordenadas[1:]:
        if f == anterior + 1:
            anterior = f
            continue
        rangos.append((inicio, anterior))
        inicio = anterior = f
    rangos.append((inicio, anterior))

    # Rangos en orden descendente para no desplazar los que faltan por borrar.
    for inicio_r, fin_r in sorted(rangos, reverse=True):
        ws.delete_rows(inicio_r, fin_r - inicio_r + 1)


# ─────────────────────────────────────────────────────────────────────────────
#  CATÁLOGO DALI COMPLETO (catalogo-dali-completo.xlsx) — columnas
#
#  Única fuente de verdad del layout del catálogo. Se lee y se escribe SIEMPRE
#  por nombre de cabecera (fila 4), nunca por posición fija: así un valor no
#  cambia de columna cuando se insertan columnas nuevas (2026-09-28: se
#  insertaron "Texto UM (expl.)" y "Precio €" entre "Unidad SAP" y "Grupo
#  productos SAP", desplazando 8 columnas) y el fichero sigue siendo válido
#  con el formato antiguo de 20 columnas.
# ─────────────────────────────────────────────────────────────────────────────

CATALOGO_FILA_CABECERA = 4
CATALOGO_PRIMERA_FILA_DATOS = 5

# (clave interna, cabecera en el Excel, ancho de columna)
CATALOGO_COLUMNAS = (
    ("clave",            "Código DALI",                     17.14),
    ("cod_prov",         "Código Proveedor",                17.0),
    ("descripcion",      "Descripción - DALI",              81.0),
    ("unidad",           "Unidad - DALI",                   17.57),
    ("igic",             "IGIC - DALI",                     15.86),
    ("naturaleza",       "Naturaleza - DALI",               29.14),
    ("familia",          "Familia - DALI",                  44.0),
    ("subfamilia",       "Subfamilia - DALI",               58.86),
    ("proveedor",        "Proveedor Asignado - DALI",       63.43),
    ("cod_sap",          "Código SAP",                      15.57),
    ("desc_sap",         "Descripción SAP",                 48.43),
    ("unidad_sap",       "Unidad SAP",                      15.86),
    ("texto_um",         "Texto UM (expl.)",                20.0),
    ("precio",           "Precio €",                        12.57),
    ("grp_prod",         "Grupo productos SAP",             19.86),
    ("desc_grp_prod",    "Descripción grupo productos SAP", 35.0),
    ("ind_imp",          "Indicador impuestos SAP",         27.71),
    ("categoria",        "Categoría valoración",            19.14),
    ("tipo_prod",        "Tipo producto",                   13.29),
    ("grp_compras",      "Grupo compras SAP",               18.43),
    ("desc_grp_compras", "Nombre grupo compras SAP",        25.86),
    ("activo",           "Activo",                          11.14),
)

# Columnas cuyo valor procede del LISTADO CODIGOS DALI - SAP.xlsx.
CATALOGO_CLAVES_SAP = (
    "cod_sap", "desc_sap", "unidad_sap", "texto_um", "grp_prod",
    "desc_grp_prod", "ind_imp", "categoria", "tipo_prod", "grp_compras",
    "desc_grp_compras",
)


def normalizar_cabecera_catalogo(texto):
    """Compara cabeceras sin depender de mayúsculas ni de espacios repetidos."""
    if texto is None:
        return ""
    return " ".join(str(texto).split()).casefold()


def mapa_columnas_catalogo(cabeceras):
    """
    Devuelve {clave_interna: nº de columna (base 1)} para las cabeceras
    reconocidas de una fila de cabecera del catálogo. Las cabeceras que no se
    reconocen se ignoran, y las columnas conocidas que falten simplemente no
    aparecen en el resultado (quien escribe debe saltárselas).
    """
    por_cabecera = {normalizar_cabecera_catalogo(h): k
                    for k, h, _ancho in CATALOGO_COLUMNAS}
    mapa = {}
    for idx, h in enumerate(cabeceras, start=1):
        clave = por_cabecera.get(normalizar_cabecera_catalogo(h))
        if clave is not None and clave not in mapa:
            mapa[clave] = idx
    return mapa


def _limpiar_texto(v):
    """Quita espacios sobrantes de inicio/fin si v es texto."""
    return v.strip() if isinstance(v, str) else v


def datos_sap_desde_fila_listado(row):
    """
    Convierte una fila de LISTADO CODIGOS DALI - SAP.xlsx (valores desde la
    col A, base 0) en el diccionario con las claves SAP del catálogo.

    Columnas del LISTADO: D=Código SAP · E=Descripción SAP · F=Texto UM (expl.)
    · G=Unidad SAP · H=Grupo productos · I=Descripción grupo · J=Indicador
    impuestos · K=Categoría valoración · L=Tipo producto · M=Grupo compras ·
    N=Nombre grupo compras.
    """
    def _v(i):
        return _limpiar_texto(row[i]) if len(row) > i else None

    cod_sap = row[3] if len(row) > 3 else None
    return {
        "cod_sap":          str(cod_sap).strip() if cod_sap not in (None, "") else None,
        "desc_sap":         _v(4),
        "texto_um":         _v(5),
        "unidad_sap":       _v(6),
        "grp_prod":         _v(7),
        "desc_grp_prod":    _v(8),
        "ind_imp":          _v(9),
        "categoria":        _v(10),
        "tipo_prod":        _v(11),
        "grp_compras":      _v(12),
        "desc_grp_compras": _v(13),
    }


def es_precio_valido(v):
    """
    True si `v` sirve como "Precio €" del catálogo: un número (int/float, no
    bool) mayor que 0 y finito. Una celda vacía, un 0, un texto o un negativo
    NO deben sustituir un precio bueno del catálogo.
    """
    return (isinstance(v, (int, float)) and not isinstance(v, bool)
            and 0 < v < float("inf"))


def debe_cerrar_fin_validez(fecha_i, fecha_j, hoy, fecha_cierre):
    """
    Decide si, al dar de baja un material, hay que cerrar la fecha fin
    (col J de Hoja 3) del período más reciente a `fecha_cierre` (día anterior
    al inicio de la nueva tarifa).

    Procede solo si el período está vigente (fin vacío o >= hoy) Y empieza el
    mismo día del cierre o antes. Si empieza DESPUÉS (p. ej. un período
    01.10→30.10 con cierre 30.09), poner su fin en la fecha de cierre lo
    dejaría invertido (fin anterior al inicio): en ese caso no se toca.

    Bug real corregido (2026-10-06, MELOCOTON/FRAN Y CHEMI): el RINFO llevaba
    el período `01.10.2026 → 30.09.2026`.
    """
    vigente = fecha_j is None or fecha_j >= hoy
    empieza_a_tiempo = fecha_i is None or fecha_i <= fecha_cierre
    return bool(vigente and empieza_a_tiempo)



def leer_productos_xlsx(ruta_productos):
    """
    Lee Productos.xlsx (maestro de materiales de SAP) y devuelve (tabla_productos, filas_prod):
    {codigo_sap_str: (desc_prod, texto_um, unidad_base, grp_prod, desc_grp, tipo_producto)}.
    Clave = col B; C descripción, D tipo de producto, F grupo, G descripción del grupo, I unidad base y
    J texto UM (expl.). Datos desde la fila 2. Movido sin cambios desde AccionListadoSap._leer_productos_sap
    (v10.2.50) para usarlo también en «Asignar SAP pendientes».
    """
    import openpyxl
    wb_prod = openpyxl.load_workbook(ruta_productos, read_only=True, data_only=True)
    ws_prod = wb_prod.active
    tabla_productos = {}
    filas_prod = 0
    for fila in ws_prod.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in fila):
            continue
        clave = fila[1] if len(fila) > 1 else None
        if clave is None or str(clave).strip() == "":
            continue
        clave = str(clave).strip()
        tipo_producto = fila[3] if len(fila) > 3 else None   # D
        desc_prod     = fila[2] if len(fila) > 2 else None   # C
        grp_prod      = fila[5] if len(fila) > 5 else None   # F
        desc_grp      = fila[6] if len(fila) > 6 else None   # G
        unidad_base   = fila[8] if len(fila) > 8 else None   # I
        texto_um      = fila[9] if len(fila) > 9 else None   # J
        tabla_productos[clave] = (desc_prod, texto_um, unidad_base,
                                   grp_prod, desc_grp, tipo_producto)
        filas_prod += 1
    wb_prod.close()
    return tabla_productos, filas_prod
