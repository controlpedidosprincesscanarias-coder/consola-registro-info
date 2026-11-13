# -*- coding: utf-8 -*-
"""
catalogo_local.py — el catálogo completo (RECURSOS/catalogo-dali-completo.xlsx) como ÚNICO archivo de artículos
de la Consola (v10.2.50). Sin tkinter: se prueba en tests/test_catalogo_local.py.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Decisión de Víctor (13 nov 2026): prescindir de «LISTADO CODIGOS DALI - SAP.xlsx» y quedarse solo con el
catálogo completo, que ya tiene todo lo que tenía el LISTADO (y además proveedor, clasificación, código de
proveedor y precio). Hasta ahora las Acciones 2 y 3 rellenaban el LISTADO y el catálogo se reconstruía copiándolo.

Para cambiar lo menos posible la lógica que ya funciona, el catálogo se ofrece con la MISMA forma que tenía el
LISTADO: filas desde la 8, columnas A..O por posición (ver COLUMNAS_LISTADO). Así:
- quien solo leía el LISTADO (tarifas, RINFO) cambia la línea que abría el archivo por `abrir_como_listado`;
- quien escribía en él (Acción 3) trabaja sobre ese libro en memoria y luego se pasa al catálogo por código DALI
  (`tabla_sap_desde_libro` + regenerar el catálogo).
- La ventana de SAP faltantes escribe el código SAP directamente en el catálogo (`escribir_codigos_sap`).

Además se corrige un riesgo del LISTADO: la Acción 2 escribía los códigos DALI por POSICIÓN de fila y el código SAP
(col D) se quedaba en su fila; si DALI quitaba o intercalaba un código, los SAP de debajo se desplazaban. En el
catálogo todo va por código DALI.

`migrar_listado` pasa una sola vez los datos SAP del LISTADO al catálogo (el LISTADO manda, como hasta ahora en cada
Acción 2/3), deja un informe de diferencias y renombra el LISTADO a «(antiguo, ya no se usa)». No borra nada.
"""
import os
from datetime import datetime

from app.comun.sap_helpers import (
    CATALOGO_CLAVES_SAP, CATALOGO_FILA_CABECERA, CATALOGO_PRIMERA_FILA_DATOS,
    datos_sap_desde_fila_listado, mapa_columnas_catalogo,
)

NOMBRE_CATALOGO = "catalogo-dali-completo.xlsx"
CARPETA_RECURSOS = "RECURSOS"
NOMBRE_LISTADO = "LISTADO CODIGOS DALI - SAP.xlsx"
NOMBRE_LISTADO_ANTIGUO = "LISTADO CODIGOS DALI - SAP (antiguo, ya no se usa).xlsx"
FILA_INICIO_LISTADO = 8

# Columnas A..O del antiguo LISTADO → clave del catálogo (sap_helpers.CATALOGO_COLUMNAS).
COLUMNAS_LISTADO = (
    "clave",          # A  código DALI
    "descripcion",    # B  descripción DALI
    "unidad",         # C  unidad DALI
    "cod_sap",        # D  código SAP
    "desc_sap",       # E  descripción SAP
    "texto_um",       # F  texto UM (expl.)
    "unidad_sap",     # G  unidad SAP
    "grp_prod",       # H  grupo de productos
    "desc_grp_prod",  # I  descripción grupo
    "ind_imp",        # J  indicador de impuestos
    "categoria",      # K  categoría de valoración
    "tipo_prod",      # L  tipo de producto
    "grp_compras",    # M  grupo de compras
    "desc_grp_compras",  # N nombre grupo de compras
    "activo",         # O  activo DALI
)


# ── Localizar ─────────────────────────────────────────────────────────────────────────────────────────
def ruta_catalogo(carpeta_proyecto):
    return os.path.join(carpeta_proyecto, CARPETA_RECURSOS, NOMBRE_CATALOGO)


def localizar_catalogo(carpeta_proyecto=None, carpeta_tarifas=None):
    """Ruta del catálogo: RECURSOS del proyecto o, si no se conoce el proyecto, la carpeta padre de TARIFAS.
    None si no existe."""
    candidatas = []
    if carpeta_proyecto:
        candidatas.append(ruta_catalogo(carpeta_proyecto))
    if carpeta_tarifas:
        candidatas.append(ruta_catalogo(os.path.dirname(os.path.abspath(carpeta_tarifas))))
        candidatas.append(os.path.join(carpeta_tarifas, NOMBRE_CATALOGO))
    for ruta in candidatas:
        if os.path.isfile(ruta):
            return ruta
    return None


def ruta_listado(carpeta_proyecto):
    return os.path.join(carpeta_proyecto, NOMBRE_LISTADO)


# ── Leer ──────────────────────────────────────────────────────────────────────────────────────────────
def _como_numero(v):
    """«10100416» o 10100416.0 → 10100416 (como estaban en el LISTADO); lo demás, tal cual."""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str) and v.strip().isdigit():
        return int(v.strip())
    return v


def leer_catalogo(ruta):
    """[(clave_int, {clave_local: valor})] en el orden del archivo. Columnas por cabecera (fila 4)."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        ws = wb.active
        filas = ws.iter_rows(min_row=CATALOGO_FILA_CABECERA, values_only=True)
        mapa = mapa_columnas_catalogo(next(filas, None) or ())
        if "clave" not in mapa:
            raise ValueError(f"{os.path.basename(ruta)} no tiene la columna «Código DALI» en la fila 4")
        salida = []
        for fila in filas:
            if not fila:
                continue
            clave = _como_numero(fila[mapa["clave"] - 1] if mapa["clave"] <= len(fila) else None)
            if clave in (None, ""):
                continue
            datos = {k: (fila[i - 1] if i <= len(fila) else None) for k, i in mapa.items()}
            datos["clave"] = clave
            salida.append((clave, datos))
        return salida
    finally:
        wb.close()


_CACHE = {}


def filas_formato_listado(ruta):
    """Filas del catálogo como las del antiguo LISTADO: tuplas A..O (COLUMNAS_LISTADO). Códigos DALI y SAP como
    número, como estaban en el LISTADO. Se guarda en memoria mientras el archivo no cambie."""
    try:
        firma = (os.path.abspath(ruta), os.path.getmtime(ruta), os.path.getsize(ruta))
    except OSError:
        firma = None
    if firma and _CACHE.get("firma") == firma:
        return _CACHE["filas"]
    filas = []
    for _clave, d in leer_catalogo(ruta):
        fila = [d.get(k) for k in COLUMNAS_LISTADO]
        fila[3] = _como_numero(fila[3]) if fila[3] not in ("",) else None
        filas.append(tuple(fila))
    if firma:
        _CACHE.update(firma=firma, filas=filas)
    return filas


def libro_formato_listado(ruta):
    """Un libro de openpyxl EN MEMORIA con la forma del antiguo LISTADO (datos desde la fila 8, columnas A..O),
    sacado del catálogo. Sirve tanto para leer (`.active.iter_rows(min_row=8, values_only=True)`, `.close()`) como
    para que la Acción 3 escriba en él antes de pasar el resultado al catálogo."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Catálogo (formato LISTADO)"
    for n, fila in enumerate(filas_formato_listado(ruta), start=FILA_INICIO_LISTADO):
        for col, valor in enumerate(fila, start=1):
            if valor is not None:
                ws.cell(row=n, column=col, value=valor)
    return wb


abrir_como_listado = libro_formato_listado


def tabla_sap_desde_libro(ws):
    """{clave_str: datos SAP} de un libro con forma de LISTADO (lo que antes se leía del archivo al regenerar)."""
    tabla = {}
    for fila in ws.iter_rows(min_row=FILA_INICIO_LISTADO, values_only=True):
        if not fila or fila[0] is None:
            continue
        tabla[str(fila[0]).strip()] = datos_sap_desde_fila_listado(fila)
    return tabla


def leer_tabla_sap(ruta):
    """{clave_str: {campos SAP}} del catálogo (para regenerarlo sin el LISTADO)."""
    tabla = {}
    for clave, d in leer_catalogo(ruta):
        datos = {k: d.get(k) for k in CATALOGO_CLAVES_SAP}
        if datos.get("cod_sap") not in (None, ""):
            datos["cod_sap"] = str(_como_numero(datos["cod_sap"])).strip()
        else:
            datos["cod_sap"] = None
        tabla[str(clave).strip()] = datos
    return tabla


# ── Escribir ──────────────────────────────────────────────────────────────────────────────────────────
def abrir_para_editar(ruta):
    """(wb, ws, mapa_columnas) del catálogo para modificarlo y guardarlo con wb.save(ruta)."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta, rich_text=False)
    ws = wb.active
    cab = next(ws.iter_rows(min_row=CATALOGO_FILA_CABECERA, max_row=CATALOGO_FILA_CABECERA, values_only=True), ())
    mapa = mapa_columnas_catalogo(cab or ())
    if "clave" not in mapa or "cod_sap" not in mapa:
        wb.close()
        raise ValueError(f"{os.path.basename(ruta)} no tiene las columnas «Código DALI» y «Código SAP» en la fila 4")
    return wb, ws, mapa


def mapas_dali_sap(ws, mapa):
    """({dali_int: nº de fila}, {sap_int: (dali_int, nº de fila)}) de un catálogo abierto para editar."""
    fila_de, sap_de = {}, {}
    c_clave, c_sap = mapa["clave"], mapa["cod_sap"]
    for n, fila in enumerate(ws.iter_rows(min_row=CATALOGO_PRIMERA_FILA_DATOS, values_only=True),
                             start=CATALOGO_PRIMERA_FILA_DATOS):
        dali = _como_numero(fila[c_clave - 1] if len(fila) >= c_clave else None)
        if not isinstance(dali, int):
            continue
        fila_de[dali] = n
        sap = _como_numero(fila[c_sap - 1] if len(fila) >= c_sap else None)
        if isinstance(sap, int):
            sap_de[sap] = (dali, n)
    return fila_de, sap_de


def escribir_codigos_sap(ruta, codigos):
    """Escribe {dali_int: sap_int o None} en la columna «Código SAP» del catálogo, buscando por código DALI.
    Devuelve (escritos, no_encontrados)."""
    wb, ws, mapa = abrir_para_editar(ruta)
    try:
        fila_de, _ = mapas_dali_sap(ws, mapa)
        escritos, no_encontrados = 0, []
        for dali, sap in codigos.items():
            fila = fila_de.get(dali)
            if fila is None:
                no_encontrados.append(dali)
                continue
            ws.cell(row=fila, column=mapa["cod_sap"]).value = sap
            escritos += 1
        wb.save(ruta)
        return escritos, no_encontrados
    finally:
        wb.close()


def conservar_sap_local(ruta_nuevo, tabla_sap_local):
    """Tras descargar el catálogo de la web (que sustituye al local): los artículos que en el catálogo LOCAL tenían
    código SAP y en el descargado no tienen ninguno conservan sus datos SAP locales. Sin LISTADO, un código SAP
    escrito en la Consola que todavía no se haya enviado a la web («Ahora no», sin conexión) solo está en el
    catálogo local: así no se pierde. Si los dos tienen código y no coinciden, manda la web.
    Devuelve la lista de códigos DALI conservados."""
    con_sap = {c: d for c, d in tabla_sap_local.items() if d.get("cod_sap")}
    if not con_sap:
        return []
    wb, ws, mapa = abrir_para_editar(ruta_nuevo)
    try:
        conservados = []
        c_clave, c_sap = mapa["clave"], mapa["cod_sap"]
        for n, fila in enumerate(ws.iter_rows(min_row=CATALOGO_PRIMERA_FILA_DATOS, values_only=True),
                                 start=CATALOGO_PRIMERA_FILA_DATOS):
            clave = _como_numero(fila[c_clave - 1] if len(fila) >= c_clave else None)
            local = con_sap.get(str(clave).strip())
            if not local or (fila[c_sap - 1] if len(fila) >= c_sap else None) not in (None, ""):
                continue
            for campo in CATALOGO_CLAVES_SAP:
                if campo in mapa and local.get(campo) not in (None, ""):
                    valor = _como_numero(local[campo]) if campo == "cod_sap" else local[campo]
                    ws.cell(row=n, column=mapa[campo]).value = valor
            conservados.append(str(clave))
        if conservados:
            wb.save(ruta_nuevo)
        return conservados
    finally:
        wb.close()


# ── Paso único del LISTADO al catálogo ────────────────────────────────────────────────────────────────
def _igual(a, b):
    na = "" if a in (None, "") else str(_como_numero(a)).strip()
    nb = "" if b in (None, "") else str(_como_numero(b)).strip()
    return na == nb


def comparar_listado(ruta_listado_xlsx, ruta_cat):
    """Diferencias de datos SAP entre el LISTADO y el catálogo, por código DALI:
    [(dali, campo, valor_listado, valor_catalogo)] y los DALI del LISTADO que no están en el catálogo."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta_listado_xlsx, read_only=True, data_only=True)
    try:
        sap_listado = tabla_sap_desde_libro(wb.active)
    finally:
        wb.close()
    sap_catalogo = leer_tabla_sap(ruta_cat)
    diferencias, solo_listado = [], []
    for clave, datos in sap_listado.items():
        clave_n = str(_como_numero(clave)).strip()
        cat = sap_catalogo.get(clave_n)
        if cat is None:
            if datos.get("cod_sap"):
                solo_listado.append(clave_n)
            continue
        for campo in CATALOGO_CLAVES_SAP:
            if not _igual(datos.get(campo), cat.get(campo)):
                diferencias.append((clave_n, campo, datos.get(campo), cat.get(campo)))
    return diferencias, solo_listado, sap_listado


def migrar_listado(carpeta_proyecto, ahora=None):
    """Si existe el LISTADO en la raíz del proyecto: pasa sus datos SAP al catálogo (el LISTADO manda, como en cada
    Acción 2/3 hasta ahora), deja un informe en RECURSOS y renombra el LISTADO a «(antiguo, ya no se usa)».
    Devuelve None si no hay nada que hacer, o un dict con el resumen."""
    ruta_lst = ruta_listado(carpeta_proyecto)
    ruta_cat = ruta_catalogo(carpeta_proyecto)
    if not os.path.isfile(ruta_lst) or not os.path.isfile(ruta_cat):
        return None
    ahora = ahora or datetime.now()
    diferencias, solo_listado, sap_listado = comparar_listado(ruta_lst, ruta_cat)

    informe = None
    if diferencias or solo_listado:
        informe = os.path.join(carpeta_proyecto, CARPETA_RECURSOS,
                               f"paso_listado_a_catalogo_{ahora.strftime('%Y%m%d_%H%M%S')}.txt")
        with open(informe, "w", encoding="utf-8") as f:
            f.write("PASO DEL «LISTADO CODIGOS DALI - SAP» AL CATÁLOGO COMPLETO\n")
            f.write(f"Hecho el {ahora.strftime('%d/%m/%Y, %H:%M:%S')}\n\n")
            f.write("Desde esta versión la Consola usa solo catalogo-dali-completo.xlsx. Estas son las diferencias de\n"
                    "datos SAP que había entre los dos archivos. Se ha dejado en el catálogo lo que decía el LISTADO,\n"
                    "igual que hacían hasta ahora las Acciones 2 y 3. Revísalas por si alguna es un error.\n\n")
            f.write(f"{'Código DALI':<12} | {'Campo':<18} | {'LISTADO (queda)':<40} | Catálogo (antes)\n")
            f.write("-" * 110 + "\n")
            for dali, campo, v_lst, v_cat in diferencias:
                f.write(f"{dali:<12} | {campo:<18} | {str(v_lst or ''):<40} | {v_cat or ''}\n")
            if solo_listado:
                f.write("\nCódigos DALI con código SAP en el LISTADO que no están en el catálogo "
                        "(no se han pasado; se darán de alta con la Acción 2 si siguen en DALI):\n")
                f.write(", ".join(solo_listado) + "\n")

    if diferencias:
        wb, ws, mapa = abrir_para_editar(ruta_cat)
        try:
            fila_de, _ = mapas_dali_sap(ws, mapa)
            for dali, campo, v_lst, _v_cat in diferencias:
                fila = fila_de.get(int(dali)) if str(dali).isdigit() else None
                if fila and campo in mapa:
                    valor = _como_numero(v_lst) if campo == "cod_sap" else v_lst
                    ws.cell(row=fila, column=mapa[campo]).value = valor
            wb.save(ruta_cat)
        finally:
            wb.close()

    destino = os.path.join(carpeta_proyecto, NOMBRE_LISTADO_ANTIGUO)
    if os.path.exists(destino):
        destino = os.path.join(carpeta_proyecto,
                               f"LISTADO CODIGOS DALI - SAP (antiguo {ahora.strftime('%Y%m%d_%H%M%S')}).xlsx")
    os.replace(ruta_lst, destino)
    return {"diferencias": diferencias, "solo_listado": solo_listado, "informe": informe,
            "renombrado_a": destino, "entradas_listado": len(sap_listado),
            "articulos_cambiados": len({d for d, *_ in diferencias})}
