# -*- coding: utf-8 -*-
"""
comprobar_purchasing.py — ¿Está al día el Purchasing que va a usar «Crear / Actualizar Registro Info SAP»? (sin tkinter)

Petición de Víctor (23 nov 2026), a raíz de su pregunta «cuando una tarifa tenga todos los registros info, ¿ya no sería
necesario el Purchasing?»: el Purchasing sigue haciendo falta (periodos de precio, datos de cada centro, bajas y la
estructura que SAP acepta), pero tiene que estar descargado DESPUÉS de la última importación en SAP. Si no, los
materiales que el RINFO anterior ya creó salen otra vez como nuevos y SAP contesta «Ya existe el registro Info de
compras…» (caso GRUPO DISTRIBUIDOR GARAU, 19 nov 2026). Antes de seguir, la Consola avisa si:

  1. la TARIFA lleva en REGISTRO INFO (columna H) números que no están en el Purchasing → el Purchasing es anterior a
     esos registros (o es de otro proveedor);
  2. en TARIFAS hay un RINFO de este proveedor guardado DESPUÉS que el Purchasing → si ya se importó en SAP, el
     Purchasing no lo refleja.

Las dos fechas son las de modificación de los archivos, las dos puestas por la misma unidad (G:), así que se comparan
entre sí y nunca con la hora del PC. No se para nada: es un aviso en la ventana «Purchasing encontrado», donde ya se
decide si se sigue con él.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import datetime as _dt
import os
import re

from app.comun.datos_tarifa import CAB_REGISTRO_INFO, COL_REGISTRO_INFO
from app.comun.sap_helpers import normalizar_codigo_sap

HOJA_1 = "Info general p.reg.info"
_RE_RINFO = re.compile(r"^RINFO\s.+\.xlsx$", re.IGNORECASE)


def _txt(v):
    if v is None:
        return ""
    return str(normalizar_codigo_sap(v)).strip()


def registros_de_tarifa(ruta_tarifa):
    """{código SAP: Registro Info} de la columna REGISTRO INFO de la TARIFA ({} si no la tiene)."""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(ruta_tarifa, read_only=True, data_only=True)
    except Exception:  # noqa: BLE001
        return {}
    try:
        ws = wb["TARIFA"] if "TARIFA" in wb.sheetnames else wb.active
        salida, cabecera_ok = {}, False
        for n, fila in enumerate(ws.iter_rows(min_row=7, values_only=True), start=7):
            if n == 7:
                cabecera_ok = (len(fila) >= COL_REGISTRO_INFO and
                               _txt(fila[COL_REGISTRO_INFO - 1]).upper() == CAB_REGISTRO_INFO)
                if not cabecera_ok:
                    return {}
                continue
            if not fila or len(fila) < COL_REGISTRO_INFO:
                continue
            sap, reg = _txt(fila[0]), _txt(fila[COL_REGISTRO_INFO - 1])
            if sap and reg:
                salida[sap] = reg
        return salida
    finally:
        wb.close()


def _hoja1(ruta):
    """Filas de la hoja 1 de un Purchasing / RINFO (en solo lectura: el de SAP no se abre en modo normal)."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        if HOJA_1 not in wb.sheetnames:
            return []
        return [f for f in wb[HOJA_1].iter_rows(min_row=2, values_only=True) if f and any(f)]
    finally:
        wb.close()


def registros_de_purchasing(ruta_purchasing):
    """Los Registros Info (columna A de la hoja 1) que tiene el Purchasing."""
    return {_txt(f[0]) for f in _hoja1(ruta_purchasing) if _txt(f[0])}


def _mtime(ruta):
    try:
        return os.path.getmtime(ruta)
    except OSError:
        return None


def _es_del_proveedor(ruta_rinfo, codigo_proveedor, nombre_proveedor):
    codigo = _txt(codigo_proveedor)
    nombre = " ".join(str(nombre_proveedor or "").upper().split())
    try:
        for f in _hoja1(ruta_rinfo)[:50]:
            if codigo and len(f) > 2 and _txt(f[2]) == codigo:
                return True
            if nombre and len(f) > 3 and " ".join(str(f[3] or "").upper().split()) == nombre:
                return True
    except Exception:  # noqa: BLE001 — un RINFO que no se puede leer no cuenta
        return False
    return False


def rinfo_posteriores(carpeta_tarifas, ruta_purchasing, codigo_proveedor=None, nombre_proveedor=None):
    """[(nombre de archivo, fecha)] de los RINFO de este proveedor guardados después que el Purchasing, del más
    reciente al más antiguo."""
    t_pur = _mtime(ruta_purchasing)
    if t_pur is None:
        return []
    salida = []
    try:
        entradas = list(os.scandir(carpeta_tarifas))
    except OSError:
        return []
    for e in entradas:
        if not e.is_file() or e.name.startswith("~$") or not _RE_RINFO.match(e.name):
            continue
        t = _mtime(e.path)
        if t is None or t <= t_pur:
            continue
        if _es_del_proveedor(e.path, codigo_proveedor, nombre_proveedor):
            salida.append((e.name, _dt.datetime.fromtimestamp(t)))
    return sorted(salida, key=lambda x: x[1], reverse=True)


def comprobar(ruta_tarifa, ruta_purchasing, carpeta_tarifas, codigo_proveedor=None, nombre_proveedor=None):
    """{"faltan": [(sap, registro)] de la tarifa que no están en el Purchasing,
        "rinfo_posteriores": [(archivo, fecha)]}"""
    faltan = []
    en_tarifa = registros_de_tarifa(ruta_tarifa)
    if en_tarifa:
        try:
            en_purchasing = registros_de_purchasing(ruta_purchasing)
        except Exception:  # noqa: BLE001
            en_purchasing = None
        if en_purchasing is not None:
            faltan = sorted((s, r) for s, r in en_tarifa.items() if r not in en_purchasing)
    return {"faltan": faltan,
            "rinfo_posteriores": rinfo_posteriores(carpeta_tarifas, ruta_purchasing,
                                                   codigo_proveedor, nombre_proveedor)}


def textos_aviso(resultado, max_codigos=8):
    """Las líneas del aviso (ventana y consola). [] si todo está bien."""
    lineas = []
    faltan = resultado.get("faltan") or []
    if faltan:
        muestra = ", ".join(f"{s} ({r})" for s, r in faltan[:max_codigos])
        mas = f" y {len(faltan) - max_codigos} más" if len(faltan) > max_codigos else ""
        lineas.append(f"La TARIFA tiene {len(faltan)} Registro(s) Info que no están en este Purchasing: "
                      f"{muestra}{mas}. El Purchasing es anterior a esos registros.")
    posteriores = resultado.get("rinfo_posteriores") or []
    if posteriores:
        nombre, fecha = posteriores[0]
        extra = f" (y {len(posteriores) - 1} más)" if len(posteriores) > 1 else ""
        lineas.append(f"Hay un RINFO de este proveedor guardado después que el Purchasing: «{nombre}», "
                      f"{fecha:%d/%m/%Y %H:%M}{extra}. Si ya lo importaste en SAP, este Purchasing no lo refleja.")
    if lineas:
        lineas.append("Descarga de nuevo el Purchasing de SAP y guárdalo en TARIFAS antes de seguir; si no, SAP "
                      "puede contestar «Ya existe el registro Info de compras…».")
    return lineas
