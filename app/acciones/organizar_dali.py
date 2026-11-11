"""
Acción 1 — Organizar artículos DALI: núcleo de datos.

Extraído de tarifa_app.py::_organizar_dali (v9.0.4) como primera tanda de
la reorganización. Cubre lectura, deduplicación de códigos de origen,
cruce por CODIGO y saneamiento — es decir, todo lo que va desde "Leyendo
ARTICULOS1.xls..." hasta "Saneamiento completado" en el log.

Deliberadamente NO incluye (todavía) la generación del Excel de salida,
estilos, ni los diccionarios de SUBFAMILIA/IGIC — eso queda como
Tanda 2, para no mover de golpe las ~600 líneas restantes sin poder
probarlas contra la interfaz real.

Este módulo no importa tkinter ni depende de ninguna instancia de App:
recibe una función `log(mensaje, nivel="INFO")` para reportar progreso,
igual que el `self.log_consola` original, así que puede conectarse
directamente desde tarifa_app.py sin cambiar el formato de log.
"""

from datetime import datetime

import pandas as pd


def _procesar_duplicados_origen(df_sel, nombre_archivo, cols_comparar, log, incidencias):
    """
    Detecta códigos duplicados en un dataframe de origen (ARTICULOS1 o
    ARTICULOS2), documenta en el log si los datos del resto de columnas
    coinciden o difieren entre repeticiones, y devuelve el dataframe con
    una única fila por CODIGO (se conserva la primera aparición).

    Las incidencias con datos DISTINTOS se añaden a la lista `incidencias`
    (pasada por referencia) para que el llamador pueda volcarlas a un log
    de incidencias y avisar al usuario.
    """
    codigo_num = pd.to_numeric(
        df_sel["CODIGO"].astype(str).str.strip(), errors="coerce")
    es_valido = codigo_num.notna() & (codigo_num > 0)

    dup_mask = df_sel["CODIGO"].duplicated(keep=False) & es_valido
    codigos_dup = sorted(
        set(df_sel.loc[dup_mask, "CODIGO"]),
        key=lambda x: (len(str(x)), str(x)))

    n_identicos = 0
    n_distintos = 0
    if codigos_dup:
        log(f"Comprobando códigos duplicados en {nombre_archivo}...", "INFO")
    for cod in codigos_dup:
        filas = df_sel[df_sel["CODIGO"] == cod]
        n_rep = len(filas)
        iguales = filas[cols_comparar].nunique(dropna=False).eq(1).all() \
            if cols_comparar else True
        if iguales:
            n_identicos += 1
            log(
                f"  └ {nombre_archivo}: código {cod} repetido {n_rep} "
                f"veces — datos idénticos, se conserva 1 sola vez",
                "INFO")
        else:
            n_distintos += 1
            log(
                f"  └ {nombre_archivo}: código {cod} repetido {n_rep} "
                f"veces — ⚠ datos DISTINTOS entre repeticiones, se "
                f"conserva la 1ª aparición", "WARN")
            columnas_distintas = {}
            for col in cols_comparar:
                if filas[col].nunique(dropna=False) > 1:
                    valores = filas[col].tolist()
                    columnas_distintas[col] = valores
                    log(
                        f"      · columna {col}: valores distintos → "
                        f"{valores}", "WARN")
            incidencias.append({
                "archivo": nombre_archivo,
                "codigo": cod,
                "n_rep": n_rep,
                "columnas": columnas_distintas,
            })

    # Códigos inválidos (0 / vacío / no numérico) también se deduplican
    # por eficiencia, pero no se documentan uno a uno porque se eliminarán
    # igualmente en el saneamiento posterior.
    df_dedup = df_sel.drop_duplicates(subset="CODIGO", keep="first")

    if n_identicos or n_distintos:
        nivel_resumen = "WARN" if n_distintos else "OK"
        log(
            f"  └ {nombre_archivo}: {n_identicos + n_distintos} "
            f"código(s) único(s) duplicado(s) "
            f"({n_identicos} con datos idénticos, {n_distintos} con "
            f"datos distintos) — normalizado a 1 fila por código",
            nivel_resumen)
    else:
        log(f"  └ {nombre_archivo}: sin códigos duplicados", "OK")

    return df_dedup


def procesar_datos_dali(file1, file2, log):
    """
    Lee ARTICULOS1.xls y ARTICULOS2.xls, deduplica códigos de origen,
    cruza por CODIGO y sanea el resultado.

    Parámetros
    ----------
    file1, file2 : str
        Rutas a ARTICULOS1.xls y ARTICULOS2.xls.
    log : callable(mensaje: str, nivel: str = "INFO")
        Función de log — en tarifa_app.py, pasar self.log_consola.

    Devuelve
    --------
    dict con:
        df_final              : DataFrame ya cruzado y saneado, columnas
                                 [CODIGO, DESCRIPCION, UNIDAD, FAMILIA,
                                  SUBFAMILIA, NATURALEZA, IGIC, PROVEEDOR,
                                  ACTIVO], listo para escribir a Excel.
        df1                    : DataFrame original de ARTICULOS1 (sin
                                  seleccionar columnas), por si se necesita
                                  para calcular el último código de origen.
        incidencias_dup        : lista de incidencias de duplicados con
                                  datos distintos (para el log de
                                  incidencias y el aviso al usuario).
        filas_elim_saneamiento : nº de filas eliminadas por CODIGO
                                  vacío/no numérico.
        errores_sin_igic       : nº de filas sin coincidencia en
                                  ARTICULOS2 (IGIC vacío).
    """
    log("Leyendo ARTICULOS1.xls...", "INFO")
    df1 = pd.read_excel(file1, engine="xlrd", dtype=str, header=None)
    log("Leyendo ARTICULOS2.xls...", "INFO")
    df2 = pd.read_excel(file2, engine="xlrd", dtype=str, header=None)

    df1.columns = df1.columns.astype(str)
    df2.columns = df2.columns.astype(str)

    log(f"  └ ARTICULOS1: {len(df1)} filas  |  ARTICULOS2: {len(df2)} filas", "INFO")

    # ── Renombrar columna 0 de ambos a CODIGO ────────────────────────
    col0_df1 = df1.columns[0]
    col0_df2 = df2.columns[0]
    df1 = df1.rename(columns={col0_df1: "CODIGO"})
    df2 = df2.rename(columns={col0_df2: "CODIGO"})

    # ── Seleccionar columnas necesarias de ARTICULOS2 ────────────────
    # cols: CODIGO(0), IGIC(3), PROVEEDOR(5), ACTIVO(7)
    cols2_idx = [0] + [i for i in [3, 5, 7] if i < len(df2.columns)]
    df2_sel = df2.iloc[:, cols2_idx].copy()
    df2_sel.columns = ["CODIGO"] + ["IGIC", "PROVEEDOR", "ACTIVO"][:len(cols2_idx) - 1]

    # ── Seleccionar columnas necesarias de ARTICULOS1 ────────────────
    # cols: CODIGO(0), DESCRIPCION(1), UNIDAD(2), FAMILIA(3), SUBFAMILIA(4), NATURALEZA(5)
    cols1 = ["CODIGO"]
    for nombre, idx in [("DESCRIPCION", 1), ("UNIDAD", 2),
                         ("FAMILIA", 3), ("SUBFAMILIA", 4), ("NATURALEZA", 5)]:
        if idx < len(df1.columns):
            df1 = df1.rename(columns={df1.columns[idx]: nombre})
            cols1.append(nombre)
    df1_sel = df1[cols1]

    # ── Detección y resolución de códigos duplicados en origen ───────
    incidencias_dup = []
    df1_sel = _procesar_duplicados_origen(
        df1_sel, "ARTICULOS1",
        [c for c in ["DESCRIPCION", "UNIDAD", "FAMILIA",
                     "SUBFAMILIA", "NATURALEZA"] if c in df1_sel.columns],
        log, incidencias_dup)

    df2_sel = _procesar_duplicados_origen(
        df2_sel, "ARTICULOS2",
        [c for c in ["IGIC", "PROVEEDOR", "ACTIVO"] if c in df2_sel.columns],
        log, incidencias_dup)

    # ── Cruce en memoria con pandas ───────────────────────────────────
    log("Cruzando datos por CODIGO...", "INFO")
    df_final = df1_sel.merge(df2_sel, on="CODIGO", how="left")

    # Columnas de salida en orden fijo
    output_cols = ["CODIGO", "DESCRIPCION", "UNIDAD", "FAMILIA",
                   "SUBFAMILIA", "NATURALEZA", "IGIC", "PROVEEDOR", "ACTIVO"]
    for col in output_cols:
        if col not in df_final.columns:
            df_final[col] = ""
    df_final = df_final[output_cols].fillna("")

    # ── SANEAMIENTO DE DATOS ──────────────────────────────────────────
    log("Saneando datos...", "INFO")
    filas_antes = len(df_final)

    # 1. CODIGO (col A): forzar numérico — eliminar filas donde CODIGO
    #    no sea numérico o esté vacío.
    df_final["CODIGO"] = pd.to_numeric(
        df_final["CODIGO"].astype(str).str.strip(), errors="coerce")
    df_final = df_final.dropna(subset=["CODIGO"])
    df_final["CODIGO"] = df_final["CODIGO"].astype(int)

    # Último código válido de ARTICULOS1 (col A), capturado ANTES de
    # filtrar por CODIGO>0 — se usa más adelante (fuera de este módulo,
    # todavía en tarifa_app.py) como techo en la detección de códigos
    # faltantes/huecos entre 1 y el último código.
    try:
        _codigos_art1 = pd.to_numeric(
            df1.iloc[:, 0].astype(str).str.strip(), errors="coerce").dropna()
        ultimo_cod_art1 = int(_codigos_art1[_codigos_art1 > 0].max())
    except Exception:
        ultimo_cod_art1 = None

    # 2. Eliminar filas donde CODIGO == 0 o negativo (no válidos)
    df_final = df_final[df_final["CODIGO"] > 0]

    # 2b. Ordenar por CODIGO ascendente. El merge conserva el orden de
    #     ARTICULOS1.xls tal cual venga (puede llegar en cualquier orden,
    #     incluso descendente, según cómo se haya exportado el origen).
    #     Se fuerza aquí un orden numérico ascendente estable para que
    #     ARTICULOS_DALI.xlsx (y todo lo que se genera después a partir
    #     de él: LISTADO CODIGOS DALI-SAP, etc.) sea siempre consistente
    #     independientemente del orden del fichero de origen.
    df_final = df_final.sort_values("CODIGO", ascending=True).reset_index(drop=True)

    filas_elim = filas_antes - len(df_final)
    if filas_elim:
        log(f"  └ {filas_elim} filas eliminadas (CODIGO vacío/no numérico)", "WARN")

    # 3. Normalizar espacios múltiples en TODAS las columnas de texto
    cols_texto = ["DESCRIPCION", "UNIDAD", "FAMILIA",
                  "SUBFAMILIA", "NATURALEZA", "IGIC", "PROVEEDOR", "ACTIVO"]
    for col in cols_texto:
        if col in df_final.columns:
            df_final[col] = (df_final[col]
                              .astype(str)
                              .str.replace(r"[ \t]+", " ", regex=True)
                              .str.strip()
                              .replace("nan", ""))

    # 4. ACTIVO (col I): True → "NO" (dado de baja) / False → "SI" (activo)
    if "ACTIVO" in df_final.columns:
        df_final["ACTIVO"] = (df_final["ACTIVO"]
                               .astype(str)
                               .str.replace(r"\s+", "", regex=True))
        df_final["ACTIVO"] = (df_final["ACTIVO"]
                               .str.upper()
                               .replace({"TRUE": "NO", "FALSE": "SI",
                                         "1": "NO", "0": "SI",
                                         "1.0": "NO", "0.0": "SI"}))

    log(f"  └ Saneamiento completado — {len(df_final)} filas válidas", "OK")

    errores_sin_igic = int(df_final["IGIC"].eq("").sum())
    log(
        f"  └ {len(df_final)} filas cruzadas  |  "
        f"{errores_sin_igic} sin coincidencia en ARTICULOS2", "INFO")

    return {
        "df_final": df_final,
        "df1": df1,
        "ultimo_cod_art1": ultimo_cod_art1,
        "incidencias_dup": incidencias_dup,
        "filas_elim_saneamiento": filas_elim,
        "errores_sin_igic": errores_sin_igic,
        "output_cols": output_cols,
    }


def generar_excel_dali(df_final, output_cols, output_path, log, progreso=None):
    """
    Construye y guarda ARTICULOS_DALI.xlsx: cabecera info, logo, tabla con
    estilos, anchos de columna. No hace ninguna validación de negocio —
    eso lo cubre validar_diccionarios_sap() por separado.

    Parámetros
    ----------
    df_final    : DataFrame ya cruzado y saneado (salida de procesar_datos_dali).
    output_cols : lista de nombres de columna en el orden de escritura.
    output_path : ruta del ARTICULOS_DALI.xlsx a generar.
    log         : callable(mensaje, nivel="INFO") — pasar self.log_consola.
    progreso    : callable(texto) opcional — pasar
                  lambda t: self.after(0, lambda: self._barra.indeterminate(t))
    """
    from datetime import datetime as _dt
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from app.comun.tarifa_core import resource_path

    if progreso:
        progreso("Generando Excel...")

    wb = Workbook()
    ws = wb.active
    ws.title = "DALI"

    # ── Alturas de fila cabecera ──────────────────────────────────────
    for r in (1, 2, 3, 5):
        ws.row_dimensions[r].height = 18
    ws.row_dimensions[4].height = 6
    ws.row_dimensions[6].height = 10
    ws.row_dimensions[7].height = 18

    # ── Cabecera info (filas 1-3) ─────────────────────────────────────
    _linea  = Side(style="thin", color="000000")
    _alin_l = Alignment(horizontal="left",   vertical="center")

    _hoy  = _dt.now().strftime("%d/%m/%Y")
    _hora = _dt.now().strftime("%H:%M")

    for r in range(1, 6):
        for c in range(2, 7):
            ws.cell(r, c).value = None

    c_titulo = ws.cell(2, 2, "LISTADO ARTICULOS DALI")
    c_titulo.font      = Font(name="Arial", size=18, bold=True)
    c_titulo.alignment = _alin_l

    c_fecha = ws.cell(3, 2, f"Actualizado: {_hoy}")
    c_fecha.font      = Font(name="Arial", size=11, bold=True)
    c_fecha.alignment = _alin_l

    c_hora = ws.cell(5, 2, f"Hora: {_hora}")
    c_hora.font      = Font(name="Arial", size=11, bold=True)
    c_hora.alignment = _alin_l

    # ── Logo Princess en A1:A5 (negativo — negro sobre transparente) ──
    log("Insertando logo...", "INFO")
    try:
        xl_img        = XLImage(resource_path("assets/logo_dali_articulos.png"))
        xl_img.width  = 110
        xl_img.height = 90
        ws.add_image(xl_img, "A1")
        log("  └ Logo insertado en A1:A5", "OK")
    except Exception as e:
        log(f"  └ Logo omitido ({e})", "WARN")

    # ── Cabecera tabla (fila 7) ───────────────────────────────────────
    _font_hdr  = Font(name="Arial", size=12, bold=True)
    _fill_hdr  = PatternFill(start_color="D9D9D9",
                             end_color="D9D9D9", fill_type="solid")
    _alin_c    = Alignment(horizontal="center", vertical="center")
    _borde_all = Border(left=_linea, right=_linea, top=_linea, bottom=_linea)

    for col_idx, h in enumerate(output_cols, start=1):
        cell           = ws.cell(7, col_idx, h)
        cell.font      = _font_hdr
        cell.fill      = _fill_hdr
        cell.alignment = _alin_c
        cell.border    = _borde_all

    ws.auto_filter.ref = f"A7:{get_column_letter(len(output_cols))}7"
    ws.freeze_panes    = "A8"

    # ── Escritura masiva de datos con ws.append() ─────────────────────
    log("Escribiendo datos...", "INFO")
    for row_vals in df_final.itertuples(index=False, name=None):
        ws.append(row_vals)

    # ── Aplicar estilos POR RANGO (no celda a celda) ──────────────────
    log("Aplicando estilos...", "INFO")
    ultima = ws.max_row
    _font_dato = Font(name="Arial", size=11)
    _alin_c11  = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=8, max_row=ultima,
                             min_col=1, max_col=len(output_cols)):
        for cell in row:
            cell.font      = _font_dato
            cell.border    = _borde_all
            cell.alignment = _alin_c11
            if cell.column == 1:
                cell.number_format = "000000"

    _anchos_fijos = {
        "A": 16, "B": 90, "C": 14, "D": 33, "E": 49,
        "F": 68, "G": 15, "H": 60, "I": 13,
    }
    for letra, ancho in _anchos_fijos.items():
        ws.column_dimensions[letra].width = ancho

    for r in range(8, ultima + 1):
        ws.row_dimensions[r].height = 15

    log("Guardando archivo...", "INFO")
    wb.save(output_path)


def eliminar_archivos_origen(file1, file2, log, motivo="limpieza"):
    """
    Elimina ARTICULOS1.xls / ARTICULOS2.xls tras usarlos. `motivo` cambia
    el texto del log ("limpieza" tras éxito, "match fallido" si la
    validación de diccionarios SAP falla).
    """
    import os
    eliminados_ok = []
    for fx in (file1, file2):
        try:
            if os.path.exists(fx):
                os.remove(fx)
                eliminados_ok.append(os.path.basename(fx))
                log(f"  └ Eliminado ({motivo}): {os.path.basename(fx)}", "INFO" if motivo == "limpieza" else "WARN")
        except Exception as ex_del:
            log(f"  └ No se pudo eliminar {os.path.basename(fx)}: {ex_del}", "WARN")
    return eliminados_ok


def validar_diccionarios_sap(df_final, log):
    """
    Valida las columnas SUBFAMILIA (E) e IGIC (G) contra los catálogos
    DALI_CLAVES_SUBFAMILIA / DALI_CLAVES_IGIC de tarifa_constants.py.
    Documenta el % de coincidencia en el log, igual que antes.

    Devuelve un dict con e_pct, g_pct, e_estado, g_estado y
    columnas_sin_match (lista de strings para el mensaje de error, vacía
    si todo coincide) — la DECISIÓN de qué hacer si hay 0% de coincidencia
    (borrar ficheros, avisar al usuario) se queda en tarifa_app.py, no aquí.
    """
    from app.comun.tarifa_constants import DALI_CLAVES_SUBFAMILIA, DALI_CLAVES_IGIC

    resultado = {"e_pct": 0, "g_pct": 0, "e_estado": "SIN DATOS",
                 "g_estado": "SIN DATOS", "columnas_sin_match": []}

    _total = len(df_final)
    if _total == 0:
        return resultado

    _claves_subfam_u = {k.upper() for k in DALI_CLAVES_SUBFAMILIA}
    _claves_igic_u = {k.upper() for k in DALI_CLAVES_IGIC}

    # ── Col E — SUBFAMILIA ────────────────────────────────────────────
    _col_e = (df_final["SUBFAMILIA"].astype(str).str.strip().str.upper()
              if "SUBFAMILIA" in df_final.columns else None)
    if _col_e is not None:
        _e_total     = len(_col_e)
        _e_vacios    = int((_col_e == "").sum())
        _e_con_valor = _e_total - _e_vacios
        _e_match     = int(_col_e[_col_e != ""].isin(_claves_subfam_u).sum())
        _e_pct       = round(_e_match / _e_total * 100, 1) if _e_total else 0
        if _e_con_valor == 0:
            _e_estado = "SIN DATOS"
        elif _e_pct == 0:
            _e_estado = "0 % coincidencia"
        else:
            _e_estado = f"{_e_pct} % coincidencia"
        log(
            f"  └ Col E (SUBFAMILIA): {_e_match}/{_e_total} filas "
            f"en diccionario  [{_e_estado}]",
            "OK" if _e_pct >= 50 else "WARN" if _e_pct > 0 else "ERROR")
    else:
        _e_pct = 0
        _e_estado = "COLUMNA AUSENTE"

    # ── Col G — IGIC ──────────────────────────────────────────────────
    _col_g = (df_final["IGIC"].astype(str).str.strip().str.upper()
              if "IGIC" in df_final.columns else None)
    if _col_g is not None:
        _g_total     = len(_col_g)
        _g_vacios    = int((_col_g == "").sum())
        _g_con_valor = _g_total - _g_vacios
        _g_match     = int(_col_g[_col_g != ""].isin(_claves_igic_u).sum())
        _g_pct       = round(_g_match / _g_total * 100, 1) if _g_total else 0
        if _g_con_valor == 0:
            _g_estado = "SIN DATOS"
        elif _g_pct == 0:
            _g_estado = "0 % coincidencia"
        else:
            _g_estado = f"{_g_pct} % coincidencia"
        log(
            f"  └ Col G (IGIC):       {_g_match}/{_g_total} filas "
            f"en diccionario  [{_g_estado}]",
            "OK" if _g_pct >= 50 else "WARN" if _g_pct > 0 else "ERROR")
    else:
        _g_pct = 0
        _g_estado = "COLUMNA AUSENTE"

    # ── Muestras de valores no reconocidos (para el log) ──────────────
    if _col_e is not None and _e_pct < 100:
        _e_desconocidos = sorted(
            set(_col_e[(_col_e != "") & (~_col_e.isin(_claves_subfam_u))].tolist())
        )[:5]
        if _e_desconocidos:
            log(f"      Subfamilias no reconocidas (muestra): "
                f"{', '.join(_e_desconocidos)}", "WARN")
    if _col_g is not None and _g_pct < 100:
        _g_desconocidos = sorted(
            set(_col_g[(_col_g != "") & (~_col_g.isin(_claves_igic_u))].tolist())
        )[:5]
        if _g_desconocidos:
            log(f"      Valores IGIC no reconocidos (muestra): "
                f"{', '.join(_g_desconocidos)}", "WARN")

    columnas_sin_match = []
    if _e_pct == 0:
        columnas_sin_match.append(f"E — SUBFAMILIA  ({_e_estado})")
    if _g_pct == 0:
        columnas_sin_match.append(f"G — IGIC  ({_g_estado})")

    if columnas_sin_match:
        log("⚠  ATENCIÓN: columnas sin coincidencia con diccionarios SAP — "
            "verifique filtros opción 131 de DALI", "ERROR")

    return {
        "e_pct": _e_pct, "g_pct": _g_pct,
        "e_estado": _e_estado, "g_estado": _g_estado,
        "columnas_sin_match": columnas_sin_match,
    }


def detectar_codigos_faltantes(df_final, ultimo_cod_art1, log):
    """
    Compara los códigos transferidos contra el rango completo 1..techo
    (techo = ultimo_cod_art1, o el máximo transferido si no se conoce) y
    documenta en el log los huecos encontrados, agrupados en tramos
    consecutivos.

    Devuelve (saltos, faltantes_lista):
        saltos          : lista de tuplas (desde, hasta, n) por tramo.
        faltantes_lista : lista plana de todos los códigos faltantes.
    """
    saltos = []
    faltantes_lista = []
    techo = ultimo_cod_art1
    try:
        codigos_transferidos = set(df_final["CODIGO"].astype(int).tolist())
        techo = ultimo_cod_art1 if ultimo_cod_art1 else max(codigos_transferidos)
        for cod in range(1, techo + 1):
            if cod not in codigos_transferidos:
                faltantes_lista.append(cod)
        if faltantes_lista:
            inicio = faltantes_lista[0]
            prev   = faltantes_lista[0]
            for cod in faltantes_lista[1:]:
                if cod != prev + 1:
                    saltos.append((inicio, prev, prev - inicio + 1))
                    inicio = cod
                prev = cod
            saltos.append((inicio, prev, prev - inicio + 1))
    except Exception:
        pass

    if faltantes_lista:
        log(f"  └ ✘ {len(faltantes_lista)} código(s) faltante(s) "
            f"entre 000001 y {techo:06d}", "ERROR")
        for desde, hasta, n in saltos[:10]:
            if desde == hasta:
                log(f"      {desde:06d}  ({n} ausente)", "ERROR")
            else:
                log(f"      {desde:06d} – {hasta:06d}  ({n} ausente(s))", "ERROR")
        if len(saltos) > 10:
            log(f"      … y {len(saltos)-10} grupo(s) más", "ERROR")

    return saltos, faltantes_lista


def guardar_log_incidencias(incidencias_dup, recursos_path, log):
    """
    Vuelca las incidencias de duplicados con datos distintos a un fichero
    de texto en recursos_path. Devuelve la ruta del fichero generado, o
    None si no había incidencias o no se pudo guardar.
    """
    if not incidencias_dup:
        return None

    import os

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre = f"INCIDENCIAS_DUPLICADOS_DALI_{ts}.txt"
    ruta = os.path.join(recursos_path, nombre)
    try:
        with open(ruta, "w", encoding="utf-8") as f:
            f.write("── INCIDENCIAS DE CÓDIGOS DUPLICADOS — ORGANIZACIÓN DALI ──\n")
            f.write(f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write(f"Carpeta: {recursos_path}\n")
            f.write(f"Total incidencias: {len(incidencias_dup)}\n")
            f.write("─" * 60 + "\n\n")
            for inc in incidencias_dup:
                f.write(
                    f"Archivo: {inc['archivo']}  |  "
                    f"Código: {inc['codigo']}  |  "
                    f"Repeticiones: {inc['n_rep']}\n")
                for col, valores in inc["columnas"].items():
                    f.write(f"    · Columna {col}: valores distintos → {valores}\n")
                f.write(
                    "    → Se ha conservado la 1ª aparición al generar "
                    "ARTICULOS_DALI.xlsx.\n")
                f.write(
                    "    → Revisar y corregir el dato correcto en DALI.\n\n")
        log(f"  └ Log de incidencias guardado: {nombre}", "OK")
        return ruta
    except Exception as e:
        log(f"  └ No se pudo guardar el log de incidencias: {e}", "WARN")
        return None
