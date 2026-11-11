# -*- coding: utf-8 -*-
"""
accion_listado_dali.py
Acción 2 — Actualizar el listado de artículos DALI.

Mixin de App (tarifa_app.py): sus métodos se usan como self.metodo(). Separado de tarifa_app.py
en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""

import os
from tkinter import messagebox
import threading
from app.comun.tarifa_core import (
    FOLDER_RECURSOS,
)


class AccionListadoDali:
    def _actualizar_listado_dali(self):
        """
        PUNTO 2 — Actualizar listado artículos DALI.
        Origen : RECURSOS/ARTICULOS_DALI.xlsx  (generado por Punto 1)
        Destino: <raíz proyecto>/LISTADO CODIGOS DALI - SAP.xlsx
        Mapeo  : A→A, B→B, C→C (desde fila 8), col I (ACTIVO) → col O (ACTIVO DALI)
        (mapeo actualizado 2026-08: tras insertar las 3 columnas nuevas
        F, K, L en LISTADO CODIGOS DALI-SAP, la antigua col L (ACTIVO)
        se desplazó a la col O -- ADEMÁS, escribir en L causaba que
        Acción 3 lo borrara en su "Limpieza previa" (rango E:N) para
        cualquier material sin match SAP, dejando ACTIVO vacío)
        """
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            messagebox.showwarning("Sin proyecto",
                                   "Seleccione primero la carpeta del Proyecto.")
            return

        if not self._confirmar_catalogo_actualizado():
            self.log_consola(
                "Actualización DALI cancelada por el usuario (catálogo desactualizado).", "WARN")
            return

        # ── Rutas ────────────────────────────────────────────────
        ruta_origen  = os.path.join(carpeta, FOLDER_RECURSOS, "ARTICULOS_DALI.xlsx")
        ruta_destino = os.path.join(carpeta, "LISTADO CODIGOS DALI - SAP.xlsx")

        # ── Verificar ARTICULOS_DALI.xlsx ────────────────────────
        if not os.path.isfile(ruta_origen):
            messagebox.showwarning(
                "Archivo no encontrado",
                "\u274c  No se encontró el archivo:\n\n"
                f"   {ruta_origen}\n\n"
                "Debe ejecutar primero la opción\n"
                "\u25b6  1 — Organizar artículos DALI\n\n"
                "para generar este archivo antes de continuar.")
            return

        # ── Leer fecha y hora del encabezado (celdas B3 y B5) ───────
        try:
            import openpyxl
            wb_orig = openpyxl.load_workbook(ruta_origen, read_only=True, data_only=True)
            ws_orig = wb_orig.active
            celda_b3 = ws_orig["B3"].value or ""
            celda_b5 = ws_orig["B5"].value or ""
            wb_orig.close()
        except Exception as e:
            messagebox.showerror("Error al leer origen",
                                 f"No se pudo leer ARTICULOS_DALI.xlsx:\n{e}")
            return

        # Extraer solo el valor limpio eliminando el prefijo de texto
        # Ej: "Actualizado: 08/04/2026" -> "08/04/2026" ; "Hora: 08:47" -> "08:47"
        def _extraer_valor(texto):
            raw = str(texto).strip() if texto else ""
            return raw.split(":", 1)[-1].strip() if ":" in raw else raw

        texto_fecha    = str(celda_b3).strip() if celda_b3 else "(fecha no disponible)"
        valor_fecha_c4 = _extraer_valor(celda_b3) or "(no disponible)"
        valor_hora_c5  = _extraer_valor(celda_b5) or "(no disponible)"

        # ── Confirmación al usuario ───────────────────────────────
        if not messagebox.askyesno(
                "Confirmar actualización DALI",
                f"Archivo origen encontrado:\n"
                f"   ARTICULOS_DALI.xlsx\n\n"
                f"   {texto_fecha}\n\n"
                f"Destino:\n"
                f"   LISTADO CODIGOS DALI - SAP.xlsx\n\n"
                "¿Desea continuar?"):
            self.log_consola("Actualización DALI cancelada por el usuario.", "WARN")
            return

        # ── Verificar destino ─────────────────────────────────────
        if not os.path.isfile(ruta_destino):
            messagebox.showerror(
                "Destino no encontrado",
                "\u274c  No se encontró el archivo destino:\n\n"
                f"   {ruta_destino}\n\n"
                "Verifique que el archivo 'LISTADO CODIGOS DALI - SAP.xlsx'\n"
                "esté en la raíz de la carpeta del proyecto.")
            return

        # ── Proceso en hilo ──────────────────────────────────────
        def proceso():
            self.consola.separador("INICIO ACTUALIZACIÓN DALI")
            self.log_consola("Iniciando actualización listado artículos DALI...", "INFO")
            self.after(0, lambda: self._barra.indeterminate("Leyendo ARTICULOS_DALI..."))

            try:
                import openpyxl

                # -- Leer origen (solo valores, sin fórmulas) --
                self.log_consola(f"Leyendo origen: {os.path.basename(ruta_origen)}", "INFO")
                wb_orig = openpyxl.load_workbook(ruta_origen, read_only=True, data_only=True)
                ws_orig = wb_orig.active

                # Recorrer desde fila 8 hasta la última con valor en col A
                FILA_INI = 8
                datos = []  # lista de (cod_a, val_b, val_c, activo_i)
                for fila in ws_orig.iter_rows(min_row=FILA_INI, values_only=True):
                    val_a = fila[0]  # col A (índice 0)
                    if val_a is None or str(val_a).strip() == "":
                        break        # primera fila vacía en A → fin de datos
                    val_b    = fila[1]  # col B
                    val_c    = fila[2]  # col C
                    activo_i = fila[8] if len(fila) > 8 else None  # col I (índice 8)
                    datos.append((val_a, val_b, val_c, activo_i))

                wb_orig.close()
                n_filas = len(datos)
                self.log_consola(f"Filas leídas desde origen: {n_filas}", "OK")

                if n_filas == 0:
                    self.after(0, lambda: self._barra.reset("Sin datos en origen"))
                    self.after(0, lambda: messagebox.showwarning(
                        "Sin datos",
                        "No se encontraron filas con datos a partir de la fila 8\n"
                        "en el archivo ARTICULOS_DALI.xlsx."))
                    return

                # -- Abrir destino (lectura/escritura, conservando formato) --
                self.log_consola(f"Abriendo destino: {os.path.basename(ruta_destino)}", "INFO")
                self.after(0, lambda: self._barra.indeterminate("Escribiendo destino..."))

                wb_dest = openpyxl.load_workbook(ruta_destino)
                ws_dest = wb_dest.active  # primera hoja

                # -- Escribir datos fila a fila --
                for i, (val_a, val_b, val_c, activo_i) in enumerate(datos):
                    fila_dest = FILA_INI + i
                    ws_dest.cell(row=fila_dest, column=1,  value=val_a)    # A → A
                    ws_dest.cell(row=fila_dest, column=2,  value=val_b)    # B → B
                    ws_dest.cell(row=fila_dest, column=3,  value=val_c)    # C → C
                    ws_dest.cell(row=fila_dest, column=15, value=activo_i) # I → O (col 15)

                # -- Eliminar físicamente filas sobrantes en destino (si había más que ahora) --
                #    ANTES se vaciaban solo las columnas A/B/C/L celda a celda, pero eso
                #    dejaba intactas las columnas D-K de ejecuciones anteriores y, además,
                #    ws.max_row seguía «viendo» esas filas como ocupadas (arrastraba el
                #    recuento a Acción 3, que volvía a cruzar filas fantasma). Se eliminan
                #    las filas por completo con delete_rows(), que limpia TODAS las columnas
                #    de una vez y corrige max_row para las acciones posteriores.
                ultima_fila_dest = ws_dest.max_row
                fila_fin_nuevo   = FILA_INI + n_filas - 1
                if ultima_fila_dest > fila_fin_nuevo:
                    n_sobrantes = ultima_fila_dest - fila_fin_nuevo
                    ws_dest.delete_rows(fila_fin_nuevo + 1, n_sobrantes)
                    self.log_consola(
                        f"Eliminadas físicamente {n_sobrantes} fila(s) sobrante(s) en "
                        f"destino (todas las columnas, no solo A/B/C/L).",
                        "WARN")

                # -- Ajustar Tabla1: rango fila 7 (cabecera) hasta última fila con datos --
                # -- y aplicar formato uniforme a todas las filas de datos (fila 8 en adelante) --
                try:
                    from openpyxl.styles import Font, Alignment, Border, Side

                    # ── 1. Redimensionar Tabla1 ────────────────────────────────
                    _fila_cab   = 7                    # cabecera siempre en fila 7
                    _fila_datos = fila_fin_nuevo       # última fila con datos reales

                    if "Tabla1" in ws_dest.tables:
                        tbl = ws_dest.tables["Tabla1"]
                        # Obtener número de columnas actual del rango de la tabla
                        import re as _re_tbl
                        _m = _re_tbl.match(r"([A-Z]+)\d+:([A-Z]+)\d+", tbl.ref)
                        if _m:
                            _col_ini_ltr = _m.group(1)
                            _col_fin_ltr = _m.group(2)
                        else:
                            _col_ini_ltr = "A"
                            _col_fin_ltr = "L"   # fallback: 12 columnas (A..L)
                        _nuevo_ref = (
                            f"{_col_ini_ltr}{_fila_cab}:"
                            f"{_col_fin_ltr}{_fila_datos}"
                        )
                        tbl.ref = _nuevo_ref
                        # El AutoFilter interno de la tabla no se actualiza solo al
                        # cambiar tbl.ref; si se deja con el rango antiguo, Excel
                        # sigue "recordando" filas que ya no existen en la tabla.
                        if tbl.autoFilter is not None:
                            tbl.autoFilter.ref = _nuevo_ref
                        self.log_consola(
                            f"  └ Tabla1 redimensionada → {_nuevo_ref}", "OK")
                    else:
                        self.log_consola(
                            "  ⚠  'Tabla1' no encontrada en el destino; "
                            "se omite el ajuste de tabla.", "WARN")
                        _col_fin_ltr = "L"   # fallback para el formato

                    # ── 2. Aplicar formato uniforme a filas de datos (fila 8 → última) ──
                    # Obtener el número de columna final de la tabla para acotar el rango
                    from openpyxl.utils import column_index_from_string
                    try:
                        _n_col_fin = column_index_from_string(_col_fin_ltr)
                    except Exception:
                        _n_col_fin = 12   # fallback

                    _linea_fmt  = Side(border_style="thin", color="000000")
                    _borde_fmt  = Border(
                        left   = _linea_fmt, right  = _linea_fmt,
                        top    = _linea_fmt, bottom = _linea_fmt,
                    )
                    _font_dato  = Font(name="Arial", size=11)
                    _alin_dato  = Alignment(horizontal="left", vertical="center",
                                            wrap_text=False)

                    _filas_fmt = 0
                    for _r in ws_dest.iter_rows(
                            min_row=FILA_INI, max_row=_fila_datos,
                            min_col=1, max_col=_n_col_fin):
                        for _c in _r:
                            _c.font      = _font_dato
                            _c.alignment = _alin_dato
                            _c.border    = _borde_fmt
                        _filas_fmt += 1

                    self.log_consola(
                        f"  └ Formato aplicado a {_filas_fmt} fila(s) "
                        f"(filas {FILA_INI}–{_fila_datos})", "OK")

                except Exception as _e_tbl:
                    self.log_consola(
                        f"  ⚠  No se pudo ajustar Tabla1 / formato: {_e_tbl}", "WARN")

                # ── 3. Ampliar el formato condicional de D (duplicados) y O ─────────
                #    (NO activos) a toda la tabla. Ambas reglas están fijadas a un
                #    rango concreto en la plantilla original (ej. D8:D30070) y no se
                #    extienden solas cuando la tabla crece — así que con el tiempo
                #    dejan de cubrir las filas nuevas. Se reutilizan las mismas
                #    reglas (mismo color/estilo) sobre el rango actualizado.
                try:
                    _cf = ws_dest.conditional_formatting
                    _rango_d = f"D{FILA_INI}:D{fila_fin_nuevo}"
                    _rango_o = f"O{FILA_INI}:O{fila_fin_nuevo}"
                    _cf_ampliados = []
                    for _cf_range in list(_cf._cf_rules.keys()):
                        _sqref = str(_cf_range.sqref)
                        if _sqref.startswith(f"D{FILA_INI}:D") and _sqref != _rango_d:
                            _reglas = _cf._cf_rules[_cf_range]
                            del _cf._cf_rules[_cf_range]
                            for _regla in _reglas:
                                _cf.add(_rango_d, _regla)
                            _cf_ampliados.append(f"D ({_sqref} → {_rango_d})")
                        elif _sqref.startswith(f"O{FILA_INI}:O") and _sqref != _rango_o:
                            _reglas = _cf._cf_rules[_cf_range]
                            del _cf._cf_rules[_cf_range]
                            for _regla in _reglas:
                                _cf.add(_rango_o, _regla)
                            _cf_ampliados.append(f"O ({_sqref} → {_rango_o})")
                    if _cf_ampliados:
                        self.log_consola(
                            f"  └ Formato condicional ampliado: {', '.join(_cf_ampliados)}",
                            "OK")
                except Exception as _e_cf:
                    self.log_consola(
                        f"  ⚠  No se pudo ampliar el formato condicional: {_e_cf}", "WARN")

                # -- Actualizar celdas de encabezado con fecha y hora de origen --
                try:
                    ws_dest["C4"] = valor_fecha_c4
                    ws_dest["C5"] = valor_hora_c5
                    self.log_consola(f"  └ Celda C4 actualizada: {valor_fecha_c4}", "OK")
                    self.log_consola(f"  └ Celda C5 actualizada: {valor_hora_c5}", "OK")
                except Exception as e_cab:
                    self.log_consola(f"  ⚠  No se pudo actualizar C4/C5: {e_cab}", "WARN")

                self._proteger_hoja_destino(ws_dest, os.path.basename(ruta_destino))

                wb_dest.save(ruta_destino)
                wb_dest.close()

                self.log_consola(f"Destino guardado: {os.path.basename(ruta_destino)}", "OK")
                self.log_consola(f"  └ {n_filas} artículos actualizados (A, B, C, O)", "OK")
                self.log_consola(f"  └ {texto_fecha}", "INFO")
                self.consola.separador("FIN ACTUALIZACIÓN DALI")

                try:
                    self._regenerar_catalogo_dali_completo()
                except Exception as _e_cat:
                    self.log_consola(
                        f"  ⚠  No se pudo regenerar catalogo-dali-completo.xlsx: {_e_cat}", "WARN")

                self.after(0, lambda: self._barra.completado(
                    f"DALI actualizado — {n_filas} artículos ✔"))

                self.after(0, lambda n=n_filas: messagebox.showinfo(
                    "Actualizacion completada",
                    f"\u2714  Listado DALI actualizado correctamente\n\n"
                    f"   Art\u00edculos DALI actualizados : {n}\n"
                    f"   {texto_fecha}\n\n"
                    f"\U0001f4c1  {ruta_destino}"))

            except Exception as e:
                self.log_consola(f"ERROR: {e}", "ERROR")
                self.consola.separador("FIN ACTUALIZACIÓN DALI (con error)")
                self.after(0, lambda: self._barra.reset("Error en actualización DALI"))
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error en actualización DALI",
                    f"\u2718  El proceso no se completó correctamente.\n\n"
                    f"Detalle:\n{err}\n\n"
                    "Revise la ventana de log para más información."))

        threading.Thread(target=proceso, daemon=True).start()
