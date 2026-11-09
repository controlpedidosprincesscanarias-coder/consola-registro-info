# -*- coding: utf-8 -*-
"""
FRAGMENTO DE INTEGRACIÓN — _organizar_dali
===========================================
Pasos para integrar en tarifa_app.py:

  1. En _build_main() → botón acción 1, reemplaza  cmd=self._wip  por  cmd=self._organizar_dali
     y elimina  wip=True  (o ponla en False / quita el parámetro).

  2. Pega el método _organizar_dali (ver más abajo) dentro de la clase App,
     justo después del método _wip (línea ~1139) o donde prefieras.

  3. Sin dependencias nuevas: usa openpyxl, pandas, threading, base64, io y PIL
     que ya están disponibles en el proyecto.

Cambios en _build_main() — botón 1 (líneas ~902-912 de tarifa_app.py):
──────────────────────────────────────────────────────────────────────────
ANTES:
    self._boton_accion(
        frm_acc,
        numero="1",
        texto="Organizar artículos DALI",
        icono="📋",
        cmd=self._wip,
        color="#6B7A8D",
        fg="white",
        wip=True
    ).pack(fill="x", pady=(0, 2))

DESPUÉS:
    self._boton_accion(
        frm_acc,
        numero="1",
        texto="Organizar artículos DALI",
        icono="📋",
        cmd=self._organizar_dali,
        color="#2A7A5C",           # verde oscuro — ya no es WIP
        fg="white",
    ).pack(fill="x", pady=(0, 2))
──────────────────────────────────────────────────────────────────────────
"""


# ─────────────────────────────────────────────────────────────────────────────
#  MÉTODO A AÑADIR DENTRO DE LA CLASE App
# ─────────────────────────────────────────────────────────────────────────────

def _organizar_dali(self):
    """
    Acción 1 — Organizar artículos DALI
    ─────────────────────────────────────
    Lee ARTICULOS1.xls y ARTICULOS2.xls desde la carpeta RECURSOS/ del
    proyecto configurado, cruza los datos por CODIGO (col A) y genera
    ARTICULOS_DALI.xlsx con cabecera con logo Princess en A1:A6, igual
    que la tarifa generada por transformar_tarifa.py.

    Archivos esperados en <carpeta_proyecto>/RECURSOS/:
        ARTICULOS1.xls  →  columnas: CODIGO, DESCRIPCION, UNIDAD,
                           FAMILIA, SUBFAMILIA, NATURALEZA
        ARTICULOS2.xls  →  columnas: CODIGO, …, …, IGIC, …, PROVEEDOR,
                           …, ACTIVO   (cols D, F, H del original)

    Salida:
        <carpeta_proyecto>/RECURSOS/ARTICULOS_DALI.xlsx
    """
    import os
    import threading
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    # ── Validar carpeta de proyecto ───────────────────────────────────────────
    carpeta = self._carpeta_proyecto.get().strip()
    if not carpeta or not os.path.isdir(carpeta):
        from tkinter import messagebox
        messagebox.showwarning(
            "Sin proyecto",
            "Seleccione primero la carpeta del Proyecto.")
        return

    recursos_path = os.path.join(carpeta, "RECURSOS")
    file1  = os.path.join(recursos_path, "ARTICULOS1.xls")
    file2  = os.path.join(recursos_path, "ARTICULOS2.xls")
    output = os.path.join(recursos_path, "ARTICULOS_DALI.xlsx")

    # ── Proceso en hilo secundario ────────────────────────────────────────────
    def proceso():
        try:
            self.after(0, lambda: self._barra.indeterminate("Procesando DALI..."))
            self.consola.separador("INICIO ORGANIZACIÓN DALI")

            # Validaciones
            if not os.path.isdir(recursos_path):
                raise Exception(
                    f"Carpeta RECURSOS no encontrada en:\n{carpeta}")
            if not os.path.exists(file1):
                raise Exception(
                    "Falta ARTICULOS1.xls en RECURSOS/")
            if not os.path.exists(file2):
                raise Exception(
                    "Falta ARTICULOS2.xls en RECURSOS/")

            self.log_consola("Leyendo ARTICULOS1.xls...", "INFO")
            df1 = pd.read_excel(file1, engine="xlrd")

            self.log_consola("Leyendo ARTICULOS2.xls...", "INFO")
            df2 = pd.read_excel(file2, engine="xlrd")

            # Normalizar nombres de columna a str
            df1.columns = df1.columns.astype(str)
            df2.columns = df2.columns.astype(str)

            # Asumimos CODIGO = columna 0 en ambos
            df1["CODIGO"] = df1.iloc[:, 0].astype(str)
            df2["CODIGO"] = df2.iloc[:, 0].astype(str)

            self.log_consola(
                f"  └ ARTICULOS1: {len(df1)} filas  |  "
                f"ARTICULOS2: {len(df2)} filas", "INFO")

            # Índice para búsqueda rápida
            df2_index = df2.set_index("CODIGO")

            # ── Crear libro ───────────────────────────────────────────────────
            wb = Workbook()
            ws = wb.active
            ws.title = "DALI"

            # ── Alturas de fila cabecera (igual que tarifa) ───────────────────
            # Filas 1-5: 18pt; fila 4: 6pt (separador); fila 6: vacía; fila 7: cabecera tabla
            for r in (1, 2, 3, 5):
                ws.row_dimensions[r].height = 18
            ws.row_dimensions[4].height = 6
            ws.row_dimensions[6].height = 10
            ws.row_dimensions[7].height = 18

            # ── Cabecera info (filas 1-5, columnas B-D) ───────────────────────
            _linea  = Side(style="thin", color="000000")
            _borde  = Border(left=_linea, right=_linea,
                             top=_linea, bottom=_linea)
            _font_label = Font(name="Arial", size=11, bold=True)
            _font_valor = Font(name="Arial", size=11, bold=False)
            _alin_c     = Alignment(horizontal="center", vertical="center")
            _alin_l     = Alignment(horizontal="left",   vertical="center")

            ws.cell(1, 2).value = "Documento:"
            ws.cell(1, 3).value = "ARTÍCULOS DALI"
            ws.cell(2, 2).value = "Origen:"
            ws.cell(2, 3).value = "ARTICULOS1.xls + ARTICULOS2.xls"
            ws.cell(3, 2).value = "Carpeta:"
            ws.cell(3, 3).value = recursos_path

            for r in (1, 2, 3):
                ws.cell(r, 2).font      = _font_label
                ws.cell(r, 2).alignment = _alin_c
                ws.cell(r, 3).font      = _font_valor
                ws.cell(r, 3).alignment = _alin_l

            # Borde externo en cada fila de cabecera
            def _borde_externo(r1, c1, r2, c2):
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        top    = _linea if r == r1 else Side(style=None)
                        bottom = _linea if r == r2 else Side(style=None)
                        left   = _linea if c == c1 else Side(style=None)
                        right  = _linea if c == c2 else Side(style=None)
                        eb = ws.cell(r, c).border
                        ws.cell(r, c).border = Border(
                            top    = top    if r == r1 else eb.top,
                            bottom = bottom if r == r2 else eb.bottom,
                            left   = left   if c == c1 else eb.left,
                            right  = right  if c == c2 else eb.right,
                        )

            for r in (1, 2, 3):
                _borde_externo(r, 2, r, 4)

            # ── Logo Princess en A1:A5 ────────────────────────────────────────
            self.log_consola("Insertando logo en A1:A5...", "INFO")
            try:
                import base64
                from io import BytesIO
                from openpyxl.drawing.image import Image as XLImage
                from PIL import Image as PILImage
                from openpyxl.drawing.spreadsheet_drawing import (
                    AnchorMarker, TwoCellAnchor
                )

                # Logo embebido idéntico al de tarifa_app.py y transformar_tarifa.py
                LOGO_B64 = (
                    "iVBORw0KGgoAAAANSUhEUgAAAJAAAACQCAIAAABoJHXvAAABCGlDQ1BJQ0MgUHJvZmlsZQAAeJxj"
                    "YGA8wQAELAYMDLl5JUVB7k4KEZFRCuwPGBiBEAwSk4sLGHADoKpv1yBqL+viUYcLcKakFicD6Q9A"
                    "rFIEtBxopAiQLZIOYWuA2EkQtg2IXV5SUAJkB4DYRSFBzkB2CpCtkY7ETkJiJxcUgdT3ANk2uTml"
                    "yQh3M/Ck5oUGA2kOIJZhKGYIYnBncAL5H6IkfxEDg8VXBgbmCQixpJkMDNtbGRgkbiHEVBYwMPC3"
                    "MDBsO48QQ4RJQWJRIliIBYiZ0tIYGD4tZ2DgjWRgEL7AwMAVDQsIHG5TALvNnSEfCNMZchhSgSKe"
                    "DHkMyQx6QJYRgwGDIYMZAKbWPz9HbOBQAAApwklEQVR42u1deVgT29qfmSQESAwgW2RHRaqiKK64"
                    "oVXUutty9Vq11lqXaqtWbWn7ea3VurXaau3jrXUp17qC+1Z3a0UqiqIWwR1QQA2yQwgkM/P98Xv6"
                    "PnNDiIjaijfnDx+cnDnnzHnPuy+HYWzN1mzN1mzN1l6k9v/FR53Zmux1+wAAAABJRU5ErkJggg=="
                )

                # Intentar cargar logo desde archivo externo primero
                ruta_logo = resource_path("assets/logo.png")
                if os.path.isfile(ruta_logo):
                    pil_img = PILImage.open(ruta_logo)
                    pil_img = pil_img.resize((95, 60), PILImage.LANCZOS)
                    buf = BytesIO()
                    pil_img.save(buf, format="PNG")
                    buf.seek(0)
                    self.log_consola("  └ Logo cargado desde assets/logo.png", "OK")
                else:
                    buf = BytesIO(base64.b64decode(LOGO_B64))
                    self.log_consola("  └ Logo cargado desde base64 embebido", "OK")

                EMU_PX   = 9525
                FILA_MID = 12 * EMU_PX
                ANCHO_PX = 95

                xl_img        = XLImage(buf)
                xl_img.width  = ANCHO_PX
                xl_img.height = 60

                # Pre-calcular ancho col A para centrar el logo
                _max_len_a = max(
                    (len(str(ws.cell(r, 1).value or ""))
                     for r in range(1, max(ws.max_row, 2) + 1)),
                    default=0
                )
                _col_a_units = max(round(_max_len_a * 1.60) + 4, 14)
                _col_a_px    = int(_col_a_units * 7 + 5)
                _margen_px   = max(0, (_col_a_px - ANCHO_PX) // 2) + 20
                _margen_emu  = _margen_px * EMU_PX
                ANCHO_EMU    = ANCHO_PX * EMU_PX

                _from  = AnchorMarker(col=0, colOff=_margen_emu,             row=0, rowOff=FILA_MID)
                _to    = AnchorMarker(col=0, colOff=_margen_emu + ANCHO_EMU, row=4, rowOff=FILA_MID)
                anchor = TwoCellAnchor(editAs="oneCell", _from=_from, to=_to)
                anchor.pic    = None
                xl_img.anchor = anchor
                ws.add_image(xl_img)
                self.log_consola(
                    f"  └ Logo centrado en col A ({ANCHO_PX}px), margen={_margen_px}px", "OK")

            except ImportError as e:
                self.log_consola(f"  └ AVISO: Pillow no disponible ({e}). Logo omitido.", "WARN")
            except Exception as e:
                self.log_consola(f"  └ AVISO: error al insertar logo ({e}). Logo omitido.", "WARN")

            # ── Cabecera tabla (fila 7) ───────────────────────────────────────
            headers = ["CODIGO", "DESCRIPCION", "UNIDAD", "FAMILIA",
                       "SUBFAMILIA", "NATURALEZA", "IGIC", "PROVEEDOR", "ACTIVO"]

            font_hdr  = Font(name="Arial", size=12, bold=True)
            fill_hdr  = PatternFill(start_color="D3D3D3",
                                    end_color="D3D3D3", fill_type="solid")
            alin_hdr  = Alignment(horizontal="center", vertical="center")
            borde_hdr = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"),  bottom=Side(style="thin"))

            for col, h in enumerate(headers, start=1):
                c = ws.cell(row=7, column=col, value=h)
                c.font      = font_hdr
                c.fill      = fill_hdr
                c.alignment = alin_hdr
                c.border    = borde_hdr

            # Filtro automático sobre los headers
            from openpyxl.utils import get_column_letter
            ws.auto_filter.ref = (
                f"A7:{get_column_letter(len(headers))}7"
            )

            # ── Cruce de datos ────────────────────────────────────────────────
            self.log_consola("Cruzando datos por CODIGO...", "INFO")
            errores   = 0
            font_dato = Font(name="Arial", size=12)
            alin_dato = Alignment(horizontal="left", vertical="center")
            borde_dato = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"),  bottom=Side(style="thin"))

            for i in range(len(df1)):
                r      = 8 + i
                codigo = str(df1.iloc[i, 0])

                # Datos de ARTICULOS1
                for col_idx, col_val in enumerate(
                        [codigo,
                         df1.iloc[i, 1] if df1.shape[1] > 1 else "",
                         df1.iloc[i, 2] if df1.shape[1] > 2 else "",
                         df1.iloc[i, 3] if df1.shape[1] > 3 else "",
                         df1.iloc[i, 4] if df1.shape[1] > 4 else "",
                         df1.iloc[i, 5] if df1.shape[1] > 5 else ""],
                        start=1):
                    cell            = ws.cell(r, col_idx, col_val)
                    cell.font       = font_dato
                    cell.alignment  = alin_dato
                    cell.border     = borde_dato

                # Buscar en ARTICULOS2
                if codigo in df2_index.index:
                    fila2 = df2_index.loc[codigo]
                    igic      = fila2.iloc[3] if len(fila2) > 3 else ""   # col D → IGIC
                    proveedor = fila2.iloc[5] if len(fila2) > 5 else ""   # col F → PROVEEDOR
                    activo    = fila2.iloc[7] if len(fila2) > 7 else ""   # col H → ACTIVO
                    for col_idx, col_val in enumerate(
                            [igic, proveedor, activo], start=7):
                        cell            = ws.cell(r, col_idx, col_val)
                        cell.font       = font_dato
                        cell.alignment  = alin_dato
                        cell.border     = borde_dato
                else:
                    errores += 1

            # ── Autoajuste de columnas ────────────────────────────────────────
            self.log_consola("Autoajustando anchos de columna...", "INFO")
            for col_cells in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col_cells[0].column)
                for cell in col_cells:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                # Col A reservada para logo; mínimo razonable
                if col_letter == "A":
                    ws.column_dimensions[col_letter].width = max(max_len * 1.2 + 4, 14)
                else:
                    ws.column_dimensions[col_letter].width = min(max(max_len * 1.2 + 4, 10), 60)

            # ── Guardar ───────────────────────────────────────────────────────
            wb.save(output)

            self.log_consola("Archivo generado correctamente", "OK")
            if errores:
                self.log_consola(
                    f"Códigos sin coincidencia en ARTICULOS2: {errores}", "WARN")
            self.log_consola(f"Ruta: {output}", "INFO")
            self.consola.separador("FIN ORGANIZACIÓN DALI")
            self.after(0, lambda: self._barra.completado("DALI generado ✔"))

            # Abrir el resultado automáticamente
            from tkinter import messagebox
            self.after(0, lambda: messagebox.showinfo(
                "DALI generado",
                f"Archivo generado correctamente:\n{output}"
                + (f"\n\n⚠ Códigos sin coincidencia: {errores}" if errores else "")))
            try:
                import os as _os
                self.after(200, lambda: _os.startfile(output))
            except Exception:
                pass

        except Exception as e:
            self.log_consola(f"ERROR: {e}", "ERROR")
            self.consola.separador("FIN ORGANIZACIÓN DALI (con error)")
            self.after(0, lambda: self._barra.reset("Error en DALI"))
            from tkinter import messagebox
            self.after(0, lambda err=str(e): messagebox.showerror(
                "Error en DALI", f"Se produjo un error:\n{err}"))

    threading.Thread(target=proceso, daemon=True).start()
