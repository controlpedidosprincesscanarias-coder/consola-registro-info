# -*- coding: utf-8 -*-
"""
creador_rinfo.py
CreadorRINFO: crea un Registro Info SAP nuevo (sin Purchasing existente).

Separado de crear_registro_info.py en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""

import os
import re
from tkinter import messagebox
import datetime as _dt_mod          # módulo completo — usado en lógica de fechas de baja
from datetime import datetime
import openpyxl
from app.comun.catalogo_local import abrir_como_listado  # v10.2.54: catálogo con forma de LISTADO
# normalizar_codigo_sap vive en sap_helpers.py (sin tkinter) para que
# también pueda importarla test_crear_registro_info.py sin arrastrar
# la UI. Ver ese fichero para la implementación y el porqué.
from app.comun.sap_helpers import (
    normalizar_codigo_sap, ultima_fila_con_datos)
from app.rinfo.comun_rinfo import (  # noqa: F401
    ARCHIVO_TEMPLATE,
    FECHA_FIN_FIJA_HOJA1,
    FECHA_INICIO_FIJA_HOJA1,
    FOLDER_RECURSOS,
    HOJAS_PURCHASING,
    HOJA_TARIFA,
    _CENTROS_LIST,
    _localizar_ruta_listado_dali,
    _log,
    _postprocesar_para_sap,
    _safe_after,
    aplicar_formato_fila,
    copiar_template_registro_info,
)
from app.rinfo.ventanas_rinfo import (  # noqa: F401
    VentanaConfirmacionRINFO,
)


# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4b — CREADOR DE RINFO DESDE CERO
# ─────────────────────────────────────────────────────────────────────────────

class CreadorRINFO:
    """
    Genera el nuevo archivo RINFO desde cero (sin depender de plantilla):

      1. Llama a copiar_template_registro_info() para obtener el xlsx limpio
         (copia binaria del Purchasing Info Records.xlsx de RECURSOS/).
      2. Escribe D2 = nombre SAP y C2 = código SAP.
      3. Copia columna A de la tarifa → columna E del RINFO desde fila 2,
         aplicando formato a cada nueva fila con aplicar_formato_fila() [ya no activo — formato preservado del template SAP].
      4. Propaga C y D hasta la última fila con datos.
      5. Guarda y muestra confirmación.
    """

    HOJA_PRINCIPAL = "Info general p.reg.info"

    def __init__(self, parent_ventana, ruta_tarifa: str, nombre_archivo: str,
                 nombre_sap: str, codigo_sap: str, carpeta_tarifas: str,
                 app_root=None, telefono_proveedor: str = None,
                 fecha_inicio: str = "01.01.2025", fecha_fin: str = "31.12.2030",
                 centros_activos: list = None):
        self._parent              = parent_ventana
        self._ruta_tarifa         = ruta_tarifa
        self._nombre_archivo      = nombre_archivo
        self._nombre_sap          = nombre_sap
        self._codigo_sap          = codigo_sap
        self._carpeta_tarifas     = carpeta_tarifas
        # Teléfono opcional (col AD = columna 30 de la hoja principal).
        # None → no se escribe nada en esa columna.
        self._telefono_proveedor  = telefono_proveedor
        # Fechas de validez de la tarifa — formato "dd.mm.yyyy".
        # Se usan en Hoja1 (cols I/J) y Hoja3 (cols I/J).
        self._fecha_inicio        = fecha_inicio or "01.01.2025"
        self._fecha_fin           = fecha_fin    or "31.12.2030"
        # Centros SAP activos para Hoja2 — lista de (codigo, nombre).
        # None o lista vacía → se usan los 10 centros por defecto.
        self._centros_activos     = centros_activos if centros_activos else list(_CENTROS_LIST)
        # Referencia directa a App para la barra de progreso.
        # Si no se pasa, se intenta localizar subiendo la jerarquía (fallback).
        self._app_root            = app_root

        # Lanzar en hilo secundario para no bloquear la UI de Tkinter
        # y permitir que la barra de progreso de tarifa_app se actualice.
        import threading
        threading.Thread(target=self._ejecutar, daemon=True).start()

    def _barra(self, modo: str, texto: str = "", v: int = 0, total: int = 100):
        """
        Envía actualizaciones de barra al hilo principal de Tkinter vía after().
        Usa _app_root si está disponible; en caso contrario sube por la jerarquía.
        Es seguro llamarlo desde cualquier hilo.
        """
        # 1. Referencia directa (camino rápido y fiable)
        app = self._app_root

        # 2. Fallback: subir por la jerarquía desde parent_ventana
        if app is None or not hasattr(app, "_barra"):
            app = self._parent
            for _ in range(8):
                if hasattr(app, "_barra"):
                    break
                app = getattr(app, "master", None) or getattr(app, "_parent", None)
                if app is None:
                    return

        if not hasattr(app, "_barra"):
            return

        barra = app._barra
        try:
            if modo == "indeterminate":
                app.after(0, lambda t=texto: barra.indeterminate(t))
            elif modo == "set":
                app.after(0, lambda: barra.set(v, total, texto))
            elif modo == "completado":
                app.after(0, lambda t=texto: barra.completado(t))
            elif modo == "reset":
                app.after(0, lambda t=texto: barra.reset(t))
        except Exception:
            pass

    def _ejecutar(self):
        """Orquesta todos los pasos de creación del RINFO (corre en hilo secundario)."""

        # ── 1. Construir nombre del RINFO de destino ──────────────────────────
        # Se usa la fecha de hoy (día en que se genera el RINFO),
        # no la fecha de vigencia que trae el nombre de la tarifa.
        nombre_base_tarifa = os.path.splitext(self._nombre_archivo)[0]
        fecha_hoy          = _dt_mod.date.today().strftime("%d-%m-%Y")
        nombre_base  = re.sub(r'\[\d{2}-\d{2}-\d{4}\]', f'[{fecha_hoy}]',
                               nombre_base_tarifa)
        nombre_nuevo = f"RINFO {nombre_base}.xlsx"
        ruta_nuevo   = os.path.join(self._carpeta_tarifas, nombre_nuevo)

        _log(self._parent, f"Nombre RINFO: {nombre_nuevo}", "INFO")
        self._barra("set", "Preparando archivo RINFO...", 5, 100)

        # ── 2. Crear el archivo copiando el template SAP desde RECURSOS/ ──────
        #    Esto preserva los metadatos internos, tipos de celda y configuración
        #    SAP que openpyxl no puede replicar al construir desde cero.
        #    Si el template no existe, FileNotFoundError → mensaje al usuario.
        self._carpeta_proyecto = os.path.dirname(self._carpeta_tarifas)

        try:
            wb = copiar_template_registro_info(ruta_nuevo, self._carpeta_proyecto)
            _log(self._parent,
                 f"Template SAP copiado y limpiado: {nombre_nuevo}", "OK")
        except FileNotFoundError as e:
            ruta_recursos = os.path.join(self._carpeta_proyecto, FOLDER_RECURSOS)
            msg = (
                f"No se encontró el archivo de plantilla SAP.\n\n{e}\n\n"
                f"Coloque el archivo «{ARCHIVO_TEMPLATE}» en:\n"
                f"   {ruta_recursos}\n\n"
                "y vuelva a intentarlo."
            )
            self._parent.after(0, lambda m=msg: messagebox.showerror(
                "Plantilla SAP no disponible", m, parent=self._parent))
            _log(self._parent,
                 "Template SAP no encontrado en RECURSOS/ — creación abortada.", "ERROR")
            self._barra("reset", "Plantilla SAP no encontrada")
            _safe_after(self._parent, self._parent.withdraw)
            return
        except Exception as e:
            self._parent.after(0, lambda err=str(e): messagebox.showerror(
                "Error al crear el archivo",
                f"No se pudo crear el archivo RINFO:\n\n{err}",
                parent=self._parent))
            _log(self._parent, f"Error al crear RINFO: {e}", "ERROR")
            self._barra("reset", "Error al crear RINFO")
            _safe_after(self._parent, self._parent.withdraw)
            return

        # ── 3. El workbook ya está abierto: devuelto por copiar_template_registro_info()
        #    No es necesario volver a abrir con load_workbook.

        # ── 4. Verificar hoja principal y escribir valores obligatorios ────────
        self._barra("set", "Escribiendo datos proveedor...", 20, 100)
        if self.HOJA_PRINCIPAL not in wb.sheetnames:
            self._parent.after(0, lambda: messagebox.showerror(
                "Hoja no encontrada",
                f"❌  La hoja principal «{self.HOJA_PRINCIPAL}»\n"
                "no existe en el archivo generado.",
                parent=self._parent))
            _log(self._parent,
                 f"Hoja «{self.HOJA_PRINCIPAL}» no encontrada en el RINFO creado.", "ERROR")
            wb.close()
            self._intentar_eliminar(ruta_nuevo)
            self._barra("reset", "Error: hoja no encontrada")
            _safe_after(self._parent, self._parent.withdraw)
            return

        ws_principal = wb[self.HOJA_PRINCIPAL]
        ws_principal["D2"] = self._nombre_sap
        ws_principal["C2"] = self._codigo_sap if self._codigo_sap else None
        _log(self._parent,
             f"  D2 ← «{self._nombre_sap}»  |  C2 ← «{self._codigo_sap}»", "OK")

        # ── 5. Copiar columna A de la tarifa → columna E del RINFO ───────────
        self._barra("set", "Copiando artículos de la tarifa...", 35, 100)
        articulos_copiados = self._copiar_articulos(wb)

        # ── 6. Numeración correlativa en columna A ────────────────────────────
        self._barra("set", "Numerando correlativos...", 55, 100)
        self._escribir_correlativos(wb)

        # ── 7. Cruzar con LISTADO CODIGOS DALI - SAP y rellenar columnas ─────
        self._barra("indeterminate", "Cruzando con el catálogo DALI-SAP...")
        self._cruzar_con_listado_dali(wb)

        # ── 7b. Rellenar Hoja 2 — Datos org.compr.p.reg.info ─────────────────
        self._barra("set", "Rellenando Hoja 2 (Datos org.compr.)...", 75, 100)
        self._rellenar_hoja2(wb)

        # ── 7c. Rellenar Hoja 3 — Det.precios p.registros info ───────────────
        self._barra("set", "Rellenando Hoja 3 (Det. precios)...", 83, 100)
        self._rellenar_hoja3(wb)

        # ── 8. Ajustar vistas: A1 activo y sin selección en todas las hojas ────
        #    Se establece via sheet_view para que al abrir el archivo en Excel
        #    no haya celdas resaltadas y la Hoja1 sea la visible de inicio.
        #    El orden importa: la hoja activa del libro debe establecerse al final,
        #    después de haber configurado todas las sheet_view.
        for nombre_hoja in HOJAS_PURCHASING:
            if nombre_hoja not in wb.sheetnames:
                continue
            ws_v = wb[nombre_hoja]
            # Asegurar que existe al menos una SheetView
            if not ws_v.sheet_view:
                from openpyxl.worksheet.views import SheetView
                ws_v.sheet_view.append(SheetView())
            sv = ws_v.sheet_view
            sv.topLeftCell   = "A1"
            sv.selection[0].activeCell = "A1"
            sv.selection[0].sqref     = "A1"
            sv.tabSelected = False           # ninguna hoja marcada aún como activa

        # Activar la Hoja1 (Info general p.reg.info) como hoja visible al abrir
        if self.HOJA_PRINCIPAL in wb.sheetnames:
            wb[self.HOJA_PRINCIPAL].sheet_view.tabSelected = True
            wb.active = wb[self.HOJA_PRINCIPAL]

        # ── 7d. Renumeración final col K Hoja3 (norma SAP) ──────────────────
        # Garantiza secuencia 1..N sin huecos en col K para todos los materiales.
        self._renumerar_col_k_h3(wb)

        # ── 8. Guardar ────────────────────────────────────────────────────────
        self._barra("set", "Guardando archivo RINFO...", 90, 100)
        ruta_template_ref = os.path.join(
            self._carpeta_proyecto, FOLDER_RECURSOS, ARCHIVO_TEMPLATE)
        try:
            wb.save(ruta_nuevo)
            wb.close()
            self._barra("set", "Adaptando estructura OOXML para SAP...", 95, 100)
            n_sst = _postprocesar_para_sap(ruta_nuevo, ruta_template_ref)
            _log(self._parent,
                 f"Post-procesado SAP OK: sharedStrings ({n_sst} cadena(s)), "
                 f"estructura OOXML restaurada del template Purchasing.", "OK")
        except Exception as e:
            self._parent.after(0, lambda err=str(e): messagebox.showerror(
                "Error al guardar",
                f"❌  No se pudo guardar el archivo:\n\n{nombre_nuevo}\n\n{err}",
                parent=self._parent))
            _log(self._parent, f"Error al guardar RINFO: {e}", "ERROR")
            self._barra("reset", "Error al guardar RINFO")
            self._intentar_eliminar(ruta_nuevo)
            _safe_after(self._parent, self._parent.withdraw)
            return

        _log(self._parent,
             f"RINFO guardado correctamente: {nombre_nuevo} "
             f"({articulos_copiados} artículo(s) copiado(s))", "OK")
        self._barra("completado", f"RINFO creado: {articulos_copiados} artículo(s)")

        # ── 9. Confirmación final ─────────────────────────────────────────────
        hojas_creadas = list(HOJAS_PURCHASING)
        nom = nombre_nuevo
        ruta = ruta_nuevo
        nom_sap = self._nombre_sap
        cod_sap = self._codigo_sap
        art = articulos_copiados
        hojas = hojas_creadas

        def _mostrar_confirmacion():
            parent_root = (self._parent.master
                           if hasattr(self._parent, "master")
                           else self._parent)
            VentanaConfirmacionRINFO(parent_root, nom, ruta, nom_sap, cod_sap,
                                     hojas, art)

        _safe_after(self._parent, self._parent.withdraw)
        self._parent.after(80, _mostrar_confirmacion)

    def _copiar_articulos(self, wb) -> int:
        """
        Lee desde la hoja TARIFA (fila 8 en adelante):
          · col A (CODIGO SAP)         -> RINFO col E  (Material)           [clave de cruce]
          · col C (CODIGO PROVEEDOR)   -> RINFO col AB (Num. mat. proveedor)

        Para cada fila propaga ademas en RINFO:
          · col C (3)  = codigo SAP del proveedor  (self._codigo_sap)
          · col D (4)  = nombre SAP del proveedor  (self._nombre_sap)

        Aplica el formato de fila 2 a cada nueva fila generada.
        Devuelve el numero de articulos copiados.

        col AB = columna 28 = "Numero de material de proveedor"
        """
        if self.HOJA_PRINCIPAL not in wb.sheetnames:
            _log(self._parent,
                 "Hoja principal no disponible para copiar articulos.", "WARN")
            return 0

        # ── Leer TARIFA: col A (CODIGO SAP) y col C (CODIGO PROVEEDOR) ────────
        # La fila 7 es la cabecera real; los datos empiezan en fila 8.
        articulos = []   # lista de (codigo_sap_str, codigo_proveedor_str|None)
        try:
            wb_tarifa = openpyxl.load_workbook(
                self._ruta_tarifa, read_only=True, data_only=True)
            hoja_tarifa = (wb_tarifa[HOJA_TARIFA]
                           if HOJA_TARIFA in wb_tarifa.sheetnames
                           else wb_tarifa.active)

            for fila in hoja_tarifa.iter_rows(min_row=8, min_col=1,
                                               max_col=3, values_only=True):
                cod_sap  = fila[0]   # col A: CODIGO SAP
                cod_prov = fila[2]   # col C: CODIGO PROVEEDOR

                if cod_sap is None:
                    continue

                # Normalizar CODIGO SAP -> str sin decimales
                cod_sap = normalizar_codigo_sap(cod_sap)
                cod_sap_str = str(cod_sap).strip()

                # Normalizar CODIGO PROVEEDOR -> str (puede ser alfanumerico)
                if cod_prov is not None:
                    cod_prov = normalizar_codigo_sap(cod_prov)
                    cod_prov_str = str(cod_prov).strip()
                else:
                    cod_prov_str = None

                articulos.append((cod_sap_str, cod_prov_str))

            wb_tarifa.close()
        except Exception as e:
            _log(self._parent,
                 f"Error al leer columna A/C de la tarifa: {e}", "ERROR")
            messagebox.showwarning(
                "Aviso — Articulos no copiados",
                f"⚠  No se pudieron leer los articulos de la tarifa:\n\n{e}\n\n"
                "El RINFO se ha creado igualmente, pero la columna E estara vacia.",
                parent=self._parent)
            return 0

        if not articulos:
            _log(self._parent,
                 "Columna A de la tarifa sin datos — columna E del RINFO vacia.", "WARN")
            return 0

        ws_rinfo = wb[self.HOJA_PRINCIPAL]

        # ── Escribir filas en el RINFO ─────────────────────────────────────────
        # col E  ( 5) = Material (CODIGO SAP normalizado)
        # col C  ( 3) = Proveedor (codigo SAP del proveedor)
        # col D  ( 4) = Nombre del proveedor (nombre SAP)
        # col AB (28) = Numero de material de proveedor (CODIGO PROVEEDOR tarifa)
        for idx, (cod_sap_str, cod_prov_str) in enumerate(articulos, start=2):
            ws_rinfo.cell(row=idx, column=5,  value=cod_sap_str)
            ws_rinfo.cell(row=idx, column=4,  value=self._nombre_sap)
            ws_rinfo.cell(row=idx, column=3,
                          value=self._codigo_sap if self._codigo_sap else None)
            ws_rinfo.cell(row=idx, column=28, value=cod_prov_str)
            # ⚠ NO se aplica formato: el template SAP ya contiene el formato
            #   correcto en cada celda. Reescribirlo con openpyxl corrompe los
            #   índices de estilo internos y SAP rechaza o ignora el archivo.

        _log(self._parent,
             f"  └ {len(articulos)} articulo(s) copiado(s) "
             f"(col E <- CODIGO SAP, col AB <- CODIGO PROVEEDOR).", "OK")
        return len(articulos)


    # ── Paso 6: Correlativos en columna A ─────────────────────────────────────
    def _escribir_correlativos(self, wb) -> None:
        """
        Escribe "1", "2", "3"... (texto forzado, number_format="@") en la columna
        A (col 1) desde la fila 2 hasta la ultima fila que tenga valor en col E.

        SAP descarga la columna «Registro info de compras» con formato de texto
        forzado (indicador de advertencia verde en Excel).  Para que la carga
        masiva funcione correctamente, los valores nuevos deben ser str y la celda
        debe tener number_format="@".  Escribir un entero puro provoca que SAP
        lo trate como numero y rechace o ignore la fila.

        Solo se toca .value y .number_format -- el resto del formato de cada celda
        (fill, border, font, alignment) no se modifica.
        """
        if self.HOJA_PRINCIPAL not in wb.sheetnames:
            return

        ws = wb[self.HOJA_PRINCIPAL]

        ultima_fila = ultima_fila_con_datos(ws, columna=5)

        if ultima_fila < 2:
            _log(self._parent,
                 "Correlativos col A: sin filas con datos en E -- nada que numerar.",
                 "WARN")
            return

        for fila_num in range(2, ultima_fila + 1):
            celda = ws.cell(row=fila_num, column=1)
            celda.value         = str(fila_num - 1)   # texto, no entero
            celda.number_format = "@"                  # fuerza triangulito verde SAP

        _log(self._parent,
             f"  Correlativos col A: \"1\" - \"{ultima_fila - 1}\" "
             f"(texto/@, filas 2-{ultima_fila}).", "OK")

    # ── Paso 7: Cruce con LISTADO CODIGOS DALI - SAP ──────────────────────────
    def _cruzar_con_listado_dali(self, wb) -> None:
        """
        Cruza col E (Material) del RINFO con col D (CODIGO SAP - TEST) del
        LISTADO CODIGOS DALI - SAP.xlsx y escribe en RINFO (solo .value):

        Cruce LISTADO -> RINFO (clave: RINFO col E = LISTADO col D):
            LISTADO col E (DESCRIPCION SAP-TEST)        -> RINFO col F  ( 6) Texto breve material
            LISTADO col G (UNIDAD SAP)                  -> RINFO col N  (14) Unidad medida pedido
            LISTADO col G (UNIDAD SAP)                  -> RINFO col P  (16) Unidad de medida
            LISTADO col H (GRUPO PRODUCTOS SAP)         -> RINFO col G  ( 7) Grupo de articulos
            LISTADO col I (Desc. grupo productos SAP)   -> RINFO col H  ( 8) Desc. grupo articulos

        (mapeo actualizado 2026-08: tras insertar las 3 columnas nuevas
        F, K, L en LISTADO CODIGOS DALI-SAP, lo que había se desplazó:
        antigua F/G/H -> G/H/I)

        Valores fijos (spec seccion 3) para todas las filas con dato en col E:
            RINFO col I  (9)  (Disponible desde)        = "01.01.2025"
            RINFO col J  (10) (Disponible hasta)        = "31.12.2030"
            RINFO col O  (15) (Denominador conversion)  = "1"
            RINFO col Q  (17) (Numerador conversion)    = "1"
            RINFO col AK (37) (Nueva)                   = "X"

        Solo se escribe .value -- el formato de destino no se modifica nunca.
        La clave de cruce real es LISTADO col D (int) normalizada a str,
        que coincide con los codigos SAP de col A de la TARIFA / col E del RINFO.
        """
        if self.HOJA_PRINCIPAL not in wb.sheetnames:
            return

        ws_rinfo = wb[self.HOJA_PRINCIPAL]
        nombre_listado = "catalogo-dali-completo.xlsx"   # v10.2.54 (antes, el LISTADO)

        # -- Localizar LISTADO CODIGOS DALI - SAP.xlsx ────────────────────────
        # (antes ~35 líneas propias; consolidado en _localizar_ruta_listado_dali)
        ruta_listado = _localizar_ruta_listado_dali(
            self._carpeta_tarifas, getattr(self, "_carpeta_proyecto", None), self._parent)
        if ruta_listado is None:
            _log(self._parent,
                 f"'{nombre_listado}' no encontrado -- "
                 "columnas F, G, H, N, P quedaran vacias.", "WARN")
            self._aplicar_valores_fijos(ws_rinfo)
            return

        _log(self._parent, f"Catálogo DALI-SAP localizado: {ruta_listado}", "INFO")

        # -- Cargar diccionario codigo_SAP --> (desc, unidad, grp, desc_grp) ──
        # Datos reales desde fila 8 (fila 7 = cabeceras).
        # La clave col D (indice 3) es int en el LISTADO; se normaliza a string.
        diccionario = {}
        try:
            wb_listado = abrir_como_listado(ruta_listado)
            ws_listado = wb_listado.active
            for row in ws_listado.iter_rows(min_row=8, values_only=True):
                if not row or len(row) < 4 or row[3] is None:
                    continue
                raw = row[3]
                if isinstance(raw, float):
                    raw = int(raw)
                clave = str(raw).strip()
                if not clave:
                    continue
                desc_sap = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
                unidad   = str(row[6]).strip() if len(row) > 6 and row[6] is not None else ""
                grp      = str(row[7]).strip() if len(row) > 7 and row[7] is not None else ""
                desc_grp = str(row[8]).strip() if len(row) > 8 and row[8] is not None else ""
                diccionario[clave] = (desc_sap, unidad, grp, desc_grp)
            wb_listado.close()
            _log(self._parent,
                 f"Catálogo DALI-SAP cargado: {len(diccionario)} entradas.", "INFO")
        except Exception as e:
            _log(self._parent,
                 f"Error al leer '{nombre_listado}': {e}", "ERROR")
            self._aplicar_valores_fijos(ws_rinfo)
            return

        if not diccionario:
            _log(self._parent,
                 f"'{nombre_listado}' sin datos -- cols F, G, H, N, P vacias.", "WARN")
            self._aplicar_valores_fijos(ws_rinfo)
            return

        # -- Recorrer filas del RINFO y escribir valores ───────────────────────
        total_e = 0
        matches = 0
        for fila in ws_rinfo.iter_rows(min_row=2, min_col=5, max_col=5):
            celda_e = fila[0]
            if celda_e.value is None:
                continue
            total_e += 1
            raw_e = celda_e.value
            # Normalizar a string limpia: float 10800044.0 → "10800044"
            raw_e = normalizar_codigo_sap(raw_e)
            clave_material = str(raw_e).strip()
            fila_num = celda_e.row

            datos = diccionario.get(clave_material)
            if datos:
                desc_sap, unidad, grp, desc_grp = datos
                # LISTADO col E -> RINFO col F  (6)  Texto breve material
                ws_rinfo.cell(row=fila_num, column=6).value  = desc_sap or None
                # LISTADO col G -> RINFO col G  (7)  Grupo de articulos
                ws_rinfo.cell(row=fila_num, column=7).value  = grp      or None
                # LISTADO col H -> RINFO col H  (8)  Desc. grupo de articulos
                ws_rinfo.cell(row=fila_num, column=8).value  = desc_grp or None
                # LISTADO col F -> RINFO col N  (14) Unidad medida de pedido
                ws_rinfo.cell(row=fila_num, column=14).value = unidad   or None
                # LISTADO col F -> RINFO col P  (16) Unidad de medida
                ws_rinfo.cell(row=fila_num, column=16).value = unidad   or None
                matches += 1

        _log(self._parent,
             f"  Cruce DALI-SAP: {matches} coincidencia(s) de {total_e} articulo(s).",
             "OK")

        self._aplicar_valores_fijos(ws_rinfo)

    # ── Paso 7b: Relleno completo de Hoja 2 ──────────────────────────────────
    def _rellenar_hoja2(self, wb) -> None:
        """
        Rellena la hoja «Datos org.compr.p.reg.info» según la especificación.

        Estructura generada (10 bloques = Hoja1 × 10):
        ─────────────────────────────────────────────────────────────────────
        Col A  ( 1)  Registro info de compras     ← Hoja1 col A × 10 bloques
        Col B  ( 2)  Tipo registro                = "0"
        Col C  ( 3)  Desc. categoría registro     = "Normal"
        Col D  ( 4)  Organización de compras      = "IC00"
        Col E  ( 5)  Desc. org. compras           = "OC ISLAS CANANARIAS"
        Col F  ( 6)  Centro                       ← CENTROS_SAP_CANARIAS (código, por bloque)
        Col G  ( 7)  Nombre de centro             ← CENTROS_SAP_CANARIAS (nombre, por bloque)
        Col H  ( 8)  Grupo de compras             ← LISTADO col M
        Col I  ( 9)  Desc. grupo compras          ← LISTADO col N
        Col J  (10)  Material                     ← Hoja1 col E × 10 bloques
        Col K  (11)  Texto breve material         ← LISTADO col E
        Col L  (12)  Grupo de artículos           ← LISTADO col H
        Col M  (13)  Cantidad pedido estándar     = 1  (entero, sin triángulo verde)
        Col N  (14)  Plazo de entrega en días     = 1  (entero)
        Col S  (19)  Verif. factura basada en EM  = "X"
        Col W  (23)  Cantidad pedido máxima       = 0  (entero, sin triángulo verde)
        Col X  (24)  Cantidad mínima de pedido    = 0  (entero, sin triángulo verde)
        Col AD (30)  Tolerancia exceso suminst. % = 0  (entero, sin triángulo verde)
        Col AE (31)  Tolerancia faltas suminst. % = 0  (entero, sin triángulo verde)
        Col AF (32)  Control fecha del precio     = 1  (entero, sin triángulo verde)
        Col AH (34)  Indicador de impuestos       ← LISTADO col J
        Col AJ (36)  Inic.val.tipo impos.         = "01.01.1900"
        Col AO (41)  Nueva                        = "X"
        ─────────────────────────────────────────────────────────────────────
        NOTA (auditoría 2026-08, manual SAP "Crear Registros Info" 3.1.2):
        se retiró la columna AK (37, "Entrega ilimitada permitida"), que no
        figura entre las columnas fijas de Hoja2 documentadas en el manual.
        N (14) y W/X (23/24) SÍ se mantienen con sus valores fijos: el manual
        no da un valor concreto, pero están confirmados como correctos por
        información notificada a posteriori (central/SAP).
        Solo se toca .value — el formato del template SAP se preserva intacto.
        """
        NOMBRE_H2   = "Datos org.compr.p.reg.info"
        centros     = self._centros_activos          # lista [(codigo, nombre), ...]
        NUM_BLOQUES = len(centros)

        if NOMBRE_H2 not in wb.sheetnames:
            _log(self._parent,
                 f"Hoja «{NOMBRE_H2}» no encontrada en el RINFO — omitida.", "WARN")
            return

        ws_h1 = wb[self.HOJA_PRINCIPAL]
        ws_h2 = wb[NOMBRE_H2]

        # ── 1. Leer Hoja1: col A (Registro info) y col E (Material) ──────────
        filas_h1 = []   # [(registro_info, material), ...]
        for row in ws_h1.iter_rows(min_row=2, min_col=1, max_col=5, values_only=True):
            reg_info = row[0]   # col A – Registro info de compras
            material = row[4]   # col E – Material
            if reg_info is None and material is None:
                continue
            filas_h1.append((reg_info, material))

        if not filas_h1:
            _log(self._parent,
                 "Hoja1 sin datos — Hoja 2 queda vacía.", "WARN")
            return

        # ── 2. Localizar y cargar LISTADO con las columnas extra ──────────────
        #    (antes ~25 líneas propias; consolidado en _localizar_ruta_listado_dali)
        ruta_listado = _localizar_ruta_listado_dali(
            self._carpeta_tarifas, getattr(self, "_carpeta_proyecto", None), self._parent)

        # diccionario: codigo_SAP → (desc_sap, grp_prod, ind_imp, grp_comp, desc_grp_comp)
        #   LISTADO col D (idx 3)  = CODIGO SAP - TEST        ← clave de cruce
        #   LISTADO col E (idx 4)  = DESCRIPCION SAP - TEST   → K  (Texto breve)
        #   LISTADO col H (idx 7)  = GRUPO PRODUCTOS SAP      → L  (Grupo artículos)
        #   LISTADO col J (idx 9)  = Indicador impuestos SAP  → AH
        #   LISTADO col M (idx 12) = GRUPO COMPRAS SAP        → H  (Grupo compras)
        #   LISTADO col N (idx 13) = NOMBRE GRUPO COMPRAS SAP → I  (Desc. grupo compras)
        # (mapeo actualizado 2026-08: tras insertar las 3 columnas nuevas
        #  F, K, L en LISTADO CODIGOS DALI-SAP, lo que había se desplazó)
        diccionario_h2 = {}
        if ruta_listado:
            try:
                wb_l  = abrir_como_listado(ruta_listado)
                ws_l  = wb_l.active
                for row in ws_l.iter_rows(min_row=8, values_only=True):
                    if not row or len(row) < 4 or row[3] is None:
                        continue
                    raw = row[3]
                    if isinstance(raw, float):
                        raw = int(raw)
                    clave = str(raw).strip()
                    if not clave:
                        continue
                    desc_sap    = str(row[4]).strip()  if len(row) > 4  and row[4]  is not None else ""
                    grp_prod    = str(row[7]).strip()  if len(row) > 7  and row[7]  is not None else ""
                    ind_imp     = str(row[9]).strip()  if len(row) > 9  and row[9]  is not None else ""
                    grp_comp    = str(row[12]).strip() if len(row) > 12 and row[12] is not None else ""
                    desc_grp_c  = str(row[13]).strip() if len(row) > 13 and row[13] is not None else ""
                    diccionario_h2[clave] = (desc_sap, grp_prod, ind_imp, grp_comp, desc_grp_c)
                wb_l.close()
                _log(self._parent,
                     f"LISTADO para Hoja2 cargado: {len(diccionario_h2)} entradas.", "INFO")
            except Exception as e:
                _log(self._parent,
                     f"Error al leer LISTADO para Hoja2: {e} — columnas de cruce vacías.", "WARN")
        else:
            _log(self._parent,
                 "LISTADO no encontrado — Hoja2 sin datos de cruce (H, I, K, L, AH).", "WARN")

        # ── 3. Expandir a 10 bloques ──────────────────────────────────────────
        #    La especificación indica que los registros de Hoja1 se repiten
        #    10 veces completos (bloque 1: 1–N, bloque 2: 1–N, ..., bloque 10: 1–N).
        filas_expandidas = filas_h1 * NUM_BLOQUES

        # ── 4. Escribir filas en Hoja2 ────────────────────────────────────────
        #
        #  Deduplicación funcional H2 (Punto B):
        #    clave = proveedor|org|material|centro
        #  Blinda contra duplicados aunque el mismo material aparezca varias veces
        #  en el input o en distintas iteraciones del proceso.
        _prov_h2_cr = str(self._codigo_sap).strip() if self._codigo_sap else ""
        _ORG_H2_CR  = "IC00"
        _escritas_h2_cr: set = set()
        saltadas_dup_h2_cr = 0

        n_filas_h1 = len(filas_h1)
        matches = 0
        fila_destino_h2 = 1   # apuntador real de escritura (salta si hay dup)
        for idx_exp, (reg_info, material) in enumerate(filas_expandidas):

            # Determinar bloque (0..N-1) y centro correspondiente
            bloque            = idx_exp // n_filas_h1
            cod_centro, nom_centro = centros[bloque]

            # Normalizar clave de cruce
            if material is not None:
                raw_m = material
                raw_m = normalizar_codigo_sap(raw_m)
                clave_mat = str(raw_m).strip()
            else:
                clave_mat = ""

            # Guard funcional: omitir si ya existe esta combinación
            _clave_h2_cr = f"{_prov_h2_cr}|{_ORG_H2_CR}|{clave_mat}|{cod_centro}"
            if _clave_h2_cr in _escritas_h2_cr:
                saltadas_dup_h2_cr += 1
                continue
            _escritas_h2_cr.add(_clave_h2_cr)

            fila_destino_h2 += 1
            idx = fila_destino_h2   # fila real en la hoja (empieza en 2)

            datos = diccionario_h2.get(clave_mat, ("", "", "", "", ""))
            desc_sap, grp_prod, ind_imp, grp_comp, desc_grp_c = datos
            if datos != ("", "", "", "", ""):
                matches += 1

            # ── Valores de Hoja1 / fijos ──────────────────────────────────────
            ws_h2.cell(row=idx, column=1,  value=reg_info)              # A  Registro info
            ws_h2.cell(row=idx, column=2,  value="0")                   # B  Tipo registro
            ws_h2.cell(row=idx, column=3,  value="Normal")              # C  Desc. categoría
            ws_h2.cell(row=idx, column=4,  value="IC00")                # D  Org. compras
            ws_h2.cell(row=idx, column=5,  value="OC ISLAS CANANARIAS") # E  Desc. org. compras
            ws_h2.cell(row=idx, column=6,  value=cod_centro)            # F  Centro
            ws_h2.cell(row=idx, column=7,  value=nom_centro)            # G  Nombre de centro

            # ── Valores de cruce con LISTADO ──────────────────────────────────
            ws_h2.cell(row=idx, column=8,  value=grp_comp   or None)   # H  Grupo compras
            ws_h2.cell(row=idx, column=9,  value=desc_grp_c or None)   # I  Desc. grupo compras

            # ── Material (clave de cruce, desde Hoja1 col E) ──────────────────
            ws_h2.cell(row=idx, column=10, value=material)              # J  Material

            # ── Más valores de cruce ──────────────────────────────────────────
            ws_h2.cell(row=idx, column=11, value=desc_sap  or None)    # K  Texto breve material
            ws_h2.cell(row=idx, column=12, value=grp_prod  or None)    # L  Grupo de artículos

            # ── Valores fijos adicionales ─────────────────────────────────────
            # IMPORTANTE: se usan enteros (no strings) para que Excel no muestre
            # el triángulo verde de "número almacenado como texto", igualando
            # el comportamiento del Purchasing descargado directamente desde SAP.
            # N (14) y W/X (23/24): el manual SAP no fija un valor concreto para
            # estas columnas (solo indica su propósito). Los valores 1 / 0 / 0
            # están confirmados como correctos — información notificada a
            # posteriori por central/SAP para el circuito de Princess Canarias.
            ws_h2.cell(row=idx, column=13, value=1.0)                    # M  Cant. pedido estándar  (float — formato SAP #,##0.0)
            ws_h2.cell(row=idx, column=14, value=1)                     # N  Plazo entrega días (confirmado a posteriori)
            ws_h2.cell(row=idx, column=19, value="X")                   # S  Verif. factura EM
            ws_h2.cell(row=idx, column=23, value=0.0)                   # W  Cant. pedido máxima (confirmado a posteriori)
            ws_h2.cell(row=idx, column=24, value=0.0)                   # X  Cant. mínima pedido (confirmado a posteriori)
            ws_h2.cell(row=idx, column=30, value=0.0)                   # AD Tolerancia exceso %    (float — formato SAP #,##0.000)
            ws_h2.cell(row=idx, column=31, value=0.0)                   # AE Tolerancia faltas %    (float — formato SAP #,##0.000)
            ws_h2.cell(row=idx, column=32, value="1")                   # AF Control fecha precio   (string — formato SAP @)
            ws_h2.cell(row=idx, column=34, value=ind_imp    or None)   # AH Indicador impuestos
            ws_h2.cell(row=idx, column=36, value="01.01.1900")          # AJ Inic.val.tipo impos.
            # AK (37) "Entrega ilimitada permitida" — ELIMINADA 2026-08: no
            # figura en el manual SAP "Crear Registros Info" punto 3.1.2
            # (ver nota en COLUMNAS_FIJAS_H2_SOLO_NUEVA).
            ws_h2.cell(row=idx, column=41, value="X")                   # AO Nueva

            # Copiar formato de fila 2 de esta hoja a la fila recién escrita.
            # La fila 2 del template contiene los estilos reales de datos (fill,
            # border, font). Se respeta el formato propio de Hoja2, distinto al
            # de Hoja1 y Hoja3.
            aplicar_formato_fila(ws_h2, idx, "Datos org.compr.p.reg.info", fila_ref=2)

        codigos_str = ", ".join(c for c, _ in centros)
        _log(self._parent,
             f"Hoja2 rellenada: {len(filas_expandidas)} filas "
             f"({len(filas_h1)} artículos × {NUM_BLOQUES} bloques) — "
             f"centros: {codigos_str} — {matches} cruces LISTADO.", "OK")

    # ── Paso 7c: Relleno completo de Hoja 3 ──────────────────────────────────
    def _renumerar_col_k_h3(self, wb) -> None:
        """
        Renumera la columna K ("ID grupo condiciones") de Hoja3 desde la fila 2
        hasta la última fila con datos, con enteros consecutivos desde 1.

        NORMA SAP (confirmada): 1, 2, 3 ... para TODOS los materiales sin excepción.
        Heredados y nuevos reciben la misma lógica, sin distinción.
        """
        HOJA_3 = "Det.precios p.registros info"
        if HOJA_3 not in wb.sheetnames:
            return
        ws_h3 = wb[HOJA_3]
        contador = 0
        for fila_row in ws_h3.iter_rows(min_row=2, min_col=1, max_col=11):
            val_a = fila_row[0].value
            val_f = fila_row[5].value if len(fila_row) > 5 else None
            if val_a is not None or val_f is not None:
                contador += 1
                celda_k = fila_row[10]          # col K (índice 10, 0-based)
                celda_k.value         = str(contador)   # texto, no entero
                celda_k.data_type     = "s"             # fuerza t="s" en XML → SAP lo lee como shared string
                celda_k.number_format = "@"             # triangulito verde SAP
        _log(self._parent,
             f"Hoja3 col K — renumeración final: 1..{contador} "
             f"(norma SAP, consecutivo global sin huecos).", "OK")

    def _rellenar_hoja3(self, wb) -> None:
        """
        Rellena la hoja «Det.precios p.registros info» según la especificación (sección 3).

        Fuentes de datos
        ────────────────
        3.1  Hoja2 (Datos org.compr.p.reg.info):
             H3-A  ← H2-A  (Registro info de compras)
             H3-B  ← H2-B  (Tipo de registro info de compras)
             H3-C  ← H2-C  (Descripción de categoría de registro info)
             H3-D  ← H2-D  (Organización de compras)
             H3-E  ← H2-F  (Centro)
             H3-F  ← H2-J  (Material)
             H3-G  ← H2-K  (Texto breve material)
             H3-H  ← H2-L  (Grupo de artículos)

        3.2  Valores fijos:
             L(12) = "Condición base"     M(13) = "PPR0"
             N(14) = "Precio bruto predet."  P(16) = "EUR"
             Q(17) = "1"                  T(20) = "C"
             U(21) = "Cantidad"           V(22) = "1"
             W(23) = "1"                  X(24) = "No"
             Z(26) = "X"

        3.3  Cruce LISTADO CODIGOS DALI-SAP (col D ↔ H3 col F):
             R(18) ← LISTADO col G (UNIDAD SAP)
             S(19) ← LISTADO col G (UNIDAD SAP)
             (mapeo actualizado 2026-08: antigua col F -> G, tras insertar
             las 3 columnas nuevas F, K, L en LISTADO CODIGOS DALI-SAP)

        3.4  Cruce tarifa "nombre proveedor [dd-mm-yyyy].xlsx" (col A ↔ H3 col F):
             O(15) ← tarifa col F (PRECIO)

        Solo se escribe .value — el formato del template SAP se preserva intacto.
        """
        NOMBRE_H3 = "Det.precios p.registros info"
        NOMBRE_H2 = "Datos org.compr.p.reg.info"

        if NOMBRE_H3 not in wb.sheetnames:
            _log(self._parent,
                 f"Hoja «{NOMBRE_H3}» no encontrada en el RINFO — omitida.", "WARN")
            return

        if NOMBRE_H2 not in wb.sheetnames:
            _log(self._parent,
                 f"Hoja «{NOMBRE_H2}» no disponible — Hoja3 queda vacía.", "WARN")
            return

        ws_h2 = wb[NOMBRE_H2]
        ws_h3 = wb[NOMBRE_H3]

        # ── 1. Leer filas de Hoja2 ────────────────────────────────────────────
        # Columnas necesarias (índices 0-based dentro del rango max_col=12):
        #   idx 0  = col A  Registro info de compras
        #   idx 1  = col B  Tipo de registro
        #   idx 2  = col C  Desc. categoría registro
        #   idx 3  = col D  Organización de compras
        #   idx 5  = col F  Centro
        #   idx 9  = col J  Material
        #   idx 10 = col K  Texto breve material
        #   idx 11 = col L  Grupo de artículos
        filas_h2 = []
        for row in ws_h2.iter_rows(min_row=2, max_col=12, values_only=True):
            reg_info = row[0]
            material = row[9]
            if reg_info is None and material is None:
                continue
            filas_h2.append((
                reg_info,   # H3-A
                row[1],     # H3-B  tipo registro
                row[2],     # H3-C  desc. categoría
                row[3],     # H3-D  org. compras
                row[5],     # H3-E  centro  (Hoja2 col F)
                material,   # H3-F  material (Hoja2 col J)
                row[10],    # H3-G  texto breve material (Hoja2 col K)
                row[11],    # H3-H  grupo artículos (Hoja2 col L)
            ))

        if not filas_h2:
            _log(self._parent, "Hoja2 sin datos — Hoja3 queda vacía.", "WARN")
            return

        # ── 2. Diccionario LISTADO: codigo_SAP → UNIDAD SAP (col G) ──────────
        #    (antes ~25 líneas propias; consolidado en _localizar_ruta_listado_dali)
        ruta_listado = _localizar_ruta_listado_dali(
            self._carpeta_tarifas, getattr(self, "_carpeta_proyecto", None), self._parent)

        dic_listado = {}   # clave_SAP → unidad_SAP
        if ruta_listado:
            try:
                wb_l = abrir_como_listado(ruta_listado)
                ws_l = wb_l.active
                for row in ws_l.iter_rows(min_row=8, values_only=True):
                    if not row or len(row) < 7 or row[3] is None:
                        continue
                    raw = row[3]
                    if isinstance(raw, float):
                        raw = int(raw)
                    clave = str(raw).strip()
                    if not clave:
                        continue
                    unidad = str(row[6]).strip() if row[6] is not None else ""
                    dic_listado[clave] = unidad
                wb_l.close()
                _log(self._parent,
                     f"LISTADO para Hoja3 cargado: {len(dic_listado)} entradas.", "INFO")
            except Exception as exc:
                _log(self._parent,
                     f"Error al leer LISTADO para Hoja3: {exc} — cols R/S vacías.", "WARN")
        else:
            _log(self._parent,
                 "LISTADO no encontrado — Hoja3 sin cruce R/S (UNIDAD SAP).", "WARN")

        # ── 3. Diccionario Tarifa: codigo_SAP → precio (col F) ───────────────
        # Fuente: self._ruta_tarifa — hoja TARIFA — datos desde fila 8
        #   col A (idx 0) = CODIGO SAP   ← clave de cruce con H3-F
        #   col F (idx 5) = PRECIO       → H3-O (Cantidad)
        dic_tarifa = {}   # clave_SAP → precio
        try:
            wb_t = openpyxl.load_workbook(
                self._ruta_tarifa, read_only=True, data_only=True)
            hoja_t = (wb_t[HOJA_TARIFA]
                      if HOJA_TARIFA in wb_t.sheetnames
                      else wb_t.active)
            for row in hoja_t.iter_rows(min_row=8, min_col=1,
                                         max_col=6, values_only=True):
                cod   = row[0]   # col A  CODIGO SAP
                precio = row[5]  # col F  PRECIO
                if cod is None:
                    continue
                cod = normalizar_codigo_sap(cod)
                clave = str(cod).strip()
                if clave and precio is not None:
                    dic_tarifa[clave] = precio
            wb_t.close()
            _log(self._parent,
                 f"Tarifa para Hoja3 cargada: {len(dic_tarifa)} precios.", "INFO")
        except Exception as exc:
            _log(self._parent,
                 f"Error al leer precios de la tarifa para Hoja3: {exc} — col O vacía.",
                 "WARN")

        # ── 4. Escribir filas en Hoja3 ────────────────────────────────────────
        #
        #  Deduplicación funcional H3 (Punto B):
        #    clave = proveedor|org|material|centro|fecha_inicio|fecha_fin
        #  Blinda contra duplicados de cualquier origen (distintas hojas,
        #  iteraciones o materiales repetidos en el input).
        _prov_h3_cr = str(self._codigo_sap).strip() if self._codigo_sap else ""
        _ORG_H3_CR  = "IC00"
        _fi_h3_cr   = str(self._fecha_inicio).strip() if self._fecha_inicio else ""
        _ff_h3_cr   = str(self._fecha_fin).strip()    if self._fecha_fin    else ""
        _escritas_h3_cr: set = set()
        saltadas_dup_h3_cr = 0

        matches_listado = 0
        matches_tarifa  = 0
        fila_destino_h3 = 1   # apuntador real de escritura

        for (reg_info, tipo_reg, desc_cat, org_comp,
             centro, material, txt_mat, grp_art) in filas_h2:

            # Normalizar clave de cruce (material) — antes del guard
            if material is not None:
                raw_m = material
                raw_m = normalizar_codigo_sap(raw_m)
                clave_mat = str(raw_m).strip()
            else:
                clave_mat = ""

            # Guard funcional H3
            _centro_h3_cr = str(centro).strip() if centro is not None else ""
            _clave_h3_cr = (
                f"{_prov_h3_cr}|{_ORG_H3_CR}"
                f"|{clave_mat}|{_centro_h3_cr}"
                f"|{_fi_h3_cr}|{_ff_h3_cr}"
            )
            if _clave_h3_cr in _escritas_h3_cr:
                saltadas_dup_h3_cr += 1
                continue
            _escritas_h3_cr.add(_clave_h3_cr)

            fila_destino_h3 += 1
            idx = fila_destino_h3   # fila real en la hoja (empieza en 2)

            # NORMA SAP — ID grupo condiciones (col K): entero consecutivo desde 1.
            # Fila 2 → 1, fila 3 → 2, ... sin excepción para ningún material.
            # Confirmado por SAP: aplica igual a heredados y nuevos.
            id_grupo_k = fila_destino_h3 - 1

            # ── Sección 3.1: valores copiados de Hoja2 ────────────────────────
            ws_h3.cell(row=idx, column=1,  value=reg_info)   # A  Registro info
            ws_h3.cell(row=idx, column=2,  value=tipo_reg)   # B  Tipo registro
            ws_h3.cell(row=idx, column=3,  value=desc_cat)   # C  Desc. categoría
            ws_h3.cell(row=idx, column=4,  value=org_comp)   # D  Org. compras
            ws_h3.cell(row=idx, column=5,  value=centro)     # E  Centro
            ws_h3.cell(row=idx, column=6,  value=material)   # F  Material
            ws_h3.cell(row=idx, column=7,  value=txt_mat)    # G  Texto breve material
            ws_h3.cell(row=idx, column=8,  value=grp_art)    # H  Grupo artículos

            # ── Sección 3.2: valores fijos ─────────────────────────────────────
            # Fechas de validez (cols I/J):
            #   En la creación desde cero, todos los materiales son nuevos:
            #   I = self._fecha_inicio  (introducida por el usuario)
            #   J = self._fecha_fin     (introducida por el usuario)
            #   No hay fechas heredadas de un Purchasing previo, por lo que
            #   no existe riesgo de solapamiento en este flujo.
            #   (La gestión de solapamiento aplica solo en GeneradorDesdeExistente.)
            _precio_fila = dic_tarifa.get(clave_mat)
            ws_h3.cell(row=idx, column=9,  value=self._fecha_inicio)   # I  Inicio validez
            ws_h3.cell(row=idx, column=10, value=self._fecha_fin)       # J  Fin validez
            _celda_k = ws_h3.cell(row=idx, column=11)
            _celda_k.value         = str(id_grupo_k)   # texto, no entero
            _celda_k.data_type     = "s"               # fuerza t="s" en XML → SAP lo lee como shared string
            _celda_k.number_format = "@"               # K  ID grupo condiciones (entero consecutivo desde 1)
            ws_h3.cell(row=idx, column=12, value="Condición base")       # L  Orden condiciones
            ws_h3.cell(row=idx, column=13, value="PPR0")                 # M  Tipo condición
            ws_h3.cell(row=idx, column=14, value="Precio bruto predet.") # N  Desc. clase cond.
            ws_h3.cell(row=idx, column=16, value="EUR")                  # P  Unidad importe
            ws_h3.cell(row=idx, column=17, value=1)                     # Q  Determinación precios
            ws_h3.cell(row=idx, column=20, value="C")                    # T  Clase de cálculo
            ws_h3.cell(row=idx, column=21, value="Cantidad")             # U  Desc. clase cálculo
            ws_h3.cell(row=idx, column=22, value=1)                    # V  Numerador conversión
            ws_h3.cell(row=idx, column=23, value=1)                    # W  Denominador conversión
            ws_h3.cell(row=idx, column=24, value="No")                   # X  Escala
            ws_h3.cell(row=idx, column=26, value="X")                    # Z  Nueva

            # ── Sección 3.3: cruce LISTADO → R(18) y S(19) ────────────────────
            unidad = dic_listado.get(clave_mat, "")
            if unidad:
                ws_h3.cell(row=idx, column=18, value=unidad)   # R  Unidad det. precio
                ws_h3.cell(row=idx, column=19, value=unidad)   # S  Unidad medida base
                matches_listado += 1

            # ── Sección 3.4: cruce tarifa → O(15) ─────────────────────────────
            # Reutiliza _precio_fila calculado en sección 3.2 (evita segunda búsqueda)
            if _precio_fila is not None:
                ws_h3.cell(row=idx, column=15, value=_precio_fila)   # O  Cantidad (precio)
                matches_tarifa += 1

            # Copiar formato de fila 2 de esta hoja a la fila recién escrita.
            # La fila 2 del template contiene los estilos reales de datos (fill,
            # border, font). Se respeta el formato propio de Hoja3, distinto al
            # de Hoja1 y Hoja2.
            aplicar_formato_fila(ws_h3, idx, "Det.precios p.registros info", fila_ref=2)

        _n_h3_escritas = fila_destino_h3 - 1
        _log(self._parent,
             f"Hoja3 rellenada: {_n_h3_escritas} fila(s)"
             + (f" ({saltadas_dup_h3_cr} omitida(s) por duplicado funcional)" if saltadas_dup_h3_cr else "")
             + f", I={self._fecha_inicio} / J={self._fecha_fin},"
             f" {matches_listado} cruces LISTADO (R/S),"
             f" {matches_tarifa} cruces tarifa (O).", "OK")

    def _aplicar_valores_fijos(self, ws_rinfo) -> None:
        """
        Escribe los valores fijos obligatorios (spec seccion 3) en todas las
        filas desde la 2 hasta la ultima con valor en col E (Material).
        Solo se toca .value -- el formato de destino no se modifica nunca.

        Columnas afectadas (Hoja 1 — Info general p.reg.info):
            col I  ( 9) Disponible desde       = FECHA_INICIO_FIJA_HOJA1  (dd.mm.yyyy) — fija, NO la del usuario
            col J  (10) Disponible hasta        = FECHA_FIN_FIJA_HOJA1     (dd.mm.yyyy) — fija, NO la del usuario
            col O  (15) Denominador conversion  = "1"
            col Q  (17) Numerador conversion    = "1"
            col AK (37) Nueva                   = "X"

        NOTA: Las fechas de Hoja 1 (I/J) son SIEMPRE fijas, tomadas de las
        constantes FECHA_INICIO_FIJA_HOJA1 / FECHA_FIN_FIJA_HOJA1 (actualmente
        01.01.2025 / 31.12.2030), independientemente de las fechas de validez
        introducidas por el usuario en VentanaFechasRINFO (self._fecha_inicio /
        self._fecha_fin). Esas fechas del usuario solo se aplican en Hoja 3
        (precios/materiales).

        AUDITORÍA 2026-08: valor confirmado como correcto — corresponde a la
        información notificada a posteriori por central/SAP para grabar el
        registro. Si dicha información cambia, actualizar únicamente las
        constantes FECHA_INICIO_FIJA_HOJA1 / FECHA_FIN_FIJA_HOJA1 (no este
        docstring, que debe reflejar siempre el criterio, no el valor literal).
        """
        ultima_fila = ultima_fila_con_datos(ws_rinfo, columna=5)

        if ultima_fila < 2:
            return

        telefono = getattr(self, "_telefono_proveedor", None)

        for fila_num in range(2, ultima_fila + 1):
            ws_rinfo.cell(row=fila_num, column=9).value  = FECHA_INICIO_FIJA_HOJA1  # I  Disponible desde (fija)
            ws_rinfo.cell(row=fila_num, column=10).value = FECHA_FIN_FIJA_HOJA1     # J  Disponible hasta (fija)
            ws_rinfo.cell(row=fila_num, column=15).value = 1                   # O  Denominador conversion
            ws_rinfo.cell(row=fila_num, column=17).value = 1                   # Q  Numerador conversion
            ws_rinfo.cell(row=fila_num, column=37).value = "X"                 # AK Nueva
            # col AD (30) = Número de teléfono de proveedor — solo si fue indicado
            if telefono:
                ws_rinfo.cell(row=fila_num, column=30).value = telefono  # AD Teléfono

        if telefono:
            _log(self._parent,
                 f"  Teléfono proveedor (col AD) escrito en filas 2-{ultima_fila}: "
                 f"«{telefono}».", "OK")
        _log(self._parent,
             f"  Valores fijos (I={FECHA_INICIO_FIJA_HOJA1} / J={FECHA_FIN_FIJA_HOJA1} / O/Q/AK) "
             f"escritos en filas 2-{ultima_fila} (Hoja1).", "OK")

    def _intentar_eliminar(self, ruta: str):
        """Limpieza: elimina el archivo de destino si existio un error."""
        try:
            if os.path.isfile(ruta):
                os.remove(ruta)
        except Exception:
            pass
