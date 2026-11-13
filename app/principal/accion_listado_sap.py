# -*- coding: utf-8 -*-
"""
accion_listado_sap.py
Acción 3 — Actualizar el listado de materiales SAP (lecturas de Productos, categoría de
valoración y artículos DALI, auditoría de duplicados y cruce).

Mixin de App (tarifa_app.py): sus métodos se usan como self.metodo(). Separado de tarifa_app.py
en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""

import os
from tkinter import messagebox
from datetime import datetime
import threading
from app.comun.tarifa_constants import (
    TABLA_SUBFAMILIA_GRUPO_COMPRAS, TABLA_IGIC_GRUPO_COMPRAS,
)
from app.comun.tarifa_core import (
    FOLDER_RECURSOS,
)
from app.comun.sap_helpers import leer_productos_xlsx
from app.ui.tarifa_dialogs import (
    VentanaResolverSAPDuplicados,
)


class AccionListadoSap:
    def _leer_productos_sap(self, ruta_productos):
        """
        LECTURA 1 (extraida de _actualizar_listado_sap -> proceso()):
        lee Productos.xlsx y construye tabla_productos, clave=codigo SAP
        (str), valor=(desc_prod, texto_um, unidad_base, grp_prod, desc_grp,
        tipo_producto).

        Mapeo de columnas Productos.xlsx -> destino (actualizado 2026-08:
        se insertaron 3 columnas nuevas F, K, L en LISTADO CODIGOS DALI-SAP,
        desplazando las que había):
            C (Descripción producto)              -> E  (DESCRIPCION SAP)
            J (Texto UM expl.)                     -> F  (nueva)
            I (Unidad medida base)                 -> G  (UNIDAD SAP)
            F (Grupo de productos)                 -> H  (GRUPO PRODUCTOS SAP)
            G (Descripción de grupo de productos)  -> I  (Desc. grupo productos SAP)
            D (Tipo de producto)                   -> L  (nueva)

        Devuelve (tabla_productos, filas_prod) o (None, 0) si el archivo
        no tiene filas validas -- en ese caso ya deja la barra reseteada
        y el aviso mostrado al usuario, igual que hacia el bloque inline.
        filas_prod se devuelve por separado (no basta con len(tabla_productos))
        porque si Productos.xlsx trae códigos SAP duplicados, el dict los
        deduplica pero el recuento de filas real debe conservarse tal cual
        -- se usa más adelante en la celda H5 y en el resumen final.
        """
        # ── LECTURA 1: Productos.xlsx ─────────────────────────────────
        # Clave cruce: col B (índice 1) → Código producto SAP
        self.log_consola("Leyendo Productos.xlsx...", "INFO")
        # (v10.2.50) La lectura está en sap_helpers.leer_productos_xlsx (sin cambios), para compartirla.
        tabla_productos, filas_prod = leer_productos_xlsx(ruta_productos)
        self.log_consola(f"  └ {filas_prod} productos cargados (Productos.xlsx)", "OK")

        if filas_prod == 0:
            self.after(0, lambda: self._barra.reset("Sin datos en Productos.xlsx"))
            self.after(0, lambda: messagebox.showwarning(
                "Sin datos",
                "No se encontraron filas con datos\n"
                "en el archivo Productos.xlsx."))
            return None, 0

        return tabla_productos, filas_prod

    def _leer_categoria_valoracion(self, ruta_listas, tabla_productos=None):
        """
        Lee "Listas de productos.xlsx" (sap_helpers.leer_listas_productos): {código SAP: Categoría valoración}
        -> columna K del catálogo. Una fila por código y centro: si un código tiene valores distintos según el
        centro, se toma el más frecuente y se devuelve en conflictos (clave, {valor: n}).

        (13 nov 2026) Las columnas se buscan por su cabecera («Producto», «Categoría valoración») y el código se
        saca de «DESCRIPCIÓN (código)». Con `tabla_productos`, si el listado no casa con Productos.xlsx lanza
        ValueError ANTES de escribir nada (sap_helpers.comprobar_listas_productos): así no se vacía la categoría
        de todo el catálogo ni se envía a la web, como pasó ese día con una exportación con otras columnas.
        Devuelve (tabla_categoria, conflictos).
        """
        from app.comun.sap_helpers import comprobar_listas_productos, leer_listas_productos

        self.log_consola("Leyendo Listas de productos.xlsx...", "INFO")
        tabla_categoria, conflictos, filas_listas, por_cabecera = leer_listas_productos(ruta_listas)
        self.log_consola(
            f"  └ {filas_listas} fila(s) leídas, {len(tabla_categoria)} "
            f"código(s) SAP con Categoría valoración (Listas de productos.xlsx)", "OK")
        if not por_cabecera:
            self.log_consola("  ⚠  No se encontraron las columnas «Producto» y «Categoría valoración» por su "
                             "nombre: se usan las de siempre (A y F).", "WARN")
        motivo = comprobar_listas_productos(tabla_categoria, tabla_productos) if tabla_productos is not None else None
        if motivo:
            self.log_consola(f"  ✘  {motivo}", "ERROR")
            raise ValueError(motivo)
        if conflictos:
            self.log_consola(
                f"  ⚠  {len(conflictos)} código(s) con Categoría valoración "
                f"distinta según el centro — se tomó el valor más frecuente:", "WARN")
            for clave, valores in conflictos[:20]:
                self.log_consola(f"      · {clave}: {valores}", "WARN")
            if len(conflictos) > 20:
                self.log_consola(f"      … y {len(conflictos) - 20} más (omitidos del log)", "WARN")

        return tabla_categoria, conflictos

    def _leer_articulos_dali_sap(self, ruta_dali):
        """
        LECTURA 2 (extraida de _actualizar_listado_sap -> proceso()):
        lee ARTICULOS_DALI.xlsx y construye tabla_dali, clave=codigo DALI
        normalizado (str), valor=(subfamilia, igic).
        """
        import openpyxl
        # ── LECTURA 2: ARTICULOS_DALI.xlsx ───────────────────────────
        # Clave cruce: col A (índice 0) → Código DALI
        # Col E (índice 4) → SUBFAMILIA  → lookup en TABLA_SUBFAMILIA
        # Col G (índice 6)→ IGIC        → lookup en TABLA_IGIC
        # (Datos desde fila 8, igual que el resto del sistema)
        self.after(0, lambda: self._barra.indeterminate("Leyendo ARTICULOS_DALI..."))
        self.log_consola("Leyendo ARTICULOS_DALI.xlsx...", "INFO")
        wb_dali = openpyxl.load_workbook(ruta_dali, read_only=True, data_only=True)
        ws_dali = wb_dali.active

        # tabla_dali: clave = CODIGO DALI (col A), valor = (subfamilia, igic)
        # Estructura de ARTICULOS_DALI.xlsx (generado por Acción 1):
        #   A(0)=CODIGO  B(1)=DESCRIPCION  C(2)=UNIDAD  D(3)=FAMILIA
        #   E(4)=SUBFAMILIA  F(5)=NATURALEZA  G(6)=IGIC  H(7)=PROVEEDOR  I(8)=ACTIVO
        #
        # Cruce I   : A destino <-> A ARTICULOS_DALI -> leer G(IGIC)       -> col I destino
        # Cruce J/K : A destino <-> A ARTICULOS_DALI -> leer E(SUBFAMILIA) -> cols J+K destino
        tabla_dali = {}   # { cod_dali_norm: (subfamilia, igic) }
        filas_dali = 0
        FILA_INI_DALI = 8
        for fila in ws_dali.iter_rows(min_row=FILA_INI_DALI, values_only=True):
            if all(v is None for v in fila):
                continue
            cod_raw = fila[0] if len(fila) > 0 else None   # col A — CODIGO
            if cod_raw is None or str(cod_raw).strip() == "":
                continue
            # Normalizar clave: ARTICULOS_DALI almacena CODIGO como entero;
            # el destino puede tenerlo como "001234" (con ceros) -> unificar a "1234"
            try:
                cod_norm = str(int(float(str(cod_raw).strip())))
            except (ValueError, TypeError):
                cod_norm = str(cod_raw).strip()

            subfamilia = str(fila[4]).strip() if len(fila) > 4 and fila[4] is not None else ""  # col E — SUBFAMILIA
            igic       = str(fila[6]).strip() if len(fila) > 6 and fila[6] is not None else ""  # col G — IGIC
            tabla_dali[cod_norm] = (subfamilia, igic)
            filas_dali += 1

        wb_dali.close()
        self.log_consola(f"  └ {filas_dali} artículos cargados (ARTICULOS_DALI.xlsx)", "OK")

        return tabla_dali

    def _auditar_sap_duplicados_listado(self, ws):
        """
        Escanea TODO el catálogo (con forma de LISTADO, v10.2.50) (no solo lo que
        se esté tecleando en ese momento) buscando códigos SAP asignados
        a más de un código DALI distinto -- el problema real reportado
        por el usuario no es un DALI concreto, sino que dos DALI
        diferentes compartan el mismo SAP en cualquier parte del listado
        (caso real: DALI 15660 y 26046, ambos con SAP 10101038).

        Se llama desde _cruzar_y_escribir_listado_sap() (Acción 3), justo
        antes de proteger/guardar, con la hoja ya actualizada y aún
        editable. Si encuentra duplicados, reutiliza la misma ventana de
        resolución que VentanaSAPFaltantes (VentanaResolverSAPDuplicados)
        -- mostrada de forma thread-safe, ya que _actualizar_listado_sap
        corre en el hilo de fondo de su propio proceso().

        Los códigos marcados por el usuario para borrar se limpian
        directamente en `ws` (col D = None); el guardado posterior de
        _cruzar_y_escribir_listado_sap ya incluye este cambio.

        Devuelve el número de códigos SAP limpiados (0 si no había
        ninguno, o si el usuario canceló la resolución).
        """
        from collections import defaultdict
        _por_sap = defaultdict(list)   # {sap: [(dali, fila), ...]}
        for idx, row in enumerate(
                ws.iter_rows(min_row=8, min_col=1, max_col=4, values_only=True),
                start=8):
            v_dali = row[0]
            v_sap  = row[3]
            if v_dali is None or v_sap in (None, ""):
                continue
            try:
                dali_int = int(float(v_dali))
                sap_int  = int(float(v_sap))
            except (ValueError, TypeError):
                continue
            _por_sap[sap_int].append((dali_int, idx))

        conflictos = {
            cod_sap: [(dali, True, fila, "") for dali, fila in entradas]
            for cod_sap, entradas in _por_sap.items()
            if len(entradas) > 1
        }
        if not conflictos:
            return 0

        self.log_consola(
            f"  ⚠ {len(conflictos)} código(s) SAP duplicados detectados en "
            f"el catálogo completo -- pidiendo resolución al usuario...",
            "WARN")

        def _mostrar():
            dlg = VentanaResolverSAPDuplicados(self, conflictos)
            self.wait_window(dlg)
            return dlg.resultado

        a_borrar = self._ejecutar_en_hilo_principal_y_esperar(_mostrar)

        if not a_borrar:
            self.log_consola(
                "  ⚠ Auditoría de duplicados del catálogo cancelada o sin "
                "cambios -- quedan sin resolver.", "WARN")
            return 0

        limpiados = 0
        for cod_sap, candidatos in conflictos.items():
            for cod_dali, _es, fila, _desc in candidatos:
                if cod_dali in a_borrar:
                    ws.cell(fila, 4).value = None
                    limpiados += 1
        self.log_consola(
            f"  ✔ {limpiados} código(s) SAP duplicados limpiados del "
            f"catálogo.", "OK")
        return limpiados

    def _cruzar_y_escribir_listado_sap(self, ruta_destino, tabla_productos,
                                        tabla_dali, tabla_categoria, conflictos_categoria,
                                        filas_prod,
                                        fecha_productos_solo, hora_productos_solo):
        """
        CRUCE Y ESCRITURA EN DESTINO + Resumen en consola (extraido de
        _actualizar_listado_sap -> proceso()). Es literalmente el resto
        del cuerpo del try de proceso() tras las lecturas: abre el
        destino, limpia columnas E:N, cruza Productos.xlsx,
        ARTICULOS_DALI.xlsx y Listas de productos.xlsx (Cruce 1, Cruce 1b
        y Cruce 2), actualiza cabecera y contadores H1-H3/H5, protege y
        guarda, y muestra el resumen final en consola + mensaje al usuario.

        Mapeo de columnas destino (actualizado 2026-08, tras insertar las
        3 columnas nuevas F, K, L en LISTADO CODIGOS DALI-SAP -- las que
        ya existían antes se desplazaron: antigua F/G/H/I -> G/H/I/J,
        antigua J/K/L -> M/N/O):
            E, F, G, H, I, L  <- Productos.xlsx (Cruce 1, ver _leer_productos_sap)
            K                 <- Listas de productos.xlsx (Cruce 1b, Categoría valoración)
            J                 <- IGIC (Cruce 2, vía ARTICULOS_DALI + TABLA_IGIC_GRUPO_COMPRAS)
            M, N              <- Grupo de Compras (Cruce 2, vía ARTICULOS_DALI + TABLA_SUBFAMILIA_GRUPO_COMPRAS)
            O                 <- ACTIVO DALI (no se escribe aquí; solo se lee para H1/H2/H3)

        No devuelve nada -- es el ultimo tramo del try de proceso(),
        nada se usa despues (solo el except que envuelve a proceso()).
        """
        from app.comun import catalogo_local
        # ── CRUCE Y ESCRITURA EN DESTINO ─────────────────────────────
        # v10.2.50: ya no hay «LISTADO CODIGOS DALI - SAP.xlsx». El cruce de siempre se hace sobre un libro EN
        # MEMORIA con la misma forma (filas desde la 8, columnas A..O) sacado del catálogo, y al final los datos
        # SAP se pasan al catálogo por código DALI (regenerándolo). La lógica de abajo no cambia.
        self.log_consola("Abriendo destino: catalogo-dali-completo.xlsx", "INFO")
        self.after(0, lambda: self._barra.indeterminate("Cruzando y escribiendo..."))

        # Pre-normalizar tablas de lookup a uppercase para acceso O(1).
        # Se hace una sola vez aquí en lugar de comparar k.upper() en cada fila.
        TABLA_SUBFAMILIA_U = {k.upper(): v for k, v in TABLA_SUBFAMILIA_GRUPO_COMPRAS.items()}
        TABLA_IGIC_U       = {k.upper(): v for k, v in TABLA_IGIC_GRUPO_COMPRAS.items()}

        wb_dest = catalogo_local.libro_formato_listado(ruta_destino)
        ws_dest = wb_dest.active

        # ── LIMPIEZA PREVIA: borrar columnas E:N desde fila 8 hasta la última con datos ──
        # Se eliminan los valores de E(5) a N(14) antes de rellenarlos con los
        # datos cruzados actualizados, para evitar valores residuales de
        # ejecuciones anteriores. No se toca O (15, ACTIVO DALI): esa columna
        # no se escribe en este proceso, solo se lee para H1/H2/H3.
        self.log_consola("Limpiando columnas E:N (fila 8 → última fila)...", "INFO")
        self.after(0, lambda: self._barra.indeterminate("Limpiando columnas E:N..."))
        _ultima_fila_dest = ws_dest.max_row
        _COLS_LIMPIAR = (5, 6, 7, 8, 9, 10, 11, 12, 13, 14)   # E, F, G, H, I, J, K, L, M, N
        _celdas_borradas = 0
        for _fila_lim in range(8, _ultima_fila_dest + 1):
            for _col_lim in _COLS_LIMPIAR:
                _celda = ws_dest.cell(row=_fila_lim, column=_col_lim)
                if _celda.value is not None:
                    _celda.value = None
                    _celdas_borradas += 1
        self.log_consola(
            f"  └ {_celdas_borradas} celda(s) borradas en E:N "
            f"(filas 8–{_ultima_fila_dest})", "OK")
        # ─────────────────────────────────────────────────────────────────

        act_productos       = 0   # filas con match en Productos.xlsx (E,F,G,H,I,L)
        sin_productos       = 0
        lista_sin_productos = []  # códigos SAP sin match (para log detallado)
        act_categoria       = 0   # filas con match en Listas de productos.xlsx (K)
        sin_categoria       = 0
        lista_sin_categoria = []  # códigos SAP sin match en Listas de productos
        act_dali            = 0   # filas con match en ARTICULOS_DALI (J, M, N)
        sin_dali            = 0
        lista_sin_dali      = []  # códigos DALI sin match (para log detallado)
        sin_subfam          = 0   # naturalezas no encontradas en el diccionario
        dict_sin_subfam     = {} # {naturaleza: [(cod_dali, cod_sap)]} para log agrupado
        sin_igic            = 0   # IGIC no encontrado en el diccionario
        # Contadores para celdas de resumen H1–H3
        h1 = 0   # D con valor  y L = "SI"
        h2 = 0   # D con valor  y L = "NO"
        h3 = 0   # D sin valor  y L = "SI"

        FILA_INI = 8
        for fila_num in range(FILA_INI, ws_dest.max_row + 1):

            # ── Cruce 1: col D destino ↔ col B Productos ──────────────
            cod_sap = ws_dest.cell(row=fila_num, column=4).value   # col D
            if cod_sap is not None and str(cod_sap).strip() != "":
                clave_sap = str(cod_sap).strip()
                if clave_sap in tabla_productos:
                    (desc_prod, texto_um, unidad_base,
                     grp_prod, desc_grp, tipo_producto) = tabla_productos[clave_sap]
                    ws_dest.cell(row=fila_num, column=5,  value=desc_prod)      # E
                    ws_dest.cell(row=fila_num, column=6,  value=texto_um)       # F
                    ws_dest.cell(row=fila_num, column=7,  value=unidad_base)    # G
                    ws_dest.cell(row=fila_num, column=8,  value=grp_prod)       # H
                    ws_dest.cell(row=fila_num, column=9,  value=desc_grp)       # I
                    ws_dest.cell(row=fila_num, column=12, value=tipo_producto)  # L
                    act_productos += 1
                else:
                    sin_productos += 1
                    lista_sin_productos.append(clave_sap)

                # ── Cruce 1b: col D destino ↔ col A Listas de productos ──
                if clave_sap in tabla_categoria:
                    ws_dest.cell(row=fila_num, column=11,
                                 value=tabla_categoria[clave_sap])   # K
                    act_categoria += 1
                else:
                    sin_categoria += 1
                    lista_sin_categoria.append(clave_sap)

            # ── Cruce 2: col A destino <-> col A ARTICULOS_DALI ─────────
            # Rellena col J (IGIC) y cols M+N (GRUPO COMPRAS) del destino.
            # La clave se normaliza igual que en tabla_dali (int -> str sin ceros).
            cod_dali = ws_dest.cell(row=fila_num, column=1).value   # col A destino
            if cod_dali is None or str(cod_dali).strip() == "":
                continue   # fila sin codigo DALI, pasar a la siguiente

            try:
                clave_dali = str(int(float(str(cod_dali).strip())))
            except (ValueError, TypeError):
                clave_dali = str(cod_dali).strip()

            if clave_dali not in tabla_dali:
                sin_dali += 1
                lista_sin_dali.append(clave_dali)
                continue

            subfamilia_raw, igic_raw = tabla_dali[clave_dali]
            act_dali += 1

            # ── Condición: sólo rellenar I/J/K si col D tiene valor ──────
            # El Cruce 2 escribe datos SAP derivados (IGIC, Grupo Compras)
            # únicamente en filas que ya disponen de código SAP en col D.
            val_d_cruce2 = ws_dest.cell(row=fila_num, column=4).value
            if val_d_cruce2 is None or str(val_d_cruce2).strip() == "":
                # Sin código SAP: no se escriben I, J ni K
                pass
            else:
                # col G de ARTICULOS_DALI (IGIC) -> col J destino
                igic_match = TABLA_IGIC_U.get(igic_raw.upper().strip())
                if igic_match:
                    ws_dest.cell(row=fila_num, column=10, value=igic_match[0])
                else:
                    sin_igic += 1

                # col E de ARTICULOS_DALI (SUBFAMILIA) -> cols M y N destino
                subfam_match = TABLA_SUBFAMILIA_U.get(subfamilia_raw.upper().strip())
                if subfam_match:
                    ws_dest.cell(row=fila_num, column=13, value=subfam_match[0])  # M
                    ws_dest.cell(row=fila_num, column=14, value=subfam_match[1])  # N
                else:
                    sin_subfam += 1
                    if subfamilia_raw:
                        cod_sap_str = str(val_d_cruce2).strip()
                        if subfamilia_raw not in dict_sin_subfam:
                            dict_sin_subfam[subfamilia_raw] = []
                        dict_sin_subfam[subfamilia_raw].append((clave_dali, cod_sap_str))

            # ── Contadores resumen I1/I2/I3/I5 ───────────────────────────
            # Se evalúan aquí, tras escribir, para usar los valores finales de D y O
            val_d = ws_dest.cell(row=fila_num, column=4).value    # col D (SAP)
            val_o = ws_dest.cell(row=fila_num, column=15).value   # col O (ACTIVO DALI)
            d_tiene_valor = val_d is not None and str(val_d).strip() != ""
            o_es_si  = str(val_o).strip().upper() == "SI" if val_o is not None else False
            o_es_no  = str(val_o).strip().upper() == "NO" if val_o is not None else False
            if d_tiene_valor and o_es_si:
                h1 += 1
            elif d_tiene_valor and o_es_no:
                h2 += 1
            elif not d_tiene_valor and o_es_si:
                h3 += 1
        # ═══════════════════════════════════════════════════════════════
        #  ACTUALIZAR CELDAS DE CABECERA
        # ═══════════════════════════════════════════════════════════════
        try:
            ws_dest["D4"] = fecha_productos_solo
            ws_dest["D5"] = hora_productos_solo
            ws_dest["I1"] = h1          # D con valor  y O = "SI"
            ws_dest["I2"] = h2          # D con valor  y O = "NO"
            ws_dest["I3"] = h3          # D sin valor  y O = "SI"
            ws_dest["I5"] = filas_prod  # total filas con valor en col B de Productos.xlsx
            self.log_consola(f"  └ Celda D4 actualizada: {fecha_productos_solo}", "OK")
            self.log_consola(f"  └ Celda D5 actualizada: {hora_productos_solo}", "OK")
            self.log_consola(f"  └ I1={h1}  I2={h2}  I3={h3}  I5={filas_prod}", "OK")
        except Exception as e_d4:
            self.log_consola(f"  ⚠  No se pudo actualizar celdas de cabecera: {e_d4}", "WARN")

        # ── Auditoría completa de códigos SAP duplicados en el LISTADO ──
        # (pedida por el usuario tras el caso real de DALI 15660/26046
        # compartiendo el mismo SAP) -- a diferencia del control de
        # VentanaSAPFaltantes (que solo mira los códigos que se están
        # tecleando en ese momento), esto revisa TODO el LISTADO ya
        # actualizado, buscando cualquier SAP asignado a más de un DALI.
        self._auditar_sap_duplicados_listado(ws_dest)

        # v10.2.50: los datos SAP cruzados pasan al catálogo (por código DALI) al regenerarlo, más abajo.
        tabla_sap = catalogo_local.tabla_sap_desde_libro(ws_dest)
        wb_dest.close()

        # ── Resumen en consola ────────────────────────────────────────
        # Constante: máximo de entradas detalle a mostrar por categoría.
        # El resto se resume en "… y N más" para evitar miles de repaints.
        _CAP = 30

        self.log_consola("Cruce terminado: se guarda en catalogo-dali-completo.xlsx", "OK")
        self.log_consola(f"  └ {act_productos} materiales actualizados desde Productos.xlsx", "OK")
        self.log_consola(f"  └ {act_categoria} materiales con Categoría valoración desde Listas de productos.xlsx", "OK")
        self.log_consola(f"  └ {act_dali} artículos cruzados con ARTICULOS_DALI", "OK")

        if lista_sin_categoria:
            lineas_batch = [
                (f"  └ ⚠  {sin_categoria} código(s) SAP sin coincidencia en Listas de productos.xlsx:", "WARN")
            ]
            for cod in lista_sin_categoria[:_CAP]:
                lineas_batch.append((f"      · {cod}", "WARN"))
            if len(lista_sin_categoria) > _CAP:
                lineas_batch.append(
                    (f"      … y {len(lista_sin_categoria) - _CAP} más (omitidos del log)", "WARN"))
            self.log_consola_batch(lineas_batch)

        if lista_sin_productos:
            lineas_batch = [
                (f"  └ ⚠  {sin_productos} código(s) SAP sin coincidencia en Productos.xlsx:", "WARN")
            ]
            for cod in lista_sin_productos[:_CAP]:
                lineas_batch.append((f"      · {cod}", "WARN"))
            if len(lista_sin_productos) > _CAP:
                lineas_batch.append(
                    (f"      … y {len(lista_sin_productos) - _CAP} más (omitidos del log)", "WARN"))
            self.log_consola_batch(lineas_batch)

        if lista_sin_dali:
            lineas_batch = [
                (f"  └ ⚠  {sin_dali} código(s) DALI sin coincidencia en ARTICULOS_DALI:", "WARN")
            ]
            for cod in lista_sin_dali[:_CAP]:
                lineas_batch.append((f"      · {cod}", "WARN"))
            if len(lista_sin_dali) > _CAP:
                lineas_batch.append(
                    (f"      … y {len(lista_sin_dali) - _CAP} más (omitidos del log)", "WARN"))
            self.log_consola_batch(lineas_batch)

        if sin_subfam:
            lineas_batch = [
                (f"  └ ⚠  {sin_subfam} artículo(s) sin mapeo de SUBFAMILIA "
                 f"({len(dict_sin_subfam)} subfamilia(s) distintas):", "WARN")
            ]
            total_mostrados = 0
            for nat, pares in dict_sin_subfam.items():
                if total_mostrados >= _CAP:
                    restantes = sum(len(v) for k, v in dict_sin_subfam.items()
                                    if k != nat) + len(pares)
                    lineas_batch.append(
                        (f"      … y {restantes} artículo(s) más en otras subfamilias (omitidos)", "WARN"))
                    break
                lineas_batch.append(
                    (f"      SUBFAMILIA '{nat}'  ({len(pares)} artículo(s)):", "WARN"))
                for cod_dali, cod_sap in pares:
                    if total_mostrados >= _CAP:
                        lineas_batch.append(
                            (f"          … y {len(pares) - (total_mostrados - 1)} más en esta subfamilia", "WARN"))
                        break
                    lineas_batch.append(
                        (f"          DALI: {cod_dali}   SAP: {cod_sap}", "WARN"))
                    total_mostrados += 1
            self.log_consola_batch(lineas_batch)

        if sin_igic:
            self.log_consola(f"  └ ⚠  {sin_igic} IGIC(s) sin mapeo en diccionario — revise log", "WARN")

        self.consola.separador("FIN ACTUALIZACIÓN SAP")
        self.after(0, lambda: self._barra.completado(
            f"SAP actualizado — {act_dali} filas ✔"))

        def _mostrar_resultado(ap=act_productos, ad=act_dali,
                               sp=sin_productos, sd=sin_dali,
                               ssf=sin_subfam, si=sin_igic):
            hay_warn = sp or sd or ssf or si
            lineas = [
                "\u2714  Actualizaci\u00f3n SAP completada correctamente",
                "",
                f"   Art\u00edculos SAP actualizados  : {ap}",
                f"   Art\u00edculos DALI actualizados : {ad}",
            ]
            if sp:
                lineas.append(f"   \u26a0 C\u00f3digos SAP sin coincidencia  : {sp}")
            if sd:
                lineas.append(f"   \u26a0 C\u00f3digos DALI sin coincidencia : {sd}")
            if ssf:
                lineas.append(f"   \u26a0 Subfamilia(s) sin mapeo       : {ssf}")
            if si:
                lineas.append(f"   \u26a0 IGIC(s) sin mapeo            : {si}")
            if hay_warn:
                lineas += [
                    "",
                    "Revise el log de consola para ver el detalle.",
                ]
            lineas += ["", f"\U0001f4c1  {ruta_destino}"]
            titulo = "SAP \u2014 Actualizaci\u00f3n completada" + (" con advertencias" if hay_warn else "")
            if hay_warn:
                messagebox.showwarning(titulo, "\n".join(lineas))
            else:
                messagebox.showinfo(titulo, "\n".join(lineas))

        try:
            self._regenerar_catalogo_dali_completo(tabla_sap=tabla_sap)
        except PermissionError:
            self.log_consola(
                "ERROR: No se pudo guardar — catalogo-dali-completo.xlsx está abierto en Excel.", "ERROR")
            self.after(0, lambda: self._barra.reset("Error al guardar"))
            self.after(0, lambda: messagebox.showerror(
                "Archivo bloqueado por Excel",
                "❌  No se pudo guardar:\n\n"
                "   RECURSOS/catalogo-dali-completo.xlsx\n\n"
                "El archivo está abierto en Excel.\n\n"
                "📋  Ciérrelo y vuelva a ejecutar la Acción 3."))
            return
        except Exception as _e_cat:
            self.log_consola(
                f"  ⚠  No se pudo guardar catalogo-dali-completo.xlsx: {_e_cat}", "ERROR")
            self.after(0, lambda: self._barra.reset("Error al guardar"))
            self.after(0, lambda err=str(_e_cat): messagebox.showerror(
                "Error al guardar el catálogo", f"No se pudo guardar el catálogo:\n\n{err}"))
            return

        if conflictos_categoria:
            def _mostrar_conflictos_categoria(conf=conflictos_categoria):
                lineas = [
                    "\u26a0  Se han encontrado códigos SAP con Categoría",
                    "   valoración distinta según el centro en",
                    "   \"Listas de productos.xlsx\".",
                    "",
                    "   Se ha usado el valor más frecuente en cada caso,",
                    "   asumiendo que el resto son errores. Revise y",
                    "   corrija estos códigos en el origen si procede:",
                    "",
                ]
                for clave, valores in conf[:20]:
                    detalle = ", ".join(f"{v}×{n}" for v, n in valores.items())
                    lineas.append(f"   • {clave}:  {detalle}")
                if len(conf) > 20:
                    lineas.append(f"   … y {len(conf) - 20} código(s) más (ver log de consola).")
                messagebox.showwarning(
                    "Categoría valoración ambigua — revisar origen", "\n".join(lineas))
            self.after(0, _mostrar_conflictos_categoria)

        self.after(0, _mostrar_resultado)

    def _avisar_listados_sap(self, codigos_sap, desde=None):
        """(13 nov 2026, petición de Víctor) Tras guardar códigos SAP (Acción 8 o códigos SAP faltantes de la
        Acción 4): si alguno no está en Productos.xlsx, avisa de que faltan los listados de SAP actualizados, guía la
        descarga (app/ui/ventana_listados_sap.py) y, cuando están en RECURSOS, lanza sola la Acción 3. Si se deja para
        más tarde, avisa de que la web queda parcialmente actualizada. Llamar desde el hilo principal.
        Devuelve "actualizar", "despues" o None (no hacía falta avisar)."""
        from app.comun import listados_sap
        from app.ui.ventana_listados_sap import VentanaListadosSap
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta or not codigos_sap:
            return None
        recursos = os.path.join(carpeta, FOLDER_RECURSOS)
        ruta_prod = listados_sap.rutas(recursos)[listados_sap.PRODUCTOS]
        tabla = None
        if os.path.isfile(ruta_prod):
            try:
                tabla, _filas = leer_productos_xlsx(ruta_prod)
            except Exception as e:  # noqa: BLE001
                self.log_consola(f"  ⚠  No se pudo leer Productos.xlsx: {e}", "WARN")
        fuera = listados_sap.codigos_sin_datos(codigos_sap, tabla)
        if not fuera:
            return None
        ven = VentanaListadosSap(self, fuera, recursos, ahora=desde)
        self.wait_window(ven)
        if ven.resultado == "actualizar":
            self.log_consola("Listados de SAP actualizados en RECURSOS: se lanza la Acción 3 para completar los "
                             "datos SAP.", "INFO")
            self.after(300, self._actualizar_listado_sap)
        else:
            self.log_consola("  ⚠  " + listados_sap.AVISO_PARCIAL, "WARN")
        return ven.resultado

    def _actualizar_listado_sap(self):
        """
        PUNTO 3 — Actualizar listado materiales SAP.

        Origen 1 : RECURSOS/Productos.xlsx           (fecha = última modificación del archivo)
        Origen 2 : RECURSOS/ARTICULOS_DALI.xlsx       (fecha = celda B3, texto "ACTUALIZADO …")
        Origen 3 : RECURSOS/Listas de productos.xlsx
        Destino  : RECURSOS/catalogo-dali-completo.xlsx (v10.2.50; antes, el LISTADO CODIGOS DALI - SAP)

        Cruce 1 — col D destino ↔ col B Productos.xlsx:
          C (Descripción producto)              → E  (DESCRIPCION SAP)
          J (Texto UM expl.)                    → F  (nueva)
          I (Unidad medida base)                → G  (UNIDAD SAP)
          F (Grupo de productos)                → H  (GRUPO PRODUCTOS SAP)
          G (Descripción de grupo de productos) → I  (Descripción de grupo de productos SAP)
          D (Tipo de producto)                  → L  (nueva)

        Cruce 1b — col D destino ↔ col A Listas de productos.xlsx:
          F (Categoría valoración) → K (nueva). Si un código SAP trae
          valores distintos según el centro, se usa el más frecuente y
          se avisa al usuario (mensaje visual + log detallado).

        Cruce 2 — col A destino ↔ col A ARTICULOS_DALI.xlsx:
          E (SUBFAMILIA) → diccionario SUBFAMILIA → M (GRUPO COMPRAS SAP) + N (NOMBRE GRUPO COMPRAS SAP)
          G (IGIC)       → diccionario IGIC       → J (Indicador de impuestos SAP)
        """
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            messagebox.showwarning("Sin proyecto",
                                   "Seleccione primero la carpeta del Proyecto.")
            return


        recursos_path  = os.path.join(carpeta, FOLDER_RECURSOS)
        ruta_productos = os.path.join(recursos_path, "Productos.xlsx")
        ruta_dali      = os.path.join(recursos_path, "ARTICULOS_DALI.xlsx")
        ruta_listas    = os.path.join(recursos_path, "Listas de productos.xlsx")
        # v10.2.50: el destino es el catálogo completo (ya no hay LISTADO)
        ruta_destino   = os.path.join(recursos_path, "catalogo-dali-completo.xlsx")

        # ── 1. Verificar archivos de origen ──────────────────────────
        faltantes = []
        if not os.path.isfile(ruta_productos):
            faltantes.append("Productos.xlsx")
        if not os.path.isfile(ruta_dali):
            faltantes.append("ARTICULOS_DALI.xlsx")
        if not os.path.isfile(ruta_listas):
            faltantes.append("Listas de productos.xlsx")

        if faltantes:
            # ── Aviso específico si falta Productos.xlsx ──────────────
            if "Productos.xlsx" in faltantes:
                messagebox.showwarning(
                    "Archivo Productos.xlsx no encontrado",
                    "❌  No se encontró el archivo  Productos.xlsx\n"
                    f"    en la carpeta RECURSOS.\n\n"
                    "📥  Descárguelo desde SAP siguiendo estos pasos:\n\n"
                    "  1.  En SAP, abra la carpeta  CREAR ARTÍCULOS\n"
                    "      (Aplicaciones → Favoritos)\n\n"
                    "  2.  Entre en  Gestionar datos maestros de producto\n\n"
                    "  3.  Haga clic en el icono  Exportar  (parte superior)\n"
                    "      y elija  Exportar como…\n\n"
                    "  4.  En el diálogo:\n"
                    "        • Nombre de fichero  →  Productos\n"
                    "        • Formato            →  Microsoft Excel (*.xlsx)\n"
                    "        • ☑  Dividir celdas con varios valores\n\n"
                    "  5.  Pulse  Exportar  y guarde el archivo descargado\n"
                    f"      dentro de la carpeta:\n"
                    f"      📁  {recursos_path}\n\n"
                    "      A continuación vuelva a ejecutar este proceso.")
                return
            # ── Resto de archivos faltantes (aviso genérico) ─────────
            otros = [f for f in faltantes if f != "Productos.xlsx"]
            if otros:
                messagebox.showwarning(
                    "Archivos no encontrados",
                    "❌  No se encontraron los siguientes archivos\n"
                    f"    en la carpeta RECURSOS:\n\n"
                    + "".join(f"   • {f}\n" for f in otros) +
                    f"\n📁  Ruta buscada:\n   {recursos_path}")
            return

        # ── 2. Obtener fechas de los archivos de origen ───────────────
        # Productos.xlsx → última modificación del archivo
        try:
            ts_prod  = os.path.getmtime(ruta_productos)
            dt_prod  = datetime.fromtimestamp(ts_prod)
            fecha_productos      = dt_prod.strftime("%d/%m/%Y %H:%M")
            fecha_productos_solo = dt_prod.strftime("%d/%m/%Y")   # para D4
            hora_productos_solo  = dt_prod.strftime("%H:%M")      # para D5
        except Exception:
            fecha_productos      = "(no disponible)"
            fecha_productos_solo = "(no disponible)"
            hora_productos_solo  = "(no disponible)"

        # ARTICULOS_DALI.xlsx → texto de celdas B3 y B5
        try:
            import openpyxl
            wb_tmp = openpyxl.load_workbook(ruta_dali, read_only=True, data_only=True)
            ws_tmp = wb_tmp.active
            celda_b3 = ws_tmp["B3"].value or ""
            wb_tmp.close()
            fecha_dali = str(celda_b3).strip() if celda_b3 else "(fecha no disponible)"
        except Exception as e:
            messagebox.showerror("Error al leer ARTICULOS_DALI",
                                 f"No se pudo leer la celda B3:\n{e}")
            return

        # ── 3. Confirmación al usuario ────────────────────────────────
        if not messagebox.askyesno(
                "Confirmar actualización SAP",
                "Se han localizado los archivos de origen:\n\n"
                f"   📄 Productos.xlsx\n"
                f"      Última modificación: {fecha_productos}\n\n"
                f"   📄 ARTICULOS_DALI.xlsx\n"
                f"      {fecha_dali}\n\n"
                f"   📄 Listas de productos.xlsx\n\n"
                "Destino:\n"
                f"   RECURSOS/catalogo-dali-completo.xlsx\n\n"
                "¿Desea continuar?"):
            self.log_consola("Actualización SAP cancelada por el usuario.", "WARN")
            return

        # ── 4. Verificar archivo destino ──────────────────────────────
        if not os.path.isfile(ruta_destino):
            messagebox.showerror(
                "Destino no encontrado",
                "❌  No se encontró el archivo destino:\n\n"
                f"   {ruta_destino}\n\n"
                "Ejecute primero la Acción 1 (artículos DALI) y pulse «Actualizar\n"
                "Catálogo Asignaciones» al final: así se crea el catálogo\n"
                "RECURSOS/catalogo-dali-completo.xlsx.")
            return

        # ── 5. Diccionarios de referencia (equivalentes al VBA) ───────
        # TABLA_SUBFAMILIA_GRUPO_COMPRAS / TABLA_IGIC_GRUPO_COMPRAS
        # (movidas a tarifa_constants.py — eran datos estáticos sin lógica,
        #  mismo patrón que DALI_CLAVES_SUBFAMILIA/DALI_CLAVES_IGIC)

        # ── 6. Proceso en hilo ────────────────────────────────────────
        def proceso():
            self.consola.separador("INICIO ACTUALIZACIÓN SAP")
            self.log_consola("Iniciando actualización listado materiales SAP...", "INFO")
            self.after(0, lambda: self._barra.indeterminate("Leyendo Productos.xlsx..."))

            try:
                # ── LECTURA 1: Productos.xlsx ─────────────────────────────────
                #    Extraido a _leer_productos_sap() -- ver ese metodo para el
                #    mapeo exacto de columnas (identico al que habia aqui antes).
                tabla_productos, filas_prod = self._leer_productos_sap(ruta_productos)
                if tabla_productos is None:
                    return

                # ── LECTURA 2: ARTICULOS_DALI.xlsx ───────────────────────────
                #    Extraido a _leer_articulos_dali_sap().
                tabla_dali = self._leer_articulos_dali_sap(ruta_dali)

                # ── LECTURA 3: Listas de productos.xlsx ──────────────────────
                #    Extraido a _leer_categoria_valoracion().
                tabla_categoria, conflictos_categoria = \
                    self._leer_categoria_valoracion(ruta_listas, tabla_productos)

                # ── CRUCE Y ESCRITURA EN DESTINO ─────────────────────────────
                #    Extraido a _cruzar_y_escribir_listado_sap() -- ver ese
                #    metodo para el detalle (identico al que habia aqui antes).
                self._cruzar_y_escribir_listado_sap(
                    ruta_destino, tabla_productos, tabla_dali,
                    tabla_categoria, conflictos_categoria, filas_prod,
                    fecha_productos_solo, hora_productos_solo)

            except Exception as e:
                self.log_consola(f"ERROR: {e}", "ERROR")
                self.consola.separador("FIN ACTUALIZACIÓN SAP (con error)")
                self.after(0, lambda: self._barra.reset("Error en actualización SAP"))
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error en actualización SAP",
                    "✘  El proceso no se completó correctamente.\n\n"
                    f"Detalle:\n{err}\n\n"
                    "Revise la ventana de log para más información."))

        threading.Thread(target=proceso, daemon=True).start()
