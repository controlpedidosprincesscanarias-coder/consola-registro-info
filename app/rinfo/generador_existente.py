# -*- coding: utf-8 -*-
"""
generador_existente.py
GeneradorDesdeExistente: genera el RINFO a partir de un «Purchasing Info Records» existente
(el orquestador _ejecutar() y sus pasos 1–12; ver HISTORIAL_REORGANIZACION_RINFO.md).

Separado de crear_registro_info.py en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""

import os
import re
from dataclasses import dataclass, field
import shutil
from tkinter import messagebox
import datetime as _dt_mod          # módulo completo — usado en lógica de fechas de baja
from datetime import datetime
from copy import copy
import openpyxl
# normalizar_codigo_sap vive en sap_helpers.py (sin tkinter) para que
# también pueda importarla test_crear_registro_info.py sin arrastrar
# la UI. Ver ese fichero para la implementación y el porqué.
from app.comun.sap_helpers import (  # noqa: E402
    normalizar_codigo_sap, ultima_fila_con_datos, parsear_fecha_celda,
    formatear_fecha_dmy, clave_pir_h3, material_de_clave_pir, formatear_clave_pir,
    debe_cerrar_fin_validez)
from app.rinfo.comun_rinfo import (  # noqa: F401
    ESTRUCTURA_HOJAS,
    FECHA_FIN_FIJA_HOJA1,
    FECHA_INICIO_FIJA_HOJA1,
    HOJA_TARIFA,
    _CENTROS_LIST,
    _LogAgrupado,
    _eliminar_filas_agrupadas,
    _localizar_ruta_listado_dali,
    _localizar_ruta_productos,
    _log,
    _parchear_column_dimension,
    _postprocesar_para_sap,
    _restaurar_column_dimension,
    _safe_after,
    aplicar_formato_fila,
)
from app.rinfo.ventanas_rinfo import (  # noqa: F401
    VentanaConfirmacionGenerado,
    VentanaMaterialesParaReactivar,
    _ventana_sin_tarifa,
)


@dataclass
class _EstadoIndicesH3:
    """
    Estado compartido entre _construir_indices_h3_heredados() (PRIORIDAD 2)
    y _aplicar_actualizaciones_h3_heredados() (PRIORIDAD 1) del paso 6c.

    Agrupar aquí las ~10 estructuras que antes eran variables sueltas
    (5 índices por clave PIR, 5 contadores, 1 lista, 1 dict de retorno y
    4 dicts de warnings) es lo que permite trocear el bucle de 619 líneas
    de forma segura: un solo objeto va y viene entre las dos fases, en vez
    de ~15 parámetros sueltos -- que es justo el tipo de contrato donde ya
    se coló una fuga real antes (ver HISTORIAL, filas_prod en el punto 2
    del roadmap). Con un solo objeto, olvidarse de un campo es un
    AttributeError inmediato en pruebas, no un silencio en producción.
    """
    periodos_h3: dict = field(default_factory=dict)
    precio_heredado_h3: dict = field(default_factory=dict)
    i_heredada_h3: dict = field(default_factory=dict)
    j_heredada_h3: dict = field(default_factory=dict)
    fila_mas_reciente_h3: dict = field(default_factory=dict)
    absorbidas_con_tarifa: dict = field(default_factory=dict)
    filas_a_eliminar_6c: list = field(default_factory=list)
    n_solo_fecha: int = 0
    n_precio_y_fecha: int = 0
    n_solapamiento: int = 0
    n_reduccion: int = 0
    n_sin_fecha_heredada: int = 0
    warn_solap: dict = field(default_factory=dict)
    warn_gap: dict = field(default_factory=dict)
    warn_reduccion: dict = field(default_factory=dict)
    warn_sin_parsear: dict = field(default_factory=dict)

class GeneradorDesdeExistente:
    """
    Genera un nuevo archivo XLSX basado en un Purchasing Info Records existente,
    añadiendo al final solo los materiales de la tarifa que no están en el
    Purchasing (cruce col A tarifa vs col E Purchasing).

    Los materiales nuevos reciben exactamente el mismo relleno de columnas
    que en la creación desde cero (CreadorRINFO).
    """

    HOJA_PRINCIPAL = "Info general p.reg.info"
    HOJA_2         = "Datos org.compr.p.reg.info"
    HOJA_3         = "Det.precios p.registros info"

    # ── Columnas de valor fijo obligatorio SAP en Hoja 2 ──────────────────────
    #   ÚNICO punto de la clase donde se definen estos literales. Cualquier
    #   cambio de valor por defecto se hace aquí y en ningún otro sitio;
    #   _escribir_fijas_h2_fila() y _normalizar_obligatorios_fila_h2() son los
    #   dos únicos métodos autorizados a escribir estas columnas.
    #
    #   Se aplican tanto a filas realmente nuevas (forzadas, ver
    #   _escribir_fijas_h2_fila) como al saneado de filas heredadas que
    #   llegaron incompletas desde el Purchasing original (solo si vienen
    #   vacías, ver _normalizar_obligatorios_fila_h2).
    COLUMNAS_FIJAS_H2 = {
        2:  "0",              # B  Tipo registro
        3:  "Normal",         # C  Desc. categoría
        4:  "IC00",           # D  Org. compras
        13: 1.0,               # M  Cant. pedido estándar
        14: 1,                  # N  Plazo entrega días
        19: "X",                 # S  Verif. factura EM
        23: 0.0,                  # W  Cant. pedido máxima
        24: 0.0,                   # X  Cant. mínima pedido
        30: 0.0,                    # AD Tolerancia exceso %
        31: 0.0,                     # AE Tolerancia faltas %
        32: "1",                      # AF Control fecha precio
        36: "01.01.1900",              # AJ Inic.val.tipo impos.
    }
    # AF (32) no aplica a filas marcadas para borrado.
    COLUMNAS_FIJAS_H2_EXCLUIR_BORRADO = {32}

    # Columnas que SOLO se fuerzan al crear una fila realmente nueva.
    # Nunca deben tocarse al normalizar una heredada: forzar AO en una
    # fila que ya existía la marcaría incorrectamente como "Nueva" ante SAP.
    #
    # AUDITORÍA 2026-08 (manual SAP "Crear Registros Info", punto 3.1.2):
    # se eliminó la escritura de AK (37, "Entrega ilimitada permitida") en
    # Hoja2. El manual detalla explícitamente todas las columnas fijas de
    # la segunda pestaña (A, B, C, D/E, F/G, H/I, J/K, L, M, N, S, W/X,
    # AD/AE, AF, AH, AJ, AO) y AK no figura entre ellas. Se retira por no
    # estar respaldada por el manual; si en el futuro se confirma con SAP/
    # central que es necesaria, reintroducir aquí con la referencia exacta
    # del punto del manual que la ampare.
    COLUMNAS_FIJAS_H2_SOLO_NUEVA = {
        41: "X",   # AO Nueva
    }

    def _escribir_fijas_h2_fila(self, ws_h2, fila_num: int) -> None:
        """
        Escribe TODAS las columnas de valor fijo obligatorio SAP (incluidas
        AK/AO) en una fila NUEVA de Hoja2. Único punto de escritura para
        estos literales al crear una fila — usado tanto por el bloque de
        materiales nuevos como por el de absorbidos recuperados.
        """
        for col, valor in self.COLUMNAS_FIJAS_H2.items():
            ws_h2.cell(row=fila_num, column=col, value=valor)
        for col, valor in self.COLUMNAS_FIJAS_H2_SOLO_NUEVA.items():
            ws_h2.cell(row=fila_num, column=col, value=valor)

    def _normalizar_obligatorios_fila_h2(self, ws_h2, fila_num: int,
                                          es_borrado: bool) -> list:
        """
        Rellena en una fila HEREDADA de Hoja2 cualquier columna de valor fijo
        obligatorio que haya quedado vacía (heredada del Purchasing original),
        sin pisar valores ya existentes. AF se omite en filas de borrado.
        AK/AO nunca se tocan aquí — ver docstring de COLUMNAS_FIJAS_H2_SOLO_NUEVA.

        Devuelve la lista de columnas (1-based) que se han rellenado.
        """
        normalizadas = []
        for col, valor in self.COLUMNAS_FIJAS_H2.items():
            if es_borrado and col in self.COLUMNAS_FIJAS_H2_EXCLUIR_BORRADO:
                continue
            celda = ws_h2.cell(row=fila_num, column=col)
            if celda.value is None or str(celda.value).strip() == "":
                celda.value = valor
                normalizadas.append(col)
        return normalizadas

    def __init__(self, parent_ventana, ruta_tarifa: str, nombre_archivo: str,
                 nombre_sap: str, codigo_sap: str,
                 purchasing: dict, carpeta_tarifas: str, app_root=None,
                 telefono_proveedor=None, fecha_inicio=None, fecha_fin=None,
                 centros_activos=None):
        self._parent              = parent_ventana
        self._ruta_tarifa         = ruta_tarifa
        self._nombre_archivo      = nombre_archivo
        self._nombre_sap          = nombre_sap
        self._codigo_sap          = codigo_sap
        self._purchasing          = purchasing
        self._carpeta_tarifas     = carpeta_tarifas
        self._app_root            = app_root
        self._telefono_proveedor  = telefono_proveedor
        self._fecha_inicio        = fecha_inicio or "01.01.2025"
        self._fecha_fin           = fecha_fin    or "31.12.2030"
        self._centros_activos     = centros_activos or list(_CENTROS_LIST)

        import threading
        threading.Thread(target=self._ejecutar, daemon=True).start()

    # ── Barra de progreso ─────────────────────────────────────────────────────
    def _barra(self, modo: str, texto: str = "", v: int = 0, total: int = 100):
        app = self._app_root
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

    # ── UTILIDADES COMPARTIDAS ────────────────────────────────────────────────

    def _cargar_listado_dali(self):
        """
        Localiza y carga el LISTADO CODIGOS DALI - SAP.xlsx.
        Devuelve un dict {clave_SAP: (desc_sap, unidad, grp_prod, desc_grp,
                                      ind_imp, grp_comp, desc_grp_comp)}
        o {} si no se encuentra o hay error.

        Mapeo de columnas (actualizado 2026-08: tras insertar las 3
        columnas nuevas F, K, L en LISTADO CODIGOS DALI-SAP, lo que había
        se desplazó: antigua F/G/H/I -> G/H/I/J, antigua J/K -> M/N):
            E (desc SAP), G (unidad), H (grp productos), I (desc grp),
            J (indicador impuestos / IGIC), M (grupo compras),
            N (desc grupo compras)

        FALLBACK A Productos.xlsx (añadido 2026-09-18, caso real registro
        5300002798): el LISTADO se construye cruzando cada código DALI con
        su código SAP -- si un material SAP nunca se vinculó a un código
        DALI (típicamente: material descatalogado y sustituido por otro,
        o duplicado que quedó huérfano), el LISTADO no tiene ninguna fila
        para él, y toda sincronización de UM que dependa de este
        diccionario (ver _sincronizar_um_hoja1/_sincronizar_um_hoja3) se
        salta silenciosamente por "sin match" -- dejando una UM obsoleta
        que SAP rechaza igualmente al procesar el registro, incluso si es
        una baja ("Por favor, adecúe la UM-base y el factor de
        conversión"). Tras cargar el LISTADO, se completa con
        Productos.xlsx (maestro de materiales SAP/MM03, indexado por
        código SAP, sin necesitar ningún cruce DALI) para cualquier
        código SAP que no haya aparecido en el LISTADO. Solo se puede
        rellenar así desc_sap/unidad/grp_prod/desc_grp: ind_imp (IGIC) y
        grupo de compras dependen de la SUBFAMILIA del lado DALI (tablas
        TABLA_IGIC_GRUPO_COMPRAS/TABLA_SUBFAMILIA_GRUPO_COMPRAS en
        tarifa_app.py), que no existe para un material sin código DALI
        -- esos dos campos quedan vacíos igual que antes del fallback.
        El LISTADO manda siempre que tenga la entrada; Productos.xlsx
        solo rellena los huecos que el LISTADO deja vacíos.
        """
        ruta = _localizar_ruta_listado_dali(
            self._carpeta_tarifas, getattr(self, "_carpeta_proyecto", None), self._parent)
        if ruta is None:
            _log(self._parent,
                 "LISTADO CODIGOS DALI-SAP no encontrado — "
                 "columnas de cruce quedarán vacías.", "WARN")
            return {}

        diccionario = {}
        try:
            wb_l = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
            ws_l = wb_l.active
            for row in ws_l.iter_rows(min_row=8, values_only=True):
                if not row or len(row) < 4 or row[3] is None:
                    continue
                raw = row[3]
                if isinstance(raw, float):
                    raw = int(raw)
                clave = str(raw).strip()
                if not clave:
                    continue
                desc_sap   = str(row[4]).strip()  if len(row) > 4  and row[4]  is not None else ""
                unidad     = str(row[6]).strip()  if len(row) > 6  and row[6]  is not None else ""
                grp_prod   = str(row[7]).strip()  if len(row) > 7  and row[7]  is not None else ""
                desc_grp   = str(row[8]).strip()  if len(row) > 8  and row[8]  is not None else ""
                ind_imp    = str(row[9]).strip()  if len(row) > 9  and row[9]  is not None else ""
                grp_comp   = str(row[12]).strip() if len(row) > 12 and row[12] is not None else ""
                desc_grp_c = str(row[13]).strip() if len(row) > 13 and row[13] is not None else ""
                diccionario[clave] = (desc_sap, unidad, grp_prod, desc_grp,
                                      ind_imp, grp_comp, desc_grp_c)
            wb_l.close()
            _log(self._parent,
                 f"LISTADO DALI-SAP cargado: {len(diccionario)} entradas.", "INFO")
        except Exception as e:
            _log(self._parent, f"Error al leer LISTADO: {e}", "ERROR")
            return diccionario

        # ── Fallback a Productos.xlsx para materiales sin cruce DALI ──────
        ruta_prod = _localizar_ruta_productos(
            self._carpeta_tarifas, getattr(self, "_carpeta_proyecto", None), self._parent)
        if ruta_prod is None:
            _log(self._parent,
                 "Productos.xlsx no encontrado — materiales sin cruce DALI "
                 "no tendrán respaldo (UM podría quedar sin sincronizar).", "WARN")
            return diccionario

        n_respaldo = 0
        try:
            wb_p = openpyxl.load_workbook(ruta_prod, read_only=True, data_only=True)
            ws_p = wb_p.active
            for row in ws_p.iter_rows(min_row=2, values_only=True):
                if not row or len(row) < 2 or row[1] is None:
                    continue
                raw = row[1]
                if isinstance(raw, float):
                    raw = int(raw)
                clave = str(raw).strip()
                if not clave or clave in diccionario:
                    continue   # el LISTADO ya tiene esta clave: no se toca
                desc_sap = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
                grp_prod = str(row[5]).strip() if len(row) > 5 and row[5] is not None else ""
                unidad   = str(row[8]).strip() if len(row) > 8 and row[8] is not None else ""
                desc_grp = str(row[6]).strip() if len(row) > 6 and row[6] is not None else ""
                if not unidad:
                    continue   # Productos.xlsx tampoco tiene UM: no hay nada útil que aportar
                diccionario[clave] = (desc_sap, unidad, grp_prod, desc_grp, "", "", "")
                n_respaldo += 1
            wb_p.close()
            if n_respaldo:
                _log(self._parent,
                     f"Productos.xlsx — {n_respaldo} material(es) sin cruce DALI "
                     f"completados como respaldo (desc./UM/grupo; sin IGIC ni "
                     f"grupo de compras, que dependen del cruce DALI).", "OK")
        except Exception as e:
            _log(self._parent,
                 f"Error al leer Productos.xlsx (respaldo omitido): {e}", "WARN")

        return diccionario

    def _cargar_precios_tarifa(self):
        """
        Carga el diccionario {clave_SAP: precio} desde col A y col F de la tarifa.
        """
        dic = {}
        try:
            wb_t = openpyxl.load_workbook(
                self._ruta_tarifa, read_only=True, data_only=True)
            hoja_t = (wb_t[HOJA_TARIFA]
                      if HOJA_TARIFA in wb_t.sheetnames
                      else wb_t.active)
            for row in hoja_t.iter_rows(min_row=8, min_col=1,
                                         max_col=6, values_only=True):
                cod    = row[0]
                precio = row[5]
                if cod is None:
                    continue
                cod = normalizar_codigo_sap(cod)
                clave = str(cod).strip()
                if clave and precio is not None:
                    dic[clave] = precio
            wb_t.close()
            _log(self._parent,
                 f"Precios tarifa cargados: {len(dic)} entradas.", "INFO")
        except Exception as e:
            _log(self._parent, f"Error al leer precios tarifa: {e}", "WARN")
        return dic

    # ── RELLENO HOJA 1 (nuevas filas) ────────────────────────────────────────

    def _rellenar_h1_nuevas_filas(self, ws, primera_fila: int,
                                   nuevos: list, diccionario: dict,
                                   id_offset: int = 0) -> None:
        """
        Escribe en Hoja 1 las filas correspondientes a los materiales nuevos,
        comenzando en primera_fila.

        nuevos : lista de (cod_sap_str, cod_prov_str|None)

        Col A (Registro info de compras):
            Se asignan valores consecutivos en texto con number_format="@"
            (triangulito verde SAP), continuando desde el maximo correlativo
            numerico ya presente en las filas preexistentes del Purchasing.
            Si no hay valores numericos previos en col A, la numeracion empieza
            en 1 — identica logica que _escribir_correlativos en CreadorRINFO.
        """
        # Fila de referencia para el formato: la ultima fila existente con datos
        # (primera_fila - 1). Si primera_fila == 2, no hay filas previas y se
        # usa la fila 2 como plantilla (comportamiento original).
        fila_ref_formato = max(2, primera_fila - 1)

        # ── Punto C: IDs col A para materiales nuevos ─────────────────────────
        # SAP usa el campo «Registro info de compras» (col A) como clave interna
        # del PIR. Si los nuevos materiales reciben IDs ya presentes en filas
        # heredadas, SAP puede rechazarlos (ME807) o sobreescribir registros
        # existentes de forma silenciosa.
        #
        # id_offset determina el punto de inicio:
        #   · Heredadas con correlativos 1..N (id_offset=N) → los nuevos reciben N+1, N+2, …
        #   · Heredadas con IDs SAP reales o hoja vacía (id_offset=0) → empieza en 1
        #     SAP asigna sus propios IDs en la carga; los correlativos 1, 2, 3... son
        #     solo marcadores de posición que SAP sustituye por IDs reales al importar.
        # La numeración preexistente (col A) se deja completamente intacta.
        _id_inicio = id_offset + 1
        _id_fin    = id_offset + len(nuevos)
        _log(self._parent,
             f"  Nuevos materiales: {len(nuevos)}. "
             f"Correlativos asignados (col A): {_id_inicio}–{_id_fin} "
             f"(continuando desde max heredado={id_offset}).",
             "INFO")

        for idx, (cod_sap_str, cod_prov_str) in enumerate(nuevos, start=1):
            fila = primera_fila + idx - 1

            # Aplicar formato desde la ultima fila existente (no desde fila 2)
            # para garantizar que los estilos coinciden con las filas anteriores.
            aplicar_formato_fila(ws, fila, self.HOJA_PRINCIPAL, fila_ref_formato)

            datos = diccionario.get(cod_sap_str,
                                     ("", "", "", "", "", "", ""))
            desc_sap, unidad, grp_prod, desc_grp, ind_imp, grp_comp, desc_grp_c = datos

            # Col A: ID que CONTINÚA desde el máximo heredado.
            # number_format "@" fuerza el triangulito verde SAP (texto forzado).
            celda_a = ws.cell(row=fila, column=1)
            celda_a.value         = str(id_offset + idx)   # continua: max+1, max+2, …
            celda_a.number_format = "@"
            ws.cell(row=fila, column=3,
                    value=self._codigo_sap if self._codigo_sap else None)   # C  código SAP proveedor
            ws.cell(row=fila, column=4,
                    value=self._nombre_sap)                                 # D  nombre SAP proveedor
            ws.cell(row=fila, column=5,
                    value=cod_sap_str)                                      # E  código SAP material
            ws.cell(row=fila, column=6,
                    value=desc_sap or None)                                 # F  descripción SAP
            ws.cell(row=fila, column=7,
                    value=grp_prod or None)                                 # G  grupo artículos
            ws.cell(row=fila, column=8,
                    value=desc_grp or None)                                 # H  desc. grupo artículos
            # I (col 9)  Disponible desde  → _aplicar_fechas_hoja1 (todas las filas)
            # J (col 10) Disponible hasta  → _aplicar_fechas_hoja1 (todas las filas)
            ws.cell(row=fila, column=14,
                    value=unidad or None)                                   # N  Unidad medida pedido
            ws.cell(row=fila, column=15,
                    value=1)                                                # O  Denominador conversión
            ws.cell(row=fila, column=16,
                    value=unidad or None)                                   # P  Unidad de medida
            ws.cell(row=fila, column=17,
                    value=1)                                                # Q  Numerador conversión
            if cod_prov_str:
                ws.cell(row=fila, column=28,
                        value=cod_prov_str)                                 # AB  Nº mat. proveedor
            # AD (col 30) Teléfono         → _aplicar_telefono_hoja1 (todas las filas)
            ws.cell(row=fila, column=37,
                    value="X")                                              # AK  Nueva

    # ── SINCRONIZACIÓN UM EN TODA HOJA 1 (cols N y P) ──────────────────────────

    def _sincronizar_um_hoja1(self, ws, diccionario: dict) -> None:
        """
        Sincroniza la Unidad de medida (col N «Unidad medida pedido» y
        col P «Unidad de medida») en TODA la Hoja 1, incluidas las filas
        HEREDADAS del Purchasing existente.

        Hasta ahora, N/P solo se rellenaban al escribir materiales realmente
        NUEVOS (ver _rellenar_h1_nuevas_filas). Las filas heredadas se dejaban
        con la UM que traía el Purchasing original, aunque el LISTADO CODIGOS
        DALI-SAP (que a su vez se alimenta de Productos.xlsx / MM03) hubiera
        cambiado de unidad para ese material — dejando una UM obsoleta o
        incoherente con la ficha de material en SAP (causa típica del error
        de UM base al cargar el registro de info).

        Ahora se recorren TODAS las filas con material, se compara N y P
        frente a la unidad del LISTADO y, si difiere, se corrige.

        IMPORTANTE — a diferencia de AO (Hoja2) y Z (Hoja3), la columna
        AK (37, «Nueva») de Hoja 1 NO es un simple indicador informativo:
        SAP la usa para decidir si debe CREAR un registro de info nuevo.
        En una fila heredada, la col A ya contiene el número real de
        registro de info existente en SAP (p.ej. 5300001809); si AK se
        marca "X" en esa fila, SAP intenta CREAR un registro con un ID
        que ya existe → error "Ya existe el registro Info de compras
        NNNNNNNNNN" y la fila queda sin procesar (la corrección de UM no
        llega a aplicarse). Por eso esta sincronización NUNCA toca AK:
        igual que _aplicar_telefono_hoja1 y _aplicar_fechas_hoja1, se
        limita a corregir el valor en las filas heredadas, que SAP ya
        recoge al modificar el registro existente (proceso "modificar",
        determinado por col A, no por AK).

        Las filas ya marcadas para borrado (AJ=36="X") se excluyen
        (revertido 2026-09-22, ver HISTORIAL): el 2026-09-18 se corrigió
        esta exclusión asumiendo que SAP no validaría la UM de un
        registro que se va a borrar -- un caso real (5300002798)
        demostró que sí la valida, así que se quitó la exclusión. Pero
        eso solo "funcionó" en ese caso concreto porque el usuario borró
        el registro en SAP a mano antes de recargar: cambiar la UM de un
        registro EXISTENTE (no ese caso puntual) sigue dando el mismo
        error, ya que SAP compara contra lo que ya tiene almacenado y no
        reconcilia un cambio de unidad vía carga masiva. La solución de
        fondo, más segura: para materiales de baja, NO tocar ningún
        valor heredado -- solo marcar para borrado donde corresponda
        (AJ/Y), dejando la UM tal cual estaba. Así se evita el error sin
        depender de borrar el registro en SAP cada vez.
        """
        COL_E_MATERIAL     = 5    # E  Material (código SAP)
        COL_N_UM_PEDIDO    = 14   # N  Unidad medida pedido
        COL_P_UM           = 16   # P  Unidad de medida
        COL_AJ_H1_BORRADO  = 36   # AJ Registro info marcado para borrado

        n_um_actualizados = 0

        for fila_row in ws.iter_rows(min_row=2):
            if len(fila_row) < COL_P_UM:
                continue

            val_mat = fila_row[COL_E_MATERIAL - 1].value \
                if len(fila_row) >= COL_E_MATERIAL else None
            if val_mat is None:
                continue

            fila_num = fila_row[0].row

            # No tocar filas ya marcadas para borrado -- ningún valor
            # heredado se modifica en materiales de baja, solo se marcan.
            val_aj = fila_row[COL_AJ_H1_BORRADO - 1].value \
                if len(fila_row) >= COL_AJ_H1_BORRADO else None
            if val_aj is not None and str(val_aj).strip().upper() == "X":
                continue

            raw_m = val_mat
            raw_m = normalizar_codigo_sap(raw_m)
            clave_mat = str(raw_m).strip()

            datos_listado = diccionario.get(clave_mat)
            if not datos_listado:
                continue   # material sin match en LISTADO: no tocar lo existente

            unidad_correcta = (datos_listado[1] or "").strip() \
                if len(datos_listado) > 1 else ""
            if not unidad_correcta:
                continue   # LISTADO sin UM informada: no borrar lo que ya había

            val_n_actual = fila_row[COL_N_UM_PEDIDO - 1].value \
                if len(fila_row) >= COL_N_UM_PEDIDO else None
            val_p_actual = fila_row[COL_P_UM - 1].value \
                if len(fila_row) >= COL_P_UM else None
            val_n_str = str(val_n_actual).strip() if val_n_actual is not None else ""
            val_p_str = str(val_p_actual).strip() if val_p_actual is not None else ""

            if val_n_str == unidad_correcta and val_p_str == unidad_correcta:
                continue   # sin cambio — nada que sincronizar

            if val_n_str != unidad_correcta:
                ws.cell(row=fila_num, column=COL_N_UM_PEDIDO, value=unidad_correcta)
            if val_p_str != unidad_correcta:
                ws.cell(row=fila_num, column=COL_P_UM, value=unidad_correcta)
            n_um_actualizados += 1

            _log(self._parent,
                 f"  H1 fila {fila_num} [material {clave_mat}]: UM "
                 f"N={val_n_str!r}/P={val_p_str!r} → {unidad_correcta!r} actualizado "
                 f"(sin tocar AK — SAP la recoge al modificar el registro existente).",
                 "INFO")

        if n_um_actualizados:
            _log(self._parent,
                 f"Hoja1 — UM sincronizada (N/P): {n_um_actualizados} fila(s) "
                 f"actualizada(s).", "OK")
        else:
            _log(self._parent,
                 "Hoja1 — UM (N/P): sin discrepancias frente al LISTADO.", "INFO")

        if hasattr(self, "_audit"):
            self._audit["n_um_h1_actualizados"] = n_um_actualizados

    # ── RELLENO HOJA 2 (nuevas filas + borrado lógico centros descartados) ──────

    def _marcar_borrado_logico_y_escanear_h2(self, ws_h2, centros_activos_set):
        """
        PASO 1 (extraído de _rellenar_h2_nuevas_filas): recorre las filas
        heredadas de Hoja2, elimina físicamente las de centro descartado/
        vacío (evita el error SAP «Introduzca Plaz.entr.prev.» en filas
        inválidas), sanea la col M (Cant. pedido estándar) si llega vacía
        o a 0, y construye los sets de deduplicación que necesitan los
        pasos siguientes.

        Devuelve (combinaciones_existentes, _escritas_h2, _claves_borrado_h2,
        ultima_fila_h2).
        """
        # ── PASO 1: Escanear filas heredadas del Purchasing ───────────────────
        #   a) Marcar borrado lógico en filas con centro descartado.
        #   b) Construir el set de combinaciones PIR ya presentes para no
        #      duplicarlas al añadir las nuevas.
        #
        #  Deduplicación funcional H2 — clave completa PIR:
        #    proveedor|org|material|centro
        #  Esto evita duplicados aunque provengan de distintas hojas, iteraciones
        #  o materiales repetidos en el input (Punto B del spec de calidad).
        _prov_str = str(self._codigo_sap).strip() if self._codigo_sap else ""
        _ORG_H2   = "IC00"
        combinaciones_existentes = set()   # {(clave_mat_str, cod_centro_str), ...}  ← compat. heredadas
        _escritas_h2: set = set()          # clave funcional "prov|org|mat|centro" — guard universal
        ultima_fila_h2 = 1
        # Filas heredadas marcadas para borrado en Hoja2, para propagar la
        # misma marca a Hoja3 (mismo material + mismo centro, incluido vacío).
        _claves_borrado_h2: set = set()    # {(clave_mat_str, cod_centro_str), ...}

        # ── Primera pasada: identificar filas a eliminar y filas válidas ─────
        #   Las filas con centro descartado/vacío se eliminan FÍSICAMENTE en
        #   lugar de marcarse para borrado. Esto evita que SAP exija campos
        #   obligatorios (Plaz.entr.prev.) en filas inválidas que no aportan
        #   nada al RINFO y generan errores en la carga.
        _filas_eliminar_h2: list = []   # números de fila a eliminar físicamente
        _agr_centro_descartado_h2 = _LogAgrupado(self._parent)

        for fila_row in ws_h2.iter_rows(min_row=2):
            val_a = fila_row[0].value if len(fila_row) > 0 else None
            val_f = fila_row[5].value if len(fila_row) > 5 else None   # col F – Centro
            val_j = fila_row[9].value if len(fila_row) > 9 else None   # col J – Material
            if val_a is None and val_j is None:
                continue

            fila_num = fila_row[0].row
            ultima_fila_h2 = fila_num

            # c) Saneado col M (13) — «Cant. pedido estándar» heredada del
            #    Purchasing SAP nunca puede ser 0 o estar vacía: si llega así
            #    se fuerza a 1 (regla de negocio obligatoria).
            val_m_heredado = fila_row[12].value if len(fila_row) > 12 else None
            try:
                _m_num = float(val_m_heredado) if val_m_heredado not in (None, "") else 0.0
            except (TypeError, ValueError):
                _m_num = 0.0
            if _m_num <= 0:
                ws_h2.cell(row=fila_num, column=13).value = 1.0
                _log(self._parent,
                     f"  H2 fila {fila_num}: col M heredada vacía/0 "
                     f"({val_m_heredado!r}) → forzada a 1.", "WARN")

            # Normalizar centro
            cod_centro_heredado = str(val_f).strip() if val_f is not None else ""

            # Normalizar material
            if val_j is not None:
                raw_m = val_j
                raw_m = normalizar_codigo_sap(raw_m)
                clave_mat_heredada = str(raw_m).strip()
            else:
                clave_mat_heredada = ""

            # a) Centro descartado/vacío → programar eliminación física.
            #    No se marca para borrado (col AN) porque SAP igualmente
            #    exigiría campos obligatorios y generaría error «Introduzca
            #    Plaz.entr.prev.». Al eliminar la fila desaparece el problema
            #    de raíz en todas las actualizaciones futuras.
            if clave_mat_heredada and cod_centro_heredado not in centros_activos_set:
                _filas_eliminar_h2.append(fila_num)
                _claves_borrado_h2.add((clave_mat_heredada, cod_centro_heredado))
                _agr_centro_descartado_h2.registrar(
                    clave=(clave_mat_heredada,), fila=fila_num,
                    centro=cod_centro_heredado if cod_centro_heredado else "(vacío)",
                    nivel="INFO",
                    plantilla=("  H2 material {mat}: centro(s) descartado(s) → "
                               "programada eliminación física — {resumen}."),
                    mat=clave_mat_heredada)
                continue   # no registrar en combinaciones_existentes

            # b) Registrar combinación existente (centros válidos)
            if clave_mat_heredada and cod_centro_heredado:
                combinaciones_existentes.add((clave_mat_heredada, cod_centro_heredado))
                _clave_h2_her = f"{_prov_str}|{_ORG_H2}|{clave_mat_heredada}|{cod_centro_heredado}"
                _escritas_h2.add(_clave_h2_her)

        # SIEMPRE al terminar el bucle exterior, para no perder el último grupo.
        _agr_centro_descartado_h2.flush()

        # ── Eliminar físicamente las filas descartadas de Hoja 2 ──────────────
        #   Se eliminan en orden inverso para no desplazar los índices de fila.
        if _filas_eliminar_h2:
            _eliminar_filas_agrupadas(ws_h2, _filas_eliminar_h2)
            _log(self._parent,
                 f"Hoja2 — {len(_filas_eliminar_h2)} fila(s) con centro "
                 f"descartado/vacío eliminada(s) físicamente (evita error SAP "
                 f"'Introduzca Plaz.entr.prev.').", "OK")
            ultima_fila_h2 = ws_h2.max_row

        return combinaciones_existentes, _escritas_h2, _claves_borrado_h2, ultima_fila_h2


    def _propagar_borrado_h2_a_h3(self, wb, _claves_borrado_h2):
        """
        PASO 1b (extraído de _rellenar_h2_nuevas_filas): elimina
        físicamente en Hoja3 las filas cuyo material+centro coincide con
        las eliminadas en Hoja2 (paso 1), manteniendo coherencia entre
        ambas hojas. No devuelve nada -- efecto directo sobre `wb`.
        """
        # ── PASO 1b: Eliminar físicamente filas correspondientes en Hoja 3 ────
        #   Las filas de Hoja3 cuyo material+centro coinciden con los eliminados
        #   en Hoja2 también se eliminan, manteniendo coherencia entre hojas.
        if _claves_borrado_h2 and self.HOJA_3 in wb.sheetnames:
            ws_h3_propagacion = wb[self.HOJA_3]
            _filas_eliminar_h3: list = []
            _agr_propagacion_h3 = _LogAgrupado(self._parent)
            for fila_row_h3 in ws_h3_propagacion.iter_rows(min_row=2):
                val_centro_h3 = fila_row_h3[4].value if len(fila_row_h3) > 4 else None
                val_mat_h3    = fila_row_h3[5].value if len(fila_row_h3) > 5 else None
                if val_mat_h3 is None:
                    continue
                raw_m_h3 = val_mat_h3
                raw_m_h3 = normalizar_codigo_sap(raw_m_h3)
                clave_mat_h3  = str(raw_m_h3).strip()
                cod_centro_h3 = str(val_centro_h3).strip() if val_centro_h3 is not None else ""
                if (clave_mat_h3, cod_centro_h3) in _claves_borrado_h2:
                    _filas_eliminar_h3.append(fila_row_h3[0].row)
                    # Agrupado (una línea por material, no una por fila/centro):
                    # con miles de filas dispersas, llamar a _log() una vez por
                    # fila fuerza una repintada de la consola por cada línea y
                    # eso —no el delete_rows, ya corregido— es lo que dominaba
                    # el tiempo total en archivos grandes.
                    _agr_propagacion_h3.registrar(
                        clave=(clave_mat_h3,), fila=fila_row_h3[0].row,
                        centro=cod_centro_h3 if cod_centro_h3 else "(vacío)",
                        nivel="INFO",
                        plantilla=("  H3 material {mat}: → programada eliminación "
                                   "física (coherencia con Hoja2) — {resumen}."),
                        mat=clave_mat_h3)
            _agr_propagacion_h3.flush()
            if _filas_eliminar_h3:
                _eliminar_filas_agrupadas(ws_h3_propagacion, _filas_eliminar_h3)
                _log(self._parent,
                     f"Hoja3 — {len(_filas_eliminar_h3)} fila(s) eliminada(s) "
                     f"físicamente (coherencia con Hoja2).", "OK")


    def _leer_h1_para_expansion_h2(self, ws_h1):
        """
        PASO 2 (extraído de _rellenar_h2_nuevas_filas): lee TODOS los
        registros info y materiales de Hoja1 (heredados + nuevos, sin
        límite de fila artificial), excluyendo los marcados para borrado
        (AJ=X) -- no tiene sentido regenerar centros para un material que
        se va a dar de baja en esta misma carga.

        Devuelve filas_nuevas_h1: lista de (reg_info, material).
        """
        # ── PASO 2: Leer TODOS los registros info y materiales de Hoja 1 ──────
        #   (heredados + nuevos), sin límite de fila artificial.
        #
        #   No se puede usar «primera_fila_nueva_h1 + len(nuevos_h1) - 1» como
        #   tope porque cuando nuevos_h1 está vacío ese cálculo da exactamente la
        #   última fila heredada solo por casualidad numérica, y si hubiera filas
        #   en blanco intermedias o la numeración difiriera, se perderían materiales.
        #
        #   El guard de PASO 3 (_escritas_h2) filtra lo que ya existe correctamente
        #   en H2: materiales con todos sus centros cubiertos no generan filas nuevas;
        #   materiales con centros parciales o ausentes generan solo los que faltan.
        #   Esto resuelve el caso de Purchasing con H2 incompleta (filas con
        #   Centro="" o materiales sin ninguna fila en H2).
        #
        #   Se excluyen materiales con AJ=X en H1: no tiene sentido regenerar
        #   centros para un material que se va a dar de baja en esta misma carga.
        filas_nuevas_h1 = []
        for fila_row in ws_h1.iter_rows(min_row=2):
            reg_info = fila_row[0].value   # A
            material = fila_row[4].value   # E
            if material is None:
                continue
            # Excluir materiales marcados para borrado en H1 (col AJ = col 36)
            col_aj = fila_row[35].value if len(fila_row) > 35 else None
            if col_aj is not None and str(col_aj).strip().upper() == "X":
                continue
            filas_nuevas_h1.append((reg_info, material))

        return filas_nuevas_h1


    def _escribir_h2_bloques_nuevos(self, ws_h2, filas_nuevas_h1, ultima_fila_h2,
                                    _escritas_h2, combinaciones_existentes,
                                    diccionario, codigos_tarifa_set):
        """
        PASO 3 (extraído de _rellenar_h2_nuevas_filas): escribe las
        filas realmente nuevas de Hoja2, un bloque por cada centro
        activo, solo para materiales de la tarifa activa
        (codigos_tarifa_set) -- un material heredado MANTENIDO (sin AJ,
        pero ausente de la tarifa) no se expande a los demás centros,
        se conserva únicamente donde ya existía. Aplica deduplicación
        funcional por clave proveedor|org|material|centro.

        Devuelve filas_h2_escritas: lista de tuplas con los datos
        escritos, en el mismo formato que usa Hoja3 después (es el
        valor de retorno final de _rellenar_h2_nuevas_filas).
        """
        _prov_str = str(self._codigo_sap).strip() if self._codigo_sap else ""
        _ORG_H2   = "IC00"

        # ── PASO 3: Escribir filas nuevas (un bloque por cada centro activo) ───
        centros     = self._centros_activos
        NUM_BLOQUES = len(centros)
        n_nuevos    = len(filas_nuevas_h1)

        filas_h2_escritas  = []
        siguiente_fila_h2  = ultima_fila_h2 + 1
        idx_escritura      = 0          # desplazamiento relativo desde siguiente_fila_h2
        saltadas_duplicado = 0

        saltadas_sin_tarifa = 0
        for bloque, (cod_centro, nom_centro) in enumerate(centros):
            for reg_info, material in filas_nuevas_h1:

                # Normalizar clave de material para la comprobación
                if material is not None:
                    raw_m = material
                    raw_m = normalizar_codigo_sap(raw_m)
                    clave_mat = str(raw_m).strip()
                else:
                    clave_mat = ""

                # ── Filtro «solo materiales de tarifa se expanden» ───────────
                # Un material heredado MANTENIDO (sin AJ, pero ausente de la
                # tarifa activa) no debe recibir filas nuevas en centros donde
                # no existía: debe permanecer solo en su(s) centro(s) original(es)
                # (esos ya están cubiertos por combinaciones_existentes/PASO 1,
                # no por este bloque). Si no se pasó codigos_tarifa_set se
                # mantiene el comportamiento previo (expandir todo) por
                # compatibilidad con otros llamadores.
                if (codigos_tarifa_set is not None
                        and clave_mat not in codigos_tarifa_set):
                    saltadas_sin_tarifa += 1
                    continue

                # ── Deduplicación funcional (Punto B) ────────────────────────
                # Clave completa: proveedor|org|material|centro
                # Blinda contra duplicados de distintas hojas, iteraciones o
                # materiales repetidos en el input, no solo contra heredadas SAP.
                _clave_h2_nueva = f"{_prov_str}|{_ORG_H2}|{clave_mat}|{cod_centro}"
                if _clave_h2_nueva in _escritas_h2:
                    saltadas_duplicado += 1
                    continue
                _escritas_h2.add(_clave_h2_nueva)
                # Compatibilidad con check heredadas (por si se usa en otro punto)
                combinaciones_existentes.add((clave_mat, cod_centro))

                fila_h2 = siguiente_fila_h2 + idx_escritura
                idx_escritura += 1

                datos = diccionario.get(clave_mat, ("", "", "", "", "", "", ""))
                desc_sap, unidad, grp_prod, desc_grp, ind_imp, grp_comp, desc_grp_c = datos

                # Columnas específicas de esta fila (dependen del material/centro).
                ws_h2.cell(row=fila_h2, column=1,  value=reg_info)               # A  Registro info
                ws_h2.cell(row=fila_h2, column=5,  value="OC ISLAS CANANARIAS")  # E  Desc. org. compras
                ws_h2.cell(row=fila_h2, column=6,  value=cod_centro)             # F  Centro
                ws_h2.cell(row=fila_h2, column=7,  value=nom_centro)             # G  Nombre centro
                ws_h2.cell(row=fila_h2, column=8,  value=grp_comp   or None)    # H  Grupo compras
                ws_h2.cell(row=fila_h2, column=9,  value=desc_grp_c or None)    # I  Desc. grupo compras
                ws_h2.cell(row=fila_h2, column=10, value=material)               # J  Material
                ws_h2.cell(row=fila_h2, column=11, value=desc_sap   or None)    # K  Texto breve mat.
                ws_h2.cell(row=fila_h2, column=12, value=grp_prod   or None)    # L  Grupo artículos
                ws_h2.cell(row=fila_h2, column=34, value=ind_imp    or None)    # AH Indicador impuestos
                # col AN (40) queda vacía en filas nuevas — solo se escribe en heredadas descartadas

                # Columnas de valor fijo obligatorio SAP (B,C,D,M,N,S,W,X,AD,AE,
                # AF,AJ,AK,AO) — único punto de escritura, ver COLUMNAS_FIJAS_H2.
                self._escribir_fijas_h2_fila(ws_h2, fila_h2)

                # Copiar formato desde la última fila preexistente de Hoja2.
                # Garantiza que las nuevas filas tienen el mismo estilo visual
                # (fill, border, font) que las filas heredadas del Purchasing,
                # respetando el formato propio de Hoja2 (distinto al de Hoja1/3).
                fila_ref_h2 = max(2, siguiente_fila_h2 - 1)
                aplicar_formato_fila(ws_h2, fila_h2, "Datos org.compr.p.reg.info", fila_ref=fila_ref_h2)

                filas_h2_escritas.append((
                    reg_info,         # A → H3-A
                    "0",              # B → H3-B  tipo registro
                    "Normal",         # C → H3-C  desc. categoría
                    "IC00",           # D → H3-D  org. compras
                    cod_centro,       # F → H3-E  centro
                    material,         # J → H3-F  material
                    desc_sap or None, # K → H3-G  texto breve mat.
                    grp_prod or None, # L → H3-H  grupo artículos
                ))

        codigos_str = ", ".join(c for c, _ in centros)
        _log(self._parent,
             f"Hoja2 — {idx_escritura} fila(s) nueva(s) añadida(s) "
             f"({n_nuevos} material(es) evaluado(s) × {NUM_BLOQUES} bloques"
             f"{f', {saltadas_duplicado} omitida(s) por duplicado' if saltadas_duplicado else ''}"
             f"{f', {saltadas_sin_tarifa} omitida(s) por no estar en tarifa (heredado mantenido)' if saltadas_sin_tarifa else ''}) "
             f"— centros activos: {codigos_str}.", "OK")

        return filas_h2_escritas


    def _sincronizar_igic_h2(self, ws_h2, diccionario):
        """
        PASO 4 (extraído de _rellenar_h2_nuevas_filas): recorre TODA
        Hoja2 (heredadas + nuevas) y sincroniza el indicador de
        impuestos (AH) contra el LISTADO DALI-SAP -- antes solo se
        aplicaba a las filas nuevas del paso 3, dejando las heredadas
        con un IGIC obsoleto sin aviso. Si la fila corregida es
        heredada, se marca AO="X" para que SAP reprocese el cambio.
        Filas ya marcadas para borrado lógico (AN=X) se excluyen.
        No devuelve nada -- efecto directo sobre `ws_h2` y self._audit.
        """
        # ── PASO 4: Sincronizar IGIC (columna AH) en TODA Hoja2 ───────────────
        #
        #  Hasta ahora, el indicador de impuestos (AH) del LISTADO solo se
        #  aplicaba a las filas realmente NUEVAS (paso 3 de arriba). Las filas
        #  HEREDADAS del Purchasing se dejaban intactas aunque el LISTADO
        #  hubiera cambiado de IGIC para ese material — dejando un valor
        #  obsoleto en SAP sin ningún aviso ni corrección.
        #
        #  Ahora, con Hoja2 ya completa (heredadas + nuevas), se recorren
        #  TODAS las filas, se compara el IGIC actual con el del LISTADO y,
        #  si difiere, se actualiza. Si la fila corregida es HEREDADA (no
        #  estaba ya marcada como Nueva), se marca AO (41) = "X" para que SAP
        #  reprocese la fila y el cambio surta efecto — sin esta marca SAP no
        #  detectaría la modificación.
        #
        #  Filas ya marcadas para borrado lógico (AN=40="X") se excluyen: no
        #  tiene sentido corregir el IGIC de una fila que se va a dar de baja.
        COL_AH_H2          = 34   # AH — Indicador de impuestos
        COL_AO_H2          = 41   # AO — Nueva
        COL_AN_H2_BORRADO  = 40   # AN — Organización de compras marcada para borrado

        n_igic_actualizados    = 0
        n_igic_marcados_nuevo  = 0

        for fila_row in ws_h2.iter_rows(min_row=2):
            if len(fila_row) < COL_AH_H2:
                continue

            val_mat_sync = fila_row[9].value if len(fila_row) > 9 else None   # col J – Material
            if val_mat_sync is None:
                continue

            fila_num_sync = fila_row[0].row

            # No tocar filas ya marcadas para borrado lógico
            val_an_sync = fila_row[COL_AN_H2_BORRADO - 1].value \
                if len(fila_row) >= COL_AN_H2_BORRADO else None
            if val_an_sync is not None and str(val_an_sync).strip().upper() == "X":
                continue

            raw_m_sync = val_mat_sync
            raw_m_sync = normalizar_codigo_sap(raw_m_sync)
            clave_mat_sync = str(raw_m_sync).strip()

            datos_listado_sync = diccionario.get(clave_mat_sync)
            if not datos_listado_sync:
                continue   # material sin match en LISTADO: no tocar lo existente

            ind_imp_listado = (datos_listado_sync[4] or "").strip() \
                if len(datos_listado_sync) > 4 else ""
            if not ind_imp_listado:
                continue   # LISTADO sin IGIC informado: no borrar lo que ya había

            val_ah_actual = fila_row[COL_AH_H2 - 1].value if len(fila_row) >= COL_AH_H2 else None
            val_ah_actual_str = str(val_ah_actual).strip() if val_ah_actual is not None else ""

            if val_ah_actual_str == ind_imp_listado:
                continue   # sin cambio — nada que sincronizar

            ws_h2.cell(row=fila_num_sync, column=COL_AH_H2, value=ind_imp_listado)
            n_igic_actualizados += 1

            val_ao_actual = fila_row[COL_AO_H2 - 1].value if len(fila_row) >= COL_AO_H2 else None
            ya_marcada_nueva = str(val_ao_actual).strip().upper() == "X" \
                if val_ao_actual is not None else False

            if not ya_marcada_nueva:
                ws_h2.cell(row=fila_num_sync, column=COL_AO_H2, value="X")
                n_igic_marcados_nuevo += 1
                _log(self._parent,
                     f"  H2 fila {fila_num_sync} [material {clave_mat_sync}]: IGIC "
                     f"{val_ah_actual_str!r} → {ind_imp_listado!r} (heredado desactualizado) "
                     f"— marcada como Nueva para que SAP aplique el cambio.", "WARN")
            else:
                _log(self._parent,
                     f"  H2 fila {fila_num_sync} [material {clave_mat_sync}]: IGIC "
                     f"{val_ah_actual_str!r} → {ind_imp_listado!r} actualizado.", "INFO")

        if n_igic_actualizados:
            _log(self._parent,
                 f"Hoja2 — IGIC sincronizado: {n_igic_actualizados} fila(s) actualizada(s) "
                 f"({n_igic_marcados_nuevo} heredada(s) marcada(s) como Nueva para SAP).",
                 "OK")
        else:
            _log(self._parent, "Hoja2 — IGIC: sin discrepancias frente al LISTADO.", "INFO")

        if hasattr(self, "_audit"):
            self._audit["n_igic_actualizados"] = n_igic_actualizados
            self._audit["n_igic_marcados_nuevo"] = n_igic_marcados_nuevo


    def _normalizar_obligatorios_h2_completo(self, ws_h2, diccionario):
        """
        PASO 5 (extraído de _rellenar_h2_nuevas_filas): recorre TODA
        Hoja2 (heredadas incluidas) y rellena columnas obligatorias que
        nunca pasan por el relleno de filas nuevas del paso 3 -- AF solo
        en filas activas, AJ en todas las filas con material (incluidas
        las de borrado, que SAP también exige), y AH en filas de borrado
        que quedaron sin indicador de impuestos (recuperado del LISTADO
        DALI-SAP, igual que en el paso 4). No devuelve nada -- efecto
        directo sobre `ws_h2` y self._audit.
        """
        COL_AH_H2          = 34   # AH — Indicador de impuestos
        COL_AN_H2_BORRADO  = 40   # AN — Organización de compras marcada para borrado

        # ── PASO 5: Normalizar obligatorios en TODA Hoja2 (heredadas incluidas) ───
        #   Las filas heredadas del Purchasing existente nunca pasan por el
        #   relleno de columnas obligatorias que sí reciben las filas nuevas
        #   (PASO 3 más arriba). Si el Purchasing original traía AF, AH o AJ
        #   vacías, se quedaban así para siempre, aunque la fila se tocara
        #   después en otros pasos (p. ej. marcada como Nueva por el sync
        #   de IGIC). Este paso cierra ese hueco para cualquier fila con
        #   material, sin pisar valores ya existentes.
        #
        #   Alcance por columna:
        #     · AF (32) — solo en filas ACTIVAS (no marcadas para borrado).
        #       No tiene sentido fijar el control de fecha del precio de una
        #       fila que se va a dar de baja.
        #     · AJ (36) — en TODAS las filas con material, estén o no
        #       marcadas para borrado (AN=X). SAP exige este campo también
        #       para procesar correctamente la baja del registro.
        #     · AH (34) — en filas marcadas para borrado (AN=X) que hayan
        #       quedado sin indicador de impuestos (el sync de IGIC del
        #       PASO 4 las excluye a propósito). Se recupera del LISTADO
        #       DALI-SAP igual que en el PASO 4; si el material no tiene
        #       entrada en el LISTADO no se inventa un valor y se deja
        #       constancia en el log.
        #
        #   AK (Entrega ilimitada) queda excluida a propósito: no es una
        #   columna de relleno obligatorio para filas heredadas.
        n_af_normalizadas  = 0
        n_aj_normalizadas  = 0
        n_ah_borrado_ok    = 0
        n_ah_borrado_fail  = 0

        for fila_row in ws_h2.iter_rows(min_row=2):
            val_j = fila_row[9].value if len(fila_row) > 9 else None
            if val_j is None:
                continue

            fila_num = fila_row[0].row
            val_an = fila_row[COL_AN_H2_BORRADO - 1].value \
                if len(fila_row) >= COL_AN_H2_BORRADO else None
            es_borrado = val_an is not None and str(val_an).strip().upper() == "X"

            # AF/AJ (columnas de valor fijo) — único punto de normalización,
            # ver COLUMNAS_FIJAS_H2 / _normalizar_obligatorios_fila_h2.
            cols_rellenadas = self._normalizar_obligatorios_fila_h2(
                ws_h2, fila_num, es_borrado)
            if 32 in cols_rellenadas:
                n_af_normalizadas += 1
            if 36 in cols_rellenadas:
                n_aj_normalizadas += 1

            # AH — solo filas de borrado que quedaron sin indicador de impuestos
            if es_borrado:
                celda_ah = fila_row[COL_AH_H2 - 1] if len(fila_row) >= COL_AH_H2 \
                    else ws_h2.cell(row=fila_num, column=COL_AH_H2)
                if celda_ah.value is None or str(celda_ah.value).strip() == "":
                    raw_m_borrado = val_j
                    raw_m_borrado = normalizar_codigo_sap(raw_m_borrado)
                    clave_mat_borrado = str(raw_m_borrado).strip()
                    datos_borrado = diccionario.get(clave_mat_borrado)
                    ind_imp_borrado = (datos_borrado[4] or "").strip() \
                        if datos_borrado and len(datos_borrado) > 4 else ""
                    if ind_imp_borrado:
                        ws_h2.cell(row=fila_num, column=COL_AH_H2, value=ind_imp_borrado)
                        n_ah_borrado_ok += 1
                    else:
                        n_ah_borrado_fail += 1
                        _log(self._parent,
                             f"  H2 fila {fila_num} [material {clave_mat_borrado}, borrado]: "
                             f"AH vacía y sin match en LISTADO DALI-SAP — no se pudo "
                             f"completar el indicador de impuestos.", "WARN")

        n_normalizadas = n_af_normalizadas + n_aj_normalizadas + n_ah_borrado_ok
        if n_normalizadas:
            _log(self._parent,
                 f"Hoja2 — Normalización obligatorios: {n_af_normalizadas} AF, "
                 f"{n_aj_normalizadas} AJ, {n_ah_borrado_ok} AH (borrado) "
                 f"rellenada(s)"
                 f"{f' | {n_ah_borrado_fail} AH (borrado) sin match en LISTADO' if n_ah_borrado_fail else ''}.",
                 "WARN")
        else:
            _log(self._parent,
                 "Hoja2 — Normalización obligatorios: sin huecos detectados "
                 "(AF/AJ/AH ya completas).", "INFO")

        if hasattr(self, "_audit"):
            self._audit["n_normalizadas_obligatorios_h2"] = n_normalizadas
            self._audit["n_ah_borrado_sin_listado"] = n_ah_borrado_fail


    def _rellenar_h2_nuevas_filas(self, wb, nuevos_h1: list,
                                   primera_fila_nueva_h1: int,
                                   diccionario: dict,
                                   codigos_tarifa_set: set = None) -> list:
        """
        Gestiona la Hoja 2 cuando se trabaja desde un Purchasing existente.

        Aplica tres reglas respecto al comportamiento «desde cero»:

        1. BORRADO LÓGICO — Filas heredadas del Purchasing cuyo centro (col F)
           NO está en self._centros_activos reciben el valor
           «Organización de compras marcada para borrado» en la col AN (40).
           La fila NO se elimina físicamente.

        2. ANTI-DUPLICADO — Para cada combinación (material, centro) que ya
           exista entre las filas heredadas NO se escribe una fila nueva,
           aunque el material figure en nuevos_h1.

        3. CREACIÓN POR BLOQUES — Para las combinaciones (material, centro)
           que sí son realmente nuevas se generan filas con la misma estructura
           que usa _rellenar_hoja2 en CreadorRINFO (un bloque por cada centro
           activo), pero SOLO para materiales presentes en la tarifa activa
           (codigos_tarifa_set). Un material heredado que el usuario decidió
           MANTENER en la ventana "sin tarifa" (AJ sin marcar, pero ausente
           de la tarifa) NO debe expandirse a los demás centros activos: debe
           conservarse únicamente en el/los centro(s) en que ya existía en el
           Purchasing heredado. Expandirlo a todos los centros seleccionados
           al inicio crearía PIRs nuevos en centros donde ese material nunca
           existió, lo cual es incorrecto — esa expansión por centro solo
           tiene sentido para materiales que vienen de la tarifa.

        Parámetros
        ----------
        nuevos_h1             : lista de (cod_sap_str, cod_prov_str|None)
        primera_fila_nueva_h1 : primera fila de datos nuevos en Hoja 1
        codigos_tarifa_set    : set de códigos SAP presentes en la tarifa activa.
                                 Determina qué materiales de Hoja 1 se expanden a
                                 todos los centros activos (PASO 3). Si es None,
                                 se conserva el comportamiento antiguo (expandir
                                 todos los materiales) por compatibilidad.

        Devuelve lista de tuplas con los datos escritos en H2 (para H3):
            [(reg_info, tipo_reg, desc_cat, org_comp, centro, material,
              txt_mat, grp_art), ...]
        """
        COL_AN_BORRADO = 40   # AN — «Organización de compras marcada para borrado»

        if self.HOJA_2 not in wb.sheetnames:
            _log(self._parent,
                 f"Hoja «{self.HOJA_2}» no encontrada — omitida.", "WARN")
            return []

        ws_h1 = wb[self.HOJA_PRINCIPAL]
        ws_h2 = wb[self.HOJA_2]

        centros_activos_set = {cod for cod, _ in self._centros_activos}

        (combinaciones_existentes, _escritas_h2, _claves_borrado_h2,
         ultima_fila_h2) = self._marcar_borrado_logico_y_escanear_h2(
            ws_h2, centros_activos_set)

        self._propagar_borrado_h2_a_h3(wb, _claves_borrado_h2)

        filas_nuevas_h1 = self._leer_h1_para_expansion_h2(ws_h1)

        filas_h2_escritas = self._escribir_h2_bloques_nuevos(
            ws_h2, filas_nuevas_h1, ultima_fila_h2, _escritas_h2,
            combinaciones_existentes, diccionario, codigos_tarifa_set)

        self._sincronizar_igic_h2(ws_h2, diccionario)

        self._normalizar_obligatorios_h2_completo(ws_h2, diccionario)

        return filas_h2_escritas

    # ── RELLENO HOJA 3 (nuevas filas) ────────────────────────────────────────

    def _rellenar_h3_nuevas_filas(self, wb, filas_h2_nuevas: list,
                                   diccionario: dict,
                                   dic_tarifa: dict) -> int:
        """
        Añade al final de la Hoja 3 una fila por cada fila nueva de la Hoja 2.
        Devuelve el número de la primera fila nueva insertada (o None si no se
        insertó ninguna), para que el ajuste de coherencia sepa qué filas
        pertenecen al batch actual y no deben ser modificadas.
        """
        if self.HOJA_3 not in wb.sheetnames:
            _log(self._parent,
                 f"Hoja «{self.HOJA_3}» no encontrada — omitida.", "WARN")
            return

        if not filas_h2_nuevas:
            return None

        ws_h3 = wb[self.HOJA_3]

        # ── NORMA SAP — Col K "ID grupo condiciones" en Hoja3 ───────────────
        # SAP EXIGE que la columna K (ID grupo condiciones) contenga SIEMPRE
        # números enteros consecutivos comenzando en 1, sin huecos ni saltos.
        # La secuencia es GLOBAL para todas las filas de la hoja (fila 2 → 1,
        # fila 3 → 2, ...) y aplica POR IGUAL a materiales heredados y nuevos,
        # SIN ninguna distinción entre ellos.
        # Regla confirmada por SAP: 1, 2, 3 ... hasta la última fila con datos.
        # ─────────────────────────────────────────────────────────────────────
        # En este paso se reescriben las filas heredadas (1..N) antes de añadir
        # los nuevos (N+1, N+2...). Al final del proceso completo (después de
        # ajustes de fechas y eliminaciones por absorción) se ejecuta una
        # renumeración final para garantizar que no queden huecos.
        # ─────────────────────────────────────────────────────────────────────
        #
        # Deduplicación funcional H3 — clave completa PIR:
        #   proveedor|org|material|centro|desde|hasta
        # Se puebla primero con las filas heredadas y luego se comprueba en cada
        # fila nueva antes de escribirla. Blinda contra duplicados de cualquier
        # origen: distintas hojas, iteraciones o materiales repetidos en el input.
        _prov_str_h3 = str(self._codigo_sap).strip() if self._codigo_sap else ""
        _ORG_H3      = "IC00"
        _fi_str      = str(self._fecha_inicio).strip() if self._fecha_inicio else ""
        _ff_str      = str(self._fecha_fin).strip()    if self._fecha_fin    else ""
        _escritas_h3: set = set()   # "prov|org|mat|centro|desde|hasta"

        ultima_fila_h3 = 1
        contador_k_heredadas = 0

        for fila_row in ws_h3.iter_rows(min_row=2, min_col=1, max_col=11):
            val_a = fila_row[0].value   # col A
            val_f = fila_row[5].value if len(fila_row) > 5 else None  # col F Material
            if val_a is not None or val_f is not None:
                ultima_fila_h3 = fila_row[0].row
                contador_k_heredadas += 1
                # Reescribir K consecutivo incluyendo filas heredadas
                celda_k = fila_row[10]
                celda_k.value         = str(contador_k_heredadas)  # texto, no entero
                celda_k.data_type     = "s"                        # fuerza t="s" en XML → SAP lo lee como shared string
                celda_k.number_format = "@"                        # triangulito verde SAP
                # Registrar en set funcional para detectar duplicados en filas nuevas
                _val_e_h = fila_row[4].value if len(fila_row) > 4 else None  # E Centro
                _val_f_h = fila_row[5].value if len(fila_row) > 5 else None  # F Material
                _val_i_h = fila_row[8].value if len(fila_row) > 8 else None  # I Inicio
                _val_j_h = fila_row[9].value if len(fila_row) > 9 else None  # J Fin
                if _val_f_h is not None:
                    _rm = _val_f_h
                    _rm = normalizar_codigo_sap(_rm)
                    _clave_her_h3 = (
                        f"{_prov_str_h3}|{_ORG_H3}"
                        f"|{str(_rm).strip()}"
                        f"|{str(_val_e_h).strip() if _val_e_h is not None else ''}"
                        f"|{str(_val_i_h).strip() if _val_i_h is not None else ''}"
                        f"|{str(_val_j_h).strip() if _val_j_h is not None else ''}"
                    )
                    _escritas_h3.add(_clave_her_h3)

        # n_filas_h3_existentes = número de filas heredadas ya renumeradas (1..N)
        n_filas_h3_existentes = contador_k_heredadas

        siguiente_fila_h3 = ultima_fila_h3 + 1
        matches_listado   = 0
        matches_tarifa    = 0
        saltadas_dup_h3   = 0
        idx_escritura_h3  = 0   # filas efectivamente escritas (excluye duplicados)

        for idx_local, (reg_info, tipo_reg, desc_cat, org_comp,
                         centro, material, txt_mat, grp_art) in enumerate(filas_h2_nuevas):

            # Normalizar clave antes del guard para poder construirla
            if material is not None:
                _rm_h3 = material
                _rm_h3 = normalizar_codigo_sap(_rm_h3)
                _clave_mat_h3 = str(_rm_h3).strip()
            else:
                _clave_mat_h3 = ""
            _centro_h3 = str(centro).strip() if centro is not None else ""

            # ── Deduplicación funcional H3 (Punto B) ─────────────────────────
            # Clave: proveedor|org|material|centro|fecha_inicio|fecha_fin
            # Blinda contra duplicados de distintas hojas, iteraciones o
            # materiales repetidos en el input.
            _clave_h3_nueva = (
                f"{_prov_str_h3}|{_ORG_H3}"
                f"|{_clave_mat_h3}|{_centro_h3}"
                f"|{_fi_str}|{_ff_str}"
            )
            if _clave_h3_nueva in _escritas_h3:
                saltadas_dup_h3 += 1
                continue
            _escritas_h3.add(_clave_h3_nueva)

            fila_h3    = siguiente_fila_h3 + idx_escritura_h3
            idx_escritura_h3 += 1
            # ID grupo condiciones (col K): valor provisional que continúa la
            # secuencia tras los heredados. La renumeración final (_renumerar_col_k_h3)
            # ejecutada después de todos los ajustes garantiza la secuencia 1..N sin huecos.
            id_grupo_k = n_filas_h3_existentes + idx_escritura_h3

            # Normalizar clave (ya normalizada arriba; reasignamos para compatibilidad)
            if material is not None:
                raw_m = material
                raw_m = normalizar_codigo_sap(raw_m)
                clave_mat = str(raw_m).strip()
            else:
                clave_mat = ""

            # ── Valores copiados de Hoja2 ──────────────────────────────────────
            ws_h3.cell(row=fila_h3, column=1,  value=reg_info)   # A  Registro info
            ws_h3.cell(row=fila_h3, column=2,  value=tipo_reg)   # B  Tipo registro
            ws_h3.cell(row=fila_h3, column=3,  value=desc_cat)   # C  Desc. categoría
            ws_h3.cell(row=fila_h3, column=4,  value=org_comp)   # D  Org. compras
            ws_h3.cell(row=fila_h3, column=5,  value=centro)     # E  Centro
            ws_h3.cell(row=fila_h3, column=6,  value=material)   # F  Material
            ws_h3.cell(row=fila_h3, column=7,  value=txt_mat)    # G  Texto breve material
            ws_h3.cell(row=fila_h3, column=8,  value=grp_art)    # H  Grupo artículos

            # ── Valores fijos ──────────────────────────────────────────────────
            ws_h3.cell(row=fila_h3, column=9,  value=self._fecha_inicio)     # I  Inicio validez
            ws_h3.cell(row=fila_h3, column=10, value=self._fecha_fin)        # J  Fin validez
            _celda_k = ws_h3.cell(row=fila_h3, column=11)
            _celda_k.value         = str(id_grupo_k)   # texto, no entero
            _celda_k.data_type     = "s"               # fuerza t="s" en XML → SAP lo lee como shared string
            _celda_k.number_format = "@"               # K  ID grupo condiciones (entero consecutivo desde 1)
            ws_h3.cell(row=fila_h3, column=12, value="Condición base")       # L  Orden condiciones
            ws_h3.cell(row=fila_h3, column=13, value="PPR0")                 # M  Tipo condición
            ws_h3.cell(row=fila_h3, column=14, value="Precio bruto predet.") # N  Desc. clase cond.
            ws_h3.cell(row=fila_h3, column=16, value="EUR")                  # P  Unidad importe
            ws_h3.cell(row=fila_h3, column=17, value=1)                     # Q  Determinación precios
            ws_h3.cell(row=fila_h3, column=20, value="C")                    # T  Clase de cálculo
            ws_h3.cell(row=fila_h3, column=21, value="Cantidad")             # U  Desc. clase cálculo
            ws_h3.cell(row=fila_h3, column=22, value=1)                    # V  Numerador conversión
            ws_h3.cell(row=fila_h3, column=23, value=1)                    # W  Denominador conversión
            ws_h3.cell(row=fila_h3, column=24, value="No")                   # X  Escala
            ws_h3.cell(row=fila_h3, column=26, value="X")                    # Z  Nueva

            # ── Cruce LISTADO → R(18) y S(19) UNIDAD SAP ──────────────────────
            datos_listado = diccionario.get(clave_mat)
            if datos_listado:
                _, unidad, _, _, _, _, _ = datos_listado
                if unidad:
                    ws_h3.cell(row=fila_h3, column=18, value=unidad)  # R  Unidad det. precio
                    ws_h3.cell(row=fila_h3, column=19, value=unidad)  # S  Unidad medida base
                    matches_listado += 1

            # ── Cruce tarifa → O(15) PRECIO ────────────────────────────────────
            precio = dic_tarifa.get(clave_mat)
            if precio is not None:
                ws_h3.cell(row=fila_h3, column=15, value=precio)     # O  Cantidad (precio)
                matches_tarifa += 1

            # Copiar formato desde la última fila preexistente de Hoja3.
            # Garantiza que las nuevas filas tienen el mismo estilo visual
            # (fill, border, font) que las filas heredadas del Purchasing,
            # respetando el formato propio de Hoja3 (distinto al de Hoja1/2).
            fila_ref_h3 = max(2, siguiente_fila_h3 - 1)
            aplicar_formato_fila(ws_h3, fila_h3, "Det.precios p.registros info", fila_ref=fila_ref_h3)

        _n_escritas_h3 = idx_escritura_h3
        _log(self._parent,
             f"Hoja3 — {_n_escritas_h3} fila(s) nueva(s) añadida(s)"
             + (f", {saltadas_dup_h3} omitida(s) por duplicado funcional" if saltadas_dup_h3 else "")
             + f", K={n_filas_h3_existentes + 1}–{n_filas_h3_existentes + _n_escritas_h3}"
             f" (consecutivo global desde 1, entero),"
             f" I={self._fecha_inicio} / J={self._fecha_fin},"
             f" {matches_listado} cruces LISTADO (R/S),"
             f" {matches_tarifa} cruces tarifa (O).", "OK")

        return siguiente_fila_h3

    # ── SINCRONIZACIÓN UM EN TODA HOJA 3 (cols R y S) ──────────────────────────

    def _sincronizar_um_hoja3(self, wb, diccionario: dict) -> None:
        """
        Sincroniza la Unidad de medida (col R «Unidad det. precio» y
        col S «Unidad medida base») en TODA la Hoja 3, incluidas las filas
        HEREDADAS del Purchasing existente.

        Mismo criterio que _sincronizar_um_hoja1 (Hoja 1, cols N/P): hasta
        ahora R/S solo se rellenaban al escribir filas realmente NUEVAS. Con
        este paso se recorren TODAS las filas con material, se compara R y S
        frente a la unidad del LISTADO CODIGOS DALI-SAP y, si difiere, se
        corrige.

        IMPORTANTE — igual que ocurría con AK en Hoja 1, la columna Z (26,
        «Nueva») de Hoja 3 NO es un simple indicador informativo: le dice a
        SAP que la condición de precio de esa fila es NUEVA. En una fila
        heredada, la col A ya contiene el registro info REAL existente en
        SAP (p.ej. 5300001868), con su condición de precio ya creada para
        esa organización/centro.

        MATIZ CONFIRMADO CON LOG REAL (v8.12.0): a diferencia de AK en
        Hoja 1, Z=X por sí solo SÍ es un mecanismo legítimo que SAP acepta
        con normalidad — es como _marcar_modificado() (más abajo) señala
        cambios reales de fecha/precio en filas heredadas cientos de veces
        por carga sin incidencias. El problema aparece SOLO cuando esa
        misma fila, en la misma pasada, tiene TAMBIÉN un cambio de UM
        (R/S): ahí SAP no puede conciliar "modificar in situ" con un
        cambio de unidad base y devuelve "Por favor, adecúe la UM-base y
        el factor de conversión". Esto se confirmó con el material 10103364:
        _marcar_modificado() ya había marcado Z=X en sus 10 filas de Hoja 3
        por un cambio genuino de Fin validez (nada que ver con la UM); al
        llegar aquí y corregir R/S de PI a KG en las mismas filas, dejar Z
        intacto (como hacía la versión anterior de este fix) no bastó: Z
        seguía en "X" desde el paso de fechas y el error persistió.

        Por eso esta sincronización, al corregir R/S en una fila heredada,
        FUERZA Z a vacío en esa fila — sobrescribiendo cualquier "X" que
        _marcar_modificado() haya dejado antes por fecha/precio.

        CORRECCIÓN (2026-10-06, caso real FRAN Y CHEMI / COCOS 10101648):
        este docstring afirmaba que, sin Z, "SAP recoge igualmente el cambio
        de fecha y de UM al modificar el registro existente". Es CIERTO para
        la UM, pero FALSO para fecha y precio: SAP ignora en silencio una
        modificación in situ de validez/precio sin Z=X. COCOS (condición
        antigua en KG, cabecera y LISTADO en PI) quedó sin precio de octubre
        en SAP (seguía 01.11.2025→30.11.2025 @1,20 tras cargar un RINFO con
        01.10→31.10 @1,35), y el único rastro era una línea de resumen.
        Por eso ahora cada fila afectada se registra y se emite un WARN
        agregado explícito: el cambio de fecha/precio de esas filas NO queda
        registrado en SAP y hay que resolverlo a mano.

        Filas ya marcadas con Indicador borrado (Y=25="X") se excluyen
        (revertido 2026-09-22, ver HISTORIAL y _sincronizar_um_hoja1):
        para materiales de baja, no se toca ningún valor heredado --
        solo se marcan para borrado donde corresponda. Cambiar la UM de
        un registro que ya existe en SAP sigue dando el error de UM-base
        aunque se vaya a borrar, salvo que el registro se elimine antes
        en SAP a mano -- no es algo con lo que se pueda contar en cada
        carga, así que la UM heredada se deja tal cual estaba.
        """
        if self.HOJA_3 not in wb.sheetnames:
            return

        ws_h3 = wb[self.HOJA_3]

        COL_F_MATERIAL_H3 = 6    # F  Material (código SAP)
        COL_R_UM_PRECIO    = 18  # R  Unidad det. precio
        COL_S_UM_BASE       = 19  # S  Unidad medida base
        COL_Y_H3_BORRADO    = 25  # Y  Indicador borrado
        COL_Z_H3_NUEVA      = 26  # Z  Nueva

        n_um_actualizados = 0
        n_z_forzados_vacio = 0
        z_perdidos: dict = {}   # {material: [(centro, inicio, fin, precio), ...]}

        for fila_row in ws_h3.iter_rows(min_row=2):
            if len(fila_row) < COL_S_UM_BASE:
                continue

            val_mat = fila_row[COL_F_MATERIAL_H3 - 1].value \
                if len(fila_row) >= COL_F_MATERIAL_H3 else None
            if val_mat is None:
                continue

            fila_num = fila_row[0].row

            # No tocar filas ya marcadas con indicador de borrado --
            # ningún valor heredado se modifica en materiales de baja.
            val_y = fila_row[COL_Y_H3_BORRADO - 1].value \
                if len(fila_row) >= COL_Y_H3_BORRADO else None
            if val_y is not None and str(val_y).strip().upper() == "X":
                continue

            raw_m = val_mat
            raw_m = normalizar_codigo_sap(raw_m)
            clave_mat = str(raw_m).strip()

            datos_listado = diccionario.get(clave_mat)
            if not datos_listado:
                continue   # material sin match en LISTADO: no tocar lo existente

            unidad_correcta = (datos_listado[1] or "").strip() \
                if len(datos_listado) > 1 else ""
            if not unidad_correcta:
                continue   # LISTADO sin UM informada: no borrar lo que ya había

            val_r_actual = fila_row[COL_R_UM_PRECIO - 1].value \
                if len(fila_row) >= COL_R_UM_PRECIO else None
            val_s_actual = fila_row[COL_S_UM_BASE - 1].value \
                if len(fila_row) >= COL_S_UM_BASE else None
            val_r_str = str(val_r_actual).strip() if val_r_actual is not None else ""
            val_s_str = str(val_s_actual).strip() if val_s_actual is not None else ""

            if val_r_str == unidad_correcta and val_s_str == unidad_correcta:
                continue   # sin cambio — nada que sincronizar

            if val_r_str != unidad_correcta:
                ws_h3.cell(row=fila_num, column=COL_R_UM_PRECIO, value=unidad_correcta)
            if val_s_str != unidad_correcta:
                ws_h3.cell(row=fila_num, column=COL_S_UM_BASE, value=unidad_correcta)
            n_um_actualizados += 1

            # Forzar Z a vacío: si _marcar_modificado() ya lo había puesto a
            # "X" por un cambio de fecha/precio en esta misma fila, un
            # cambio de UM simultáneo hace que SAP rechace la carga con
            # "adecúe la UM-base y el factor de conversión". Se prioriza
            # que la carga no reviente sobre conservar el flag informativo.
            val_z_actual = fila_row[COL_Z_H3_NUEVA - 1].value \
                if len(fila_row) >= COL_Z_H3_NUEVA else None
            if val_z_actual is not None and str(val_z_actual).strip().upper() == "X":
                # OJO: ws_h3.cell(..., value=None) NO borra la celda — openpyxl
                # solo escribe si value is not None (ver Worksheet.cell()), así
                # que pasar None se ignora y deja la "X" anterior intacta. Hay
                # que coger la celda y asignar .value = None directamente.
                celda_z = ws_h3.cell(row=fila_num, column=COL_Z_H3_NUEVA)
                celda_z.value = None
                n_z_forzados_vacio += 1
                z_perdidos.setdefault(clave_mat, []).append((
                    fila_row[4].value,       # E centro
                    fila_row[8].value,       # I inicio
                    fila_row[9].value,       # J fin
                    fila_row[14].value))     # O precio
                if hasattr(self, "_audit"):
                    self._audit.get("materiales_modificados", set()).discard(clave_mat)
                _log(self._parent,
                     f"  H3 fila {fila_num} [material {clave_mat}]: UM "
                     f"R={val_r_str!r}/S={val_s_str!r} → {unidad_correcta!r} actualizado "
                     f"— Z forzada a vacío: el cambio de fecha/precio de esta fila "
                     f"NO se registrará en SAP (SAP ignora una modificación de "
                     f"validez/precio sin Z=X; combinada con el cambio de UM, Z=X "
                     f"rompe la carga).",
                     "WARN")
            else:
                _log(self._parent,
                     f"  H3 fila {fila_num} [material {clave_mat}]: UM "
                     f"R={val_r_str!r}/S={val_s_str!r} → {unidad_correcta!r} actualizado "
                     f"(sin tocar Z — SAP la recoge al modificar la condición existente).",
                     "INFO")

        if z_perdidos:
            _detalle_z = "; ".join(
                f"{m} (centro(s) {', '.join(sorted({str(c) for c, *_ in fs}))}; "
                f"{fs[0][1]}→{fs[0][2]} @ {fs[0][3]})"
                for m, fs in sorted(z_perdidos.items())[:15])
            _log(self._parent,
                 f"  ⚠ {len(z_perdidos)} material(es) con el cambio de fecha/PRECIO "
                 f"SIN REGISTRAR en SAP: la unidad de medida de su condición "
                 f"difiere de la del LISTADO y, al corregirla en la misma fila, "
                 f"se ha quitado la marca Nueva para que la carga no falle -- SAP "
                 f"ignora una modificación de validez/precio sin esa marca, así "
                 f"que SU PRECIO NUEVO NO ESTARÁ VIGENTE tras cargar este RINFO. "
                 f"Hay que crear a mano esa condición en SAP (o corregir la UM "
                 f"antes): {_detalle_z}"
                 + (" …" if len(z_perdidos) > 15 else ""),
                 "WARN")

        if n_um_actualizados:
            _log(self._parent,
                 f"Hoja3 — UM sincronizada (R/S): {n_um_actualizados} fila(s) "
                 f"actualizada(s)" +
                 (f" ({n_z_forzados_vacio} con Z forzada a vacío por conflicto "
                  f"fecha+UM)." if n_z_forzados_vacio else "."), "OK")
        else:
            _log(self._parent,
                 "Hoja3 — UM (R/S): sin discrepancias frente al LISTADO.", "INFO")

        if hasattr(self, "_audit"):
            self._audit["n_um_h3_actualizados"] = n_um_actualizados
            self._audit["n_um_h3_z_forzados_vacio"] = n_z_forzados_vacio

    # ── TELÉFONO EN TODAS LAS FILAS DE HOJA 1 ────────────────────────────────

    def _aplicar_telefono_hoja1(self, ws) -> None:
        """
        Escribe el teléfono del proveedor en la columna AD (col 30) de la Hoja 1,
        desde la fila 2 hasta la última fila con valor en la columna E (Material).
        Se aplica a TODOS los materiales: tanto los preexistentes como los nuevos.
        Solo actúa si self._telefono_proveedor tiene valor.
        """
        telefono = getattr(self, "_telefono_proveedor", None)
        if not telefono:
            _log(self._parent,
                 "  Teléfono no indicado — col AD de Hoja1 sin cambios.", "INFO")
            return

        ultima_fila = ultima_fila_con_datos(ws, columna=5)

        if ultima_fila < 2:
            _log(self._parent,
                 "  Hoja1 sin datos en col E — teléfono no escrito.", "WARN")
            return

        for fila_num in range(2, ultima_fila + 1):
            ws.cell(row=fila_num, column=30).value = telefono  # AD  Nº teléfono proveedor

        _log(self._parent,
             f"  Teléfono proveedor (col AD) escrito en filas 2-{ultima_fila}: "
             f"«{telefono}».", "OK")

    # Fechas fijas para Hoja 1 (Info general): independientes de lo que pida el usuario.
    # Las fechas del usuario van exclusivamente en Hoja 3 (Det. precios), solo nuevos materiales.
    # Reutilizan las mismas constantes globales que CreadorRINFO (única fuente de verdad).
    _FECHA_INICIO_H1 = FECHA_INICIO_FIJA_HOJA1
    _FECHA_FIN_H1    = FECHA_FIN_FIJA_HOJA1

    def _aplicar_fechas_hoja1(self, ws) -> None:
        """
        Escribe las fechas fijas de disponibilidad en las columnas I (col 9) y
        J (col 10) de la Hoja 1, desde la fila 2 hasta la ultima fila con valor
        en la columna E (Material).

        Los valores son SIEMPRE fijos (01.01.2025 / 31.12.2030), independientes
        de las fechas solicitadas al usuario. Esas fechas van en las cols I/J de
        la Hoja 3 (Det. precios), pero solo para los materiales nuevos.

        Se aplica a TODOS los materiales: tanto los preexistentes del Purchasing
        como los nuevos insertados en este proceso.
        Solo se toca .value — el formato de destino no se modifica.
        """
        ultima_fila = ultima_fila_con_datos(ws, columna=5)

        if ultima_fila < 2:
            _log(self._parent,
                 "  Hoja1 sin datos en col E — fechas I/J no escritas.", "WARN")
            return

        for fila_num in range(2, ultima_fila + 1):
            ws.cell(row=fila_num, column=9).value  = self._FECHA_INICIO_H1  # I  Disponible desde
            ws.cell(row=fila_num, column=10).value = self._FECHA_FIN_H1     # J  Disponible hasta

        _log(self._parent,
             f"  Fechas fijas Hoja1 (col I={self._FECHA_INICIO_H1} / "
             f"J={self._FECHA_FIN_H1}) escritas en filas 2-{ultima_fila}.", "OK")

    # ── REPLICAR FORMATO DE ÚLTIMA FILA EXISTENTE A FILAS NUEVAS (HOJA 1) ──────────────────────────────────

    def _aplicar_formato_plantilla_hoja1(self, ws, primera_fila_nueva: int) -> None:
        """
        Replica el formato de la última fila EXISTENTE (anterior a las nuevas)
        a todas las filas nuevas insertadas en Hoja 1.

        Usa primera_fila_nueva - 1 como referencia (no la fila 2), que es la
        última fila con datos originales del Purchasing. Esto garantiza que el
        formato copiado coincide exactamente con las filas preexistentes
        (azul claro, bordes, fuentes, etc.).

        Parámetros
        ----------
        ws                 : hoja de trabajo (Worksheet) ya abierta.
        primera_fila_nueva : primera fila insertada por este proceso (>= 3).
        """
        nombre_hoja = self.HOJA_PRINCIPAL
        if nombre_hoja in ESTRUCTURA_HOJAS:
            n_cols = len(ESTRUCTURA_HOJAS[nombre_hoja]["cabeceras"])
        else:
            n_cols = ws.max_column or 1

        ultima_fila = ultima_fila_con_datos(ws, columna=5)

        if ultima_fila < primera_fila_nueva:
            _log(self._parent,
                 "  No hay filas nuevas en Hoja1 — formato plantilla no aplicado.", "INFO")
            return

        # Fila de referencia: la última fila ANTERIOR a las nuevas insertadas.
        # Tiene los estilos reales del Purchasing existente.
        fila_ref = max(2, primera_fila_nueva - 1)

        # Leer formato de referencia una sola vez
        formato_ref = []
        for col_idx in range(1, n_cols + 1):
            c = ws.cell(row=fila_ref, column=col_idx)
            formato_ref.append({
                "font":          copy(c.font),
                "fill":          copy(c.fill),
                "border":        copy(c.border),
                "alignment":     copy(c.alignment),
                "number_format": c.number_format,
            })

        # Aplicar a cada fila nueva
        filas_formateadas = 0
        for fila_num in range(primera_fila_nueva, ultima_fila + 1):
            for col_idx, fmt in enumerate(formato_ref, start=1):
                cd = ws.cell(row=fila_num, column=col_idx)
                cd.font          = copy(fmt["font"])
                cd.fill          = copy(fmt["fill"])
                cd.border        = copy(fmt["border"])
                cd.alignment     = copy(fmt["alignment"])
                cd.number_format = fmt["number_format"]
            filas_formateadas += 1

        _log(self._parent,
             f"  Formato replicado desde fila {fila_ref} a "
             f"{filas_formateadas} fila(s) nueva(s) "
             f"(filas {primera_fila_nueva}-{ultima_fila}, {n_cols} columnas).", "OK")


    def _ajustar_fechas_fin_h3(self, wb, primera_fila_nueva: int | None = None) -> None:
        """
        Ajusta la columna J (Fin validez) de la Hoja 3 (Det.precios p.registros info)
        para materiales heredados repetidos por centro, eliminando solapamientos de fechas.

        Regla:
        - Las filas del batch actual (row >= primera_fila_nueva) NUNCA se modifican,
          aunque sí pueden actuar como referencia de «siguiente inicio» para ajustar
          las filas heredadas que las preceden.
        - Para el resto de registros del mismo material+centro, col J se calcula como
          (Inicio validez de la siguiente tarifa) - 1 día.
        - Si col Y (Indicador borrado) = 'X', también se aplica la misma lógica.
        - El objetivo es que no existan solapamientos de fechas entre tarifas del mismo
          material y centro.

        Parámetros:
            primera_fila_nueva : número de fila (1-based) a partir de la cual las filas
                                 pertenecen al batch actual y no deben modificarse.
                                 Si es None, no se excluye ninguna fila por número.

        Columnas (1-based):
            E=5  Centro
            F=6  Material
            I=9  Inicio de validez
            J=10 Fin validez
            Y=25 Indicador borrado
            Z=26 Nueva
        """
        import datetime as _dt_local

        HOJA_3 = "Det.precios p.registros info"
        if HOJA_3 not in wb.sheetnames:
            return

        ws_h3 = wb[HOJA_3]

        # Antes definida aquí mismo (12 líneas); consolidada en sap_helpers.py
        # junto con las otras 4 copias idénticas del proyecto (ver ese módulo).
        _parse_fecha = parsear_fecha_celda

        # Antes definida aquí mismo; consolidada en sap_helpers.py.
        _date_str = formatear_fecha_dmy

        # ── 1. Leer todas las filas de Hoja 3 relevantes ─────────────────────
        #
        #  PRIORIDAD 2 — Clave SAP PIR completa: (Registro_info, Org_compras, Material, Centro)
        #  columnas (A=1, D=4, E=5, F=6) → índices 0-based (0, 3, 4, 5).
        #
        #  Usar solo (material, centro) mezclaba registros de distintos proveedores
        #  u organizaciones, produciendo ajustes de fecha incorrectos o absorciones
        #  cruzadas entre PIRs independientes.
        filas_info = []
        for fila_row in ws_h3.iter_rows(min_row=2):
            if len(fila_row) < 26:
                continue
            val_a   = fila_row[0].value   # A  Registro info
            val_d   = fila_row[3].value   # D  Org. compras
            val_e   = fila_row[4].value   # E  Centro
            val_f   = fila_row[5].value   # F  Material
            val_i   = fila_row[8].value   # I  Inicio validez
            val_j   = fila_row[9].value   # J  Fin validez
            val_y   = fila_row[24].value  # Y  Indicador borrado
            val_z   = fila_row[25].value  # Z  Nueva
            if val_f is None and val_e is None:
                continue
            raw_m = val_f
            raw_m = normalizar_codigo_sap(raw_m)
            clave_pir = (
                str(val_a).strip() if val_a is not None else "",
                str(val_d).strip() if val_d is not None else "",
                str(raw_m).strip() if raw_m is not None else "",
                str(val_e).strip() if val_e is not None else "",
            )
            filas_info.append({
                "row":         fila_row[0].row,
                "clave_pir":   clave_pir,
                "inicio":      _parse_fecha(val_i),
                "fin":         _parse_fecha(val_j),
                "es_nueva":    (str(val_z).strip().upper() == "X") if val_z is not None else False,
                "ind_borrado": (str(val_y).strip().upper() == "X") if val_y is not None else False,
                "cell_j":      fila_row[9],
            })

        if not filas_info:
            return

        # ── 2. Agrupar por clave PIR completa (reg_info, org, material, centro) ──
        from collections import defaultdict
        grupos = defaultdict(list)
        for fi in filas_info:
            grupos[fi["clave_pir"]].append(fi)

        n_ajustes    = 0
        n_absorbidos = 0
        # Agrupadores de log: condensan las líneas repetitivas (una por
        # centro, hasta 10 por material) en una única línea resumen.
        _agr_omitidas   = _LogAgrupado(self._parent)
        _agr_absorbidos = _LogAgrupado(self._parent)
        _agr_fin_ajustado    = _LogAgrupado(self._parent)
        _agr_periodo_invalido = _LogAgrupado(self._parent)
        _agr_parcial = _LogAgrupado(self._parent)
        # Filas a eliminar (absorción total): se acumulan y se borran al final
        # en orden inverso para no desplazar índices durante la iteración.
        filas_a_eliminar_paso4: list[int] = []

        for clave_pir, registros in grupos.items():
            mat    = clave_pir[2]
            centro = clave_pir[3]
            # Solo procesar grupos con más de un registro, o con indicador borrado
            if len(registros) == 1 and not registros[0]["ind_borrado"]:
                continue

            # Ordenar por fecha de inicio (None al final).
            # En caso de empate, los registros nuevos (Z=X) van primero para
            # evitar que un heredado quede delante de una nueva tarifa con la
            # misma fecha y reciba un cierre anterior a su propio inicio.
            registros_ord = sorted(
                registros,
                key=lambda r: (r["inicio"] is None, r["inicio"], not r["es_nueva"])
            )

            # ── Identificar rangos de los períodos del batch actual (Z=X) ────
            # Distinción crítica entre dos tipos de absorción:
            #   · TOTAL   — nuevo_inicio <= viejo_inicio AND nuevo_fin >= viejo_fin
            #               El nuevo ENGLOBA completamente al heredado → ELIMINAR FILA.
            #   · PARCIAL — el nuevo solapa pero no engloba completamente → AJUSTAR FIN.
            rangos_nuevos = [
                (r["inicio"], r["fin"])
                for r in registros_ord
                if r["es_nueva"] and r["inicio"] is not None
            ]

            def _absorbido_total(reg) -> bool:
                """True si algún rango nuevo engloba COMPLETAMENTE al heredado.
                Condición: nuevo_inicio <= viejo_inicio AND nuevo_fin >= viejo_fin.
                Esto garantiza que el viejo no aporta ningún día de cobertura propio.
                """
                if reg["es_nueva"] or reg["inicio"] is None:
                    return False
                for (i_nueva, j_nueva) in rangos_nuevos:
                    # j_nueva=None significa vigencia abierta (engloba cualquier fin)
                    engloba_inicio = i_nueva <= reg["inicio"]
                    engloba_fin   = (j_nueva is None) or (
                        reg["fin"] is None or j_nueva >= reg["fin"]
                    )
                    if engloba_inicio and engloba_fin:
                        return True
                return False

            def _absorbido_parcial(reg) -> bool:
                """True si el inicio del período heredado cae dentro de algún rango nuevo
                pero NO es absorción total (el heredado tiene días fuera del nuevo rango).
                Solo se usa para el encadenado de J en el paso 3.
                """
                if reg["es_nueva"] or reg["inicio"] is None:
                    return False
                for (i_nueva, j_nueva) in rangos_nuevos:
                    if j_nueva is None:
                        if reg["inicio"] >= i_nueva:
                            return True
                    else:
                        if i_nueva <= reg["inicio"] <= j_nueva:
                            return True
                return False

            # ── 3. Para cada registro (excepto el último), ajustar J ─────────
            for i, reg in enumerate(registros_ord[:-1]):
                siguiente = registros_ord[i + 1]
                inicio_siguiente = siguiente["inicio"]
                if inicio_siguiente is None:
                    continue  # sin inicio en la siguiente, no se puede calcular

                fecha_fin_correcta = inicio_siguiente - _dt_local.timedelta(days=1)

                # REGLA CRÍTICA: las filas del batch actual nunca se modifican.
                # Solo pueden actuar como referencia de «siguiente inicio».
                if primera_fila_nueva is not None and reg["row"] >= primera_fila_nueva:
                    _agr_omitidas.registrar(
                        clave=("omitida_batch", mat, clave_pir[0], clave_pir[1]),
                        fila=reg["row"], centro=centro, nivel="INFO",
                        plantilla=("  H3 [{reg_info}|{org}|material {mat}]: "
                                   "omitida (fila nueva batch actual) — {resumen}."),
                        mat=mat, reg_info=clave_pir[0], org=clave_pir[1])
                    continue

                # GUARD: evitar periodo inválido (Inicio > Fin).
                # Ocurre cuando el registro siguiente tiene la misma fecha de
                # inicio que el actual (ej. heredado y nuevo coinciden en I).
                if reg["inicio"] is not None and fecha_fin_correcta < reg["inicio"]:
                    _agr_periodo_invalido.registrar(
                        clave=("periodo_invalido", mat, reg["inicio"], fecha_fin_correcta),
                        fila=reg["row"], centro=centro, nivel="WARN",
                        plantilla=("  H3 material {mat}: periodo inválido evitado — "
                                   "Inicio {inicio} > Fin calculado {fin_calc}. "
                                   "Filas omitidas — {resumen}."),
                        mat=mat, inicio=reg["inicio"], fin_calc=fecha_fin_correcta)
                    continue

                # Aplicar solo si hay diferencia (evita escrituras innecesarias)
                if reg["fin"] != fecha_fin_correcta:
                    reg["cell_j"].value = _date_str(fecha_fin_correcta)
                    n_ajustes += 1
                    _agr_fin_ajustado.registrar(
                        clave=("fin_ajustado", mat, reg["fin"], fecha_fin_correcta, inicio_siguiente),
                        fila=reg["row"], centro=centro, nivel="INFO",
                        plantilla=("  H3 material {mat}: Fin validez ajustado "
                                   "{fin_ant} → {fin_nuevo} (inicio siguiente tarifa: "
                                   "{inicio_sig}) — {resumen}."),
                        mat=mat, fin_ant=reg["fin"], fin_nuevo=fecha_fin_correcta,
                        inicio_sig=inicio_siguiente)

            # ── 4. Gestionar períodos heredados absorbidos por la nueva tarifa ──
            #
            # El bucle anterior procesa hasta el penúltimo registro. Si el último
            # registro heredado (mayor I) cae dentro del rango de una fila nueva,
            # queda solapado sin que ningún paso lo corrija.
            # También pueden existir períodos intermedios absorbidos si la nueva
            # tarifa arranca antes de lo esperado (ej. I_nueva < I_heredado tardío).
            #
            # CASO A — Absorción TOTAL (nuevo engloba completamente al heredado):
            #   nuevo_inicio <= viejo_inicio AND nuevo_fin >= viejo_fin
            #   → ELIMINAR FILA directamente (se acumula para borrado en orden inverso).
            #   El heredado no aporta ningún día propio; eliminarlo es más limpio,
            #   SAP-correct y evita períodos inválidos (Fin < Inicio) en ME13/MEK3/BW.
            #
            # CASO B — Absorción PARCIAL (solapamiento pero no total):
            #   El heredado tiene días fuera del nuevo rango → AJUSTAR FIN (J = I - 1).
            #   Esta rama ya queda cubierta mayoritariamente por el paso 3 anterior;
            #   aquí se deja como salvaguarda para el último registro del grupo.
            for reg in registros_ord:
                if primera_fila_nueva is not None and reg["row"] >= primera_fila_nueva:
                    continue   # nunca tocar filas del batch actual

                if _absorbido_total(reg):
                    # CASO A — eliminar fila (se acumula, se borra al final en reversa)
                    filas_a_eliminar_paso4.append(reg["row"])
                    n_absorbidos += 1
                    _agr_absorbidos.registrar(
                        clave=("absorbido_total", mat, reg["inicio"], reg["fin"]),
                        fila=reg["row"], centro=centro, nivel="WARN",
                        plantilla=("  H3 material {mat}: período heredado TOTALMENTE "
                                   "ABSORBIDO — I={inicio} J={fin} cubierto por nueva "
                                   "tarifa. Filas marcadas para eliminar — {resumen}."),
                        mat=mat, inicio=reg["inicio"], fin=reg["fin"])

                elif _absorbido_parcial(reg):
                    # CASO B — solapamiento parcial: ajustar fin (salvaguarda)
                    j_ajustada = reg["inicio"] - _dt_local.timedelta(days=1)
                    if reg["fin"] != j_ajustada:
                        reg["cell_j"].value = _date_str(j_ajustada)
                        n_absorbidos += 1
                        _agr_parcial.registrar(
                            clave=("parcial", mat, reg["inicio"], reg["fin"], j_ajustada),
                            fila=reg["row"], centro=centro, nivel="WARN",
                            plantilla=("  H3 material {mat}: período heredado "
                                       "PARCIALMENTE solapado — J ajustada {fin_ant} → "
                                       "{fin_nuevo} (I={inicio} cae dentro del rango "
                                       "nuevo) — {resumen}."),
                            mat=mat, fin_ant=reg["fin"], fin_nuevo=j_ajustada, inicio=reg["inicio"])

        # Volcar el último grupo pendiente de cada agrupador (el bucle anterior
        # solo vuelca al cambiar de material; el último grupo se queda en buffer).
        _agr_omitidas.flush()
        _agr_absorbidos.flush()
        _agr_fin_ajustado.flush()
        _agr_periodo_invalido.flush()
        _agr_parcial.flush()

        # ── 5. Reconstruir tramos posteriores perdidos (heredado engloba nueva) ──
        #
        #  Caso "inverso": un período heredado [I_per, J_per] contiene POR
        #  COMPLETO a un rango de la tarifa nueva [I_nueva, J_nueva], es decir:
        #      I_nueva > I_per   Y   J_nueva < J_per   (ambos límites estrictos)
        #
        #  Ejemplo: heredado 01/01/2025-31/12/2025 a 10€, nueva tarifa
        #  01/04/2025-30/06/2025 a 12€ (promoción temporal). Los pasos 3 y 4
        #  anteriores cierran correctamente el tramo INICIAL del heredado
        #  (J_per = I_nueva - 1 → 01/01-31/03 a 10€) y la fila nueva cubre el
        #  tramo central (01/04-30/06 a 12€), pero el tramo FINAL
        #  (01/07-31/12 a 10€) desaparecía sin generar ninguna fila: se perdía
        #  histórico de precio para SAP (pedidos futuros, BW, auditoría).
        #
        #  Aquí se reconstruye copiando íntegra la fila heredada ORIGINAL
        #  (mismo Registro info / misma PIR, mismo precio) y ajustando solo
        #  las fechas I/J al tramo final. La fila se añade al final de la
        #  hoja — igual que el resto de filas nuevas del batch — y queda
        #  correctamente renumerada por _renumerar_col_k_h3, que se ejecuta
        #  después de este paso.
        filas_tail_a_crear: list[tuple] = []   # (fila_origen, f_inicio, f_fin, clave_pir)
        n_tramos_reconstruidos = 0

        for clave_pir, registros in grupos.items():
            if len(registros) == 1 and not registros[0]["ind_borrado"]:
                continue

            rangos_nuevos_grp = [
                (r["inicio"], r["fin"])
                for r in registros
                if r["es_nueva"] and r["inicio"] is not None
            ]
            if not rangos_nuevos_grp:
                continue

            for reg in registros:
                if reg["es_nueva"]:
                    continue
                if reg["inicio"] is None or reg["fin"] is None:
                    continue   # vigencia abierta o sin fechas: no hay tramo final que reconstruir
                if primera_fila_nueva is not None and reg["row"] >= primera_fila_nueva:
                    continue   # nunca tocar filas del batch actual

                for (i_nueva, j_nueva) in rangos_nuevos_grp:
                    if j_nueva is None:
                        continue   # nueva tarifa con vigencia abierta: no deja tramo posterior
                    if i_nueva > reg["inicio"] and j_nueva < reg["fin"]:
                        f_inicio_tail = j_nueva + _dt_local.timedelta(days=1)
                        f_fin_tail    = reg["fin"]
                        filas_tail_a_crear.append(
                            (reg["row"], f_inicio_tail, f_fin_tail, clave_pir)
                        )
                        n_tramos_reconstruidos += 1
                        break   # un tramo posterior por heredado es suficiente

        if filas_tail_a_crear:
            max_col = ws_h3.max_column
            for (fila_origen, f_inicio_tail, f_fin_tail, clave_pir) in filas_tail_a_crear:
                nueva_fila_num = ws_h3.max_row + 1
                for col_idx in range(1, max_col + 1):
                    ws_h3.cell(row=nueva_fila_num, column=col_idx).value = \
                        ws_h3.cell(row=fila_origen, column=col_idx).value
                ws_h3.cell(row=nueva_fila_num, column=9,  value=_date_str(f_inicio_tail))  # I Inicio validez
                ws_h3.cell(row=nueva_fila_num, column=10, value=_date_str(f_fin_tail))     # J Fin validez
                ws_h3.cell(row=nueva_fila_num, column=26, value="X")                        # Z marca NUEVO
                aplicar_formato_fila(ws_h3, nueva_fila_num, HOJA_3, fila_ref=fila_origen)
                _log(self._parent,
                     f"  H3 fila {nueva_fila_num} "
                     f"[{clave_pir[0]}|{clave_pir[1]}|{clave_pir[2]}|{clave_pir[3]}]: "
                     f"tramo POSTERIOR reconstruido — [{_date_str(f_inicio_tail)}"
                     f"→{_date_str(f_fin_tail)}] (la tarifa nueva quedaba englobada "
                     f"por el heredado; copiado desde fila {fila_origen}, mismo precio "
                     f"y Registro info).", "OK")
            _log(self._parent,
                 f"  H3: {n_tramos_reconstruidos} tramo(s) posterior(es) reconstruido(s) "
                 f"para no perder histórico de precio.", "OK")

        if hasattr(self, "_audit"):
            self._audit["n_tramos_reconstruidos"] = n_tramos_reconstruidos

        # ── Eliminar filas absorbidas totalmente (orden inverso) ─────────────
        # CRÍTICO: borrar en orden descendente para que los índices de fila
        # superiores no se desplacen al borrar las inferiores.
        if filas_a_eliminar_paso4:
            _eliminar_filas_agrupadas(ws_h3, filas_a_eliminar_paso4)
            _log(self._parent,
                 f"  H3: {len(filas_a_eliminar_paso4)} fila(s) absorbidas totalmente "
                 f"ELIMINADAS (filas: {sorted(filas_a_eliminar_paso4)}).", "INFO")

        resumen = (
            f"Hoja3 — coherencia fechas: "
            f"{n_ajustes} ajuste(s) de Fin validez (encadenado)"
        )
        if n_absorbidos:
            resumen += (
                f" | {n_absorbidos} período(s) heredado(s) ABSORBIDO(S) "
                f"({len(filas_a_eliminar_paso4)} fila(s) eliminadas, "
                f"{n_absorbidos - len(filas_a_eliminar_paso4)} fin ajustado)."
            )
        else:
            resumen += " | sin períodos absorbidos."
        if n_tramos_reconstruidos:
            resumen += f" | {n_tramos_reconstruidos} tramo(s) posterior(es) reconstruido(s)."

        _log(self._parent, resumen,
             "OK" if (n_ajustes or n_absorbidos or n_tramos_reconstruidos) else "INFO")
        # Paso 10a resuelve los gaps detectados en paso 6c: reportar al audit
        if hasattr(self, "_audit"):
            self._audit["gaps_corregidos"] = n_ajustes
            self._audit["n_absorbidos"]    = n_absorbidos

 # ── ORQUESTADOR PRINCIPAL ─────────────────────────────────────────────────

    def _preparar_archivo_destino(self):
        """
        Pasos 1-3 de _ejecutar: construye el nombre del RINFO destino
        (RINFO <nombre_base_tarifa>.xlsx, con la fecha de la tarifa
        sustituida por la de hoy), copia el Purchasing existente a esa
        ruta, lo abre con openpyxl y valida que contiene la hoja
        principal.

        Devuelve (wb, ws, ruta_nuevo, nombre_nuevo) si todo va bien.

        Devuelve None si algo falla — en ese caso, este método ya se ha
        encargado de mostrar el messagebox de error, loguear, resetear la
        barra de progreso y (si aplica) intentar eliminar el fichero a
        medio copiar/abrir. El llamador (_ejecutar) solo necesita
        comprobar "if resultado is None: return".
        """
        # ── 1. Nombre del archivo destino ─────────────────────────────────────
        # Regla idéntica a CreadorRINFO: RINFO <nombre_base_tarifa>.xlsx
        # El nombre base ya contiene la fecha de la tarifa, e.g.:
        #   "PROVEEDOR X [01-01-2025]"  →  "RINFO PROVEEDOR X [01-01-2025].xlsx"
        # El nombre base de la tarifa contiene la fecha de vigencia de la tarifa.
        # Se reemplaza por la fecha del día en que se genera el RINFO.
        nombre_base_tarifa = os.path.splitext(self._nombre_archivo)[0]
        fecha_hoy          = _dt_mod.date.today().strftime("%d-%m-%Y")
        nombre_base  = re.sub(r'\[\d{2}-\d{2}-\d{4}\]', f'[{fecha_hoy}]',
                               nombre_base_tarifa)
        nombre_nuevo = f"RINFO {nombre_base}.xlsx"
        ruta_nuevo   = os.path.join(self._carpeta_tarifas, nombre_nuevo)

        _log(self._parent,
             f"Generando archivo desde Purchasing: {nombre_nuevo}", "INFO")
        self._barra("set", "Preparando archivo base...", 5, 100)

        # ── 2. Copiar el Purchasing existente al destino ──────────────────────
        ruta_purchasing = self._purchasing["ruta"]
        try:
            shutil.copy2(ruta_purchasing, ruta_nuevo)
        except Exception as e:
            self._parent.after(0, lambda err=str(e): messagebox.showerror(
                "Error al copiar archivo base",
                f"No se pudo copiar el Purchasing Info Records:\n\n{err}",
                parent=self._parent))
            _log(self._parent, f"Error al copiar Purchasing: {e}", "ERROR")
            self._barra("reset", "Error al copiar archivo base")
            _safe_after(self._parent, self._parent.withdraw)
            return None

        # ── 3. Abrir el archivo copiado ───────────────────────────────────────
        init_original = _parchear_column_dimension()
        try:
            wb = openpyxl.load_workbook(ruta_nuevo)
        except Exception as e:
            _restaurar_column_dimension(init_original)
            self._parent.after(0, lambda err=str(e): messagebox.showerror(
                "Error al abrir archivo base",
                f"No se pudo abrir el Purchasing Info Records copiado:\n\n{err}",
                parent=self._parent))
            _log(self._parent, f"Error al abrir Purchasing copiado: {e}", "ERROR")
            self._barra("reset", "Error al abrir archivo base")
            self._intentar_eliminar(ruta_nuevo)
            _safe_after(self._parent, self._parent.withdraw)
            return None
        finally:
            _restaurar_column_dimension(init_original)

        if self.HOJA_PRINCIPAL not in wb.sheetnames:
            self._parent.after(0, lambda: messagebox.showerror(
                "Hoja no encontrada",
                f"La hoja «{self.HOJA_PRINCIPAL}» no existe en el Purchasing base.",
                parent=self._parent))
            _log(self._parent,
                 f"Hoja principal no encontrada en: {self._purchasing['archivo']}",
                 "ERROR")
            self._barra("reset", "Hoja no encontrada en Purchasing")
            wb.close()
            self._intentar_eliminar(ruta_nuevo)
            _safe_after(self._parent, self._parent.withdraw)
            return None

        ws = wb[self.HOJA_PRINCIPAL]
        return wb, ws, ruta_nuevo, nombre_nuevo

    def _construir_indices_h3_heredados(self, ws_h3_act,
                                         heredados_con_match) -> "_EstadoIndicesH3":
        """
        PRIORIDAD 2 (extraída de _actualizar_h3_heredados): recorre Hoja 3
        UNA vez y construye, por clave SAP PIR completa (Registro_info,
        Org_compras, Material, Centro), todos los períodos heredados y el
        más reciente de cada clave. También detecta materiales con tarifa
        activa que nunca tuvieron ninguna fila en Hoja 3 (hueco ya presente
        en el Purchasing original) y los deja preparados en
        estado.absorbidas_con_tarifa, listos para que el paso 9b les genere
        su fila.

        Devuelve el _EstadoIndicesH3 que después mutará
        _aplicar_actualizaciones_h3_heredados() (PRIORIDAD 1).
        """
        _clave_pir_h3 = clave_pir_h3
        _a_date_h3 = parsear_fecha_celda

        estado = _EstadoIndicesH3()

        # ── PRIORIDAD 2: leer Hoja 3 agrupando por clave SAP completa ────────
        #
        #  Clave SAP PIR correcta: (Registro_info, Org_compras, Material, Centro)
        #  que corresponde a las columnas (A=1, D=4, E=5, F=6) de Hoja 3,
        #  es decir índices 0-based (0, 3, 4, 5).
        #
        #  Usar solo (material, centro) era insuficiente: el mismo material puede
        #  tener condiciones para distintos proveedores u organizaciones de compra,
        #  y mezclarlas produce solapamientos fantasma o actualizaciones cruzadas.
        #
        #  estado.precio_heredado_h3, i/estado.j_heredada_h3, estado.fila_mas_reciente_h3 ahora se
        #  indexan por esta clave de 4 campos (tupla de str).

        # Antes definidas aquí mismo; consolidadas en sap_helpers.py
        # (lógica pura, sin dependencias de self/wb/ws).
        _clave_pir_h3 = clave_pir_h3
        _mat_de_clave = material_de_clave_pir

        # ── Punto D: guardar TODOS los períodos heredados por clave PIR ──────
        #
        #  La lógica anterior guardaba solo el período con la I (inicio) más
        #  reciente por clave PIR. Esto dejaba sin detectar solapamientos entre
        #  la nueva tarifa y períodos intermedios o antiguos:
        #
        #    Período 1:  01.01.2023 – 31.12.2023   (antiguo)
        #    Período 2:  01.01.2024 – 31.12.2024   (intermedio)   ← no se comprobaba
        #    Período 3:  01.01.2025 – 31.12.2025   (más reciente) ← único guardado
        #    Nueva tarifa: 15.06.2024 – 31.12.2026
        #      → solapamiento con P2 y P3; P2 quedaba sin cerrar
        #
        #  Ahora se guardan TODOS los períodos en una lista ordenada por I.
        #  La detección de solapamiento itera sobre todos y cierra cada uno
        #  que solape con el nuevo rango [I_nueva, J_nueva].
        #
        #  Estructura:
        #    estado.periodos_h3[clave_pir] = [
        #        {"i": date|None, "j": date|None, "precio": val, "row": int, "cell_j": cell},
        #        ...  (ordenados por i ascendente)
        #    ]
        #    estado.precio_heredado_h3[clave_pir]   = precio del período más reciente
        #    estado.fila_mas_reciente_h3[clave_pir] = fila del período más reciente
        #    estado.i_heredada_h3[clave_pir]        = I del más reciente (compatibilidad)
        #    estado.j_heredada_h3[clave_pir]        = J del más reciente (compatibilidad)

        estado.periodos_h3          = {}   # clave_pir → lista de dicts de período
        estado.precio_heredado_h3   = {}   # clave_pir → precio período más reciente
        estado.i_heredada_h3        = {}   # clave_pir → date | None (más reciente)
        estado.j_heredada_h3        = {}   # clave_pir → date | None (más reciente)
        estado.fila_mas_reciente_h3 = {}   # clave_pir → fila (1-based) del más reciente

        for fila_row in ws_h3_act.iter_rows(min_row=2):
            if len(fila_row) < 15:
                continue
            clave_pir = _clave_pir_h3(fila_row)
            if clave_pir is None or not clave_pir[2]:
                continue
            d_i   = _a_date_h3(fila_row[8].value)   # col I
            d_j   = _a_date_h3(fila_row[9].value)   # col J
            precio = fila_row[14].value              # col O
            fila_n = fila_row[8].row                 # número de fila real
            periodo = {"i": d_i, "j": d_j, "precio": precio,
                       "row": fila_n, "cell_j": fila_row[9]}
            if clave_pir not in estado.periodos_h3:
                estado.periodos_h3[clave_pir] = []
            estado.periodos_h3[clave_pir].append(periodo)
            # Mantener el más reciente para compatibilidad con el resto del código
            if clave_pir not in estado.i_heredada_h3 or (
                    d_i is not None and (
                        estado.i_heredada_h3[clave_pir] is None or
                        d_i > estado.i_heredada_h3[clave_pir])):
                estado.precio_heredado_h3[clave_pir]   = precio
                estado.i_heredada_h3[clave_pir]         = d_i
                estado.j_heredada_h3[clave_pir]         = d_j
                estado.fila_mas_reciente_h3[clave_pir]  = fila_n

        # Ordenar cada clave por I ascendente (None al final)
        for clave_pir in estado.periodos_h3:
            estado.periodos_h3[clave_pir].sort(
                key=lambda p: (p["i"] is None, p["i"]))

        estado.n_solo_fecha         = 0
        estado.n_precio_y_fecha     = 0
        estado.n_solapamiento       = 0   # overlap universal — período cerrado
        estado.n_reduccion          = 0   # precio igual, J_nueva < J_heredada
        estado.n_sin_fecha_heredada = 0   # sin fechas heredadas → actualización directa

        # Filas con absorción TOTAL en paso 6c (Punto D):
        # se acumulan y se eliminan al final en orden inverso.
        estado.filas_a_eliminar_6c: list[int] = []

        # Claves PIR absorbidas totalmente que tienen match en la tarifa.
        # Estructura: {(reg_info, org, material, centro): True}
        # Se usa después del borrado para generar las filas H2/H3 faltantes
        # (un material absorbido que sigue en tarifa necesita su fila nueva).
        estado.absorbidas_con_tarifa: dict = {}

        # ── Materiales con match en tarifa que NUNCA tuvieron ninguna
        #    fila en Hoja 3 (hueco de datos ya presente en el Purchasing
        #    original, no causado por este proceso: el material existe en
        #    Hoja1/Hoja2 pero no tenía ninguna condición de precio previa).
        #    Sin este bloque, esos materiales se quedaban con precio en
        #    Hoja1/Hoja2 pero SIN NINGUNA fila en Hoja3 — es decir, sin la
        #    condición de precio que SAP necesita para aplicarlo.
        #    Se reutiliza el mismo mecanismo de estado.absorbidas_con_tarifa /
        #    _generar_filas_absorbidas_con_tarifa: para esa función da
        #    igual si la fila nueva hace falta porque se absorbió una
        #    existente o porque nunca hubo ninguna — el resultado que
        #    necesita generar es el mismo.
        _materiales_en_periodos_h3 = {clave[2] for clave in estado.periodos_h3}
        _materiales_sin_ninguna_fila_h3 = (
            heredados_con_match - _materiales_en_periodos_h3)
        if _materiales_sin_ninguna_fila_h3:
            _log(self._parent,
                 f"  ⚠ {len(_materiales_sin_ninguna_fila_h3)} material(es) "
                 f"con tarifa activa nunca tuvieron ninguna fila en Hoja 3 "
                 f"(hueco ya presente en el Purchasing original) — "
                 f"generando fila nueva: "
                 f"{', '.join(sorted(_materiales_sin_ninguna_fila_h3))}",
                 "WARN")
            for _mat_sin_h3 in _materiales_sin_ninguna_fila_h3:
                for _cod_centro, _ in self._centros_activos:
                    estado.absorbidas_con_tarifa[("", "IC00", _mat_sin_h3, _cod_centro)] = True

        # ── Materiales con tarifa que SÍ tienen filas en Hoja 3, pero no en
        #    TODOS los centros activos (hueco por CENTRO, no por material).
        #
        #    BUG REAL corregido (2026-09-28, caso real FRAN Y CHEMI SL,
        #    material 10101531 y otros 5): el bloque de arriba razona solo
        #    a nivel de material -- un material con condiciones de precio en
        #    0810 pero ninguna en 0910 no se detectaba (su material sí
        #    aparece en periodos_h3), así que el RINFO salía con la fila de
        #    Hoja 2 de 0910 pero sin su condición de precio en Hoja 3. El
        #    hueco venía ya en el Purchasing original (6 de 268
        #    combinaciones material×centro), pero como Hoja 2 sí añade
        #    filas para todo material de tarifa × centro activo, Hoja 3
        #    debe tener también su precio en cada uno de ellos.
        #    Mismo mecanismo (absorbidas_con_tarifa → paso 9b → paso 10):
        #    no duplica Hoja 2 si ya existe y el paso 10 deduplica por
        #    clave completa, así que es seguro solaparlo con otros orígenes.
        _pares_en_periodos_h3 = {(clave[2], clave[3]) for clave in estado.periodos_h3}
        _huecos_por_centro: dict = {}
        for _mat_c in sorted(heredados_con_match - _materiales_sin_ninguna_fila_h3):
            for _cod_centro, _ in self._centros_activos:
                if (_mat_c, _cod_centro) not in _pares_en_periodos_h3:
                    _huecos_por_centro.setdefault(_mat_c, []).append(_cod_centro)
                    estado.absorbidas_con_tarifa[("", "IC00", _mat_c, _cod_centro)] = True
        if _huecos_por_centro:
            _n_huecos = sum(len(v) for v in _huecos_por_centro.values())
            _detalle_huecos = "; ".join(
                f"{m} ({','.join(cs)})" for m, cs in _huecos_por_centro.items())
            _log(self._parent,
                 f"  ⚠ {_n_huecos} combinación(es) material×centro con tarifa "
                 f"activa sin ninguna fila en Hoja 3 (hueco ya presente en el "
                 f"Purchasing original: el material tiene precio en otros "
                 f"centros pero no en éstos) — generando fila nueva: "
                 f"{_detalle_huecos}", "WARN")

        return estado


    def _aplicar_actualizaciones_h3_heredados(
            self, wb, ws, ws_h3_act, estado, heredados_con_match,
            dic_tarifa, d_inicio_nueva, d_fin_nueva) -> dict:
        """
        PRIORIDAD 1 (extraída de _actualizar_h3_heredados): recorre Hoja 3
        una SEGUNDA vez y aplica la actualización de fechas/precio a cada
        fila heredada con match en la tarifa nueva, usando los índices ya
        construidos en `estado` por _construir_indices_h3_heredados()
        (PRIORIDAD 2). Muta `estado` en el proceso (contadores, filas a
        eliminar, dicts de warnings, y estado.absorbidas_con_tarifa) y,
        al terminar, elimina físicamente las filas absorbidas totalmente y
        vuelca los WARNs agrupados en el log -- idéntico comportamiento al
        bloque inline original, solo que ahora opera sobre `estado` en vez
        de variables sueltas.

        Devuelve estado.absorbidas_con_tarifa (dict) para el paso 9b --
        mismo contrato de salida que ya exponía _actualizar_h3_heredados().
        """
        # Antes definidas aquí mismo; consolidadas en sap_helpers.py.
        # (necesarias también aquí: esta fase vuelve a recorrer Hoja 3 y
        #  reconstruye la misma clave PIR por fila, y usa _mat_de_clave al
        #  consolidar self._audit al final del proceso).
        _clave_pir_h3 = clave_pir_h3
        _mat_de_clave = material_de_clave_pir
        _date_a_str_h3 = formatear_fecha_dmy
        _fmt_clave = formatear_clave_pir

        # ── Helper: marcar NUEVO=X solo ante cambio REAL respecto al ──────
        #  Purchasing original (regla: NUEVO se reserva a SAP procesando
        #  algo distinto). Se marca ÚNICAMENTE en:
        #    · Hoja 3 col Z  (26) — fila concreta que se actualiza in-place.
        #
        #  IMPORTANTE — Hoja 1 col AK (37) NUNCA se marca aquí. AK=X en
        #  Hoja 1 está reservado EXCLUSIVAMENTE para materiales que NO
        #  existían en el Purchasing original (ver _rellenar_h1_nuevas_filas).
        #  Un material heredado puede tener cambios reales de precio/fecha
        #  en Hoja 3 sin que eso implique marcar su fila de Hoja 1 como
        #  "Nueva" — la fila de Hoja 1 ya existía y sigue existiendo.
        #
        #  EXCEPCIÓN: si el material está marcado para baja (AJ=X en Hoja1,
        #  detectado en el paso 6b ANTERIOR a este bloque), NUEVO se deja
        #  vacío — regla 4: BORRAR y NUEVO son mutuamente excluyentes.
        #  El propio material queda registrado en materiales_modificados
        #  para el resumen final, independientemente de si hay baja.
        COL_AJ_H1_0BASED = 35   # AJ Hoja1 (1-based 36) — marca de baja
        COL_Z_H3         = 26  # Z  Hoja3 (1-based)     — marca NUEVO

        def _material_en_baja(clave_m: str) -> bool:
            fila_h1 = self._fila_h1_por_codigo.get(clave_m)
            if fila_h1 is None:
                return False
            val_aj = ws.cell(row=fila_h1, column=COL_AJ_H1_0BASED + 1).value
            return str(val_aj).strip().upper() == "X" if val_aj is not None else False

        def _marcar_modificado(clave_m: str, fila_h3_num: int) -> None:
            self._audit["materiales_modificados"].add(clave_m)
            if _material_en_baja(clave_m):
                return   # regla 4 — BORRAR excluye NUEVO
            ws_h3_act.cell(row=fila_h3_num, column=COL_Z_H3).value = "X"
        # ── PRIORIDAD 1: overlap universal ────────────────────────────────
        #
        #  Lógica anterior (3 casos A/B/C):
        #    CASO A  I_nueva <= I_her                 → NO tocar (dejaba registro sin resolver)
        #    CASO B  I_her < I_nueva <= J_her         → cerrar J_her = I_nueva - 1
        #    CASO C  I_nueva > J_her                  → actualización directa
        #
        #  Lógica nueva — overlap universal:
        #    Hay solapamiento si los dos intervalos se superponen en cualquier punto:
        #      [I_her, J_her] ∩ [I_nueva, J_nueva] ≠ ∅
        #    Condición equivalente (negada = sin solapamiento):
        #      J_her < I_nueva  OR  J_nueva < I_her
        #    Si hay solapamiento → cerrar J_her = I_nueva - 1 (sea cual sea la relación).
        #    Esto cubre tanto el antiguo CASO A como el B, sin dejar registros "sin tocar".
        #
        #  Precio igual vs. distinto solo determina si se actualiza col O (precio):
        #    · Precio igual   → solo J (extensión/reducción de vigencia)
        #    · Precio distinto → I + J + O (nuevo período)

        # Guard de deduplicación para la rama PRECIO DISTINTO (Punto D):
        # garantiza que cada clave PIR se procese una sola vez aunque tenga
        # varias filas físicas heredadas en Hoja 3 (varios períodos históricos).
        # Sin este guard, el bloque "Punto D" se ejecutaría una vez por cada
        # fila física que comparta la clave, duplicando entradas en
        # estado.filas_a_eliminar_6c y corrompiendo el borrado posterior (delete_rows
        # con el mismo número de fila más de una vez borra filas distintas
        # tras el primer desplazamiento de índices).
        _claves_precio_distinto_procesadas: set = set()

        # Agrupadores de log: condensan las líneas repetitivas (una por
        # centro, hasta 10 por material) en una única línea resumen.
        _agr_j_omitida    = _LogAgrupado(self._parent)
        _agr_absorcion_6c = _LogAgrupado(self._parent)

        # ── Estado heredado incoherente: periodos solapados con otro precio ──
        #
        # BUG REAL corregido (2026-10-06, caso real FRAN Y CHEMI SL /
        # ESCAROLA 10101544): la decisión "MISMO PRECIO" comparaba la
        # tarifa SOLO con el precio del período más reciente (el de mayor
        # inicio). El Purchasing de ese día traía la escarola con dos filas
        # de octubre incoherentes -- 31.10→31.10 @1,80 y 01.10→30.10 @1,50
        # (restos de varias cargas repetidas del mismo RINFO que SAP había
        # ido recortando por su cuenta) -- y la tarifa decía 1,80 para todo
        # octubre. El "más reciente" por inicio era la de 1 día @1,80, igual
        # que la tarifa, así que se daba por buen precio y no se miraba nada
        # más: el 1,50 siguió cubriendo 30 de los 31 días sin ningún aviso.
        # Ahora "mismo precio" exige además que NINGÚN período heredado que
        # solape con [I_nueva, J_nueva] tenga un precio distinto; si lo hay,
        # se trata como cambio de precio (los períodos solapados quedan
        # absorbidos y se genera la fila correcta). En datos coherentes
        # (sin solapes con otro precio) el resultado es idéntico al de antes.
        _claves_incoherentes: set = set()

        def _precios_iguales(a, b) -> bool:
            try:
                return abs(float(a) - float(b)) < 0.005
            except (TypeError, ValueError):
                return a == b

        def _solapa_con_precio_distinto(clave_pir, precio_nuevo) -> bool:
            if precio_nuevo is None or d_inicio_nueva is None:
                return False
            for per in estado.periodos_h3.get(clave_pir, []):
                d_i_per, d_j_per = per["i"], per["j"]
                if d_i_per is None:
                    continue
                sin_solap = (
                    (d_j_per is not None and d_j_per < d_inicio_nueva) or
                    (d_fin_nueva is not None and d_fin_nueva < d_i_per))
                if sin_solap:
                    continue
                if not _precios_iguales(per["precio"], precio_nuevo):
                    _claves_incoherentes.add(clave_pir)
                    return True
            return False

        # Recorrer Hoja 3 fila a fila y aplicar la actualización
        for fila_row in ws_h3_act.iter_rows(min_row=2):
            if len(fila_row) < 15:
                continue
            clave_pir = _clave_pir_h3(fila_row)
            if clave_pir is None or not clave_pir[2]:
                continue
            clave_m = clave_pir[2]   # solo el campo Material para cruce con tarifa

            if clave_m not in heredados_con_match:
                continue

            precio_nuevo    = dic_tarifa.get(clave_m)
            precio_heredado = estado.precio_heredado_h3.get(clave_pir)
            fila_num        = fila_row[5].row

            if (precio_nuevo == precio_heredado
                    and not _solapa_con_precio_distinto(clave_pir, precio_nuevo)):
                # ── MISMO PRECIO: solo actualizar Fin validez (col J) ─────
                # Actualizar únicamente la fila más reciente (mayor I) de esta clave.
                fila_mas_rec = estado.fila_mas_reciente_h3.get(clave_pir)
                if fila_mas_rec is not None and fila_num != fila_mas_rec:
                    continue   # fila histórica — no tocar
                # GUARD: evitar periodo inválido (Inicio > nuevo Fin).
                # Ocurre cuando self._fecha_fin cae antes del Inicio de la fila
                # más reciente (ej. tarifa anterior con fecha_fin < I_heredado).
                d_i_mas_rec = estado.i_heredada_h3.get(clave_pir)
                if d_fin_nueva is not None and d_i_mas_rec is not None                             and d_fin_nueva < d_i_mas_rec:
                    _agr_j_omitida.registrar(
                        clave=("j_omitida_6c", clave_m, d_i_mas_rec),
                        fila=fila_num, centro=clave_pir[3], nivel="WARN",
                        plantilla=("  H3 material {mat}: actualización J omitida "
                                   "en paso 6c (mismo precio) — fecha_fin={fecha_fin} "
                                   "< I_heredado={i_heredado}. Periodo ya cubierto "
                                   "por tarifa posterior — {resumen}."),
                        mat=clave_m, fecha_fin=self._fecha_fin, i_heredado=d_i_mas_rec)
                    self._audit["materiales_sin_cambios"].add(clave_m)
                    continue
                ws_h3_act.cell(row=fila_num, column=10, value=self._fecha_fin)
                estado.n_solo_fecha += 1
                # Cambio real solo si la fecha fin difiere de la heredada.
                d_j_actual_mat = estado.j_heredada_h3.get(clave_pir)
                if d_j_actual_mat != d_fin_nueva:
                    _marcar_modificado(clave_m, fila_num)
                else:
                    self._audit["materiales_sin_cambios"].add(clave_m)
                d_j_her_mat = estado.j_heredada_h3.get(clave_pir)
                if d_fin_nueva is not None and d_j_her_mat is not None                             and d_fin_nueva < d_j_her_mat:
                    estado.n_reduccion += 1
                    if clave_pir not in estado.warn_reduccion:
                        estado.warn_reduccion[clave_pir] = (self._fecha_fin, d_j_her_mat)

            else:
                # ── GUARD: procesar cada clave PIR una sola vez ───────────
                #  El bucle externo recorre FILAS físicas de Hoja 3, pero una
                #  misma clave PIR (proveedor|org|material|centro) puede tener
                #  varias filas heredadas (varios períodos históricos). Sin
                #  este guard, el bloque de abajo (que ya opera sobre TODOS
                #  los períodos de la clave vía estado.periodos_h3[clave_pir]) se
                #  repetiría una vez por cada fila física que comparta la
                #  clave, duplicando entradas en estado.filas_a_eliminar_6c y
                #  corrompiendo el borrado posterior (delete_rows con el
                #  mismo número de fila más de una vez borra filas distintas
                #  tras el primer desplazamiento de índices).
                if clave_pir in _claves_precio_distinto_procesadas:
                    continue
                _claves_precio_distinto_procesadas.add(clave_pir)

                # ── PRECIO DISTINTO: overlap universal ────────────────────
                #
                #  Independientemente de si I_nueva cae antes, dentro o después
                #  del período heredado, la acción correcta es siempre la misma:
                #
                #  · Si hay solapamiento ([I_her,J_her] ∩ [I_nueva,J_nueva] ≠ ∅):
                #      J_her = I_nueva - 1  (cerrar el período anterior)
                #      La fila nueva con el nuevo precio se añadirá al final (paso 10).
                #
                #  · Si no hay solapamiento (J_her < I_nueva):
                #      Actualización directa de I, J y precio en la fila heredada.
                #      Se comprueba gap (días sin cobertura de precio).
                #
                #  · Si no hay fechas heredadas: actualización directa sin aviso.
                #
                #  Esta lógica unificada elimina el antiguo CASO A (solapamiento total
                #  "sin tocar") que dejaba registros con precio vigente solapado en SAP,
                #  causando errores en la carga masiva o condiciones duplicadas.

                d_i_her = estado.i_heredada_h3.get(clave_pir)
                d_j_her = estado.j_heredada_h3.get(clave_pir)

                if d_inicio_nueva is None:
                    # Sin fecha nueva válida: actualización directa (caso raro).
                    # Se aplica a TODAS las filas físicas de la clave (no solo
                    # fila_num), ya que el guard de arriba garantiza una sola
                    # pasada por clave_pir.
                    for per in estado.periodos_h3.get(clave_pir, [{"row": fila_num}]):
                        _fila_per = per["row"]
                        ws_h3_act.cell(row=_fila_per, column=9,  value=self._fecha_inicio)
                        ws_h3_act.cell(row=_fila_per, column=10, value=self._fecha_fin)
                        if precio_nuevo is not None:
                            ws_h3_act.cell(row=_fila_per, column=15, value=precio_nuevo)
                        _marcar_modificado(clave_m, _fila_per)
                    estado.n_precio_y_fecha += 1
                    if clave_pir not in estado.warn_sin_parsear:
                        estado.warn_sin_parsear[clave_pir] = (precio_heredado, precio_nuevo)

                elif d_i_her is None and d_j_her is None:
                    # Sin fechas heredadas — actualización directa.
                    # Igual que arriba: se aplica a todas las filas físicas
                    # de la clave, no solo a la que disparó la primera visita.
                    for per in estado.periodos_h3.get(clave_pir, [{"row": fila_num}]):
                        _fila_per = per["row"]
                        ws_h3_act.cell(row=_fila_per, column=9,  value=self._fecha_inicio)
                        ws_h3_act.cell(row=_fila_per, column=10, value=self._fecha_fin)
                        if precio_nuevo is not None:
                            ws_h3_act.cell(row=_fila_per, column=15, value=precio_nuevo)
                        _marcar_modificado(clave_m, _fila_per)
                    estado.n_sin_fecha_heredada += 1
                    estado.n_precio_y_fecha += 1

                else:
                    # ── Punto D: solapamiento contra TODOS los períodos ───
                    #
                    #  La lógica anterior solo comparaba contra el período MÁS
                    #  RECIENTE de la clave PIR, dejando solapamientos con
                    #  períodos intermedios sin detectar. Tipos cubiertos ahora:
                    #    · Parcial:   I_nueva cae dentro del período heredado
                    #    · Total:     nuevo rango engloba al heredado
                    #    · Inverso:   heredado engloba al nuevo rango
                    #    · Múltiple:  varios períodos de la misma clave PIR solapan
                    #    · Mismo mat+centro: garantizado por clave PIR completa
                    #
                    #  Para cada período de estado.periodos_h3[clave_pir] (ordenados I asc):
                    #    Solapa ⟺ NOT (J_per < I_nueva  OR  J_nueva < I_per)
                    #    Si solapa y I_nueva > I_per  → J_per = I_nueva - 1
                    #    Si solapa y I_nueva ≤ I_per  → J_per = I_per  - 1 (absorbido)
                    #  Si ningún período solapa → actualización directa (caso sin solap).

                    periodos_clave   = estado.periodos_h3.get(clave_pir, [])
                    n_solap_este_pir = 0
                    _alguno_sin_solap = False

                    for per in periodos_clave:
                        d_i_per = per["i"]
                        d_j_per = per["j"]
                        fila_per = per["row"]

                        if d_i_per is None:
                            # Sin inicio → actualizar directamente este período
                            ws_h3_act.cell(row=fila_per, column=9,  value=self._fecha_inicio)
                            ws_h3_act.cell(row=fila_per, column=10, value=self._fecha_fin)
                            if precio_nuevo is not None:
                                ws_h3_act.cell(row=fila_per, column=15, value=precio_nuevo)
                            estado.n_sin_fecha_heredada += 1
                            estado.n_precio_y_fecha += 1
                            _marcar_modificado(clave_m, fila_per)
                            continue

                        # Overlap universal
                        j_per_inf = d_j_per     is None
                        j_nva_inf = d_fin_nueva is None
                        sin_solap_per = (
                            (not j_per_inf and d_j_per  < d_inicio_nueva) or
                            (not j_nva_inf and d_fin_nueva < d_i_per)
                        )

                        if not sin_solap_per:
                            # ── Solapamiento detectado ───────────────────
                            #
                            # Distinción fundamental entre dos situaciones:
                            #
                            # CASO A — Absorción TOTAL:
                            #   d_inicio_nueva <= d_i_per  AND
                            #   (d_fin_nueva is None OR d_fin_nueva >= d_j_per_real)
                            #   El nuevo período ENGLOBA completamente al heredado.
                            #   → MARCAR fila para eliminar directamente.
                            #   Ventajas: no genera períodos inválidos (Fin < Inicio),
                            #   SAP queda limpio en ME13/MEK3/BW/ETL, sin ruido histórico.
                            #
                            # CASO B — Solapamiento PARCIAL:
                            #   El heredado tiene días fuera del nuevo rango.
                            #   Subcasos:
                            #     · I_nueva > I_per: cerrar J_per = I_nueva - 1
                            #     · I_nueva == I_per pero J_per > J_nueva: ya cubierto
                            #       por la condición de absorción total (no entra aquí)
                            #   → AJUSTAR FIN (comportamiento anterior conservado).

                            # Evaluar absorción total
                            j_per_real = d_j_per  # puede ser None (vigencia abierta)
                            absorbido_total = (
                                d_inicio_nueva <= d_i_per
                                and (
                                    d_fin_nueva is None
                                    or j_per_real is None
                                    or d_fin_nueva >= j_per_real
                                )
                            )

                            if absorbido_total:
                                # CASO A — eliminar fila (acumular para borrado en reversa)
                                estado.filas_a_eliminar_6c.append(fila_per)
                                estado.n_solapamiento   += 1
                                n_solap_este_pir += 1
                                # Registrar si el material sigue en la tarifa activa:
                                # necesitará una fila nueva en H2/H3 tras el borrado.
                                if clave_pir[2] in heredados_con_match:
                                    estado.absorbidas_con_tarifa[clave_pir] = True
                                _agr_absorcion_6c.registrar(
                                    clave=("absorcion_6c", clave_pir[2],
                                           d_i_per, j_per_real,
                                           d_inicio_nueva, d_fin_nueva),
                                    fila=fila_per, centro=clave_pir[3], nivel="WARN",
                                    plantilla=(
                                        "  H3 material {mat}: absorción TOTAL — nuevo "
                                        "[{ini_nueva}→{fin_nueva}] engloba "
                                        "[{ini_per}→{fin_per}]. Filas marcadas para "
                                        "eliminar — {resumen}."),
                                    mat=clave_pir[2],
                                    ini_nueva=_date_a_str_h3(d_inicio_nueva),
                                    fin_nueva=_date_a_str_h3(d_fin_nueva) if d_fin_nueva else "∞",
                                    ini_per=_date_a_str_h3(d_i_per),
                                    fin_per=_date_a_str_h3(j_per_real) if j_per_real else "∞")
                                if clave_pir not in estado.warn_solap:
                                    estado.warn_solap[clave_pir] = (
                                        f"absorbida total — fila {fila_per} eliminada"
                                    )
                            else:
                                # CASO B — solapamiento parcial: ajustar fin
                                if d_inicio_nueva > d_i_per:
                                    # I_nueva posterior: cerrar J al día anterior
                                    d_cierre = d_inicio_nueva - _dt_mod.timedelta(days=1)
                                else:
                                    # I_nueva igual o anterior pero J_per extiende más allá
                                    # del rango nuevo (parcial): cerrar J_per = I_per - 1
                                    d_cierre = d_i_per - _dt_mod.timedelta(days=1)
                                # GUARD: evitar periodo inválido (Inicio > Fin calculado).
                                if d_i_per is not None and d_cierre < d_i_per:
                                    _log(self._parent,
                                         f"  H3 fila {fila_per} [{_fmt_clave(clave_pir)}]: "
                                         f"cierre inválido evitado en paso 6c — "
                                         f"I_per={d_i_per} > cierre calculado={d_cierre}. "
                                         f"Fila omitida.", "WARN")
                                    _alguno_sin_solap = True
                                    continue
                                per["cell_j"].value = _date_a_str_h3(d_cierre)
                                estado.n_solapamiento    += 1
                                n_solap_este_pir  += 1
                                _marcar_modificado(clave_m, fila_per)
                                # BUG REAL corregido (2026-09-28, caso real MAHOU
                                # / DALI 29612, SAP 10200147): al acortar el
                                # período heredado (J = I_nueva - 1) éste SIGUE
                                # existiendo con su precio viejo -- el período
                                # nuevo (I_nueva → J_nueva, precio nuevo) tiene
                                # que crearse aparte. Antes solo se registraba la
                                # clave en absorbidas_con_tarifa para la
                                # absorción TOTAL (CASO A), así que en el
                                # solapamiento PARCIAL nadie generaba la fila
                                # nueva: el material se quedaba con el precio
                                # viejo y sin cobertura desde I_nueva, sin ningún
                                # error visible (130 registros en ese caso real,
                                # log "0 precio+fechas | 130 solapam."). El aviso
                                # "SAP recibirá la fila nueva en la siguiente
                                # importación" prometía una fila que nadie
                                # creaba. _generar_filas_absorbidas_con_tarifa
                                # (paso 9b) ya sirve tal cual: no duplica la fila
                                # de H2 si existe y devuelve la tupla para que el
                                # paso 10 escriba la de H3.
                                if clave_pir[2] in heredados_con_match:
                                    estado.absorbidas_con_tarifa[clave_pir] = True
                                if clave_pir not in estado.warn_solap:
                                    estado.warn_solap[clave_pir] = _date_a_str_h3(d_cierre)
                        else:
                            _alguno_sin_solap = True

                    # Si ningún período heredado solapó → actualizar I, J, precio
                    # en la fila más reciente de la clave (la que SAP extendería).
                    #
                    # BUG REAL corregido (2026-09-23, caso real FRAN Y CHEMI SL /
                    # CALABACIN 10101584): antes se exigía además `fila_num ==
                    # fila_mas_rec`, pero fila_num es la PRIMERA fila física de
                    # la clave que el bucle exterior encontró (por el guard de
                    # "procesar cada clave una sola vez"), no necesariamente la
                    # más reciente -- si el material tiene 2+ períodos
                    # históricos en H3 (habitual: uno antiguo cerrado + otro
                    # vigente), la fila más reciente casi nunca es la primera
                    # físicamente, así que esta condición era prácticamente
                    # siempre falsa y la actualización JAMÁS se aplicaba: el
                    # material se quedaba sin precio/período para la tarifa
                    # nueva, sin ningún error visible. Se escribe directamente
                    # en fila_mas_rec, sin depender de en qué fila física
                    # estuviera el bucle exterior en ese momento.
                    if n_solap_este_pir == 0:
                        fila_mas_rec = estado.fila_mas_reciente_h3.get(clave_pir)
                        if fila_mas_rec is not None:
                            ws_h3_act.cell(row=fila_mas_rec, column=9,  value=self._fecha_inicio)
                            ws_h3_act.cell(row=fila_mas_rec, column=10, value=self._fecha_fin)
                            if precio_nuevo is not None:
                                ws_h3_act.cell(row=fila_mas_rec, column=15, value=precio_nuevo)
                            estado.n_precio_y_fecha += 1
                            _marcar_modificado(clave_m, fila_mas_rec)

                        # Comprobar gap (días sin cobertura de precio) -- SOLO
                        # contra el período más reciente (d_j_her), no contra
                        # cada período histórico: comparar contra uno antiguo
                        # ya cerrado hace tiempo genera un "gap" enorme que no
                        # refleja ningún problema real (mismo caso real de
                        # arriba: el período de nov-2025 disparaba un falso
                        # aviso de 274 días, cuando el período vigente de
                        # ago-2026 conectaba sin ningún hueco).
                        if d_j_her is not None:
                            gap_dias = (d_inicio_nueva - d_j_her).days - 1
                            if gap_dias > 0:
                                estado.warn_gap[clave_pir] = gap_dias

        # Volcar el último grupo pendiente de cada agrupador (el bucle
        # anterior solo vuelca al cambiar de material/período; el último
        # grupo se queda en buffer hasta que se fuerza aquí).
        _agr_j_omitida.flush()
        _agr_absorcion_6c.flush()

        if _claves_incoherentes:
            _mats_inc = sorted({c[2] for c in _claves_incoherentes})
            _lista_inc = ", ".join(_mats_inc[:15]) + (
                f" … y {len(_mats_inc) - 15} más" if len(_mats_inc) > 15 else "")
            _log(self._parent,
                 f"  ⚠ {len(_mats_inc)} material(es) con períodos heredados "
                 f"SOLAPADOS en Hoja 3 con un precio distinto al de la tarifa "
                 f"(estado incoherente en el Purchasing, p. ej. por cargas "
                 f"repetidas del mismo RINFO): {_lista_inc}. Se han tratado "
                 f"como cambio de precio -- los períodos solapados se "
                 f"sustituyen por uno nuevo con el precio de la tarifa. "
                 f"Conviene revisar esos registros en SAP.", "WARN")

        # ── Eliminar filas absorbidas totalmente (Punto D, orden inverso) ──
        # CRÍTICO: borrar en orden descendente para que los índices de fila
        # superiores no se desplacen al borrar las inferiores.
        if estado.filas_a_eliminar_6c:
            _eliminar_filas_agrupadas(ws_h3_act, estado.filas_a_eliminar_6c)
            _log(self._parent,
                 f"  H3: {len(estado.filas_a_eliminar_6c)} fila(s) absorbidas totalmente "
                 f"ELIMINADAS en paso 6c "
                 f"(filas: {sorted(estado.filas_a_eliminar_6c)}).", "INFO")

        # Las claves PIR absorbidas totalmente representan un cambio real
        # (el período heredado deja de existir; SAP recibirá un período
        # nuevo en su lugar — ver estado.absorbidas_con_tarifa / paso 10).
        for _clave_abs in estado.absorbidas_con_tarifa:
            self._audit["materiales_modificados"].add(_mat_de_clave(_clave_abs))

        # ── Consolidar materiales_sin_cambios vs materiales_modificados ───
        # Un material puede tener varias filas (centros/períodos distintos);
        # si CUALQUIERA de ellas tuvo un cambio real, el material es
        # "modificado" aunque otra fila suya quedara sin tocar.
        self._audit["materiales_sin_cambios"] -= self._audit["materiales_modificados"]

        # ── Emitir WARNs agrupados por clave PIR ─────────────────────────

        if estado.warn_reduccion:
            detalle = "; ".join(
                f"{_fmt_clave(k)} (J_nueva={jn} < J_ant={jh})"
                for k, (jn, jh) in estado.warn_reduccion.items()
            )
            _log(self._parent,
                 f"  ⚠ REDUCCIÓN VIGENCIA — {len(estado.warn_reduccion)} registro(s): {detalle}",
                 "WARN")
        if estado.warn_solap:
            detalle = "; ".join(
                f"{_fmt_clave(k)} ({cierre})"
                for k, cierre in estado.warn_solap.items()
            )
            n_total_elim   = len(estado.filas_a_eliminar_6c)
            n_parcial_ajus = len(estado.warn_solap) - sum(
                1 for v in estado.warn_solap.values()
                if isinstance(v, str) and v.startswith("absorbida total")
            )
            _log(self._parent,
                 f"  ⚠ SOLAPAMIENTO — {len(estado.warn_solap)} registro(s): "
                 f"{n_total_elim} absorbidos totalmente (fila eliminada), "
                 f"{n_parcial_ajus} parciales (J ajustada). "
                 f"Detalle: {detalle}",
                 "WARN")
        if estado.warn_gap:
            detalle = "; ".join(
                f"{_fmt_clave(k)} ({dias}d)"
                for k, dias in sorted(estado.warn_gap.items(), key=lambda x: -x[1])
            )
            _log(self._parent,
                 f"  ⚠ GAP SIN COBERTURA — {len(estado.warn_gap)} registro(s): {detalle}",
                 "WARN")
        if estado.warn_sin_parsear:
            detalle = ", ".join(
                f"{_fmt_clave(k)} ({ph}→{pn})"
                for k, (ph, pn) in estado.warn_sin_parsear.items()
            )
            _log(self._parent,
                 f"  ⚠ FECHA INICIO NO PARSEABLE — {len(estado.warn_sin_parsear)} registro(s), "
                 f"actualización directa: {detalle}",
                 "WARN")

        n_avisos_total = (estado.n_solapamiento + estado.n_reduccion
                          + len(estado.warn_gap) + len(estado.warn_sin_parsear))
        _n_solap_elim   = len(estado.filas_a_eliminar_6c)
        _n_solap_ajuste = estado.n_solapamiento - _n_solap_elim
        resumen_log = (
            f"Hoja3 heredados OK — "
            f"{estado.n_solo_fecha} solo fin-validez | "
            f"{estado.n_precio_y_fecha} precio+fechas"
            + (f" ({estado.n_sin_fecha_heredada} sin fechas heredadas)"
               if estado.n_sin_fecha_heredada else "")
            + (f" | ⚠ {estado.n_solapamiento} solapam."
               f" ({_n_solap_elim} fila(s) eliminadas absorción total"
               f", {_n_solap_ajuste} J ajustada)"
               if estado.n_solapamiento else "")
            + (f" | ⚠ {estado.n_reduccion} reducción(es) vigencia"
               if estado.n_reduccion else "")
            + (f" | ⚠ {len(estado.warn_gap)} gap(s) sin cobertura"
               if estado.warn_gap else "")
            + (f" | ⚠ {len(estado.warn_sin_parsear)} fecha inicio no parseable"
               if estado.warn_sin_parsear else "")
            + "."
        )
        _log(self._parent, resumen_log, "WARN" if n_avisos_total else "OK")
        self._audit["n_solapamientos"] += estado.n_solapamiento
        self._audit["n_absorbidos"]    = (
            self._audit.get("n_absorbidos", 0) + _n_solap_elim
        )  # filas eliminadas por absorción total en paso 6c
        self._audit["n_gaps"]          += len(estado.warn_gap)
        self._audit["n_reducciones"]   += estado.n_reduccion
        self._audit["gaps_detalle"].update(estado.warn_gap)
        self._audit["solap_detalle"].update(estado.warn_solap)

        return estado.absorbidas_con_tarifa


    def _actualizar_h3_heredados(self, wb, ws, codigos_existentes,
                                  codigos_tarifa, dic_tarifa) -> dict:
        """Paso 6c (extraido de _ejecutar): actualiza fechas y precio en
        Hoja 3 para materiales heredados con match en la tarifa nueva.
        Recibe wb/ws y los diccionarios ya construidos en los pasos 4-6
        de _ejecutar. Devuelve _absorbidas_con_tarifa (dict) para el
        paso 9b -- idéntico contrato al que ya exponía el bloque inline.

        Orquestador puro (rediseño 2026-08, ver HISTORIAL 'Roadmap punto
        3'): delega en _construir_indices_h3_heredados() (PRIORIDAD 2) y
        _aplicar_actualizaciones_h3_heredados() (PRIORIDAD 1), que se
        pasan un único objeto _EstadoIndicesH3 en vez de ~15 variables
        sueltas -- el mismo bucle de 619 líneas que antes no se podía
        trocear con seguridad.
        """
        # ── 6c. Actualizar fechas (y precio si cambia) en Hoja 3 ──────────────
        #
        #  Para cada material que EXISTE en el Purchasing Y aparece en la nueva
        #  tarifa (heredado con continuidad) se actualizan sus filas en Hoja 3:
        #
        #    · Precio igual al heredado  → solo col J (Fin validez) = fecha_fin
        #                                  La col I (Inicio validez) se conserva.
        #    · Precio distinto           → col I = fecha_inicio, col J = fecha_fin
        #                                  col O (precio) = nuevo precio
        #
        #  Se procesan TODAS las filas del código en Hoja 3 (una por centro).
        #  La comparación de precios es exacta, sin redondeo.
        self._barra("set", "Actualizando fechas/precios en Hoja 3 (heredados)...", 56, 100)

        if self.HOJA_3 not in wb.sheetnames:
            _log(self._parent,
                 f"Hoja «{self.HOJA_3}» no disponible — actualización heredados omitida.", "WARN")
            return {}

        ws_h3_act = wb[self.HOJA_3]

        # Conjunto de materiales heredados que continúan en la nueva tarifa
        codigos_tarifa_set_6c = {c for c, _ in codigos_tarifa}
        heredados_con_match   = codigos_existentes & codigos_tarifa_set_6c

        _a_date_h3 = parsear_fecha_celda
        d_inicio_nueva = _a_date_h3(self._fecha_inicio)
        d_fin_nueva    = _a_date_h3(self._fecha_fin)

        estado = self._construir_indices_h3_heredados(ws_h3_act, heredados_con_match)
        return self._aplicar_actualizaciones_h3_heredados(
            wb, ws, ws_h3_act, estado, heredados_con_match, dic_tarifa,
            d_inicio_nueva, d_fin_nueva)



    def _limpiar_marca_aj_heredados(self, wb, ws, codigos_existentes,
                                    codigos_tarifa_set) -> list:
        """Paso 4c (extraido de _ejecutar): limpia AJ (Hoja1) / AN (Hoja2)
        en materiales heredados que SI continuan en la tarifa nueva, y
        gestiona los que se quedan sin tarifa (baja automatica por fecha
        vencida via Hoja3 col J, o ventana de decision del usuario para
        los vigentes). Devuelve la lista sin_tarifa (codigos heredados sin
        match en la tarifa), que _ejecutar necesita mas adelante solo para
        self._audit['n_sin_tarifa'] = len(sin_tarifa).
        """
        # ── 4c. Limpiar marca AJ=X heredada en materiales que SÍ continúan ────
        #
        #  El Purchasing descargado de SAP puede traer filas con AJ = "X"
        #  («Registro info marcado para borrado») en materiales que el usuario
        #  quiere reactivar incluyéndolos en la nueva tarifa.  Si no se limpia
        #  esta marca, el código la hereda tal cual y el paso 6b-bis la propaga
        #  a Hoja 2 (col AN) y Hoja 3 (col Y), produciendo un RINFO que da de
        #  baja materiales perfectamente válidos.
        #
        #  Regla: si un material heredado (en codigos_existentes) aparece también
        #  en codigos_tarifa_set → limpiar AJ en Hoja 1 y AN en Hoja 2.
        #  Los materiales que NO están en la tarifa conservan su AJ intacto;
        #  el paso 6b decidirá si marca baja o pregunta al usuario.
        #
        #  Solo se modifica .value = None; el formato de celda no se toca.

        self._barra("set", "Limpiando marcas de borrado en materiales con tarifa...", 48, 100)

        # Fix (2026-09-30, caso real melocotón/FRAN Y CHEMI, segunda vuelta):
        # el primer intento marcaba "Nueva" tanto en H1 (AK) como en H2 (AO)
        # al limpiar sus respectivas marcas de borrado -- SAP respondió con
        # "Ya existe el registro Info de compras" para las 3 cabeceras
        # marcadas. El motivo: lo que SAP tenía marcado para borrado era el
        # segmento organizativo por centro (H2 -- "Organización de compras
        # marcada para borrado"), NO la cabecera del registro info (H1 --
        # "Registro info marcado para borrado" venía en el Purchasing como
        # arrastre/resto, pero el registro en sí sigue existiendo en SAP).
        # Marcar H1 como Nueva le pide a SAP crear un registro con un número
        # que ya existe -> error. Solo H2 (el segmento realmente eliminado)
        # necesita el indicador Nueva para que SAP lo recree; H1 solo se
        # limpia (a None), nunca se marca Nueva.
        COL_AJ_H1 = 36   # AJ — «Registro info marcado para borrado» (1-based)
        n_aj_limpiados_h1 = 0
        self._materiales_registro_a_renumerar = {}   # {material: registro_viejo}
        self._materiales_registro_a_renumerar_desc = {}   # {material: texto_breve}
        self._materiales_centros_borrado = {}   # {material: [centro1, centro2, ...]}
        for fila_row in ws.iter_rows(min_row=2):
            val_e = fila_row[4].value if len(fila_row) > 4 else None
            if val_e is None:
                continue
            raw = val_e
            raw = normalizar_codigo_sap(raw)
            clave = str(raw).strip()
            if clave not in codigos_tarifa_set:
                continue   # no está en tarifa → dejar AJ intacto
            celda_aj = fila_row[COL_AJ_H1 - 1]   # índice 0-based = 35
            if celda_aj is not None and str(celda_aj.value).strip().upper() == "X":
                celda_aj.value = None
                self._materiales_registro_a_renumerar[clave] = fila_row[0].value
                _val_f = fila_row[5].value if len(fila_row) > 5 else None
                self._materiales_registro_a_renumerar_desc[clave] = \
                    str(_val_f).strip() if _val_f else ""
                n_aj_limpiados_h1 += 1

        COL_AN_H2 = 40   # AN — «Organización de compras marcada para borrado» (1-based)
        COL_AO_H2 = 41   # AO — «Nueva» de Hoja 2 (1-based) -- este SÍ es el
                         # segmento que SAP tenía realmente eliminado.
        n_an_limpiados_h2 = 0
        if self.HOJA_2 in wb.sheetnames:
            ws_h2_limpiar = wb[self.HOJA_2]
            for fila_row in ws_h2_limpiar.iter_rows(min_row=2):
                val_j = fila_row[9].value if len(fila_row) > 9 else None   # col J = Material
                if val_j is None:
                    continue
                raw = val_j
                raw = normalizar_codigo_sap(raw)
                clave = str(raw).strip()
                if clave not in codigos_tarifa_set:
                    continue
                celda_an = fila_row[COL_AN_H2 - 1]   # índice 0-based = 39
                if celda_an is not None and str(celda_an.value).strip().upper() == "X":
                    celda_an.value = None
                    # NO se marca AO="X" (2026-09-30): probado en real que no
                    # tiene ningún efecto -- SAP sigue sin reconocer el
                    # registro como existente ("modificado", nunca "creado").
                    # Ver _avisar_materiales_reactivados_bloqueados: el aviso
                    # al usuario es la única acción útil que queda por hacer
                    # aquí, la reactivación real es manual en SAP.
                    _val_centro = fila_row[5].value if len(fila_row) > 5 else None   # F Centro
                    if _val_centro:
                        self._materiales_centros_borrado.setdefault(
                            clave, []).append(str(_val_centro).strip())
                    n_an_limpiados_h2 += 1

        if n_aj_limpiados_h1 or n_an_limpiados_h2:
            _log(self._parent,
                 f"  Marca de borrado limpiada en materiales con tarifa: "
                 f"H1 col AJ = {n_aj_limpiados_h1} fila(s) | "
                 f"H2 col AN = {n_an_limpiados_h2} fila(s). "
                 f"(El Purchasing los traía marcados para borrado pero continúan "
                 f"en la tarifa; limpiar la marca no basta para que SAP los "
                 f"reactive -- ver aviso siguiente sobre gestión manual).",
                 "OK")

        # Materiales del Purchasing que NO tienen fila en la tarifa
        sin_tarifa = [c for c in codigos_existentes if c not in codigos_tarifa_set]

        if sin_tarifa:
            _log(self._parent,
                 f"  Materiales heredados sin tarifa: {len(sin_tarifa)} "
                 f"— verificando Fin validez en Hoja 3 (col J).", "INFO")

            # ── Diccionario código → (fila_h1, fin_validez_h3) ─────────────────
            #    Necesitamos:
            #      • fila de Hoja 1 donde está el código (para escribir col AJ)
            #      • fecha de Fin validez de Hoja 3 col J (para decidir automático/manual)
            #    También recuperamos la descripción (Hoja 1 col F = texto breve,
            #    índice 1-based=6) para mostrarlo en la ventana.

            # Mapa código → fila en Hoja 1 (reutiliza el mapa general construido
            # en el paso 4, ya disponible como self._fila_h1_por_codigo).
            fila_h1_por_codigo = self._fila_h1_por_codigo

            # Mapa código → (fin_validez datetime.date | None) desde Hoja 3 col J
            fin_validez_por_codigo = {}
            if self.HOJA_3 in wb.sheetnames:
                ws_h3_ref = wb[self.HOJA_3]
                for fila_row in ws_h3_ref.iter_rows(min_row=2, values_only=True):
                    # col F = índice 5 (material), col J = índice 9 (fin validez)
                    if not fila_row or len(fila_row) < 10:
                        continue
                    val_mat = fila_row[5]   # F Material
                    val_fin = fila_row[9]   # J Fin validez
                    if val_mat is None:
                        continue
                    raw_m = val_mat
                    raw_m = normalizar_codigo_sap(raw_m)
                    clave_m = str(raw_m).strip()
                    # Solo la primera aparición (una fecha por material)
                    if clave_m and clave_m not in fin_validez_por_codigo:
                        fin_validez_por_codigo[clave_m] = val_fin
            else:
                _log(self._parent,
                     f"  Hoja «{self.HOJA_3}» no disponible — "
                     "fechas de fin validez no consultadas.", "WARN")

            # ── Mapa código → descripción (Hoja 1 col F = texto breve) ─────────
            desc_por_codigo = {}
            for fila_row in ws.iter_rows(min_row=2, values_only=True):
                val_e = fila_row[4] if len(fila_row) > 4 else None
                val_f = fila_row[5] if len(fila_row) > 5 else None
                if val_e is None:
                    continue
                raw = val_e
                raw = normalizar_codigo_sap(raw)
                clave = str(raw).strip()
                if clave and clave not in desc_por_codigo:
                    desc_por_codigo[clave] = str(val_f).strip() if val_f else ""

            # ── Clasificar: automáticos vs pendientes de decisión usuario ───────
            hoy = _dt_mod.date.today()
            pendientes_ventana = []   # materiales vigentes → ventana tkinter
            _st_automaticos    = []   # acumular bajas automáticas para resumen
            _st_omitidos       = []   # acumular omisiones (sin fila H1) para resumen

            for codigo in sin_tarifa:
                fila_h1 = fila_h1_por_codigo.get(codigo)
                if fila_h1 is None:
                    _st_omitidos.append(codigo)
                    continue

                val_fin = fin_validez_por_codigo.get(codigo)

                # Normalizar la fecha de fin de validez a datetime.date
                # (antes ~15 líneas inline; consolidado en sap_helpers.py)
                fecha_fin_date = parsear_fecha_celda(val_fin)

                if fecha_fin_date is not None and fecha_fin_date < hoy:
                    # Vencido: marcar baja automáticamente
                    ws.cell(row=fila_h1, column=36).value = "X"   # AJ (1-based)
                    _st_automaticos.append(
                        f"{codigo} (fin {fecha_fin_date.strftime('%d/%m/%Y')})")
                else:
                    # Sin fecha o vigente: acumular para decisión del usuario
                    fecha_str = (fecha_fin_date.strftime("%d/%m/%Y")
                                 if fecha_fin_date else "Sin fecha")
                    desc = desc_por_codigo.get(codigo, "")
                    pendientes_ventana.append({
                        "codigo":    codigo,
                        "desc":      desc,
                        "fecha":     fecha_str,
                        "fila_h1":   fila_h1,
                    })

            # Emitir WARN agrupado si hay omisiones (sin fila H1)
            if _st_omitidos:
                _log(self._parent,
                     f"  Sin tarifa — {len(_st_omitidos)} código(s) no encontrado(s) en Hoja1 "
                     f"(omitidos): {', '.join(_st_omitidos)}", "WARN")

            # ── Mostrar ventana para los materiales vigentes ─────────────────
            if pendientes_ventana:
                _log(self._parent,
                     f"  {len(pendientes_ventana)} material(es) vigente(s) sin tarifa "
                     "→ abriendo ventana de confirmación.", "INFO")

                # La ventana tkinter debe ejecutarse en el hilo principal.
                # Usamos after() + un Event para sincronizar el hilo secundario.
                import threading
                resultado_ventana = {}
                evento_listo      = threading.Event()

                # Resolver el widget raíz ahora (en el hilo secundario) para
                # usarlo como parent del Toplevel. _app_root es la ventana App
                # que vive toda la sesión; self._parent puede haberse destruido.
                _tk_root_local = self._app_root
                if _tk_root_local is None or not hasattr(_tk_root_local, "after"):
                    # Traversal de seguridad: subir hasta encontrar algo con .after()
                    _obj = self._parent
                    for _ in range(8):
                        if _obj is None:
                            break
                        if hasattr(_obj, "after"):
                            _tk_root_local = _obj
                            break
                        _obj = (getattr(_obj, "_app_root", None) or
                                getattr(_obj, "master",    None) or
                                getattr(_obj, "_parent",   None))

                def _abrir_ventana():
                    # ── Guard: verificar que el widget padre sigue existiendo ──
                    # Si el parent fue destruido, usamos _tk_root_local como
                    # alternativa. Si tampoco existe, ejecutamos sin ventana modal.
                    _parent_dlg = _tk_root_local
                    try:
                        if (_parent_dlg is None or
                                not hasattr(_parent_dlg, "winfo_exists") or
                                not _parent_dlg.winfo_exists()):
                            # No hay ventana válida — decisión conservadora
                            resultado_ventana["decisiones"] = [
                                {"codigo": p["codigo"], "fila_h1": p["fila_h1"], "baja": False}
                                for p in pendientes_ventana
                            ]
                            _log(self._parent,
                                 "  ⚠ Ventana sin tarifa omitida: widget padre destruido. "
                                 "Todos los materiales sin tarifa conservados (baja=False).",
                                 "WARN")
                            return
                        decisiones = _ventana_sin_tarifa(_parent_dlg, pendientes_ventana)
                        resultado_ventana["decisiones"] = decisiones
                    except Exception as exc_v:
                        _log(self._parent,
                             f"Error en ventana sin tarifa: {exc_v}", "ERROR")
                        resultado_ventana["decisiones"] = [
                            {"codigo": p["codigo"], "fila_h1": p["fila_h1"], "baja": False}
                            for p in pendientes_ventana
                        ]
                    finally:
                        evento_listo.set()

                _safe_after(_tk_root_local, _abrir_ventana)
                evento_listo.wait()   # bloquear hilo secundario hasta que el usuario decida

                _st_baja_usuario  = []
                _st_sin_fila_usr  = []
                decisiones = resultado_ventana.get("decisiones", [])
                for d in decisiones:
                    if d.get("baja"):
                        fila_h1 = d.get("fila_h1")
                        codigo  = d.get("codigo", "?")
                        if fila_h1:
                            ws.cell(row=fila_h1, column=36).value = "X"   # AJ
                            _st_baja_usuario.append(codigo)
                        else:
                            _st_sin_fila_usr.append(codigo)

                if _st_sin_fila_usr:
                    _log(self._parent,
                         f"  Sin tarifa — DAR BAJA solicitado pero fila_h1 desconocida "
                         f"para: {', '.join(_st_sin_fila_usr)}", "WARN")

            n_automaticos   = len(_st_automaticos)
            n_baja_usuario  = len(_st_baja_usuario) if pendientes_ventana else 0
            n_sin_decision  = len(pendientes_ventana) - n_baja_usuario
            partes_resumen  = [
                f"Sin tarifa: {len(sin_tarifa)} total",
                f"{n_automaticos} baja(s) automática(s) (vencidas)",
            ]
            if n_baja_usuario:
                partes_resumen.append(f"{n_baja_usuario} baja(s) por decisión usuario")
            if n_sin_decision > 0:
                partes_resumen.append(f"{n_sin_decision} sin decisión (conservadas)")
            if _st_automaticos:
                partes_resumen.append(f"Bajas auto: {', '.join(_st_automaticos)}")
            _log(self._parent,
                 " | ".join(partes_resumen) + ".",
                 "OK" if sin_tarifa else "INFO")

        else:
            _log(self._parent,
                 "  Todos los materiales del Purchasing tienen correspondencia "
                 "en la tarifa — sin bajas pendientes.", "INFO")

        return sin_tarifa

    def _avisar_materiales_reactivados_bloqueados(self, wb) -> int:
        """
        Sustituye a _renumerar_materiales_reactivados() (2026-09-30, mismo
        día, dos intentos fallidos en real con el caso MELOCOTON/FRAN Y
        CHEMI):
          1) Limpiar AJ/AN y marcar Nueva (AK) conservando el número real
             de SAP en col A → "Ya existe el registro Info de compras
             NNNNNNNNNN".
          2) Renumerar el registro entero como si fuera genuinamente nuevo
             (correlativo provisional + AK/AO/Z marcados Nueva, igual que
             un material insertado por primera vez) → MISMO error, con el
             número REAL antiguo, aunque el Excel ya no lo contuviera en
             ningún sitio.

        La segunda prueba es la que descarta cualquier vía por Excel: SAP
        compara por PROVEEDOR + MATERIAL + ORGANIZACIÓN DE COMPRAS, no por
        el número que lleve la fila -- es una regla de unicidad de sus
        datos maestros, no algo que una carga masiva pueda evitar con
        ningún truco de numeración. El usuario confirmó en SAP la
        transacción correcta para este caso ("Marcar reg.info de compra
        para borrar", desmarcando sus dos casillas "Regtro-Info completo"
        e "Ind.borr.dat.org.co.") -- pero es una acción manual sobre SAP,
        fuera del alcance de esta app.

        Por eso este método YA NO toca el Excel (ni renumera ni marca
        Nueva) -- limpiar AJ/AN en el paso 4c es la única modificación que
        se hace, y solo para que la carga no arrastre una marca de borrado
        obsoleta si algún día SAP sí deja reactivarlo por esta vía. Aquí
        solo se avisa con claridad de qué materiales/registros necesitan
        esa gestión manual, para que el usuario no suba a ciegas una carga
        que SAP va a rechazar con toda seguridad.
        """
        materiales = getattr(self, "_materiales_registro_a_renumerar", None)
        if not materiales:
            return 0

        centros_por_mat = getattr(self, "_materiales_centros_borrado", {})
        detalle = "; ".join(
            f"{mat} (antes {reg}, centro(s) "
            f"{', '.join(centros_por_mat.get(mat, [])) or '?'})"
            for mat, reg in materiales.items())
        _log(self._parent,
             f"  ⚠ {len(materiales)} material(es) con tarifa activa que el "
             f"Purchasing traía marcado(s) para borrado (Registro info "
             f"completo): {detalle}. Confirmado en real que SAP rechaza "
             f"tanto reactivar el número existente como crear uno nuevo "
             f"para el mismo proveedor+material+organización -- hay que "
             f"reactivarlos a mano en SAP antes o después de esta carga, "
             f"con la transacción \"Marcar reg.info de compra para borrar\" "
             f"(desmarcando sus dos casillas de borrado, para cada centro "
             f"afectado). Hasta entonces, sus precios NO estarán disponibles "
             f"para pedidos nuevos aunque este RINFO se cargue sin errores "
             f"en el resto de materiales.",
             "WARN")
        return len(materiales)

    def _leer_purchasing_y_tarifa(self, wb, ws, ruta_nuevo):
        """Pasos 4-6 (extraido de _ejecutar): lee los materiales ya
        presentes en Hoja1 col E del Purchasing, limpia el campo NUEVO
        residual en las 3 hojas (Hoja1 AK / Hoja2 AO / Hoja3 Z), lee la
        tarifa nueva (col A/C) y calcula los materiales sin match (a
        insertar). Tambien construye self._fila_h1_por_codigo, usado mas
        adelante en el paso 4c y en el paso 6c.

        Devuelve la tupla (codigos_existentes, codigos_tarifa,
        codigos_tarifa_set, nuevos, ultima_fila_datos, max_id_col_a), o
        None si falla la lectura de la tarifa -- en ese caso ya deja el
        archivo destino limpio y el error mostrado al usuario, igual que
        hacia el bloque inline (mismo contrato que _preparar_archivo_destino).
        """
        # ── 4. Leer materiales ya presentes en col E del Purchasing ───────────
        self._barra("set", "Leyendo materiales del Purchasing...", 20, 100)
        codigos_existentes = set()
        _max_id_col_a      = 0   # máximo ID numérico de col A (Hoja 1) en filas heredadas
        ultima_fila_datos  = 1

        for fila in ws.iter_rows(min_row=2, values_only=False):
            val_e = fila[4].value if len(fila) > 4 else None   # col E = índice 4
            if val_e is not None:
                raw = val_e
                raw = normalizar_codigo_sap(raw)
                codigos_existentes.add(str(raw).strip())
                # Punto C: capturar max correlativo de posición (col A).
                # Col A puede contener:
                #   · Correlativo simple (1, 2, 3, …) — asignado por este código
                #   · Número de registro SAP real (530000xxxx, 10 dígitos) — asignado por SAP
                # Si col A contiene un ID SAP real (> 9999), no es un correlativo
                # de posición y no debe usarse como offset: SAP lo asigna en la
                # carga. En ese caso, el offset se deriva de n_filas_heredadas.
                val_a_scan = fila[0].value
                if val_a_scan is not None:
                    try:
                        _id_num = int(float(str(val_a_scan).strip()))
                        if 0 < _id_num <= 9999:   # solo correlativos de posición
                            if _id_num > _max_id_col_a:
                                _max_id_col_a = _id_num
                    except (ValueError, TypeError):
                        pass
                ultima_fila_datos = fila[4].row

        _log(self._parent,
             f"  Purchasing: {len(codigos_existentes)} material(es) existente(s), "
             f"última fila con datos en col E = {ultima_fila_datos}.",
             "INFO")
        self._audit["n_heredados"]   = len(codigos_existentes)
        self._audit["max_id_col_a"]  = _max_id_col_a
        self._audit["centros_activos"] = list(self._centros_activos)

        # ── Mapa general código → fila Hoja 1 (TODOS los heredados) ───────────
        # Se usa en varios puntos posteriores del flujo (marcado de NUEVO=X
        # ante cambios reales en paso 6c, baja por "sin tarifa", etc.).
        # Solo se guarda la primera fila encontrada para cada código.
        self._fila_h1_por_codigo: dict = {}
        for fila_row in ws.iter_rows(min_row=2):
            val_e = fila_row[4].value if len(fila_row) > 4 else None
            if val_e is None:
                continue
            raw = val_e
            raw = normalizar_codigo_sap(raw)
            clave = str(raw).strip()
            if clave and clave not in self._fila_h1_por_codigo:
                self._fila_h1_por_codigo[clave] = fila_row[4].row

        # ── 4b. Limpiar campo NUEVO en todas las filas heredadas ──────────────
        #
        #  Regla SAP PIR: NUEVO = "X" SOLO en registros que se crean por primera
        #  vez. Las filas que ya existen en SAP (las que vienen del Purchasing
        #  descargado) no deben llevar NUEVO = "X"; si lo tuvieran (valor residual
        #  de la exportación SAP), SAP produciría el error ME807
        #  «Registro info ya existe» o generaría condiciones duplicadas.
        #
        #  Se limpian tres hojas:
        #    · Hoja 1 col AK (37) — «Nueva»
        #    · Hoja 2 col AO (41) — «Nueva»
        #    · Hoja 3 col Z  (26) — «Nueva»
        #
        #  Solo se toca .value = None — el formato de celda no se modifica.
        #  Las filas nuevas que se insertarán en pasos posteriores recibirán
        #  NUEVO = "X" de forma normal.

        n_nuevo_limpiados_h1 = 0
        for fila_row in ws.iter_rows(min_row=2):
            celda_nuevo = fila_row[36] if len(fila_row) > 36 else None  # col AK (índice 36, 0-based)
            if celda_nuevo is not None and celda_nuevo.value == "X":
                celda_nuevo.value = None
                n_nuevo_limpiados_h1 += 1

        n_nuevo_limpiados_h2 = 0
        if self.HOJA_2 in wb.sheetnames:
            for fila_row in wb[self.HOJA_2].iter_rows(min_row=2):
                celda_nuevo = fila_row[40] if len(fila_row) > 40 else None  # col AO (índice 40, 0-based)
                if celda_nuevo is not None and celda_nuevo.value == "X":
                    celda_nuevo.value = None
                    n_nuevo_limpiados_h2 += 1

        n_nuevo_limpiados_h3 = 0
        if self.HOJA_3 in wb.sheetnames:
            for fila_row in wb[self.HOJA_3].iter_rows(min_row=2):
                celda_nuevo = fila_row[25] if len(fila_row) > 25 else None  # col Z (índice 25, 0-based)
                if celda_nuevo is not None and celda_nuevo.value == "X":
                    celda_nuevo.value = None
                    n_nuevo_limpiados_h3 += 1

        _log(self._parent,
             f"  Campo NUEVO limpiado en filas heredadas: "
             f"H1 col AK = {n_nuevo_limpiados_h1} | "
             f"H2 col AO = {n_nuevo_limpiados_h2} | "
             f"H3 col Z  = {n_nuevo_limpiados_h3} "
             f"(las filas nuevas recibirán X en pasos posteriores).", "OK")

        # ── 5. Leer col A y C de la tarifa ────────────────────────────────────
        self._barra("set", "Leyendo materiales de la tarifa...", 35, 100)
        codigos_tarifa = []
        try:
            wb_tarifa = openpyxl.load_workbook(
                self._ruta_tarifa, read_only=True, data_only=True)
            hoja_t = (wb_tarifa[HOJA_TARIFA]
                      if HOJA_TARIFA in wb_tarifa.sheetnames
                      else wb_tarifa.active)
            for fila in hoja_t.iter_rows(min_row=8, min_col=1,
                                          max_col=3, values_only=True):
                cod_sap  = fila[0]
                cod_prov = fila[2]
                if cod_sap is None:
                    continue
                cod_sap = normalizar_codigo_sap(cod_sap)
                cod_sap_str = str(cod_sap).strip()
                if cod_prov is not None:
                    cod_prov = normalizar_codigo_sap(cod_prov)
                    cod_prov_str = str(cod_prov).strip()
                else:
                    cod_prov_str = None
                codigos_tarifa.append((cod_sap_str, cod_prov_str))
            wb_tarifa.close()
        except Exception as e:
            _log(self._parent, f"Error al leer tarifa: {e}", "ERROR")
            wb.close()
            self._intentar_eliminar(ruta_nuevo)
            self._parent.after(0, lambda err=str(e): messagebox.showerror(
                "Error al leer tarifa",
                f"No se pudieron leer los materiales de la tarifa:\n\n{err}",
                parent=self._parent))
            self._barra("reset", "Error al leer tarifa")
            _safe_after(self._parent, self._parent.withdraw)
            return None

        # ── 6. Filtrar solo materiales sin match ──────────────────────────────
        self._barra("set", "Cruzando materiales...", 50, 100)
        nuevos = [(c, p) for c, p in codigos_tarifa if c not in codigos_existentes]

        _log(self._parent,
             f"  Tarifa: {len(codigos_tarifa)} material(es) total | "
             f"Con match: {len(codigos_tarifa) - len(nuevos)} | "
             f"Sin match (a insertar): {len(nuevos)}", "INFO")

        # ── 6b. Marcar baja materiales heredados SIN match en tarifa ──────────
        #
        #  Los materiales que existen en el Purchasing (col E Hoja 1) pero NO
        #  tienen correspondencia en la tarifa activa se consideran "sin tarifa".
        #  Para cada uno se consulta la col J (Fin validez, índice 1-based=10)
        #  de la Hoja 3 (Det.precios p.registros info):
        #
        #    · Fecha < hoy  →  se marca automáticamente Hoja1 col AJ (36) = "X"
        #    · Fecha vigente →  se acumula en lista para que el usuario decida
        #                        mediante la ventana VentanaSinTarifa.
        #
        #  Se busca la primera aparición de cada código en la Hoja 3; si un
        #  material no tiene fila en Hoja 3, se omite sin marcar.

        self._barra("set", "Comprobando materiales sin tarifa...", 53, 100)

        # Construir set de códigos SAP de la tarifa para cruce rápido
        codigos_tarifa_set = {c for c, _ in codigos_tarifa}

        return (codigos_existentes, codigos_tarifa, codigos_tarifa_set,
                nuevos, ultima_fila_datos, _max_id_col_a)

    def _propagar_borrado_h1_a_h2_h3(self, wb, ws) -> None:
        """Paso 6b-bis (extraido de _ejecutar): cualquier material con
        AJ='X' en Hoja1 (baja automatica o por decision del usuario)
        propaga la marca de borrado a Hoja2 (col AN) y Hoja3 (col Y) en
        todas sus filas, y cierra col J (Fin validez) solo en la fila mas
        reciente de cada grupo (material, centro).

        No devuelve nada: escribe directamente sobre wb/ws (mismos objetos
        que el resto de _ejecutar) y sobre self._audit['materiales_baja'].
        Ninguna variable local de este bloque se usaba fuera de el.
        """
        # ── 6b-bis: Propagar borrado H1 → H2 y H3 ────────────────────────────
        #
        #  Cualquier material que tenga AJ = 'X' en Hoja 1 (ya sea por baja
        #  automática por fecha vencida o por decisión del usuario) debe tener
        #  la marca de borrado también en:
        #    · Hoja 2 col AN (40) — «Organización de compras marcada para borrado»
        #    · Hoja 3 col Y  (25) — «Indicador borrado»
        #  en TODAS sus filas (una por cada centro).
        #  Si el material no tiene filas en H2/H3 (caso marginal) se omite sin
        #  error, ya que SAP no puede dar de baja lo que no existe en esas hojas.

        self._barra("set", "Propagando borrado H1 → H2/H3...", 55, 100)

        # 1. Construir set de claves con AJ=X en H1
        materiales_baja_h1 = set()
        for fila_row in ws.iter_rows(min_row=2):
            val_e  = fila_row[4].value  if len(fila_row) > 4  else None
            val_aj = fila_row[35].value if len(fila_row) > 35 else None
            if val_aj == "X" and val_e is not None:
                raw = val_e
                raw = normalizar_codigo_sap(raw)
                materiales_baja_h1.add(str(raw).strip())

        self._audit["materiales_baja"] = set(materiales_baja_h1)

        n_marcados_h2 = 0
        n_marcados_h3 = 0

        # 2. Marcar en H2 col AN (índice 1-based = 40)
        if materiales_baja_h1 and self.HOJA_2 in wb.sheetnames:
            ws_h2_baja = wb[self.HOJA_2]
            COL_AN_H2 = 40   # AN — «Organización de compras marcada para borrado»
            for fila_row in ws_h2_baja.iter_rows(min_row=2):
                val_j = fila_row[9].value if len(fila_row) > 9 else None
                if val_j is None:
                    continue
                raw = val_j
                raw = normalizar_codigo_sap(raw)
                if str(raw).strip() in materiales_baja_h1:
                    ws_h2_baja.cell(
                        row=fila_row[0].row, column=COL_AN_H2).value = "X"
                    n_marcados_h2 += 1

        # 3. Marcar en H3 col Y (índice 1-based = 25) y cerrar col J
        #
        #  Estrategia para materiales con UNA o VARIAS filas por centro:
        #
        #    a) col Y = "X" se escribe en TODAS las filas del material (cualquier centro).
        #
        #    b) col J (Fin validez) solo se cierra en la fila MÁS RECIENTE de cada
        #       grupo (material + centro), es decir la de mayor col I (Inicio validez).
        #       El cierre se aplica únicamente si esa J está vacía o es >= hoy.
        #       Valor de cierre: fecha_inicio_nueva - 1 día.
        #
        #    c) Las filas anteriores del mismo grupo (históricas) se dejan intactas
        #       en col J. El paso 10a (_ajustar_fechas_fin_h3) las encadenará
        #       correctamente después: a cada una le pondrá J = Inicio_siguiente - 1 día,
        #       garantizando que no haya solapamientos ni huecos entre períodos.
        #
        #    Separar ambas responsabilidades evita que este paso sobrescriba fechas
        #    históricas con el mismo valor de cierre, lo que confundiría a
        #    _ajustar_fechas_fin_h3 y generaría períodos con J <= I en SAP.
        #
        #  La col I (Inicio validez) nunca se modifica en ningún caso.

        if materiales_baja_h1 and self.HOJA_3 in wb.sheetnames:
            ws_h3_baja = wb[self.HOJA_3]
            COL_Y_H3 = 25    # Y  — «Indicador borrado»
            COL_J_H3 = 10    # J  — «Fin validez»
            COL_I_H3 =  9    # I  — «Inicio validez»

            # ── Calcular fecha de cierre: día anterior al inicio de la nueva tarifa.
            # Si _fecha_inicio no se puede parsear, se usa ayer como fallback seguro.
            # (parseo consolidado en sap_helpers.py; antes ~15 líneas propias aquí)
            _hoy_bis = _dt_mod.date.today()
            _fecha_inicio_parsed = (
                parsear_fecha_celda(self._fecha_inicio) if self._fecha_inicio else None)
            _fecha_cierre_h3: _dt_mod.date = (
                (_fecha_inicio_parsed - _dt_mod.timedelta(days=1))
                if _fecha_inicio_parsed is not None
                else (_hoy_bis - _dt_mod.timedelta(days=1))
            )

            _str_cierre_h3 = _fecha_cierre_h3.strftime("%d.%m.%Y")

            # Antes definida aquí mismo; consolidada en sap_helpers.py junto
            # con las otras 4 copias idénticas del proyecto (ver ese módulo).
            _parse_date_bis = parsear_fecha_celda

            # ── 1ª pasada: recopilar todas las filas de materiales en baja,
            #    agrupadas por (material, centro) para identificar la más reciente.
            #    Estructura: { (mat, centro): [ {row, cell_y, cell_j, fecha_i, fecha_j}, ... ] }
            from collections import defaultdict as _defdict
            _grupos_baja: dict = _defdict(list)

            for fila_row in ws_h3_baja.iter_rows(min_row=2):
                val_f = fila_row[5].value if len(fila_row) > 5 else None
                if val_f is None:
                    continue
                raw_f = val_f
                raw_f = normalizar_codigo_sap(raw_f)
                clave_mat = str(raw_f).strip()
                if clave_mat not in materiales_baja_h1:
                    continue

                val_e = fila_row[4].value if len(fila_row) > 4 else None
                clave_centro = str(val_e).strip() if val_e is not None else ""

                _grupos_baja[(clave_mat, clave_centro)].append({
                    "row":    fila_row[0].row,
                    "cell_y": fila_row[24] if len(fila_row) > 24 else None,
                    "cell_j": fila_row[9]  if len(fila_row) > 9  else None,
                    "fecha_i": _parse_date_bis(fila_row[8].value  if len(fila_row) > 8  else None),
                    "fecha_j": _parse_date_bis(fila_row[9].value  if len(fila_row) > 9  else None),
                })

            # ── 2ª pasada: marcar col Y en todas las filas; cerrar col J solo en
            #    la fila más reciente (mayor fecha_i) de cada grupo.
            n_j_cerradas_h3 = 0

            for (mat, centro), filas in _grupos_baja.items():
                # Ordenar por fecha_i desc: None al final (período sin inicio conocido)
                filas_ord = sorted(
                    filas,
                    key=lambda r: (r["fecha_i"] is None, r["fecha_i"]),
                    reverse=True,
                )
                fila_reciente = filas_ord[0]   # la de mayor col I

                for fi in filas:
                    # a) Marcar indicador borrado en TODAS
                    if fi["cell_y"] is not None:
                        ws_h3_baja.cell(row=fi["row"], column=COL_Y_H3).value = "X"
                        n_marcados_h3 += 1

                # b) Cerrar col J solo en la fila más reciente, si está vigente o
                #    vacía Y empieza antes (o el mismo día) que la fecha de cierre:
                #    si empieza DESPUÉS, el cierre dejaría fin < inicio (período
                #    invertido, caso MELOCOTON 2026-10-06) -> no se toca.
                fj = fila_reciente["fecha_j"]
                if debe_cerrar_fin_validez(fila_reciente["fecha_i"], fj,
                                           _hoy_bis, _fecha_cierre_h3):
                    if fila_reciente["cell_j"] is not None:
                        ws_h3_baja.cell(
                            row=fila_reciente["row"],
                            column=COL_J_H3).value = _str_cierre_h3
                        n_j_cerradas_h3 += 1

        if materiales_baja_h1:
            _log(self._parent,
                 f"  Borrado propagado: {len(materiales_baja_h1)} material(es) "
                 f"→ {n_marcados_h2} fila(s) H2 col AN | "
                 f"{n_marcados_h3} fila(s) H3 col Y "
                 f"({n_j_cerradas_h3} col J cerrada(s) en fila más reciente"
                 f" → {_str_cierre_h3}; resto encadenado por paso 10a).", "OK")

    def _ejecutar(self):
        """Orquesta la generación del nuevo archivo en hilo secundario."""

        _log(self._parent,
             f"GeneradorDesdeExistente: teléfono={self._telefono_proveedor!r}  "
             f"inicio={self._fecha_inicio}  fin={self._fecha_fin}  "
             f"centros={[c for c,_ in self._centros_activos]}", "INFO")

        # Diccionario de auditoría — se rellena a lo largo del proceso y se
        # vuelca en el bloque de validación SAP previo al guardado (paso 10d).
        self._audit: dict = {
            "centros_activos":     [],   # lista de (cod, nom)
            "n_heredados":         0,    # materiales preexistentes en Purchasing
            "n_nuevos":            0,    # materiales insertados desde tarifa
            "n_omitidos_dup_h2":   0,    # filas H2 saltadas por duplicado funcional
            "n_omitidos_dup_h3":   0,    # filas H3 saltadas por duplicado funcional
            "n_borrado_h2":        0,    # filas H2 marcadas para borrado (centros inactivos)
            "n_sin_tarifa":        0,    # materiales heredados sin precio en tarifa
            "n_solapamientos":     0,    # períodos cerrados por solapamiento
            "n_gaps":              0,    # pares (clave_pir) con gap de cobertura
            "n_reducciones":       0,    # períodos con reducción de vigencia
            "n_absorbidos":        0,    # períodos heredados absorbidos (validez cero)
            "n_tramos_reconstruidos": 0, # tramos posteriores recreados (heredado engloba nueva)
            "centros_sin_ext":     [],   # centros activos sin ninguna fila nueva (H2 vacía)
            "max_id_col_a":        0,    # mayor ID heredado en col A Hoja 1
            "gaps_detalle":        {},   # clave_pir → gap_dias (máximo)
            "solap_detalle":       {},   # clave_pir → str_cierre
            # ── Resumen NUEVO/cambios reales (regla NUEVO solo si hay cambio) ──
            "materiales_baja":           set(),  # códigos con AJ=X (para borrar)
            "materiales_modificados":    set(),  # códigos heredados con cambio real (precio/fecha)
            "materiales_sin_cambios":    set(),  # códigos heredados sin ningún cambio
        }

        # ── Pasos 1-3: nombre, copia y apertura del archivo destino ──────────
        #    Extraído a _preparar_archivo_destino() — ver ese método para el
        #    detalle línea a línea (idéntico al que había aquí antes).
        _resultado_prep = self._preparar_archivo_destino()
        if _resultado_prep is None:
            return
        wb, ws, ruta_nuevo, nombre_nuevo = _resultado_prep

        # ── Pasos 4-6: leer Purchasing, limpiar NUEVO, leer tarifa ───────────
        #    Extraido a _leer_purchasing_y_tarifa() -- ver ese metodo para el
        #    detalle linea a linea (identico al que habia aqui antes).
        _resultado_lectura = self._leer_purchasing_y_tarifa(wb, ws, ruta_nuevo)
        if _resultado_lectura is None:
            return
        (codigos_existentes, codigos_tarifa, codigos_tarifa_set,
         nuevos, ultima_fila_datos, _max_id_col_a) = _resultado_lectura


        # ── 4c. Limpiar marca AJ=X heredada en materiales que SI continuan ───
        #    Extraido a _limpiar_marca_aj_heredados() -- ver ese metodo para
        #    el detalle linea a linea (identico al que habia aqui antes).
        sin_tarifa = self._limpiar_marca_aj_heredados(
            wb, ws, codigos_existentes, codigos_tarifa_set)

        # ── 6b-bis: Propagar borrado H1 → H2 y H3 ────────────────────────────
        #    Extraido a _propagar_borrado_h1_a_h2_h3() -- ver ese metodo para
        #    el detalle linea a linea (identico al que habia aqui antes).
        self._propagar_borrado_h1_a_h2_h3(wb, ws)


        # ── 7. Cargar diccionarios auxiliares ─────────────────────────────────
        self._barra("indeterminate", "Cargando datos auxiliares...")
        diccionario = self._cargar_listado_dali()
        dic_tarifa  = self._cargar_precios_tarifa()

        # ── 6c. Actualizar fechas (y precio si cambia) en Hoja 3 ──────────────
        #    Extraido a _actualizar_h3_heredados() -- ver ese metodo para el
        #    detalle linea a linea (identico al que habia aqui antes).
        _absorbidas_con_tarifa = self._actualizar_h3_heredados(
            wb, ws, codigos_existentes, codigos_tarifa, dic_tarifa)

        # ── 8. Insertar materiales nuevos al final de Hoja 1 ─────────────────
        self._barra("set", "Insertando nuevos materiales en Hoja 1...", 60, 100)
        primera_fila_nueva = ultima_fila_datos + 1

        # Punto C: si _max_id_col_a == 0 significa que col A contiene IDs SAP
        # reales (no correlativos de posición) o la hoja está vacía.
        # En ese caso los nuevos materiales siempre reciben correlativos 1, 2, 3, …
        # empezando desde 1 (offset = 0), ya que SAP asigna sus propios IDs reales
        # en la carga y los heredados conservan los suyos intactos.
        # Si col A SÍ tenía correlativos (1..N), _max_id_col_a = N y los nuevos
        # continúan desde N+1.
        # En ambos casos col A de los heredados NO se modifica.
        _id_offset_efectivo = _max_id_col_a if _max_id_col_a > 0 else 0
        self._rellenar_h1_nuevas_filas(ws, primera_fila_nueva, nuevos, diccionario,
                                           id_offset=_id_offset_efectivo)

        insertados = len(nuevos)
        self._audit["n_nuevos"] = insertados
        _log(self._parent,
             f"  {insertados} material(es) insertado(s) en Hoja1 "
             f"a partir de fila {primera_fila_nueva}.",
             "OK")

        # ── 9. Añadir filas nuevas a Hoja 2 ───────────────────────────────────
        self._barra("set", "Rellenando Hoja 2 (Datos org.compr.)...", 72, 100)
        filas_h2_nuevas = self._rellenar_h2_nuevas_filas(
            wb, nuevos, primera_fila_nueva, diccionario,
            codigos_tarifa_set=codigos_tarifa_set)

        # ── 9b. Recuperar absorbidos con tarifa ───────────────────────────────
        #   Materiales cuya fila en H3 fue eliminada por absorción total pero
        #   que siguen activos en la tarifa nueva necesitan una fila nueva en
        #   H2 y H3. Sin este paso quedan con GAP de cobertura de precio.
        if _absorbidas_con_tarifa:
            filas_recuperadas = self._generar_filas_absorbidas_con_tarifa(
                wb, _absorbidas_con_tarifa, diccionario)
            filas_h2_nuevas = list(filas_h2_nuevas) + filas_recuperadas
            _log(self._parent,
                 f"  Recuperados/generados: {len(filas_recuperadas)} fila(s) de "
                 f"H2/H3 generadas para {len(_absorbidas_con_tarifa)} PIR(s) "
                 f"con tarifa activa (absorbidos totalmente, con solapamiento "
                 f"parcial cuyo período viejo se acortó, o sin ninguna fila "
                 f"previa en Hoja 3).", "OK")

        # ── 10. Añadir filas nuevas a Hoja 3 ──────────────────────────────────
        self._barra("set", "Rellenando Hoja 3 (Det. precios)...", 83, 100)
        primera_fila_nueva_h3 = self._rellenar_h3_nuevas_filas(wb, filas_h2_nuevas, diccionario, dic_tarifa)

        # ── 10a-bis. Avisar de los registros que el Purchasing traía ──────────
        #    marcados para borrado (ver _avisar_materiales_reactivados_bloqueados
        #    para el detalle). Confirmado en real, dos veces, con el caso
        #    MELOCOTON/FRAN Y CHEMI (2026-09-30): ni limpiar la marca y dejar
        #    el número real, ni renumerar como registro Nuevo, consiguen que
        #    SAP acepte la carga -- responde "Ya existe el registro Info de
        #    compras" comparando por proveedor+material+organización de
        #    compras, no por el número que lleve la fila del Excel. Es una
        #    regla de SAP que ninguna carga masiva puede sortear; hace falta
        #    reactivarlo a mano en SAP (transacción "Marcar reg.info de
        #    compra para borrar", desmarcando sus dos casillas). Por eso la
        #    app ya NO intenta renumerar ni marcar Nueva -- solo limpia AJ/AN
        #    (paso 4c) y avisa con claridad de cuáles necesitan esa gestión
        #    manual, para no generar una carga que SAP va a rechazar seguro.
        self._avisar_materiales_reactivados_bloqueados(wb)

        # ── 10a. Ajustar coherencia de fechas en Hoja 3 (heredados repetidos) ─
        #   Garantiza que la col J (Fin validez) de cada tarifa heredada termine
        #   el día anterior al inicio de la siguiente tarifa del mismo mat.+centro.
        #   Las filas marcadas como Nueva (actualización vigente) conservan sus
        #   fechas. También aplica a filas con Indicador borrado (col Y = 'X').
        #   Se pasa primera_fila_nueva_h3 para que el ajuste sepa qué filas
        #   pertenecen al batch actual y nunca las modifique.
        self._barra("set", "Ajustando coherencia de fechas en Hoja 3...", 85, 100)
        _log(self._parent, "Ajustando coherencia fechas col J en Hoja 3 (heredados)...", "INFO")
        self._audit["n_sin_tarifa"] = len(sin_tarifa)
        self._ajustar_fechas_fin_h3(wb, primera_fila_nueva_h3)

        # ── 10b. Escribir teléfono en col AD de TODAS las filas de Hoja 1 ─────
        #   Se aplica tanto a los materiales preexistentes como a los nuevos.
        self._barra("set", "Escribiendo teléfono proveedor (Hoja 1)...", 87, 100)
        _log(self._parent, "Aplicando teléfono proveedor en Hoja 1 (col AD)...", "INFO")
        self._aplicar_telefono_hoja1(ws)

        # ── 10b2. Escribir fechas I/J en TODAS las filas de Hoja 1 ───────────
        #   Se aplica tanto a los materiales preexistentes como a los nuevos.
        #   Garantiza que las filas heredadas del Purchasing quedan con las
        #   fechas de vigencia correctas de esta tarifa.
        _log(self._parent, "Aplicando fechas vigencia en Hoja 1 (cols I/J)...", "INFO")
        self._aplicar_fechas_hoja1(ws)

        # ── 10b3. Sincronizar UM en TODAS las filas de Hoja 1 (cols N/P) ─────
        #   Corrige la Unidad de medida de las filas heredadas si el LISTADO
        #   CODIGOS DALI-SAP (alimentado desde Productos.xlsx/MM03) trae un
        #   valor distinto. NO marca AK: en Hoja 1, AK="X" le indica a SAP
        #   que CREE un registro nuevo, y en una fila heredada la col A ya
        #   contiene el ID real existente — marcarla como Nueva provocaría
        #   el error "Ya existe el registro Info de compras NNNNNNNNNN" y
        #   dejaría la fila sin procesar. SAP recoge el cambio de N/P al
        #   modificar el registro existente, igual que ya hace con teléfono
        #   (10b) y fechas (10b2).
        self._barra("set", "Sincronizando UM en Hoja 1 (cols N/P)...", 88, 100)
        _log(self._parent, "Sincronizando UM en Hoja 1 (cols N/P) frente al LISTADO...", "INFO")
        self._sincronizar_um_hoja1(ws, diccionario)

        # ── 10c. Replicar formato de última fila existente a las filas nuevas ──
        #   primera_fila_nueva es la primera fila insertada en este proceso.
        #   El método usará primera_fila_nueva - 1 como referencia de formato,
        #   garantizando que las nuevas filas tienen el mismo estilo visual
        #   (azul claro, bordes, fuentes) que las preexistentes del Purchasing.
        self._barra("set", "Aplicando formato a filas nuevas en Hoja 1...", 90, 100)
        _log(self._parent,
             f"Replicando formato de fila {primera_fila_nueva - 1} "
             f"a filas nuevas ({primera_fila_nueva}+) en Hoja 1...", "INFO")
        self._aplicar_formato_plantilla_hoja1(ws, primera_fila_nueva)

        # ── 10d. Renumeración final col K Hoja3 (norma SAP) ─────────────────
        # DEBE ejecutarse DESPUÉS de todos los ajustes de fechas y eliminaciones
        # por absorción, para que la secuencia 1..N quede sin huecos.
        self._barra("set", "Renumerando col K Hoja3 (ID grupo cond.)...", 90, 100)
        self._renumerar_col_k_h3(wb)

        # ── 10d2. Sincronizar UM en TODAS las filas de Hoja 3 (cols R/S) ─────
        #   Mismo criterio que 10b3 (Hoja 1) e IGIC (Hoja 2): corrige la UM de
        #   las filas heredadas frente al LISTADO CODIGOS DALI-SAP y marca la
        #   fila como Nueva (Z) cuando corresponda.
        self._barra("set", "Sincronizando UM en Hoja 3 (cols R/S)...", 91, 100)
        _log(self._parent, "Sincronizando UM en Hoja 3 (cols R/S) frente al LISTADO...", "INFO")
        self._sincronizar_um_hoja3(wb, diccionario)

        # ── 10e. Resumen de validación SAP previo al guardado ────────────────
        self._emitir_resumen_validacion_sap()

        # ── 11. Guardar ───────────────────────────────────────────────────────
        self._barra("set", "Guardando nuevo archivo...", 92, 100)
        try:
            wb.save(ruta_nuevo)
            wb.close()
            self._barra("set", "Adaptando estructura OOXML para SAP...", 95, 100)
            n_sst = _postprocesar_para_sap(ruta_nuevo, self._purchasing["ruta"])
            _log(self._parent,
                 f"Post-procesado SAP OK: sharedStrings ({n_sst} cadena(s)), "
                 f"estructura OOXML restaurada del Purchasing original.", "OK")
        except Exception as e:
            self._parent.after(0, lambda err=str(e): messagebox.showerror(
                "Error al guardar",
                f"No se pudo guardar el archivo:\n\n{nombre_nuevo}\n\n{err}",
                parent=self._parent))
            _log(self._parent, f"Error al guardar: {e}", "ERROR")
            self._barra("reset", "Error al guardar")
            self._intentar_eliminar(ruta_nuevo)
            _safe_after(self._parent, self._parent.withdraw)
            return

        _log(self._parent,
             f"Archivo generado: {nombre_nuevo} "
             f"({insertados} material(es) nuevo(s) añadido(s))", "OK")
        self._barra("completado", f"Generado: {insertados} material(es) nuevo(s)")

        # ── 12. Confirmación final ────────────────────────────────────────────
        nom     = nombre_nuevo
        ruta    = ruta_nuevo
        n_sap   = self._nombre_sap
        c_sap   = self._codigo_sap
        n_exist = len(codigos_existentes)
        n_ins   = insertados

        def _mostrar():
            parent_root = (self._parent.master
                           if hasattr(self._parent, "master")
                           else self._parent)
            VentanaConfirmacionGenerado(parent_root, nom, ruta, n_sap, c_sap,
                                        n_exist, n_ins,
                                        self._purchasing["archivo"])
            _materiales_bloqueados = getattr(
                self, "_materiales_registro_a_renumerar", None)
            if _materiales_bloqueados:
                VentanaMaterialesParaReactivar(
                    parent_root, _materiales_bloqueados,
                    getattr(self, "_materiales_registro_a_renumerar_desc", None),
                    getattr(self, "_materiales_centros_borrado", None))

        _safe_after(self._parent, self._parent.withdraw)
        _safe_after(self._parent, _mostrar, 80)

    # ── RECUPERACIÓN DE ABSORBIDOS CON TARIFA ────────────────────────────────

    def _generar_filas_absorbidas_con_tarifa(
            self, wb, absorbidas_con_tarifa: dict, diccionario: dict) -> list:
        """
        Para cada clave PIR (reg_info, org, material, centro) absorbida
        totalmente en el paso 6c que además tiene match en la tarifa activa,
        genera una fila nueva en Hoja 2 y devuelve la tupla correspondiente
        para que _rellenar_h3_nuevas_filas la escriba en Hoja 3.

        Sin este paso esos materiales quedan con GAP de cobertura: su período
        heredado fue eliminado pero la fila nueva del período actual no se
        crea porque el material ya existía en combinaciones_existentes de H2.

        Devuelve lista de tuplas con el formato de filas_h2_escritas:
            [(reg_info, tipo_reg, desc_cat, org_comp, centro, material,
              txt_mat, grp_art), ...]
        """
        if self.HOJA_2 not in wb.sheetnames or not absorbidas_con_tarifa:
            return []

        ws_h2 = wb[self.HOJA_2]
        ws_h1 = wb[self.HOJA_PRINCIPAL]

        COL_AN_BORRADO = 40

        # Última fila de H2 con dato (para añadir al final)
        ultima_fila_h2 = 1
        for fila_row in ws_h2.iter_rows(min_row=2):
            val_a = fila_row[0].value if len(fila_row) > 0 else None
            val_j = fila_row[9].value if len(fila_row) > 9 else None
            if val_a is not None or val_j is not None:
                ultima_fila_h2 = fila_row[0].row

        # Mapa material → reg_info (col A Hoja 1) — buscamos el reg_info
        # heredado de H1 para usarlo en las filas nuevas de H2/H3.
        reg_info_por_mat = {}
        for fila_row in ws_h1.iter_rows(min_row=2):
            val_e = fila_row[4].value if len(fila_row) > 4 else None
            val_a = fila_row[0].value if len(fila_row) > 0 else None
            if val_e is None:
                continue
            raw = val_e
            raw = normalizar_codigo_sap(raw)
            mat_str = str(raw).strip()
            if mat_str not in reg_info_por_mat:
                reg_info_por_mat[mat_str] = val_a

        centros_activos_set = {cod for cod, _ in self._centros_activos}

        # ── Combinaciones (material, centro) YA existentes en Hoja 2 ────────
        #    Antes, esta función escribía SIEMPRE una fila nueva en Hoja2
        #    para cada material×centro recibido, sin comprobar si ya existía
        #    una. Eso estaba bien para el caso "absorción total genuina" (el
        #    material pierde su período en H3 y necesita uno nuevo), pero
        #    creaba DUPLICADOS reales en Hoja2 para el caso "material con
        #    tarifa que nunca tuvo fila en H3" (ese material YA tenía su fila
        #    H2 correcta desde el Purchasing original — no necesitaba otra).
        #    H2 no lleva fechas/períodos (eso vive solo en H3), así que un
        #    material×centro que ya está en H2 nunca necesita una segunda fila
        #    ahí, se absorba o no su período en H3.
        _combinaciones_h2_existentes: set = set()
        for fila_row in ws_h2.iter_rows(min_row=2):
            val_f = fila_row[5].value if len(fila_row) > 5 else None   # col F – Centro
            val_j = fila_row[9].value if len(fila_row) > 9 else None   # col J – Material
            if val_j is None:
                continue
            _cod_centro_h2 = str(val_f).strip() if val_f is not None else ""
            _clave_mat_h2 = str(normalizar_codigo_sap(val_j)).strip()
            if _clave_mat_h2 and _cod_centro_h2:
                _combinaciones_h2_existentes.add((_clave_mat_h2, _cod_centro_h2))

        filas_recuperadas = []
        idx_escritura = 0

        # Deduplicar por (material, centro) para no generar dos filas del mismo PIR
        _ya_generado: set = set()

        # Acumulador para loguear agrupado por material al final del bucle,
        # en vez de una línea por cada combinación material×centro.
        _log_recuperados_por_mat: dict = {}

        for clave_pir in absorbidas_con_tarifa:
            reg_info_pir = clave_pir[0]
            material_pir = clave_pir[2]
            centro_pir   = clave_pir[3]

            # Solo centros activos
            if centro_pir not in centros_activos_set:
                continue

            dedup_key = (material_pir, centro_pir)
            if dedup_key in _ya_generado:
                continue
            _ya_generado.add(dedup_key)

            # Normalizar material a int si es posible
            try:
                mat_val = int(material_pir)
            except (ValueError, TypeError):
                mat_val = material_pir

            reg_info = reg_info_pir or reg_info_por_mat.get(material_pir, reg_info_pir)

            datos = diccionario.get(material_pir, ("", "", "", "", "", "", ""))
            desc_sap, unidad, grp_prod, desc_grp, ind_imp, grp_comp, desc_grp_c = datos

            # Nombre del centro
            nom_centro = next(
                (nom for cod, nom in self._centros_activos if cod == centro_pir),
                centro_pir
            )

            # ── Escribir fila nueva en Hoja 2 SOLO si no existe ya ───────────
            #    (ver comentario junto a _combinaciones_h2_existentes más arriba)
            _clave_h2_check = (str(material_pir).strip(), centro_pir)
            if _clave_h2_check not in _combinaciones_h2_existentes:
                fila_h2 = ultima_fila_h2 + 1 + idx_escritura
                idx_escritura += 1

                # Columnas específicas de esta fila (dependen del material/centro).
                ws_h2.cell(row=fila_h2, column=1,  value=reg_info)
                ws_h2.cell(row=fila_h2, column=5,  value="OC ISLAS CANANARIAS")
                ws_h2.cell(row=fila_h2, column=6,  value=centro_pir)
                ws_h2.cell(row=fila_h2, column=7,  value=nom_centro)
                ws_h2.cell(row=fila_h2, column=8,  value=grp_comp   or None)
                ws_h2.cell(row=fila_h2, column=9,  value=desc_grp_c or None)
                ws_h2.cell(row=fila_h2, column=10, value=mat_val)
                ws_h2.cell(row=fila_h2, column=11, value=desc_sap   or None)
                ws_h2.cell(row=fila_h2, column=12, value=grp_prod   or None)
                ws_h2.cell(row=fila_h2, column=34, value=ind_imp    or None)

                # Columnas de valor fijo obligatorio SAP — único punto de
                # escritura, ver COLUMNAS_FIJAS_H2.
                self._escribir_fijas_h2_fila(ws_h2, fila_h2)

                # Copiar formato desde última fila preexistente de Hoja 2
                fila_ref_h2 = max(2, ultima_fila_h2)
                aplicar_formato_fila(ws_h2, fila_h2,
                                     "Datos org.compr.p.reg.info", fila_ref=fila_ref_h2)

                _combinaciones_h2_existentes.add(_clave_h2_check)

                # Se agrupa el log por material (ver más abajo, tras el bucle) en
                # vez de loguear una línea por cada combinación material×centro
                # — con 50+ materiales × 10 centros esto podía llegar a cientos
                # de líneas.
                _log_recuperados_por_mat.setdefault(material_pir, []).append(
                    (centro_pir, fila_h2))
            else:
                # Ya existía la fila en Hoja2 (caso normal para "nunca tuvo
                # fila en H3"): no se duplica, solo se deja constancia en el
                # log agrupado de que este material×centro solo necesitó H3.
                _log_recuperados_por_mat.setdefault(material_pir, []).append(
                    (centro_pir, None))

            filas_recuperadas.append((
                reg_info,
                "0",
                "Normal",
                "IC00",
                centro_pir,
                mat_val,
                desc_sap or None,
                grp_prod or None,
            ))

        # ── Log agrupado por material (en vez de material×centro) ─────────
        for _mat_log, _pares in sorted(_log_recuperados_por_mat.items()):
            _centros_str  = ",".join(c for c, _ in _pares)
            _filas_nuevas = [f for _, f in _pares if f is not None]
            _n_ya_h2      = sum(1 for _, f in _pares if f is None)

            if _filas_nuevas:
                _f_ini, _f_fin = min(_filas_nuevas), max(_filas_nuevas)
                _rango_filas = (f"fila H2 {_f_ini}" if _f_ini == _f_fin
                                else f"filas H2 {_f_ini}-{_f_fin}")
                _detalle_h2 = f"{_rango_filas} generada(s)"
            else:
                _detalle_h2 = "ya existían en H2, sin duplicar"

            if _n_ya_h2 and _filas_nuevas:
                _detalle_h2 += f" ({_n_ya_h2} centro(s) ya en H2, sin duplicar)"

            _log(self._parent,
                 f"  Recuperado absorbido: mat={_mat_log} → {len(_pares)} "
                 f"centro(s) ({_centros_str}), {_detalle_h2} — Hoja3 "
                 f"(período {self._fecha_inicio}/{self._fecha_fin}).", "INFO")

        return filas_recuperadas

    # ── RENUMERACIÓN FINAL COL K HOJA 3 ─────────────────────────────────────

    def _renumerar_col_k_h3(self, wb) -> None:
        """
        Renumera la columna K ("ID grupo condiciones") de Hoja3 desde la fila 2
        hasta la última fila con datos, asignando valores enteros consecutivos
        comenzando en 1 (fila 2 → 1, fila 3 → 2, ...).

        NORMA SAP (confirmada): la col K DEBE ser siempre 1, 2, 3 ... sin huecos
        ni saltos, para TODOS los materiales (heredados y nuevos) sin distinción.
        Esta renumeración se ejecuta al final del proceso, después de todos los
        ajustes de fechas y eliminaciones por absorción, para garantizar que
        cualquier hueco generado por esas operaciones quede corregido.
        """
        if self.HOJA_3 not in wb.sheetnames:
            return
        ws_h3 = wb[self.HOJA_3]
        contador = 0
        for fila_row in ws_h3.iter_rows(min_row=2, min_col=1, max_col=11):
            # Considerar fila con datos si col A o col F tienen valor
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

    # ── VALIDACIÓN SAP PREVIA AL GUARDADO ────────────────────────────────────

    def _emitir_resumen_validacion_sap(self) -> None:
        """
        Emite un bloque de validación SAP consolidado justo antes del guardado.
        Resume todos los avisos detectados durante el proceso para que el usuario
        pueda corregir manualmente antes de importar en SAP.

        Categorías emitidas:
          📊 RESUMEN NUEVO/BORRAR — registros sin cambios, modificados, nuevos
                               y marcados para borrar.
                               · Modificados  → Z=X en Hoja 3 (condición de
                                 precio/vigencia distinta a la heredada).
                                 Su fila de Hoja 1 NO se marca como Nueva,
                                 porque el material ya existía.
                               · Nuevos       → AK=X en Hoja 1 (y Z=X/AO=X en
                                 H3/H2). Solo materiales sin match en el
                                 Purchasing original.
          ⚠ SOLAPAMIENTOS   — períodos cerrados automáticamente (puede haber filas
                               nuevas pendientes de carga en SAP)
          ⚠ REGISTROS OMIT. — filas saltadas por duplicado funcional
          ⚠ GAPS DE VIGENCIA — días sin cobertura de precio entre períodos
          ⚠ CENTROS SIN EXT. — centros activos sin ninguna fila nueva en H2
        """
        a = self._audit
        _n_gaps_activos = max(0, a["n_gaps"] - a.get("gaps_corregidos", 0))
        hay_avisos = (
            a["n_solapamientos"] > 0 or
            a["n_omitidos_dup_h2"] > 0 or
            a["n_omitidos_dup_h3"] > 0 or
            _n_gaps_activos > 0 or
            a["n_reducciones"] > 0 or
            bool(a["centros_sin_ext"])
        )

        sep = "─" * 60
        _log(self._parent, sep, "INFO")
        _log(self._parent, "RESUMEN VALIDACIÓN SAP — previo al guardado", "INFO")
        _log(self._parent, sep, "INFO")

        # ── 📊 Resumen NUEVO/BORRAR (regla: NUEVO solo ante cambio real) ──────
        n_sin_cambios = len(a.get("materiales_sin_cambios", set()))
        n_modificados = len(a.get("materiales_modificados", set()))
        n_nuevos_resumen = a["n_nuevos"]
        n_borrar      = len(a.get("materiales_baja", set()))
        _log(self._parent,
             f"  📊 Registros sin cambios   : {n_sin_cambios}", "INFO")
        _log(self._parent,
             f"  📊 Registros modificados   : {n_modificados}  (Hoja3: Z=X)", "INFO")
        _log(self._parent,
             f"  📊 Registros nuevos        : {n_nuevos_resumen}  (Hoja1: AK=X)", "INFO")
        _log(self._parent,
             f"  📊 Registros para borrar   : {n_borrar}  (BORRAR=X, NUEVO vacío)", "INFO")

        # Estadísticas base
        _log(self._parent,
             f"  Materiales heredados : {a['n_heredados']}  |  "
             f"Materiales nuevos : {a['n_nuevos']}  |  "
             f"Sin tarifa : {a['n_sin_tarifa']}",
             "INFO")
        _log(self._parent,
             f"  ID col A — max heredado : {a['max_id_col_a'] or '(ninguno numérico)'}  |  "
             f"Nuevos IDs : {a['max_id_col_a']+1}–{a['max_id_col_a']+a['n_nuevos'] if a['n_nuevos'] else '—'}",
             "INFO")

        if not hay_avisos:
            _log(self._parent,
                 "  ✅ Sin incidencias detectadas. El archivo puede importarse en SAP.",
                 "OK")
            _log(self._parent, sep, "INFO")
            return

        # ── ⚠ Solapamientos ──────────────────────────────────────────────────
        if a["n_solapamientos"] > 0:
            detalle = "; ".join(
                f"{k[2]}|{k[3]} → cerrado {v}"
                for k, v in list(a["solap_detalle"].items())[:5]
            )
            sufijo = f"... (+{len(a['solap_detalle'])-5} más)" if len(a["solap_detalle"]) > 5 else ""
            _log(self._parent,
                 f"  ⚠ SOLAPAMIENTOS: {a['n_solapamientos']} período(s) cerrado(s) automáticamente. "
                 f"SAP recibirá la fila nueva en la siguiente importación. "
                 f"Detalle: {detalle}{sufijo}",
                 "WARN")

        # ── ⚠ Registros omitidos por duplicado ───────────────────────────────
        n_dup_total = a["n_omitidos_dup_h2"] + a["n_omitidos_dup_h3"]
        if n_dup_total > 0:
            _log(self._parent,
                 f"  ⚠ REGISTROS OMITIDOS: {n_dup_total} fila(s) saltada(s) por duplicado funcional "
                 f"(H2: {a['n_omitidos_dup_h2']}, H3: {a['n_omitidos_dup_h3']}). "
                 f"Sin impacto en SAP — la combinación PIR ya existe.",
                 "WARN")

        # ── ⚠ Gaps de vigencia ───────────────────────────────────────────────
        if a["n_gaps"] > 0:
            n_corr = a.get("gaps_corregidos", 0)
            n_activos_g = max(0, a["n_gaps"] - n_corr)
            top_gaps = sorted(a["gaps_detalle"].items(), key=lambda x: -x[1])[:3]
            detalle = "; ".join(f"{k[2]}|{k[3]} ({dias}d)" for k, dias in top_gaps)
            if n_activos_g == 0:
                _log(self._parent,
                     f"  ✅ GAPS: {a['n_gaps']} detectado(s) → todos corregidos "
                     f"por ajuste automático de fechas (paso 10a). Sin gaps activos.",
                     "OK")
            else:
                _log(self._parent,
                     f"  ⚠ GAPS VIGENCIA: {n_activos_g} activo(s) "
                     f"({n_corr} corregido(s) automáticamente). "
                     f"Puede causar errores en órdenes de compra SAP. Top: {detalle}",
                     "WARN")

        # ── ⚠ Reducciones de vigencia ────────────────────────────────────────
        if a["n_reducciones"] > 0:
            _log(self._parent,
                 f"  ⚠ REDUCCIONES: {a['n_reducciones']} período(s) con J_nueva < J_anterior. "
                 f"La vigencia del precio se ha acortado respecto al Purchasing original.",
                 "WARN")

        # ── ⚠ Períodos absorbidos ────────────────────────────────────────────
        if a["n_absorbidos"] > 0:
            _log(self._parent,
                 f"  ✅ PERÍODOS ABSORBIDOS: {a['n_absorbidos']} fila(s) heredada(s) "
                 f"eliminadas directamente (nueva tarifa las engloba por completo). "
                 f"Hoja3 queda temporalmente consistente y sin períodos inválidos.",
                 "OK")

        # ── ⚠ Centros sin extensión ──────────────────────────────────────────
        if a["centros_sin_ext"]:
            _log(self._parent,
                 f"  ⚠ CENTROS SIN EXTENSIÓN: {a['centros_sin_ext']} — "
                 f"centros activos que no recibieron ninguna fila nueva en Hoja 2. "
                 f"Verificar si el material tiene cobertura para estos centros.",
                 "WARN")

        n_cat = sum([
            1 if a["n_solapamientos"] > 0 else 0,
            1 if n_dup_total > 0 else 0,
            1 if _n_gaps_activos > 0 else 0,
            1 if a["n_reducciones"] > 0 else 0,
            1 if a["n_absorbidos"] > 0 else 0,
            1 if a["centros_sin_ext"] else 0,
        ])
        _log(self._parent,
             f"  ── {n_cat} categoría(s) de aviso. Revisar antes de importar en SAP.",
             "WARN")
        _log(self._parent, sep, "INFO")

    def _intentar_eliminar(self, ruta: str):
        try:
            if os.path.isfile(ruta):
                os.remove(ruta)
        except Exception:
            pass
