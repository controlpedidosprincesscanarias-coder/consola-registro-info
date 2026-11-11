# -*- coding: utf-8 -*-

"""
crear_registro_info.py
----------------------
Módulo independiente — Botón «CREAR REGISTRO INFO»

Flujo completo según especificación:

  1. Seleccionar tarifa válida de la carpeta TARIFAS (sin RINFO, sin INCOMPLETO,
     con patrón  NOMBRE PROVEEDOR [dd-mm-yyyy].xlsx).
  2. Leer la celda C2 de la hoja TARIFA → nombre del proveedor.
  3. Buscar el proveedor en RECURSOS/Proveedores.xlsx por similitud (col A → D2,
     col B → C2 del nuevo RINFO).
  4. Buscar si ya existe un Purchasing Info Records*.xlsx compatible en TARIFAS.
     · Si existe  → informar y cerrar.
     · Si no      → ofrecer:
         a) Descargar desde SAP (guía de 7 pasos).
         b) Crear RINFO desde cero a partir de la plantilla.
  5. Creación del RINFO:
     · Crea un archivo NUEVO desde cero con las mismas 4 hojas del original.
     · Inserta las mismas cabeceras en la fila 1 de cada hoja.
     · Aplica el formato de la fila 2 del archivo original a todas las filas de datos.
     · Escribe D2 = nombre SAP (col A de Proveedores.xlsx).
     · Escribe C2 = código SAP (col B de Proveedores.xlsx).
     · Copia columna A de la tarifa → columna E del RINFO (datos desde fila 2).
     · Guarda como: RINFO NOMBRE ORIGINAL.xlsx  (mismo nombre de base que la tarifa).
  6. Confirmación final con opción de abrir el archivo.

Dependencias:  openpyxl, tkinter (stdlib)
Integración:   llamar a  iniciar_crear_registro_info(app, carpeta_proyecto)
               desde tarifa_app.App.

Central Compras Princess Canarias — (c) VAMA 2026
"""

import os
import re
from dataclasses import dataclass, field
import shutil
import unicodedata
import subprocess
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox
import datetime as _dt_mod          # módulo completo — usado en lógica de fechas de baja
from datetime import datetime
from copy import copy
import zipfile

import openpyxl

from app.comun.tarifa_constants import (
    resource_path,
    C_MID, C_BG, C_PANEL, C_TEXT, C_TEXT_LIGHT, C_HOVER, C_ACCENT2,
    C_ERROR, FONT_TITLE, FONT_SUB,
    FONT_LABEL, FONT_BOLD, FONT_SMALL,
)
from app.comun.tarifa_core import cargar_icono_ventana, ajustar_geometria_pantalla



# ─────────────────────────────────────────────────────────────────────────────
#  UTILIDAD POST-GUARDADO: compatibilidad SAP (sharedStrings + estructura OOXML)
# ─────────────────────────────────────────────────────────────────────────────

def _postprocesar_para_sap(ruta_xlsx: str, ruta_purchasing_ref: str) -> int:
    """
    Post-procesa un .xlsx ya guardado por openpyxl para que SAP pueda cargarlo.

    PROBLEMA RAÍZ confirmado mediante comparación XML directa (ZIP vs ZIP):
    openpyxl 3.x reescribe innecesariamente varios archivos de estructura
    OOXML introduciendo diferencias que SAP rechaza:

      1. Escribe t="inlineStr" en lugar de t="s" + sharedStrings.xml
         → SAP no puede leer los valores de texto y los ignora.
      2. Usa rutas absolutas (/xl/worksheets/...) en workbook.xml.rels
         → El Purchasing original usa rutas relativas (worksheets/...).
      3. Reescribe [Content_Types].xml, _rels/.rels y xl/workbook.xml
         con diferencias de atributos y orden de entradas.

    ESTRATEGIA: restaurar literalmente del Purchasing de referencia los
    archivos de estructura que openpyxl reescribe pero que NO contienen
    datos variables (son idénticos en cualquier Purchasing SAP), y generar
    sharedStrings.xml a partir de las hojas ya procesadas.

    Archivos restaurados del Purchasing de referencia (sin tocar):
      · [Content_Types].xml   (+ entrada sharedStrings si falta)
      · _rels/.rels
      · xl/_rels/workbook.xml.rels  (+ relación sharedStrings si falta)
      · xl/workbook.xml

    Archivos generados por esta función:
      · xl/sharedStrings.xml  (construido desde las hojas)

    Archivos que NO se tocan (los escribe openpyxl con los datos reales):
      · xl/worksheets/sheet1-4.xml  (convertidos de inlineStr a t="s")
      · xl/styles.xml

    Parámetros:
      ruta_xlsx          Ruta del .xlsx recién guardado por openpyxl.
      ruta_purchasing_ref Ruta del Purchasing original SAP usado como base.

    Devuelve el nº de cadenas únicas en sharedStrings.xml.
    """
    tmp = ruta_xlsx + ".tmp_sap"
    shutil.copy2(ruta_xlsx, tmp)

    with zipfile.ZipFile(tmp, "r") as zin:
        contents = {n: zin.read(n) for n in zin.namelist()}

    # Leer archivos de estructura del Purchasing de referencia
    with zipfile.ZipFile(ruta_purchasing_ref, "r") as zref:
        ref_names = set(zref.namelist())
        ref = {n: zref.read(n) for n in ref_names}

    sheet_names = [n for n in contents if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]

    # ── PASO 1: Convertir inlineStr → t="s" en todas las hojas ──────────────
    shared: list = []
    shared_index: dict = {}

    def _get_idx(text: str) -> int:
        if text not in shared_index:
            shared_index[text] = len(shared)
            shared.append(text)
        return shared_index[text]

    # Celda con texto: <c ...t="inlineStr"...><is><t...>TEXTO</t></is></c>
    _pat_texto = re.compile(
        rb'<c([^>]*?) t="inlineStr"([^>]*)>\s*<is>\s*<t(?:[^>]*)>(.*?)</t>\s*</is>\s*</c>',
        re.DOTALL,
    )
    # Celda self-closing vacía: <c ... t="inlineStr" />
    _pat_sc = re.compile(rb'<c([^>]*?) t="inlineStr"([^>]*)/>'  )
    # Celda con <is/> vacío: <c ...t="inlineStr"...><is/></c>
    _pat_is_v = re.compile(
        rb'<c([^>]*?) t="inlineStr"([^>]*)>\s*<is\s*/>\s*</c>', re.DOTALL
    )

    def _repl_texto(m):
        a1, a2, txt = m.group(1), m.group(2), m.group(3).decode("utf-8")
        idx = _get_idx(txt)
        return (
            f'<c{a1.decode("utf-8")} t="s"{a2.decode("utf-8")}><v>{idx}</v></c>'
        ).encode("utf-8")

    def _repl_vacio(m):
        a1, a2 = m.group(1).decode("utf-8"), m.group(2).decode("utf-8")
        return f'<c{a1}{a2}/>'.encode("utf-8")

    new_contents = dict(contents)
    for sn in sheet_names:
        data = contents[sn]
        data = _pat_is_v.sub(_repl_vacio, data)
        data = _pat_sc.sub(_repl_vacio, data)
        data = _pat_texto.sub(_repl_texto, data)
        new_contents[sn] = data

    # ── PASO 2: Generar xl/sharedStrings.xml ─────────────────────────────────
    def _xml_esc(s: str) -> str:
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    sst_items = "".join(
        f'<si><t xml:space="preserve">{_xml_esc(s)}</t></si>' for s in shared
    )
    sst_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        f'count="{len(shared)}" uniqueCount="{len(shared)}">{sst_items}</sst>'
    ).encode("utf-8")
    new_contents["xl/sharedStrings.xml"] = sst_xml

    # ── PASO 3: Restaurar archivos de estructura del Purchasing original ──────
    # Estos archivos son idénticos en cualquier Purchasing SAP y openpyxl
    # los reescribe innecesariamente con diferencias que SAP rechaza.
    RESTAURAR = [
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
    ]
    for nombre in RESTAURAR:
        if nombre in ref:
            new_contents[nombre] = ref[nombre]

    # workbook.xml.rels: restaurar del original pero asegurar que la
    # relación de sharedStrings existe (el original ya la tiene siempre).
    if "xl/_rels/workbook.xml.rels" in ref:
        rels_ref = ref["xl/_rels/workbook.xml.rels"].decode("utf-8")
        # Si por algún motivo el original no tuviera sharedStrings (improbable),
        # la añadimos para garantizarlo.
        if "sharedStrings" not in rels_ref:
            rids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_ref)]
            nxt = max(rids) + 1 if rids else 7
            rels_ref = rels_ref.replace(
                "</Relationships>",
                f'<Relationship Id="rId{nxt}" ' +
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006' +
                f'/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
                f"</Relationships>",
            )
        new_contents["xl/_rels/workbook.xml.rels"] = rels_ref.encode("utf-8")

    # [Content_Types].xml: asegurar entrada sharedStrings (el original ya la tiene).
    ct = new_contents["[Content_Types].xml"].decode("utf-8")
    if "sharedStrings" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/xl/sharedStrings.xml" ' +
            'ContentType="application/vnd.openxmlformats-officedocument' +
            '.spreadsheetml.sharedStrings+xml"/></Types>',
        )
        new_contents["[Content_Types].xml"] = ct.encode("utf-8")

    # ── PASO 4: Escribir el ZIP final ────────────────────────────────────────
    with zipfile.ZipFile(ruta_xlsx, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in new_contents.items():
            zout.writestr(name, data)

    os.remove(tmp)
    return len(shared)

# ─────────────────────────────────────────────────────────────────────────────
#  CONSTANTES INTERNAS
# ─────────────────────────────────────────────────────────────────────────────

FOLDER_TARIFAS   = "TARIFAS"
FOLDER_RECURSOS  = "RECURSOS"
ARCHIVO_PROV     = "Proveedores.xlsx"
ARCHIVO_TEMPLATE = "Purchasing Info Records.xlsx"   # Plantilla SAP en RECURSOS/


HOJA_TARIFA      = "TARIFA"
UMBRAL_SIMILITUD = 90       # % mínimo de coincidencia en Proveedores.xlsx
UMBRAL_PURCHASING = 45      # % mínimo para considerar un Purchasing compatible
URL_SAP = (
    "https://my405727.s4hana.cloud.sap/ui#Shell-home"
)

# Fechas FIJAS para Hoja 1 (Info general p.reg.info) cols I/J al crear el RINFO
# desde cero (CreadorRINFO). NO deben confundirse con las fechas de validez que
# el usuario introduce en VentanaFechasRINFO, las cuales solo se aplican en Hoja 3
# (precios/materiales) — ver self._fecha_inicio / self._fecha_fin en CreadorRINFO.
FECHA_INICIO_FIJA_HOJA1 = "01.01.2025"
FECHA_FIN_FIJA_HOJA1    = "31.12.2030"

# Patrón de nombre de tarifa válida:  * [dd-mm-yyyy].xlsx
_RE_PATRON_TARIFA     = re.compile(r'^.+\[\d{2}-\d{2}-\d{4}\]\.xlsx$', re.IGNORECASE)
_RE_INCOMPLETO        = re.compile(r'\[INCOMPLETO\]', re.IGNORECASE)
_RE_RINFO             = re.compile(r'^RINFO\s', re.IGNORECASE)

# Formas societales a ignorar en la comparación de nombres
_RE_FORMAS_SOC = re.compile(
    r'\b(S\.?\s*L\.?\s*U?\.?|S\.?\s*A\.?\s*U?\.?|S\.?\s*A\.?\s*S\.?|'
    r'S\.?\s*L\.?|S\.?\s*A\.?|LTD\.?|B\.?\s*V\.?|N\.?\s*V\.?|'
    r'SAS|SPA|AG|INC\.?|CO\.?)\b',
    re.IGNORECASE,
)

# Centros SAP de Canarias — 10 bloques de Hoja 2 (mismo orden que la spec)
CENTROS_SAP_CANARIAS = {
    "0110": "GUAYARMINA PRINCESS",
    "0120": "PRINCESS INSPIRE TENERIFE",
    "0210": "LA PALMA&TENEGUIA PRINCESS",
    "0310": "TAURITO PRINCESS",
    "0410": "SUITE PRINCESS",
    "0510": "MOGAN PRINCESS",
    "0610": "GRAN CANARIA PRINCESS",
    "0710": "MASPALOMAS&TABAIBA PRINCESS",
    "0810": "CLUB JANDIA PRINCESS",
    "0910": "FUERTEVENTURA PRINCESS",
}
# Lista ordenada para asignar bloque 0..9 → centro correspondiente
_CENTROS_LIST = list(CENTROS_SAP_CANARIAS.items())   # [(codigo, nombre), ...]

# Hojas estándar del Purchasing Info Records que se deben vaciar
HOJAS_PURCHASING = [
    "Info general p.reg.info",
    "Datos org.compr.p.reg.info",
    "Det.precios p.registros info",
    "Escalas p.cond.precio PIR",
]

# ─────────────────────────────────────────────────────────────────────────────
#  ESTRUCTURA HARDCODED DEL ARCHIVO PURCHASING INFO RECORDS
#  (detectada automáticamente del archivo adjunto Purchasing_Info_Records.xlsx)
#
#  Formato fila 1 (cabeceras):
#    · Font: Calibri 11, bold=True, color=theme:1 (blanco)
#    · Fill: solid, fgColor=theme:0 tint:-0.2499 (gris oscuro SAP)
#    · Border: thin negro en los 4 lados
#
#  Formato fila 2 común a las 3 primeras hojas:
#    · Font: Calibri 11, bold=True, color por defecto (negro)
#    · Fill: solid, fgColor=FFEDF4F8 (azul claro — alpha FF obligatorio)
#    · Border: thin negro en los 4 lados
#    · Alignment: default (sin alineación especial)
#
#  La hoja "Escalas p.cond.precio PIR" solo tiene fila 1 (cabeceras)
#  y su fila 2 es sin formato especial (Calibri 11, sin fill, sin borde).
#
#  Propiedades de hoja: freeze_panes="A2", col A width=40 en las 4 hojas.
# ─────────────────────────────────────────────────────────────────────────────

# Cabeceras detectadas por hoja (fila 1)
ESTRUCTURA_HOJAS = {
    "Info general p.reg.info": {
        "cabeceras": [
            "Registro info de compras",
            "Descripción de registro info de compras",
            "Proveedor",
            "Nombre del proveedor",
            "Material",
            "Texto breve material",
            "Grupo de artículos",
            "Descripción del grupo de artículos",
            "Disponible desde",
            "Disponible hasta",
            "1r recordatorio/Días de reclamación",
            "2º recordatorio/Días de reclamación",
            "3r recordatorio/Días de reclamación",
            "Unidad de medida de pedido",
            "Denominador de conversión",
            "Unidad de medida",
            "Numerador de conversión",
            "Fabricante",
            "Puntuación",
            "Unidad de puntos",
            "Proveedor anterior",
            "Nombre de proveedor precedente",
            "País/región de origen (certificado de origen)",
            "Proveedor regular",
            "Clasificar número de secuencia",
            "Grupo de materiales de proveedor",
            "Vendedor",
            "Número de material de proveedor",
            "Surtido parcial de proveedor",
            "Número de teléfono de proveedor",
            "Unidad de pedido variable",
            "Descripción de unidad de medida variable",
            "Clave clasificación",
            "Autor",
            "Fecha de creación",
            "Registro info marcado para borrado",
            "Nueva",
        ],
        # Formato de cada columna en fila 2: (number_format, bold)
        # Columnas con #,##0: 11,12,13,15,17 (1-based)
        # Columna 19: General  |  resto: @
        "num_formats": {
            11: "#,##0", 12: "#,##0", 13: "#,##0",
            15: "#,##0", 17: "#,##0",
            19: "General",
        },
        "formato_fila2": "datos",   # "datos" = fill+border+bold
    },
    "Datos org.compr.p.reg.info": {
        "cabeceras": [
            "Registro info de compras",
            "Tipo de registro info de compras",
            "Descripción de categoría de registro info",
            "Organización de compras",
            "Descripción de la organización de compras",
            "Centro",
            "Nombre de centro",
            "Grupo de compras",
            "Descripción de grupo de compras",
            "Material",
            "Texto breve material",
            "Grupo de artículos",
            "Cantidad de pedido estándar",
            "Plazo de entrega en días",
            "Aprovisionamiento automático",
            "Grupo de condiciones",
            "Control de confirmaciones",
            "Versión",
            "Verificación de factura basada en entrada de mercancías",
            "Incoterms",
            "Incoterms (parte 2)",
            "Ubicación de incoterms 1",
            "Cantidad de pedido máxima",
            "Cantidad mínima de pedido",
            "Tiempo mínimo de duración restante",
            "Indicador de períodos para fecha de vencimiento",
            "Ningún descuento",
            "Liquidación por comprobante no evaluada",
            "Confirmación de pedido obligatoria",
            "Tolerancia de exceso de suministro en %",
            "Tolerancia para faltas en suministro en %",
            "Control de fecha del precio",
            "Número AMD de proveedor requerido",
            "Indicador de impuestos",
            "País/Reg.decl.",
            "Inic.val.tipo impos.",
            "Entrega ilimitada permitida",
            "Autor",
            "Fecha de creación",
            "Organización de compras marcada para borrado",
            "Nueva",
        ],
        # Col 14: #,##0 | col 25: #,##0 | cols 13,23,24,30,31: General | resto: @
        "num_formats": {
            14: "#,##0", 25: "#,##0",
            13: "General", 23: "General", 24: "General",
            30: "General", 31: "General",
        },
        "formato_fila2": "datos",
    },
    "Det.precios p.registros info": {
        "cabeceras": [
            "Registro info",
            "Tipo de registro info de compras",
            "Descripción de categoría de registro info",
            "Organización de compras",
            "Centro",
            "Material",
            "Texto breve material",
            "Grupo de artículos",
            "Inicio de validez",
            "Fin validez",
            "ID grupo condiciones",
            "Orden de condiciones",
            "Tipo de condición",
            "Descripción de clase de condición",
            "Cantidad",
            "Unidad importe condición",
            "Determinación precios condición",
            "Unidad de determinación de precio de condición",
            "Unidad de medida base",
            "Clase de cálculo de condición",
            "Descripción de clase de cálculo de condición",
            "Numerador para conversión",
            "Denominador para conversión",
            "Escala",
            "Indicador borrado",
            "Nueva",
            # Columna nueva añadida por SAT en el Purchasing Info Records
            # (sesión 2026-09-29) -- col AA, al final, sin desplazar nada.
            "Unidad medida pedido",
        ],
        # Col 15: #,##0.00 | cols 17,22,23: #,##0 | resto: @
        "num_formats": {
            15: "#,##0.00",
            17: "#,##0", 22: "#,##0", 23: "#,##0",
        },
        "formato_fila2": "datos",
    },
    "Escalas p.cond.precio PIR": {
        "cabeceras": [
            "Registro info de compras",
            "Tipo de registro info de compras",
            "Descripción de categoría de registro info",
            "Organización de compras",
            "Centro",
            "Material",
            "Grupo de artículos",
            "Proveedor",
            "ID grupo condiciones",
            "Tipo de condición",
            "Inicio de validez",
            "Fin validez",
            "Tipo base de escala",
            "Base de escala",
            "Verificar clase de escala",
            "Verificar descripción de clase de escala",
            "Clase de escala",
            "Descripción de clase de escala",
            "Cantidad escala",
            "Unidad medida escala",
            "Valor de la escala",
            "Moneda de la escala",
            "Valor de condición",
            "Unidad de valor de condición",
            "Moneda condición",
            "Cantidad de unidad de precio",
            "UM unidad de precio",
            "Indicador borrado",
            "Nuevo",
        ],
        "num_formats": {},          # Todas General (sin datos en fila 2)
        "formato_fila2": "vacio",   # Sin fill ni borde especial
    },
}


# ─────────────────────────────────────────────────────────────────────────────
#  FUNCIÓN PRINCIPAL: copiar_template_registro_info
#  Copia el template SAP original desde RECURSOS/ y lo limpia de datos,
#  preservando intactos todos los metadatos internos, tipos de celda y
#  configuración que openpyxl no puede replicar al crear desde cero.
#
#  ⚠  Requiere que exista:
#       <carpeta_proyecto>/RECURSOS/Purchasing Info Records.xlsx
#  Si el template no está disponible, se lanza FileNotFoundError con instrucciones.
# ─────────────────────────────────────────────────────────────────────────────

def _parchear_column_dimension():
    """
    Parche temporal para openpyxl: los archivos SAP S/4HANA incluyen el
    atributo 's' (style index) en los elementos <col> del XML de hoja.
    openpyxl no reconoce ese atributo en ColumnDimension.__init__() y lanza
    TypeError al abrir el workbook.

    Este parche envuelve __init__ para absorber cualquier kwarg desconocido
    sin error. Se aplica solo durante la carga y se revierte inmediatamente.

    Devuelve la función original para poder restaurarla.
    """
    from openpyxl.worksheet.dimensions import ColumnDimension
    _init_original = ColumnDimension.__init__

    def _init_tolerante(self, worksheet, index='A', width=13, bestFit=False,
                        hidden=False, outlineLevel=0, outline_level=None,
                        collapsed=False, style=None, min=None, max=None,
                        customWidth=False, visible=None, auto_size=None,
                        **kwargs_ignorados):
        # kwargs_ignorados absorbe 's' y cualquier otro atributo SAP desconocido
        _init_original(
            self, worksheet, index=index, width=width, bestFit=bestFit,
            hidden=hidden, outlineLevel=outlineLevel, outline_level=outline_level,
            collapsed=collapsed, style=style, min=min, max=max,
            customWidth=customWidth, visible=visible, auto_size=auto_size,
        )

    ColumnDimension.__init__ = _init_tolerante
    return _init_original


def _restaurar_column_dimension(init_original):
    """Restaura ColumnDimension.__init__ a su implementación original."""
    from openpyxl.worksheet.dimensions import ColumnDimension
    ColumnDimension.__init__ = init_original


def copiar_template_registro_info(ruta_destino: str,
                                   carpeta_proyecto: str) -> openpyxl.Workbook:
    """
    Crea el archivo Purchasing Info Records copiando el template SAP original.

    1. Verifica que exista <carpeta_proyecto>/RECURSOS/Purchasing Info Records.xlsx.
       Si no existe, lanza FileNotFoundError con mensaje descriptivo para el usuario.
    2. Copia el template al destino (copia binaria exacta, preserva estructura SAP).
    3. Carga el workbook aplicando el parche ColumnDimension para tolerar el
       atributo 's' que SAP S/4HANA incluye en el XML interno y que openpyxl
       no reconoce (evita TypeError: ColumnDimension.__init__() got an unexpected
       keyword argument 's').
    4. Limpia los datos de las 4 hojas poniendo .value = None en todas las
       celdas desde fila 2 en adelante — SIN delete_rows, para preservar
       intacto el XML interno de SAP (estilos, validaciones, metadatos).
    5. Devuelve el workbook abierto para escritura posterior.

    Parámetros
    ----------
    ruta_destino      : ruta completa del nuevo RINFO a generar.
    carpeta_proyecto  : ruta raíz del proyecto (contiene la subcarpeta RECURSOS/).

    Devuelve
    --------
    openpyxl.Workbook  — el workbook ya cargado, listo para escribir datos.

    Lanza
    -----
    FileNotFoundError  — si el template SAP no está en RECURSOS/.
    """
    ruta_template = os.path.join(carpeta_proyecto, FOLDER_RECURSOS, ARCHIVO_TEMPLATE)

    if not os.path.isfile(ruta_template):
        raise FileNotFoundError(
            f"No se encontró el archivo original de Purchasing Info Records en:\n\n"
            f"   {ruta_template}\n\n"
            "Para crear un Registro Info es necesario colocar un archivo\n"
            "«Purchasing Info Records.xlsx» original (descargado de SAP)\n"
            f"en la carpeta RECURSOS del proyecto:\n\n"
            f"   {os.path.join(carpeta_proyecto, FOLDER_RECURSOS)}\n\n"
            "Una vez colocado el archivo, vuelva a intentarlo."
        )

    # ── 1. Copiar template al destino (copia binaria exacta) ──────────────────
    shutil.copy2(ruta_template, ruta_destino)

    # ── 2. Cargar el workbook con parche ColumnDimension ──────────────────────
    #    Los archivos SAP incluyen el atributo 's' en <col> del XML interno.
    #    openpyxl no lo reconoce → TypeError al abrir. El parche lo absorbe.
    init_original = _parchear_column_dimension()
    try:
        wb = openpyxl.load_workbook(ruta_destino)
    finally:
        _restaurar_column_dimension(init_original)

    # ── 3. Limpiar datos de las 4 hojas ──────────────────────────────────────
    #
    #  ⚠  NO se usa delete_rows(): ese método reescribe el XML interno de la
    #     hoja y destruye metadatos SAP (estilos, validaciones, atributos <col>)
    #     que openpyxl no puede reconstruir fielmente, provocando que SAP
    #     rechace el archivo o lo marque como «sin cambios» al cargarlo.
    #
    #  En su lugar se pone .value = None celda a celda desde fila 2 en
    #  adelante.  El formato, los anchos de columna, freeze_panes y toda la
    #  estructura XML del template quedan intactos.
    for nombre_hoja in HOJAS_PURCHASING:
        if nombre_hoja not in wb.sheetnames:
            continue
        ws = wb[nombre_hoja]
        if ws.max_row and ws.max_row > 1:
            for fila in ws.iter_rows(min_row=2):
                for celda in fila:
                    celda.value = None

    # NO se guarda aquí: el save final lo hace _ejecutar() tras escribir
    # todos los datos (artículos, correlativos, cruce LISTADO).
    return wb


# ─────────────────────────────────────────────────────────────────────────────
#  NOTA: crear_registro_info() (fallback desde cero) eliminada.
#  Única vía: copiar_template_registro_info() — requiere RECURSOS/Purchasing Info Records.xlsx
# ─────────────────────────────────────────────────────────────────────────────



def aplicar_formato_fila(ws: openpyxl.worksheet.worksheet.Worksheet,
                          fila: int,
                          nombre_hoja: str,
                          fila_ref: int = 2) -> None:
    """
    Aplica a la fila indicada exactamente el mismo formato que tiene la fila
    de referencia (por defecto la fila 2) del propio worksheet.

    Lee celda a celda la fila de referencia y copia font, fill, border,
    alignment y number_format a la fila destino — sin constantes hardcoded,
    fidelidad 100% con el archivo original independientemente del template usado.

    Parámetros
    ----------
    ws          : hoja de trabajo (Worksheet) ya abierta.
    fila        : número de fila destino donde aplicar el formato (>= 2).
    nombre_hoja : nombre de la hoja (para determinar max columnas).
    fila_ref    : fila de referencia desde la que copiar el formato (default: 2).
                  · Ruta «desde cero» (CreadorRINFO): siempre 2, que es la única
                    fila con estilos reales en el template limpio. Cada hoja usa
                    su propia fila 2, respetando el formato distinto de cada una.
                  · Ruta «desde Purchasing» (GeneradorDesdeExistente): usar la
                    última fila preexistente (primera_fila_nueva - 1) para que
                    el estilo copiado coincida con el de las filas heredadas.
    """
    if nombre_hoja in ESTRUCTURA_HOJAS:
        n_cols = len(ESTRUCTURA_HOJAS[nombre_hoja]["cabeceras"])
    else:
        n_cols = ws.max_column or 1

    # Si la fila destino ES la fila de referencia, no hay nada que copiar
    if fila == fila_ref:
        return

    for col_idx in range(1, n_cols + 1):
        celda_ref = ws.cell(row=fila_ref, column=col_idx)
        celda_dst = ws.cell(row=fila,     column=col_idx)
        celda_dst.font          = copy(celda_ref.font)
        celda_dst.fill          = copy(celda_ref.fill)
        celda_dst.border        = copy(celda_ref.border)
        celda_dst.alignment     = copy(celda_ref.alignment)
        celda_dst.number_format = celda_ref.number_format


# ─────────────────────────────────────────────────────────────────────────────
#  UTILIDADES DE NORMALIZACIÓN Y SIMILITUD
# ─────────────────────────────────────────────────────────────────────────────

def _sin_tildes(texto: str) -> str:
    """Elimina diacríticos (tildes, cedillas, etc.)."""
    normalizado = unicodedata.normalize("NFD", texto)
    return "".join(c for c in normalizado if unicodedata.category(c) != "Mn")


def _normalizar(texto: str) -> str:
    """Mayúsculas, sin tildes, solo alfanumérico y espacio, sin espacios dobles."""
    t = _sin_tildes(str(texto or "").upper())
    t = re.sub(r"[^A-Z0-9 ]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _normalizar_sin_soc(texto: str) -> str:
    """Como _normalizar pero eliminando además formas societales."""
    base = _normalizar(texto)
    return re.sub(r"\s+", " ", _RE_FORMAS_SOC.sub(" ", base)).strip()


# normalizar_codigo_sap vive en sap_helpers.py (sin tkinter) para que
# también pueda importarla test_crear_registro_info.py sin arrastrar
# la UI. Ver ese fichero para la implementación y el porqué.
from app.comun.sap_helpers import (  # noqa: E402
    normalizar_codigo_sap, ultima_fila_con_datos, parsear_fecha_celda,
    formatear_fecha_dmy, clave_pir_h3, material_de_clave_pir, formatear_clave_pir,
    debe_cerrar_fin_validez)


def _palabras_significativas(texto_norm: str) -> list:
    """Palabras con ≥3 caracteres excluyendo stopwords y formas societales."""
    STOP = {
        "SA", "SL", "SLU", "SAU", "LTD", "SAS", "BV", "NV",
        "THE", "AND", "LOS", "LAS", "DEL", "DE", "LA", "EL",
        "CON", "SIN", "POR", "SPA", "AG", "INC", "CO", "SAC",
    }
    return [w for w in texto_norm.split() if len(w) >= 3 and w not in STOP]


def _similitud(ref: str, candidato: str) -> int:
    """
    Puntuación 0–100 entre dos nombres de proveedor.
    Aplica comparación con y sin formas societales, devuelve el máximo.
    """
    def _puntuar(r: str, c: str) -> int:
        if not r or not c:
            return 0
        if r == c:
            return 100
        if r in c or c in r:
            return 95
        palabras = _palabras_significativas(r)
        if not palabras:
            return 0
        presentes  = sum(1 for p in palabras if p in c)
        ratio_pal  = presentes / len(palabras)
        bigramas_r = set(r[i:i+2] for i in range(len(r) - 1))
        bigramas_c = set(c[i:i+2] for i in range(len(c) - 1))
        if bigramas_r or bigramas_c:
            jaccard = len(bigramas_r & bigramas_c) / len(bigramas_r | bigramas_c)
        else:
            jaccard = 0.0
        return round(ratio_pal * 60 + jaccard * 40)

    s1 = _puntuar(_normalizar(ref),         _normalizar(candidato))
    s2 = _puntuar(_normalizar_sin_soc(ref), _normalizar_sin_soc(candidato))
    return max(s1, s2)


# ─────────────────────────────────────────────────────────────────────────────
#  UTILIDADES DE ARCHIVOS
# ─────────────────────────────────────────────────────────────────────────────

def _listar_tarifas_validas(carpeta_tarifas: str) -> list:
    resultado = []
    try:
        for entry in sorted(os.scandir(carpeta_tarifas),
                             key=lambda e: e.name.lower()):
            if not entry.is_file():
                continue
            nombre = entry.name
            if nombre.startswith("~$"):
                # Fichero temporal de bloqueo de Excel/Office (se crea al abrir
                # el archivo real y se borra al cerrarlo correctamente). Puede
                # quedar huérfano tras un cierre anómalo o si otro usuario en
                # red lo tiene abierto. Nunca es una tarifa válida — descartar.
                continue
            if not _RE_PATRON_TARIFA.match(nombre):
                continue
            if _RE_INCOMPLETO.search(nombre):
                continue
            if _RE_RINFO.match(nombre):
                continue
            ts    = entry.stat().st_mtime
            fecha = datetime.fromtimestamp(ts).strftime("%d-%m-%Y %H:%M")
            resultado.append({
                "nombre":         os.path.splitext(nombre)[0],
                "nombre_archivo": nombre,
                "ruta":           entry.path,
                "fecha_mod":      fecha,
            })
    except PermissionError:
        pass
    return resultado


def _leer_proveedor_tarifa(ruta_tarifa: str) -> str | None:
    try:
        wb = openpyxl.load_workbook(ruta_tarifa, read_only=True, data_only=True)
        hoja = wb[HOJA_TARIFA] if HOJA_TARIFA in wb.sheetnames else wb.active
        valor = hoja["C2"].value
        wb.close()
        return str(valor).strip() if valor else None
    except Exception:
        return None


def _buscar_proveedor_en_xlsx(nombre_dali: str, ruta_proveedores: str) -> list:
    candidatos = []
    try:
        wb = openpyxl.load_workbook(ruta_proveedores, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=1, values_only=True):
            if not row or not row[0]:
                continue
            nombre_sap = str(row[0]).strip()
            codigo_sap = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            score = _similitud(nombre_dali, nombre_sap)
            if score >= UMBRAL_SIMILITUD:
                candidatos.append((score, nombre_sap, codigo_sap))
        wb.close()
    except Exception:
        return []

    visto, unicos = set(), []
    for s, n, c in sorted(candidatos, key=lambda x: -x[0]):
        if n not in visto:
            visto.add(n)
            unicos.append((s, n, c))
    return unicos


def _buscar_purchasing_compatible(nombre_proveedor: str,
                                  carpeta_tarifas: str) -> dict | None:
    HOJA_P   = "Info general p.reg.info"
    COL_PROV = 3

    mejor = {"ruta": None, "score": 0, "proveedor_purchasing": None, "archivo": None}

    try:
        for entry in os.scandir(carpeta_tarifas):
            if not entry.is_file():
                continue
            if not (entry.name.lower().startswith("purchasing info records") and
                    entry.name.lower().endswith(".xlsx")):
                continue
            try:
                wb_p = openpyxl.load_workbook(
                    entry.path, read_only=True, data_only=True)
                if HOJA_P in wb_p.sheetnames:
                    ws_p = wb_p[HOJA_P]
                    for i, row in enumerate(ws_p.iter_rows(
                            min_row=2, values_only=True)):
                        if len(row) > COL_PROV and row[COL_PROV]:
                            prov = str(row[COL_PROV]).strip()
                            if prov:
                                score = _similitud(nombre_proveedor, prov)
                                if score > mejor["score"]:
                                    mejor["score"]                = score
                                    mejor["ruta"]                 = entry.path
                                    mejor["archivo"]              = entry.name
                                    mejor["proveedor_purchasing"] = prov
                        if i > 500:
                            break
                wb_p.close()
            except Exception:
                continue
    except PermissionError:
        pass

    if mejor["ruta"] and mejor["score"] >= UMBRAL_PURCHASING:
        return mejor
    return None


# ─────────────────────────────────────────────────────────────────────────────
#  PUNTO DE ENTRADA PÚBLICO
# ─────────────────────────────────────────────────────────────────────────────

def iniciar_crear_registro_info(app, carpeta_proyecto: str):
    """
    Punto de entrada llamado desde tarifa_app.App al pulsar «CREAR REGISTRO INFO».
    """
    if not carpeta_proyecto or not carpeta_proyecto.strip():
        messagebox.showwarning(
            "Sin proyecto",
            "Seleccione primero la carpeta del Proyecto antes de continuar.")
        return

    carpeta_tarifas = os.path.join(carpeta_proyecto.strip(), FOLDER_TARIFAS)
    if not os.path.isdir(carpeta_tarifas):
        messagebox.showerror(
            "Carpeta TARIFAS no encontrada",
            f"❌  No se encontró la carpeta TARIFAS en:\n\n"
            f"   {carpeta_tarifas}\n\n"
            "Verifique que la estructura del proyecto es correcta.")
        return

    archivos = _listar_tarifas_validas(carpeta_tarifas)
    _log(app, f"CREAR REGISTRO INFO: {len(archivos)} tarifa(s) válida(s) encontrada(s).", "INFO")

    if not archivos:
        messagebox.showwarning(
            "Sin tarifas disponibles",
            "❌  No se encontraron archivos de tarifa válidos en TARIFAS/.\n\n"
            "Los archivos deben cumplir:\n"
            "  • Patrón:  NOMBRE PROVEEDOR [dd-mm-yyyy].xlsx\n"
            "  • Sin etiqueta  [INCOMPLETO]\n"
            "  • Sin prefijo  RINFO\n\n"
            f"📁  Carpeta revisada:\n   {carpeta_tarifas}")
        return

    ventana = VentanaSeleccionTarifa(app, archivos, carpeta_tarifas)
    app.wait_window(ventana)


# ─────────────────────────────────────────────────────────────────────────────
#  HELPER: LOG EN CONSOLA
# ─────────────────────────────────────────────────────────────────────────────

def _eliminar_filas_agrupadas(ws: "openpyxl.worksheet.worksheet.Worksheet",
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


def _log(app_o_consola, mensaje: str, nivel: str = "INFO"):
    """
    Registra un mensaje en la consola de tarifa_app.

    Escala automáticamente la jerarquía de ventanas (master / _parent /
    _parent_ventana / _app_root) hasta encontrar un objeto que exponga
    log_consola() o log(), de modo que los mensajes siempre llegan a la
    CONSOLA DEL SISTEMA aunque se llame desde un Toplevel intermedio.
    """
    def _encontrar_app(obj, profundidad=10):
        """Devuelve el primer ancestro con log_consola o log, o None."""
        visitados = set()
        for _ in range(profundidad):
            if obj is None or id(obj) in visitados:
                break
            visitados.add(id(obj))
            if hasattr(obj, "log_consola"):
                return obj
            if hasattr(obj, "log"):
                return obj
            # Intentar subir por distintos atributos de jerarquía
            siguiente = (
                getattr(obj, "_app_root",        None) or
                getattr(obj, "master",           None) or
                getattr(obj, "_parent",          None) or
                getattr(obj, "_parent_ventana",  None)
            )
            obj = siguiente
        return None

    try:
        destino = _encontrar_app(app_o_consola)
        if destino is not None and hasattr(destino, "log_consola"):
            destino.log_consola(mensaje, nivel)
        elif destino is not None and hasattr(destino, "log"):
            destino.log(mensaje, nivel)
        else:
            print(f"[{nivel}] {mensaje}")
    except Exception:
        print(f"[{nivel}] {mensaje}")


class _LogAgrupado:
    """
    Agrupa mensajes de log repetitivos que se generan una vez por CENTRO
    para un mismo material (hasta 10 líneas casi idénticas, una por centro),
    y los condensa en una única línea resumen cuando cambia el grupo o se
    fuerza flush().

    No se pierde contexto: se conserva el material, el motivo, y la lista
    completa de centros/filas afectadas — solo se elimina la repetición
    literal del mismo mensaje 10 veces seguidas.

    Uso:
        agr = _LogAgrupado(parent)
        ...dentro del bucle...
        agr.registrar(clave=(mat,), fila=fila_num, centro=centro,
                      nivel="WARN", plantilla="H3 material {mat}: {motivo} — {resumen}",
                      mat=mat, motivo="...")
        ...
        agr.flush()   # SIEMPRE al terminar el bucle exterior
    """
    def __init__(self, parent):
        self._parent = parent
        self._clave = None
        self._nivel = "INFO"
        self._plantilla = None
        self._ctx: dict = {}
        self._filas: list = []
        self._centros: list = []

    def registrar(self, clave, fila, centro, nivel, plantilla, **ctx):
        if clave != self._clave:
            self.flush()
            self._clave = clave
            self._nivel = nivel
            self._plantilla = plantilla
            self._ctx = ctx
        self._filas.append(fila)
        self._centros.append(str(centro))

    def flush(self):
        if self._clave is None:
            return
        n = len(self._filas)
        if n == 1:
            resumen = f"fila {self._filas[0]} (centro {self._centros[0]})"
        else:
            resumen = (f"{n} centros ({','.join(self._centros)}), "
                       f"filas {self._filas[0]}-{self._filas[-1]}")
        try:
            msg = self._plantilla.format(resumen=resumen, **self._ctx)
            _log(self._parent, msg, self._nivel)
        finally:
            self._clave = None
            self._ctx = {}
            self._filas = []
            self._centros = []


def _safe_after(widget, fn, delay: int = 0):
    """
    Llama widget.after(delay, fn) solo si el widget todavía existe.

    Evita el error «bad window path name» que se produce cuando el hilo
    secundario intenta programar una llamada en un widget ya destruido.

    Uso:
        _safe_after(self._parent, _abrir_ventana)       # inmediato
        _safe_after(self._parent, lambda: ..., 80)      # con retardo

    Si el widget está destruido, la función fn NO se ejecuta y se devuelve
    None sin lanzar ninguna excepción.
    """
    try:
        if widget is not None and widget.winfo_exists():
            widget.after(delay, fn)
        else:
            fn()   # fallback: ejecutar directamente si no hay widget válido
    except Exception:
        try:
            fn()   # último recurso
        except Exception:
            pass



# ─────────────────────────────────────────────────────────────────────────────
#  PASO 1 — VENTANA DE SELECCIÓN DE TARIFA
# ─────────────────────────────────────────────────────────────────────────────

class VentanaSeleccionTarifa(tk.Toplevel):

    def __init__(self, parent, archivos: list, carpeta_tarifas: str):
        super().__init__(parent)
        self.title("Crear Registro Info — Seleccionar tarifa")
        self.resizable(True, True)
        self.minsize(680, 360)
        self.configure(bg=C_BG)
        self._intentar_icono()

        self._parent          = parent
        self._archivos        = archivos
        self._carpeta_tarifas = carpeta_tarifas
        self._procesando      = False   # guard anti-reentrada

        self._construir_ui()
        self.grab_set()
        self.transient(parent)
        self._centrar(parent)

    def _centrar(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + parent.winfo_width()  // 2
        py = parent.winfo_rooty() + parent.winfo_height() // 2
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px - w // 2}+{py - h // 2}")

    def _intentar_icono(self):
        cargar_icono_ventana(self)

    def _construir_ui(self):
        hdr = tk.Frame(self, bg=C_MID, pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📋  Crear Registro Info SAP",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Seleccione el archivo de tarifa con el que desea trabajar:",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        frm_tabla = tk.Frame(self, bg=C_BG, padx=14, pady=10)
        frm_tabla.pack(fill="both", expand=True)

        estilo = ttk.Style()
        estilo.configure("CRI.Treeview",
                          background=C_PANEL, fieldbackground=C_PANEL,
                          foreground=C_TEXT, rowheight=26, font=FONT_LABEL)
        estilo.configure("CRI.Treeview.Heading",
                          background=C_MID, foreground="white",
                          font=FONT_BOLD, relief="flat")
        estilo.map("CRI.Treeview",
                    background=[("selected", "#285FA8")],
                    foreground=[("selected", "white")])

        cols = ("nombre", "fecha_mod")
        self.tree = ttk.Treeview(frm_tabla, columns=cols, show="headings",
                                  selectmode="browse", style="CRI.Treeview")
        self.tree.heading("nombre",    text="Archivo de tarifa")
        self.tree.heading("fecha_mod", text="Última modificación")
        self.tree.column("nombre",    width=450, anchor="w")
        self.tree.column("fecha_mod", width=170, anchor="center")

        sb = ttk.Scrollbar(frm_tabla, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.tree.tag_configure("par",   background="#F0F5FB", foreground=C_TEXT)
        self.tree.tag_configure("impar", background="white",   foreground=C_TEXT)

        primer_item = None
        for i, archivo in enumerate(self._archivos):
            tag = "par" if i % 2 == 0 else "impar"
            self.tree.insert("", "end", iid=archivo["ruta"],
                              values=(f"  {archivo['nombre']}", archivo["fecha_mod"]),
                              tags=(tag,))
            if primer_item is None:
                primer_item = archivo["ruta"]

        if primer_item:
            self.tree.selection_set(primer_item)

        self.tree.bind("<Double-1>", lambda e: self._confirmar())

        frm_btn = tk.Frame(self, bg=C_BG, padx=14, pady=10)
        frm_btn.pack(fill="x", side="bottom")

        tk.Button(frm_btn, text="✔  Continuar",
                  command=self._confirmar,
                  bg="#285FA8", fg="white", font=FONT_BOLD,
                  padx=16, pady=7, cursor="hand2", relief="flat",
                  activebackground="#3A7AC8").pack(side="right", padx=(6, 0))

        tk.Button(frm_btn, text="Cancelar", command=self.destroy,
                  font=FONT_LABEL, padx=14, pady=6, cursor="hand2",
                  relief="flat", bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(side="right")

    def _confirmar(self):
        if self._procesando:          # evita reentrada por deiconify/grab de Tk
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sin selección",
                                   "Seleccione una tarifa de la lista antes de continuar.",
                                   parent=self)
            return
        ruta_sel   = sel[0]
        nombre_sel = next(
            (a["nombre_archivo"] for a in self._archivos if a["ruta"] == ruta_sel), "")

        _log(self._parent, f"Tarifa seleccionada: {nombre_sel}", "INFO")
        self._procesando = True        # bloquear re-llamadas mientras se procesa
        self._procesar_tarifa(ruta_sel, nombre_sel)

    def _procesar_tarifa(self, ruta_tarifa: str, nombre_archivo: str):
        nombre_proveedor_tarifa = _leer_proveedor_tarifa(ruta_tarifa)

        if not nombre_proveedor_tarifa:
            messagebox.showerror(
                "Proveedor no encontrado",
                f"❌  No se pudo leer el nombre de proveedor de la celda C2\n"
                f"   de la hoja '{HOJA_TARIFA}' en el archivo:\n\n"
                f"   {nombre_archivo}\n\n"
                "Compruebe que la hoja existe y que C2 contiene el nombre del proveedor.",
                parent=self)
            _log(self._parent, f"Proveedor no legible en C2 de: {nombre_archivo}", "ERROR")
            return

        _log(self._parent,
             f"Proveedor detectado en tarifa (C2): «{nombre_proveedor_tarifa}»", "OK")

        carpeta_proyecto  = os.path.dirname(self._carpeta_tarifas)
        ruta_proveedores  = os.path.join(carpeta_proyecto, FOLDER_RECURSOS, ARCHIVO_PROV)
        nombre_sap        = None
        codigo_sap        = None

        if not os.path.isfile(ruta_proveedores):
            _log(self._parent,
                 f"Proveedores.xlsx no encontrado en: {os.path.dirname(ruta_proveedores)}",
                 "WARN")
        else:
            candidatos = _buscar_proveedor_en_xlsx(nombre_proveedor_tarifa, ruta_proveedores)
            _log(self._parent,
                 f"Proveedores.xlsx — {len(candidatos)} coincidencia(s) (≥{UMBRAL_SIMILITUD}%).",
                 "INFO")

            if len(candidatos) == 1:
                _, nombre_sap, codigo_sap = candidatos[0]
                _log(self._parent,
                     f"Proveedor SAP: «{nombre_sap}» [cód. {codigo_sap}]", "OK")
            elif len(candidatos) > 1:
                nombre_sap, codigo_sap = self._dialogo_seleccion_proveedor(
                    nombre_proveedor_tarifa, candidatos)
                if nombre_sap:
                    _log(self._parent,
                         f"Usuario seleccionó proveedor SAP: «{nombre_sap}» "
                         f"[cód. {codigo_sap}]", "OK")
                else:
                    _log(self._parent,
                         "Usuario no seleccionó ningún proveedor SAP — "
                         "se usará el nombre de la tarifa.", "WARN")
            else:
                _log(self._parent,
                     f"Sin coincidencia ≥{UMBRAL_SIMILITUD}% para "
                     f"«{nombre_proveedor_tarifa}» en Proveedores.xlsx.", "WARN")
                messagebox.showwarning(
                    "Proveedor no localizado en Proveedores.xlsx",
                    f"⚠  No se encontró ninguna coincidencia suficiente\n"
                    f"   para el proveedor:\n\n"
                    f"   «{nombre_proveedor_tarifa}»\n\n"
                    "Se usará el nombre tal como aparece en la tarifa.\n\n"
                    "Si el nombre es incorrecto, revise Proveedores.xlsx\n"
                    "o corrija la celda C2 de la tarifa.",
                    parent=self)

        if not nombre_sap:
            nombre_sap  = nombre_proveedor_tarifa
            codigo_sap  = ""

        _log(self._parent,
             "Buscando Purchasing Info Records compatible en TARIFAS/…", "INFO")
        purchasing = _buscar_purchasing_compatible(nombre_proveedor_tarifa,
                                                   self._carpeta_tarifas)

        if purchasing:
            score_txt = str(purchasing["score"]) + "%"
            _log(self._parent,
                 f"Purchasing encontrado: {purchasing['archivo']} "
                 f"(score {score_txt}) — preguntando al usuario cómo continuar.", "OK")
            self.withdraw()
            VentanaPurchasingEncontrado(
                self,
                ruta_tarifa,
                nombre_archivo,
                nombre_proveedor_tarifa,
                nombre_sap,
                codigo_sap,
                purchasing,
                self._carpeta_tarifas,
                app_root=self._parent,
            )
            return

        _log(self._parent,
             "No se encontró Purchasing Info Records compatible.", "WARN")
        self.withdraw()
        VentanaOpcionesSinPurchasing(
            self, ruta_tarifa, nombre_archivo,
            nombre_proveedor_tarifa, nombre_sap, codigo_sap,
            self._carpeta_tarifas,
            app_root=self._parent)

    def _dialogo_seleccion_proveedor(self, nombre_dali: str,
                                     candidatos: list) -> tuple:
        resultado = [None, None]

        dlg = tk.Toplevel(self)
        dlg.title("Seleccionar proveedor SAP")
        dlg.configure(bg=C_BG)
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.transient(self)

        hdr_d = tk.Frame(dlg, bg="#285FA8", pady=10, padx=16)
        hdr_d.pack(fill="x")
        tk.Label(hdr_d, text="🔍  Seleccione el proveedor SAP correcto",
                 bg="#285FA8", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr_d,
                 text=f"Se encontraron {len(candidatos)} coincidencias en Proveedores.xlsx.",
                 bg="#285FA8", fg="#C0D8F8", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        cuerpo = tk.Frame(dlg, bg=C_BG, padx=18, pady=14)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text=f"Nombre en la tarifa:  «{nombre_dali}»",
                 bg=C_BG, fg="#888",
                 font=("Segoe UI", 9, "italic")).pack(anchor="w")
        tk.Label(cuerpo,
                 text="Elija el proveedor SAP con el que desea crear el Registro Info:",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 10))

        for score, nombre, codigo in candidatos:
            etiqueta = (f"  {nombre}  [{codigo}]  —  {score}% coincidencia"
                        if codigo else f"  {nombre}  —  {score}% coincidencia")

            def _elegir(n=nombre, c=codigo):
                resultado[0] = n
                resultado[1] = c
                dlg.destroy()

            tk.Button(cuerpo, text=etiqueta,
                      command=_elegir,
                      bg="#EEF2F8", fg="#1A2E4A",
                      font=FONT_LABEL, padx=12, pady=9,
                      cursor="hand2", relief="flat",
                      activebackground="#D0D8F0",
                      anchor="w").pack(fill="x", pady=2)

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(10, 6))

        tk.Button(cuerpo,
                  text="Ninguno — usar el nombre de la tarifa",
                  command=dlg.destroy,
                  bg="#D8DEE8", font=FONT_LABEL, padx=12, pady=7,
                  cursor="hand2", relief="flat",
                  activebackground="#C0C8D8").pack(fill="x", pady=(0, 4))

        dlg.update_idletasks()
        sw = dlg.winfo_screenwidth()
        w  = min(560, sw - 60)
        dlg.geometry(f"{w}+{(sw - w) // 2}+{dlg.winfo_screenheight() // 5}")
        self.wait_window(dlg)

        return (resultado[0], resultado[1])


# ─────────────────────────────────────────────────────────────────────────────
#  VENTANA INFORMATIVA: TEMPLATE SAP AUSENTE
# ─────────────────────────────────────────────────────────────────────────────

class VentanaTemplateAusente(tk.Toplevel):
    """
    Informa al usuario de que no se encontró el template SAP e indica
    exactamente qué archivo debe colocar y dónde.
    La ventana padre permanece visible para que pueda reintentar.
    """

    def __init__(self, parent, ruta_recursos: str, nombre_template: str):
        super().__init__(parent)
        self.title("Plantilla SAP no encontrada")
        self.configure(bg=C_BG)
        self.resizable(False, False)
        cargar_icono_ventana(self)
        self._ruta_recursos   = ruta_recursos
        self._nombre_template = nombre_template
        self._construir_ui()
        self._ajustar_geometria()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grab_set()
        self.transient(parent)

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(580, sw - 40)
        h  = self.winfo_reqheight() + 20
        x  = (sw - w) // 2
        y  = max(20, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _construir_ui(self):
        hdr = tk.Frame(self, bg="#8B1A1A", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="\u274c  Plantilla SAP no encontrada",
                 bg="#8B1A1A", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="No se puede crear el Registro Info sin el archivo original de SAP.",
                 bg="#8B1A1A", fg="#F4AAAA", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=16)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text=("Para crear un Registro Info es imprescindible disponer de un\n"
                       "archivo \'Purchasing Info Records\' original descargado de SAP.\n"
                       "Este archivo se usa como plantilla y garantiza que el RINFO\n"
                       "mantiene exactamente la estructura y formato que espera SAP."),
                 bg=C_BG, fg=C_TEXT, font=FONT_LABEL,
                 justify="left").pack(anchor="w", pady=(0, 12))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(cuerpo, text="Archivo necesario:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  \U0001f4c4  {self._nombre_template}",
                 bg="#FFF3F3", fg="#8B1A1A",
                 font=("Segoe UI", 10, "bold"),
                 padx=10, pady=7, relief="flat").pack(fill="x", pady=(4, 10))

        tk.Label(cuerpo, text="Debe colocarlo en la carpeta:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  \U0001f4c1  {self._ruta_recursos}",
                 bg="#F5F5F5", fg="#333",
                 font=FONT_LABEL, padx=10, pady=6, relief="flat",
                 wraplength=520, justify="left").pack(fill="x", pady=(4, 10))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(cuerpo, text="¿Cómo obtener el archivo?",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo,
                 text=("  1.  Use el botón \'Descargar desde SAP\' para obtener\n"
                       "       el Purchasing Info Records original.\n"
                       "  2.  Copie el archivo descargado a la carpeta RECURSOS indicada.\n"
                       "  3.  Vuelva a pulsar \'Crear nuevo Registro Info\'."),
                 bg=C_BG, fg="#555", font=FONT_LABEL,
                 justify="left").pack(anchor="w", pady=(4, 12))

        tk.Button(cuerpo, text="Entendido — Cerrar",
                  command=self.destroy,
                  bg=C_MID, fg=C_TEXT_LIGHT,
                  font=FONT_BOLD, padx=14, pady=10,
                  cursor="hand2", relief="flat",
                  activebackground=C_HOVER,
                  activeforeground=C_ACCENT2).pack(fill="x")


# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4 (PURCHASING ENCONTRADO) — VentanaPurchasingEncontrado
#
#  Informa al usuario de que se ha encontrado un Purchasing compatible,
#  pregunta si desea continuar usándolo o comenzar desde cero, y a
#  continuación ejecuta SIEMPRE el flujo estándar de recogida de datos:
#      teléfono → fechas → centros
#  Una vez completado ese flujo, bifurca según la elección del usuario:
#      · CONTINUAR  → GeneradorDesdeExistente (cruza materiales con el Purchasing)
#      · DESDE CERO → CreadorRINFO (crea desde la plantilla SAP)
# ─────────────────────────────────────────────────────────────────────────────

class VentanaPurchasingEncontrado(tk.Toplevel):
    """
    Ventana que aparece cuando se detecta un Purchasing Info Records compatible.

    Muestra la información del Purchasing encontrado, pregunta al usuario si
    desea continuar usándolo o empezar desde cero, y a continuación ejecuta
    el flujo estándar (teléfono → fechas → centros) independientemente de
    la elección.  Solo al final del flujo se bifurca el destino.

    Parámetros
    ----------
    parent_ventana   : VentanaSeleccionTarifa (llamante, ya ocultada con withdraw).
    ruta_tarifa      : ruta del archivo de tarifa seleccionado.
    nombre_archivo   : nombre de archivo de la tarifa (con extensión).
    nombre_dali      : nombre de proveedor leído de C2 de la tarifa.
    nombre_sap       : nombre SAP del proveedor (de Proveedores.xlsx).
    codigo_sap       : código SAP del proveedor.
    purchasing       : dict devuelto por _buscar_purchasing_compatible().
    carpeta_tarifas  : ruta de la carpeta TARIFAS del proyecto.
    app_root         : referencia al App raíz (para log y barra de progreso).
    """

    def __init__(self, parent_ventana, ruta_tarifa: str, nombre_archivo: str,
                 nombre_dali: str, nombre_sap: str, codigo_sap: str,
                 purchasing: dict, carpeta_tarifas: str, app_root=None):
        super().__init__(parent_ventana.master if hasattr(parent_ventana, "master")
                         else parent_ventana)
        self.title("Registro Info — Purchasing Info Records encontrado")
        self.resizable(True, True)
        self.configure(bg=C_BG)
        cargar_icono_ventana(self)

        self._parent_ventana  = parent_ventana
        self._ruta_tarifa     = ruta_tarifa
        self._nombre_archivo  = nombre_archivo
        self._nombre_dali     = nombre_dali
        self._nombre_sap      = nombre_sap
        self._codigo_sap      = codigo_sap
        self._purchasing      = purchasing
        self._carpeta_tarifas = carpeta_tarifas
        self._app_root        = app_root

        # La elección del usuario se guardará aquí antes de iniciar el flujo
        self._eleccion = None   # "continuar" | "desde_cero"

        self._construir_ui()
        self.grab_set()
        self.transient(self.master)
        self._ajustar_geometria()

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(700, sw - 40)
        h  = min(sh - 60, 560)
        x  = (sw - w) // 2
        y  = 20
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(560, 420)

    def _construir_ui(self):
        etiqueta_prov = (f"{self._nombre_sap}  [{self._codigo_sap}]"
                         if self._codigo_sap else self._nombre_sap)

        hdr = tk.Frame(self, bg="#285FA8", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr,
                 text="✔  Se ha encontrado un Registro Info para este proveedor",
                 bg="#285FA8", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Se ha localizado un Purchasing Info Records compatible en la carpeta TARIFAS.",
                 bg="#285FA8", fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=14)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text="Proveedor identificado en la tarifa (C2):",
                 bg=C_BG, fg="#999",
                 font=("Segoe UI", 8, "italic")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  {self._nombre_dali}",
                 bg="#F8F4FF", fg="#555",
                 font=FONT_LABEL, padx=10, pady=5, relief="flat").pack(fill="x")

        tk.Frame(cuerpo, bg=C_BG, height=8).pack()

        tk.Label(cuerpo,
                 text="Proveedor SAP a usar:",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  {etiqueta_prov}",
                 bg="#EEF2F8", fg="#285FA8",
                 font=("Segoe UI", 11, "bold"),
                 padx=10, pady=8, relief="flat").pack(fill="x")

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=10)

        tk.Label(cuerpo,
                 text="Purchasing Info Records compatible encontrado:",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo,
                 text=f"  📋  {self._purchasing['archivo']}  "
                      f"({self._purchasing['score']}% coincidencia)",
                 bg="#EBF4FF", fg="#1A3A6A",
                 font=FONT_LABEL, padx=10, pady=7, relief="flat").pack(fill="x")

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=10)

        tk.Label(cuerpo,
                 text="¿Desea continuar usando este Purchasing o comenzar desde cero?",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 10, "bold"),
                 wraplength=600, justify="left").pack(anchor="w", pady=(0, 10))

        tk.Button(cuerpo,
                  text="📋  Continuar usando este Purchasing Info Records",
                  command=self._elegir_continuar,
                  bg="#285FA8", fg="white",
                  font=FONT_BOLD, padx=12, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#3A7AC8").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo,
                  text=f"🆕  Comenzar desde cero  —  {etiqueta_prov}",
                  command=self._elegir_desde_cero,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=12, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo, text="Cancelar",
                  command=self._cancelar,
                  font=FONT_LABEL, padx=12, pady=8,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(fill="x")

    # ── Elección del usuario ───────────────────────────────────────────────────

    def _elegir_continuar(self):
        self._eleccion = "continuar"
        _log(self._parent_ventana,
             "Usuario eligió: Continuar usando el Purchasing existente.", "INFO")
        self._iniciar_flujo_estandar()

    def _elegir_desde_cero(self):
        # Verificar template SAP antes de continuar
        carpeta_proyecto = os.path.dirname(self._carpeta_tarifas)
        ruta_template    = os.path.join(carpeta_proyecto, FOLDER_RECURSOS, ARCHIVO_TEMPLATE)
        if not os.path.isfile(ruta_template):
            ruta_recursos = os.path.join(carpeta_proyecto, FOLDER_RECURSOS)
            _log(self._parent_ventana,
                 "Template SAP no encontrado en RECURSOS/ — creación abortada.", "ERROR")
            VentanaTemplateAusente(self, ruta_recursos, ARCHIVO_TEMPLATE)
            return
        self._eleccion = "desde_cero"
        _log(self._parent_ventana,
             "Usuario eligió: Comenzar desde cero (ignorar Purchasing).", "INFO")
        self._iniciar_flujo_estandar()

    def _cancelar(self):
        _log(self._parent_ventana, "Acción cancelada por el usuario.", "WARN")
        self.destroy()
        self._parent_ventana.withdraw()

    # ── Flujo estándar: teléfono → fechas → centros → bifurcación ─────────────

    def _iniciar_flujo_estandar(self):
        """
        Lanza el flujo estándar de recogida de datos (igual que en
        VentanaOpcionesSinPurchasing._opcion_crear_desde_cero).
        Al terminar, bifurca según self._eleccion.
        """
        self._telefono_pendiente = None
        self._fechas_pendientes  = (None, None)
        self.destroy()   # Cierra esta ventana; el flujo continúa en el padre
        self._parent_ventana.withdraw()   # re-ocultar: destroy() del hijo revierte el withdraw() previo
        VentanaTelefonoRINFO(self._parent_ventana, self._tras_telefono)

    def _tras_telefono(self, resultado):
        """Callback de VentanaTelefonoRINFO."""
        if resultado is False:
            _log(self._parent_ventana,
                 "Creación cancelada en la pantalla de teléfono.", "WARN")
            self._parent_ventana.withdraw()
            return
        self._telefono_pendiente = resultado
        VentanaFechasRINFO(self._parent_ventana, self._tras_fechas)

    def _tras_fechas(self, resultado):
        """Callback de VentanaFechasRINFO."""
        if resultado is False:
            _log(self._parent_ventana,
                 "Creación cancelada en la pantalla de fechas.", "WARN")
            self._parent_ventana.withdraw()
            return
        self._fechas_pendientes = resultado
        fecha_inicio, fecha_fin = resultado
        _log(self._parent_ventana,
             f"Fechas validez: inicio={fecha_inicio}  fin={fecha_fin}", "OK")
        VentanaCentrosRINFO(self._parent_ventana, self._tras_centros)

    def _tras_centros(self, resultado):
        """Callback de VentanaCentrosRINFO. Aquí bifurca según la elección."""
        if resultado is False:
            _log(self._parent_ventana,
                 "Creación cancelada en la pantalla de centros.", "WARN")
            self._parent_ventana.withdraw()
            return

        codigos = ", ".join(c for c, _ in resultado)
        _log(self._parent_ventana,
             f"Centros seleccionados ({len(resultado)}): {codigos}", "OK")

        if self._eleccion == "continuar":
            _log(self._parent_ventana,
                 "Iniciando generación desde Purchasing existente.", "INFO")
            _pv = self._parent_ventana
            _pv.withdraw()  # ocultar sin destruir antes de lanzar el generador
            GeneradorDesdeExistente(
                self._app_root,
                self._ruta_tarifa,
                self._nombre_archivo,
                self._nombre_sap,
                self._codigo_sap,
                self._purchasing,
                self._carpeta_tarifas,
                app_root=self._app_root,
                telefono_proveedor=self._telefono_pendiente,
                fecha_inicio=self._fechas_pendientes[0],
                fecha_fin=self._fechas_pendientes[1],
                centros_activos=resultado,
            )
        else:
            _log(self._parent_ventana,
                 "Iniciando creación de RINFO desde cero.", "INFO")
            _pv = self._parent_ventana
            _pv.withdraw()  # ocultar sin destruir antes de lanzar el creador
            CreadorRINFO(
                self._app_root,
                self._ruta_tarifa,
                self._nombre_archivo,
                self._nombre_sap,
                self._codigo_sap,
                self._carpeta_tarifas,
                app_root=self._app_root,
                telefono_proveedor=self._telefono_pendiente,
                fecha_inicio=self._fechas_pendientes[0],
                fecha_fin=self._fechas_pendientes[1],
                centros_activos=resultado,
            )


# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4 (PURCHASING ENCONTRADO) — GeneradorDesdeExistente
#
#  Cuando existe un Purchasing Info Records compatible con el proveedor,
#  se usa como base para generar un nuevo archivo:
#      {nombre_base_tarifa} [dd-mm-yyyy].xlsx
#  en la misma carpeta TARIFAS.
#
#  Flujo:
#    1. Leer col E (Material) del Purchasing existente → set de códigos ya presentes.
#    2. Leer col A (CODIGO SAP) del archivo nuevo (tarifa) → lista de códigos nuevos.
#    3. Insertar al final del Purchasing SOLO los códigos que no existan en col E,
#       aplicando el formato de fila 2 (aplicar_formato_fila) y numerando
#       correlativamente en col A según la lógica SAP (max_correlativo + 1, + 2, …).
#    4. Guardar como nuevo archivo: {nombre_base_tarifa} [dd-mm-yyyy].xlsx.
#    5. Mostrar ventana de confirmación.
# ─────────────────────────────────────────────────────────────────────────────

def _localizar_ruta_listado_dali(carpeta_tarifas, carpeta_proyecto=None, parent=None):
    """
    Localiza el fichero LISTADO CODIGOS DALI - SAP.xlsx: primero por
    ruta exacta (raíz del proyecto, RECURSOS, TARIFAS) y si no aparece,
    por patrón de nombre en esas mismas 3 carpetas (cubre variantes de
    mayúsculas/guiones bajos). Devuelve la ruta encontrada, o None si
    no existe en ningún sitio. No abre ni carga el fichero.

    Función de módulo (no método) porque la usan dos clases distintas
    sin relación de herencia entre sí: GeneradorDesdeExistente (flujo
    "desde Purchasing existente") y CreadorRINFO (flujo "crear desde
    cero"). `parent` es opcional, solo para el log de "encontrado por
    patrón" (pasar self._parent si se quiere ese log).

    Antes esta misma lógica de localización (~30 líneas) estaba
    copiada y pegada en _cargar_listado_dali, _cruzar_con_listado_dali,
    _rellenar_hoja2 y _rellenar_hoja3 — con una inconsistencia real
    entre copias: solo _cargar_listado_dali NO usaba el fallback de
    carpeta_proyecto que las otras tres sí usaban. Consolidado aquí
    con la versión más defensiva.

    NOTA: una primera versión de esta consolidación la dejó como método
    de GeneradorDesdeExistente, lo que rompió el flujo "crear desde
    cero" (CreadorRINFO no tiene ese método → AttributeError silencioso
    en el hilo de fondo → app "colgada" sin error visible). Corregido
    convirtiéndola en función de módulo independiente de self.
    """
    if carpeta_proyecto is None:
        carpeta_proyecto = os.path.dirname(carpeta_tarifas)
    nombre_listado = "LISTADO CODIGOS DALI - SAP.xlsx"

    ruta = os.path.join(carpeta_proyecto, nombre_listado)
    if not os.path.isfile(ruta):
        ruta = os.path.join(carpeta_proyecto, FOLDER_RECURSOS, nombre_listado)
    if not os.path.isfile(ruta):
        ruta = os.path.join(carpeta_tarifas, nombre_listado)
    if not os.path.isfile(ruta):
        _pat = re.compile(r'listado.*codigos.*dali.*sap.*\.xlsx$', re.IGNORECASE)
        for carpeta_scan in (
            carpeta_proyecto,
            os.path.join(carpeta_proyecto, FOLDER_RECURSOS),
            carpeta_tarifas,
        ):
            try:
                for entry in os.scandir(carpeta_scan):
                    if entry.is_file() and _pat.match(entry.name):
                        ruta = entry.path
                        if parent is not None:
                            _log(parent,
                                 f"LISTADO encontrado por patrón: {entry.name}", "INFO")
                        break
            except (PermissionError, FileNotFoundError):
                continue
            if os.path.isfile(ruta):
                break

    return ruta if os.path.isfile(ruta) else None


def _localizar_ruta_productos(carpeta_tarifas, carpeta_proyecto=None, parent=None):
    """
    Localiza el fichero Productos.xlsx (maestro de materiales SAP/MM03),
    con la misma estrategia que _localizar_ruta_listado_dali: ruta exacta
    en raíz del proyecto, RECURSOS y TARIFAS, y si no aparece, por patrón
    de nombre en esas mismas 3 carpetas. Devuelve la ruta encontrada, o
    None si no existe en ningún sitio.

    Usada como origen de respaldo (ver _cargar_listado_dali) para
    materiales SAP sin cruce DALI en el LISTADO -- caso real: material
    descatalogado o duplicado que nunca llegó a vincularse a un código
    DALI, dejando su Unidad de medida base sin sincronizar y provocando
    el error SAP "adecúe la UM-base y el factor de conversión" al
    procesar el registro (2026-09-18).
    """
    if carpeta_proyecto is None:
        carpeta_proyecto = os.path.dirname(carpeta_tarifas)
    nombre_productos = "Productos.xlsx"

    ruta = os.path.join(carpeta_proyecto, nombre_productos)
    if not os.path.isfile(ruta):
        ruta = os.path.join(carpeta_proyecto, FOLDER_RECURSOS, nombre_productos)
    if not os.path.isfile(ruta):
        ruta = os.path.join(carpeta_tarifas, nombre_productos)
    if not os.path.isfile(ruta):
        _pat = re.compile(r'productos\.xlsx$', re.IGNORECASE)
        for carpeta_scan in (
            carpeta_proyecto,
            os.path.join(carpeta_proyecto, FOLDER_RECURSOS),
            carpeta_tarifas,
        ):
            try:
                for entry in os.scandir(carpeta_scan):
                    if entry.is_file() and _pat.match(entry.name):
                        ruta = entry.path
                        if parent is not None:
                            _log(parent,
                                 f"Productos.xlsx encontrado por patrón: {entry.name}", "INFO")
                        break
            except (PermissionError, FileNotFoundError):
                continue
            if os.path.isfile(ruta):
                break

    return ruta if os.path.isfile(ruta) else None


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





# ─────────────────────────────────────────────────────────────────────────────
#  VENTANA: Materiales heredados sin tarifa — decisión de baja
# ─────────────────────────────────────────────────────────────────────────────

def _ventana_sin_tarifa(parent, pendientes: list) -> list:
    """
    Muestra una ventana modal tkinter para que el usuario decida qué hacer con
    los materiales heredados del Purchasing que NO tienen correspondencia en la
    tarifa activa y cuya fecha de Fin validez (Hoja 3 col J) está aún vigente.

    Parámetros
    ----------
    parent     : ventana padre tkinter (para posicionamiento y grab_set).
    pendientes : lista de dicts con claves:
                    'codigo'  (str)  — código SAP del material
                    'desc'    (str)  — texto breve (Hoja 1 col F)
                    'fecha'   (str)  — Fin validez formateado "dd/mm/yyyy"
                    'fila_h1' (int)  — fila en Hoja 1 (para escritura posterior)

    Devuelve
    --------
    Lista de dicts {'codigo': str, 'fila_h1': int, 'baja': bool}
    con la decisión del usuario para cada material:
        · baja=True   → el usuario marcó el ítem para dar de baja (col AJ = "X")
        · baja=False  → el usuario decidió mantenerlo (col AJ sin cambios)

    Si el usuario cierra la ventana o pulsa CANCELAR, todos los ítems
    devuelven baja=False (comportamiento conservador: no se toca nada).

    Notas de implementación
    -----------------------
    · ttk.Treeview con columna de checkbox simulada mediante texto ("☑"/"☐")
      ya que Treeview nativo no soporta checkboxes reales.
    · Botón [Info] por fila muestra un messagebox con el detalle del ítem.
    · [MANTENER] → todos los ítems sin check quedan baja=False y los
      marcados quedan baja=True; en definitiva es equivalente a [DAR BAJA]
      salvo que el nombre resulta más descriptivo para el flujo.
    · [DAR BAJA]   → solo los ítems con checkbox activo → baja=True.
    · [CANCELAR]   → todos baja=False.
    """

    n = len(pendientes)

    # ── Resultado compartido entre callbacks ──────────────────────────────────
    # key = índice en la lista 'pendientes', value = bool (True = dar de baja)
    checks = {i: False for i in range(n)}
    resultado = []   # se rellena al cerrar con decisión

    # ── Ventana toplevel ──────────────────────────────────────────────────────
    dlg = tk.Toplevel(parent)
    dlg.title(f"Materiales sin tarifa — Confirmar ({n} ítem{'s' if n != 1 else ''})")
    dlg.configure(bg=C_BG)
    dlg.resizable(True, True)
    try:
        dlg.iconbitmap(resource_path("assets/logo.ico"))
    except Exception:
        pass

    # Modal: bloquea interacción con el resto de la app
    dlg.grab_set()
    dlg.focus_set()

    # ── Cabecera ──────────────────────────────────────────────────────────────
    hdr = tk.Frame(dlg, bg="#8B4513", pady=10, padx=16)   # marrón oscuro = alerta
    hdr.pack(fill="x")
    tk.Label(
        hdr,
        text=f"⚠  {n} material{'es' if n != 1 else ''} sin tarifa activa",
        bg="#8B4513", fg="white", font=FONT_TITLE,
    ).pack(anchor="w")
    tk.Label(
        hdr,
        text="Estos materiales existen en el Purchasing pero no están en la tarifa. "
             "Marque los que desea dar de baja (col AJ = 'X').",
        bg="#8B4513", fg="#FFD5A0", font=FONT_SMALL,
        wraplength=620, justify="left",
    ).pack(anchor="w", pady=(4, 0))

    # ── Zona central: Treeview + scrollbar ───────────────────────────────────
    frm_tree = tk.Frame(dlg, bg=C_BG, padx=14, pady=10)
    frm_tree.pack(fill="both", expand=True)

    # Columnas: check (ancho fijo), codigo, desc, fecha, info
    cols = ("check", "codigo", "desc", "fecha", "info")
    tree = ttk.Treeview(
        frm_tree,
        columns=cols,
        show="headings",
        selectmode="none",
        height=min(n, 12),
    )

    # Configurar cabeceras
    tree.heading("check",  text="Baja",        anchor="center")
    tree.heading("codigo", text="Código SAP",  anchor="w")
    tree.heading("desc",   text="Descripción", anchor="w")
    tree.heading("fecha",  text="Fin validez", anchor="center")
    tree.heading("info",   text="",            anchor="center")

    # Anchuras de columna
    tree.column("check",  width=55,  minwidth=45,  stretch=False, anchor="center")
    tree.column("codigo", width=110, minwidth=80,  stretch=False, anchor="w")
    tree.column("desc",   width=280, minwidth=120, stretch=True,  anchor="w")
    tree.column("fecha",  width=110, minwidth=90,  stretch=False, anchor="center")
    tree.column("info",   width=60,  minwidth=50,  stretch=False, anchor="center")

    # Insertar filas
    item_ids = []   # iid por índice
    for i, mat in enumerate(pendientes):
        iid = tree.insert(
            "", "end",
            values=("☐", mat["codigo"], mat["desc"], mat["fecha"], "ℹ Info"),
            tags=("fila_par" if i % 2 == 0 else "fila_impar",),
        )
        item_ids.append(iid)

    # Colores alternos
    tree.tag_configure("fila_par",   background="#F5F8FF")
    tree.tag_configure("fila_impar", background="#FFFFFF")
    tree.tag_configure("marcado",    foreground="#C0392B", font=("Segoe UI", 9, "bold"))

    # Scrollbar vertical
    sb = ttk.Scrollbar(frm_tree, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    # ── Toggle checkbox al hacer clic en la columna "check" ──────────────────
    def _toggle_check(event):
        """Alterna el checkbox del ítem clicado (columna 'check' o doble clic)."""
        region = tree.identify_region(event.x, event.y)
        if region not in ("cell", "tree"):
            return
        col_id = tree.identify_column(event.x)
        iid    = tree.identify_row(event.y)
        if not iid:
            return
        # Permitir toggle en cualquier columna excepto "info"
        if col_id == "#5":   # columna info → gestionado por _click_info
            return
        idx = item_ids.index(iid)
        checks[idx] = not checks[idx]
        vals = list(tree.item(iid, "values"))
        vals[0] = "☑" if checks[idx] else "☐"
        tree.item(iid, values=vals,
                  tags=("marcado",) if checks[idx] else
                       ("fila_par" if idx % 2 == 0 else "fila_impar",))

    tree.bind("<Button-1>", _toggle_check)

    # ── Botón ℹ Info por fila ─────────────────────────────────────────────────
    def _click_info(event):
        """Muestra messagebox con el detalle del ítem cuando se pulsa la col info."""
        col_id = tree.identify_column(event.x)
        iid    = tree.identify_row(event.y)
        if not iid or col_id != "#5":
            return
        idx = item_ids.index(iid)
        mat = pendientes[idx]
        messagebox.showinfo(
            f"Detalle material — {mat['codigo']}",
            f"Código SAP:   {mat['codigo']}\n"
            f"Descripción:  {mat['desc'] or '(sin descripción)'}\n"
            f"Fin validez:  {mat['fecha']}\n"
            f"Fila Hoja 1:  {mat['fila_h1']}\n\n"
            "Este material existe en el Purchasing pero no está en la tarifa activa.\n"
            "Si lo marca para baja, se escribirá 'X' en la columna AJ (Hoja 1).",
            parent=dlg,
        )

    tree.bind("<ButtonRelease-1>", _click_info)

    # ── Botones [MANTENER] / [DAR BAJA] / [CANCELAR] ─────────────────────────
    frm_pie = tk.Frame(dlg, bg=C_BG, padx=16, pady=12)
    frm_pie.pack(side="bottom", fill="x")

    tk.Frame(frm_pie, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

    # Texto informativo sobre la selección
    lbl_info = tk.Label(
        frm_pie,
        text="Marque los ítems a dar de baja y pulse DAR BAJA, "
             "o pulse MANTENER para conservar todos sin cambios.",
        bg=C_BG, fg=C_TEXT, font=FONT_SMALL,
        wraplength=600, justify="left",
    )
    lbl_info.pack(anchor="w", pady=(0, 8))

    frm_botones = tk.Frame(frm_pie, bg=C_BG)
    frm_botones.pack(fill="x")

    def _accion_mantener():
        """Conservar todos sin cambios (ningún check = baja)."""
        for i in range(n):
            checks[i] = False
        for d in pendientes:
            resultado.append({
                "codigo":  d["codigo"],
                "fila_h1": d["fila_h1"],
                "baja":    False,
            })
        dlg.destroy()

    def _accion_dar_baja():
        """Aplicar la baja solo a los ítems marcados con checkbox."""
        for i, d in enumerate(pendientes):
            resultado.append({
                "codigo":  d["codigo"],
                "fila_h1": d["fila_h1"],
                "baja":    checks[i],
            })
        dlg.destroy()

    def _accion_cancelar():
        """Cancelar: ningún ítem recibe baja (comportamiento conservador)."""
        for d in pendientes:
            resultado.append({
                "codigo":  d["codigo"],
                "fila_h1": d["fila_h1"],
                "baja":    False,
            })
        dlg.destroy()

    # Botón MANTENER (verde neutro)
    tk.Button(
        frm_botones,
        text="✔  MANTENER",
        command=_accion_mantener,
        bg="#1A7A40", fg="white",
        font=FONT_BOLD, padx=14, pady=9,
        cursor="hand2", relief="flat",
        activebackground="#228B4A",
    ).pack(side="left", padx=(0, 6))

    # Botón DAR BAJA (rojo)
    tk.Button(
        frm_botones,
        text="✖  DAR BAJA",
        command=_accion_dar_baja,
        bg="#C0392B", fg="white",
        font=FONT_BOLD, padx=14, pady=9,
        cursor="hand2", relief="flat",
        activebackground="#A93226",
    ).pack(side="left", padx=(0, 6))

    # Botón CANCELAR (gris)
    tk.Button(
        frm_botones,
        text="Cancelar",
        command=_accion_cancelar,
        font=FONT_LABEL, padx=14, pady=7,
        cursor="hand2", relief="flat",
        bg="#D8DEE8",
        activebackground="#C0C8D8",
    ).pack(side="left")

    # ── Ajustar geometría ─────────────────────────────────────────────────────
    dlg.update_idletasks()
    sw = dlg.winfo_screenwidth()
    sh = dlg.winfo_screenheight()
    w  = min(700, sw - 60)
    h  = max(380, min(dlg.winfo_reqheight() + 60, sh - 80))
    x  = (sw - w) // 2
    y  = max(20, (sh - h) // 2)
    dlg.geometry(f"{w}x{h}+{x}+{y}")
    dlg.minsize(560, 340)

    # ── Esperar a que el usuario cierre la ventana ────────────────────────────
    # wait_window() bloquea hasta que dlg.destroy() sea llamado.
    parent.wait_window(dlg)

    # Si el usuario cerró la ventana con la X (sin pulsar botón), resultado=[]
    if not resultado:
        for d in pendientes:
            resultado.append({
                "codigo":  d["codigo"],
                "fila_h1": d["fila_h1"],
                "baja":    False,
            })

    return resultado


class VentanaMaterialesParaReactivar(tk.Toplevel):
    """
    Aviso final (2026-09-30, caso real MELOCOTON/FRAN Y CHEMI): se muestra
    cuando la tarifa incluye materiales cuyo Purchasing traía el registro
    info marcado para borrado (ver _avisar_materiales_reactivados_bloqueados).

    Confirmado en real, dos veces, que ninguna carga por Excel puede
    reactivarlos (SAP compara por proveedor+material+organización de
    compras, no por el número de la fila, y rechaza tanto conservar el
    número real como renumerar como registro nuevo). La única vía que
    funcionó fue la transacción SAP "Marcar reg.info de compra para
    borrar", desmarcando sus dos casillas de borrado -- esta ventana
    muestra esa pantalla como referencia visual, junto con la lista de
    materiales/registros concretos que hay que reactivar así.
    """

    def __init__(self, parent, materiales: dict, descripciones: dict = None,
                 centros: dict = None):
        """
        materiales: {codigo_sap: registro_info_antiguo} -- ver
        self._materiales_registro_a_renumerar en GeneradorDesdeExistente.
        descripciones: {codigo_sap: texto_breve} opcional, para mostrar el
        nombre del material junto a su código.
        centros: {codigo_sap: [centro1, centro2, ...]} opcional -- un mismo
        registro/material puede tener el problema en unos centros y no en
        otros, así que se muestran explícitamente (ver
        self._materiales_centros_borrado en GeneradorDesdeExistente).
        """
        super().__init__(parent)
        self.title("Materiales a reactivar manualmente en SAP")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        cargar_icono_ventana(self)

        self._materiales    = materiales
        self._descripciones = descripciones or {}
        self._centros       = centros or {}

        self._construir_ui()
        self._ajustar_geometria()
        self.grab_set()

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(920, sw - 40)
        h  = min(max(760, self.winfo_reqheight() + 40), sh - 60)
        x  = (sw - w) // 2
        y  = max(20, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(800, 680)

    def _construir_ui(self):
        hdr = tk.Frame(self, bg="#B8860B", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⚠  Materiales a reactivar manualmente en SAP",
                 bg="#B8860B", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Su registro info seguía marcado para borrado en el "
                      "Purchasing -- este RINFO NO puede reactivarlos por sí solo.",
                 bg="#B8860B", fg="#FFF3D0", font=FONT_SUB,
                 wraplength=850, justify="left").pack(anchor="w", pady=(4, 0))

        frm_pie = tk.Frame(self, bg=C_BG, padx=22, pady=12)
        frm_pie.pack(side="bottom", fill="x")
        tk.Frame(frm_pie, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))
        tk.Button(frm_pie, text="Entendido, cerrar",
                  command=self._cerrar,
                  bg="#B8860B", fg="white",
                  font=FONT_BOLD, padx=14, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#9A7009").pack(fill="x")

        # ── Zona con scroll (lista + instrucciones + imagen) ──────────────
        contenedor = tk.Frame(self, bg=C_BG)
        contenedor.pack(fill="both", expand=True, padx=22, pady=(14, 4))
        canvas  = tk.Canvas(contenedor, bg=C_BG, highlightthickness=0)
        scrolly = ttk.Scrollbar(contenedor, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrolly.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrolly.pack(side="right", fill="y")
        cuerpo = tk.Frame(canvas, bg=C_BG)
        canvas.create_window((0, 0), window=cuerpo, anchor="nw")
        cuerpo.bind("<Configure>",
                    lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        tk.Label(cuerpo,
                 text=f"Se han detectado {len(self._materiales)} material(es) "
                      f"con tarifa activa cuyo registro info seguía marcado "
                      f"para borrado en el Purchasing:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold"),
                 wraplength=850, justify="left").pack(anchor="w", pady=(0, 8))

        for cod_sap, registro_viejo in self._materiales.items():
            desc = self._descripciones.get(cod_sap, "")
            centros_mat = self._centros.get(cod_sap, [])
            centros_txt = (f"   centro(s): {', '.join(centros_mat)}"
                           if centros_mat else "")
            texto = f"  📦  {cod_sap}" + (f" — {desc}" if desc else "") \
                    + f"   (registro {registro_viejo})" + centros_txt
            tk.Label(cuerpo, text=texto,
                     bg="#FFF8E1", fg="#7B5800",
                     font=FONT_LABEL, padx=10, pady=6, relief="flat",
                     wraplength=850, justify="left", anchor="w").pack(
                fill="x", pady=(0, 4))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(12, 10))

        tk.Label(cuerpo, text="Cómo reactivarlos en SAP:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        pasos = (
            "1.  En SAP, busca la transacción \"Marcar reg.info de compra "
            "para borrar\".\n"
            "2.  Introduce el Registro info, Proveedor, Material y Centro "
            "del material afectado (uno por cada centro, si tiene más de "
            "uno marcado).\n"
            "3.  Desmarca las dos casillas de \"Peticiones de borrado\": "
            "\"Regtro-Info completo\" e \"Ind.borr.dat.org.co.\".\n"
            "4.  Ejecuta el cambio y confirma en \"Gestión de registros "
            "info de compras\" que el material ya aparece para ese centro."
        )
        tk.Label(cuerpo, text=pasos,
                 bg="#F0F5FB", fg="#285FA8",
                 font=FONT_LABEL, padx=10, pady=8, relief="flat",
                 wraplength=850, justify="left").pack(fill="x", pady=(4, 12))

        tk.Label(cuerpo, text="Referencia — pantalla de acceso de esa transacción:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        self._img_ref = self._cargar_imagen_referencia()
        if self._img_ref is not None:
            tk.Label(cuerpo, image=self._img_ref, bg=C_BG,
                     relief="solid", borderwidth=1).pack(pady=(6, 4))

    def _cargar_imagen_referencia(self):
        try:
            from PIL import Image, ImageTk
            img = Image.open(resource_path("assets/marcar_borrado_sap.png"))   # (nov 2026) antes, base64 incrustado aquí
            ancho_destino = 760
            alto_destino  = int(img.height * (ancho_destino / img.width))
            img = img.resize((ancho_destino, alto_destino), Image.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception:
            return None

    def _cerrar(self):
        parent = self.master
        self.destroy()
        try:
            if parent and parent.winfo_exists():
                parent.deiconify()
        except Exception:
            pass


class VentanaConfirmacionGenerado(tk.Toplevel):
    """
    Confirmación final tras generar el nuevo archivo desde un Purchasing existente.
    Informa del nº de materiales ya existentes, los nuevos insertados y ofrece
    abrir el archivo generado.
    """

    def __init__(self, parent, nombre_nuevo: str, ruta_nuevo: str,
                 nombre_sap: str, codigo_sap: str,
                 n_existentes: int, n_insertados: int,
                 nombre_purchasing_base: str):
        super().__init__(parent)
        self.title("Archivo generado correctamente")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        self._ruta_nuevo = ruta_nuevo
        cargar_icono_ventana(self)
        self._construir_ui(nombre_nuevo, ruta_nuevo, nombre_sap, codigo_sap,
                            n_existentes, n_insertados, nombre_purchasing_base)
        self._ajustar_geometria()
        self.grab_set()

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(680, sw - 40)
        h  = max(540, min(self.winfo_reqheight() + 40, sh - 60))
        x  = (sw - w) // 2
        y  = max(20, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(580, 540)

    def _construir_ui(self, nombre_nuevo, ruta_nuevo, nombre_sap, codigo_sap,
                       n_existentes, n_insertados, nombre_purchasing_base):
        carpeta = os.path.dirname(ruta_nuevo)

        hdr = tk.Frame(self, bg="#1A7A40", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="✔  Archivo generado correctamente",
                 bg="#1A7A40", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Los materiales nuevos han sido añadidos al Purchasing base.",
                 bg="#1A7A40", fg="#A8E6C0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        frm_pie = tk.Frame(self, bg=C_BG, padx=22, pady=12)
        frm_pie.pack(side="bottom", fill="x")
        tk.Frame(frm_pie, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))
        tk.Label(frm_pie, text="¿Desea abrir el archivo ahora?",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 8))
        tk.Button(frm_pie, text="✔  Sí, abrir ahora",
                  command=self._abrir_y_cerrar,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=14, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))
        tk.Button(frm_pie, text="No, cerrar",
                  command=self._cerrar,
                  font=FONT_LABEL, padx=14, pady=8,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(fill="x")

        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=16)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo, text="✔  Archivo generado:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  📄  {nombre_nuevo}",
                 bg="#EEF7EE", fg="#1B5E20",
                 font=("Segoe UI", 10, "bold"),
                 padx=10, pady=8, relief="flat").pack(fill="x", pady=(4, 10))

        tk.Label(cuerpo, text="Guardado en:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  📁  {carpeta}",
                 bg="#F5F5F5", fg="#333",
                 font=FONT_LABEL, padx=10, pady=6, relief="flat",
                 wraplength=600, justify="left").pack(fill="x", pady=(4, 10))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(cuerpo, text="Archivo base (Purchasing):",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  📋  {nombre_purchasing_base}",
                 bg="#F0F5FB", fg="#285FA8",
                 font=FONT_LABEL, padx=10, pady=6, relief="flat",
                 wraplength=600, justify="left").pack(fill="x", pady=(4, 10))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(cuerpo, text="Resumen de materiales:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        frm_vals = tk.Frame(cuerpo, bg="#F0F5FB", padx=12, pady=8)
        frm_vals.pack(fill="x", pady=(4, 10))
        _fila_valor(frm_vals, "Ya existentes en Purchasing:", str(n_existentes))
        _fila_valor(frm_vals, "Nuevos insertados (tarifa):",  str(n_insertados))
        _fila_valor(frm_vals, "D/C — Proveedor SAP:",
                    f"{nombre_sap}  [{codigo_sap}]" if codigo_sap else nombre_sap)

        if n_insertados == 0:
            tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 8))
            tk.Label(cuerpo,
                     text="ℹ  Todos los materiales de la tarifa ya estaban en el Purchasing.",
                     bg="#FFF8E1", fg="#7B5800",
                     font=("Segoe UI", 9, "italic"),
                     padx=10, pady=6, relief="flat").pack(fill="x")

    def _abrir_y_cerrar(self):
        try:
            os.startfile(self._ruta_nuevo)
        except AttributeError:
            subprocess.call(["xdg-open", self._ruta_nuevo])
        except Exception:
            pass
        self._cerrar()

    def _cerrar(self):
        # Restaurar la ventana padre antes de destruir esta ventana.
        # Sin deiconify, el padre queda en withdraw() y la app parece cerrarse.
        parent = self.master
        self.destroy()
        try:
            if parent and parent.winfo_exists():
                parent.deiconify()
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4 — VENTANA DE OPCIONES CUANDO NO HAY PURCHASING
# ─────────────────────────────────────────────────────────────────────────────

class VentanaOpcionesSinPurchasing(tk.Toplevel):

    def __init__(self, parent_ventana, ruta_tarifa: str, nombre_archivo: str,
                 nombre_dali: str, nombre_sap: str, codigo_sap: str,
                 carpeta_tarifas: str, app_root=None):
        super().__init__(parent_ventana.master if hasattr(parent_ventana, "master")
                         else parent_ventana)
        self.title("Registro Info — Proveedor sin Purchasing en TARIFAS")
        self.resizable(True, True)
        self.configure(bg=C_BG)
        self._intentar_icono()

        self._parent_ventana  = parent_ventana
        self._ruta_tarifa     = ruta_tarifa
        self._nombre_archivo  = nombre_archivo
        self._nombre_dali     = nombre_dali
        self._nombre_sap      = nombre_sap
        self._codigo_sap      = codigo_sap
        self._carpeta_tarifas = carpeta_tarifas
        # Referencia directa al App raíz para actualizar la barra de progreso
        self._app_root        = app_root

        self._construir_ui()
        self.grab_set()
        self.transient(self.master)
        self._ajustar_geometria()

    def _intentar_icono(self):
        cargar_icono_ventana(self)

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(700, sw - 40)
        h  = min(sh - 60, 620)
        x  = (sw - w) // 2
        y  = 20
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(560, 420)

    def _construir_ui(self):
        etiqueta_prov = (f"{self._nombre_sap}  [{self._codigo_sap}]"
                         if self._codigo_sap else self._nombre_sap)

        hdr = tk.Frame(self, bg="#8B5A00", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr,
                 text="⚠  No se encontró Purchasing Info Records para este proveedor",
                 bg="#8B5A00", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="No hay ningún archivo Purchasing Info Records*.xlsx en la carpeta TARIFAS.",
                 bg="#8B5A00", fg="#FFD99A", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=14)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text="Proveedor identificado en la tarifa (C2):",
                 bg=C_BG, fg="#999",
                 font=("Segoe UI", 8, "italic")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  {self._nombre_dali}",
                 bg="#F8F4FF", fg="#555",
                 font=FONT_LABEL, padx=10, pady=5, relief="flat").pack(fill="x")

        tk.Frame(cuerpo, bg=C_BG, height=8).pack()

        tk.Label(cuerpo,
                 text="Proveedor SAP a usar (Proveedores.xlsx):",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  {etiqueta_prov}",
                 bg="#EEF2F8", fg="#285FA8",
                 font=("Segoe UI", 11, "bold"),
                 padx=10, pady=8, relief="flat").pack(fill="x")

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=10)

        frm_sap = tk.Frame(cuerpo, bg="#EBF4FF", padx=10, pady=8, relief="flat")
        frm_sap.pack(fill="x", pady=(0, 12))
        tk.Label(frm_sap,
                 text="🔗  Acceso directo a SAP S/4HANA — Registros Info de Compras:",
                 bg="#EBF4FF", fg="#1A3A6A",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        lnk = tk.Label(frm_sap, text=URL_SAP,
                        bg="#EBF4FF", fg="#1565C0",
                        font=("Segoe UI", 8, "underline"),
                        cursor="hand2")
        lnk.pack(anchor="w", pady=(2, 0))
        lnk.bind("<Button-1>", lambda e: webbrowser.open(URL_SAP))

        tk.Label(cuerpo,
                 text="Seleccione una de las siguientes opciones:",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 8))

        tk.Button(cuerpo,
                  text="📥  Descargar el Purchasing Info Records desde SAP primero",
                  command=self._opcion_descargar_sap,
                  bg=C_MID, fg=C_TEXT_LIGHT,
                  font=FONT_BOLD, padx=12, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground=C_HOVER,
                  activeforeground=C_ACCENT2).pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo,
                  text=f"🆕  Crear nuevo Registro Info desde cero  —  {etiqueta_prov}",
                  command=self._opcion_crear_desde_cero,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=12, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo, text="Cancelar",
                  command=self._cancelar,
                  font=FONT_LABEL, padx=12, pady=8,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(fill="x")

    def _opcion_descargar_sap(self):
        _log(self._parent_ventana,
             "Usuario eligió: Descargar Purchasing desde SAP.", "INFO")
        self.destroy()
        self._parent_ventana.withdraw()
        VentanaInstruccionesSAP(self.master, self._nombre_sap)

    def _opcion_crear_desde_cero(self):
        _log(self._parent_ventana,
             "Usuario eligió: Crear Registro Info nuevo desde cero.", "INFO")

        # ── Verificar template SAP ANTES de abrir cualquier ventana ──────────
        carpeta_proyecto = os.path.dirname(self._carpeta_tarifas)
        ruta_template    = os.path.join(carpeta_proyecto, FOLDER_RECURSOS, ARCHIVO_TEMPLATE)

        if not os.path.isfile(ruta_template):
            ruta_recursos = os.path.join(carpeta_proyecto, FOLDER_RECURSOS)
            _log(self._parent_ventana,
                 "Template SAP no encontrado en RECURSOS/ — creación abortada.", "ERROR")
            VentanaTemplateAusente(self, ruta_recursos, ARCHIVO_TEMPLATE)
            return

        self.destroy()

        def _tras_centros(resultado_centros):
            """Callback de VentanaCentrosRINFO."""
            if resultado_centros is False:
                _log(self._parent_ventana,
                     "Creación de RINFO cancelada en la pantalla de centros.", "WARN")
                self._parent_ventana.withdraw()
                return
            codigos = ", ".join(c for c, _ in resultado_centros)
            _log(self._parent_ventana,
                 f"Centros seleccionados ({len(resultado_centros)}): {codigos}", "OK")
            CreadorRINFO(
                self._parent_ventana,
                self._ruta_tarifa,
                self._nombre_archivo,
                self._nombre_sap,
                self._codigo_sap,
                self._carpeta_tarifas,
                app_root=self._app_root,
                telefono_proveedor=self._telefono_pendiente,
                fecha_inicio=self._fechas_pendientes[0],
                fecha_fin=self._fechas_pendientes[1],
                centros_activos=resultado_centros,
            )

        def _tras_fechas(resultado_fechas):
            """Callback de VentanaFechasRINFO."""
            if resultado_fechas is False:
                _log(self._parent_ventana,
                     "Creación de RINFO cancelada en la pantalla de fechas.", "WARN")
                self._parent_ventana.withdraw()
                return
            self._fechas_pendientes = resultado_fechas
            fecha_inicio, fecha_fin = resultado_fechas
            _log(self._parent_ventana,
                 f"Fechas validez: inicio={fecha_inicio}  fin={fecha_fin}", "OK")
            VentanaCentrosRINFO(self._parent_ventana, _tras_centros)

        def _tras_telefono(resultado):
            """Callback de VentanaTelefonoRINFO."""
            if resultado is False:
                # El usuario canceló — no se crea el RINFO
                _log(self._parent_ventana,
                     "Creación de RINFO cancelada en la pantalla de teléfono.", "WARN")
                self._parent_ventana.withdraw()
                return
            # Guardar teléfono temporalmente y pasar a pantalla de fechas
            self._telefono_pendiente = resultado   # str validado o None
            VentanaFechasRINFO(self._parent_ventana, _tras_fechas)

        self._telefono_pendiente = None
        self._fechas_pendientes  = (None, None)
        VentanaTelefonoRINFO(self._parent_ventana, _tras_telefono)

    def _cancelar(self):
        _log(self._parent_ventana, "Acción cancelada por el usuario.", "WARN")
        self.destroy()
        self._parent_ventana.withdraw()


# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4a — VENTANA DE INSTRUCCIONES SAP (7 pasos)
# ─────────────────────────────────────────────────────────────────────────────

class VentanaInstruccionesSAP(tk.Toplevel):

    PASOS = [
        ("1", "Acceder a la aplicación TARIFAS en SAP",
         "En el menú de favoritos de SAP S/4HANA, localice el grupo  TARIFAS\n"
         "y pulse en:  «Modificaciones en masa de registros info de compras»."),
        ("2", "Pantalla de búsqueda",
         "Se abrirá la pantalla  «Estándar COMPRAS».\n"
         "Localice el campo  Proveedor:  en la parte superior izquierda."),
        ("3", "Abrir el selector de Proveedor",
         "Pulse el icono de cuadraditos (⊞) a la derecha del campo  Proveedor:\n"
         "para abrir el buscador de proveedores."),
        ("4", "Buscar y seleccionar el proveedor",
         ""),   # Se rellena dinámicamente
        ("5", "Ejecutar la búsqueda",
         "De vuelta en la pantalla principal, pulse el botón  Ir  (azul)\n"
         "para cargar todos los Registros Info del proveedor."),
        ("6", "Seleccionar todos y descargar",
         "Una vez cargados los registros:\n"
         "  1. Marque la casilla de cabecera para  SELECCIONAR TODOS.\n"
         "  2. Pulse el botón  Descargar  (parte superior derecha)."),
        ("7", "Guardar y volver a ejecutar",
         "Se descargará el archivo:  📄  Purchasing Info Records.xlsx\n"
         "Guárdelo en la carpeta  TARIFAS  del proyecto y vuelva\n"
         "a pulsar el botón  «CREAR REGISTRO INFO»  para continuar."),
    ]

    def __init__(self, parent, nombre_proveedor: str):
        super().__init__(parent)
        self.title("Cómo descargar el Purchasing Info Records desde SAP")
        self.resizable(True, True)
        self.configure(bg=C_BG)
        self._nombre_proveedor = nombre_proveedor
        cargar_icono_ventana(self)
        self._construir_ui()
        self._ajustar_geometria()
        self.grab_set()

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw  = self.winfo_screenwidth()
        sh  = self.winfo_screenheight()
        h_disponible = sh - 48 - 30
        w = min(880, sw - 40)
        h = min(int(h_disponible * 0.88), h_disponible)
        x = (sw - w) // 2
        y = 20
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(min(w, 700), min(h, 540))

    def _construir_ui(self):
        hdr = tk.Frame(self, bg=C_MID, pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📥  Descargar Purchasing Info Records desde SAP",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Siga los pasos indicados a continuación en SAP S/4HANA Cloud:",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        frm_link = tk.Frame(self, bg="#EBF4FF", padx=16, pady=10)
        frm_link.pack(fill="x")
        tk.Label(frm_link,
                 text="🔗  Acceso directo a SAP S/4HANA — Modificaciones en masa de registros info:",
                 bg="#EBF4FF", fg="#1A3A6A",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        lnk = tk.Label(frm_link, text=URL_SAP,
                        bg="#EBF4FF", fg="#1565C0",
                        font=("Segoe UI", 9, "underline"),
                        cursor="hand2")
        lnk.pack(anchor="w", pady=(2, 0))
        lnk.bind("<Button-1>", lambda e: webbrowser.open(URL_SAP))

        frame_steps = tk.Frame(self, bg=C_BG)
        frame_steps.pack(fill="both", expand=True, padx=12, pady=(6, 0))

        pasos_renderizados = list(self.PASOS)
        pasos_renderizados[3] = (
            "4", "Buscar y seleccionar el proveedor",
            f"En el buscador que se abre:\n"
            f"  1. Escriba el nombre:  «{self._nombre_proveedor}»\n"
            f"  2. Pulse la lupa (🔍) o el botón  Ir.\n"
            f"  3. Marque la casilla junto al Acreedor (✔) para seleccionar todos.\n"
            f"  4. Pulse  OK.",
        )

        for idx, (num, titulo, desc) in enumerate(pasos_renderizados):
            frame_steps.rowconfigure(idx, weight=1)
            frame_steps.columnconfigure(0, weight=1)

            frm_fila = tk.Frame(frame_steps, bg=C_BG)
            frm_fila.grid(row=idx, column=0, sticky="nsew", padx=4, pady=1)

            tk.Label(frm_fila,
                     text=f" {num} ",
                     bg=C_MID, fg="white",
                     font=("Segoe UI", 10, "bold"),
                     padx=8, pady=4,
                     relief="flat").pack(side="left", anchor="n", padx=(0, 8))

            frm_txt = tk.Frame(frm_fila, bg="#EEF2F8", padx=10, pady=4)
            frm_txt.pack(side="left", fill="both", expand=True)
            tk.Label(frm_txt, text=titulo,
                     bg="#EEF2F8", fg="#1A2E4A",
                     font=("Segoe UI", 9, "bold"),
                     anchor="w", justify="left").pack(anchor="w")
            tk.Label(frm_txt, text=desc,
                     bg="#EEF2F8", fg=C_TEXT,
                     font=FONT_LABEL,
                     anchor="w", justify="left").pack(anchor="w", pady=(1, 0))

        frm_cerrar = tk.Frame(self, bg=C_BG, pady=6, padx=14)
        frm_cerrar.pack(fill="x")
        tk.Button(frm_cerrar, text="Cerrar",
                  command=self.destroy,
                  font=FONT_LABEL, padx=14, pady=5,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(side="right")
        self.grab_set()


# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4b.0 — VENTANA: ASIGNACIÓN OPCIONAL DE TELÉFONO DE PROVEEDOR
# ─────────────────────────────────────────────────────────────────────────────

class VentanaTelefonoRINFO(tk.Toplevel):
    """
    Pregunta al usuario si desea asignar un número de teléfono al Registro Info
    antes de iniciar la creación.

    · Si confirma con un número válido (8-10 dígitos, espacios opcionales) → llama
      al callback con el número limpio como str (sin espacios).
    · Si deja vacío o pulsa «Sin teléfono / Continuar» → llama al callback con None.
    · Si cierra la ventana o cancela → llama al callback con False (abortar).

    Los espacios que el usuario introduzca se eliminan automáticamente antes de
    validar y de escribir el valor en la columna AD del RINFO.
    """

    _RE_TELEFONO = re.compile(r'^\d{8,10}$')

    def __init__(self, parent, on_confirm):
        """
        Parámetros
        ----------
        parent      : ventana padre (Toplevel o Tk).
        on_confirm  : callable(resultado) donde resultado es:
                        str   → número validado sin espacios (p.ej. «612345678»)
                        None  → sin teléfono, continuar igualmente
                        False → cancelado / cerrar sin crear RINFO
        """
        super().__init__(parent)
        self.title("Número de teléfono del proveedor")
        self.configure(bg=C_BG)
        self.resizable(False, False)
        self._on_confirm = on_confirm

        cargar_icono_ventana(self)

        self._construir_ui()
        ajustar_geometria_pantalla(self, 480)
        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.grab_set()

    # ── Geometría ─────────────────────────────────────────────────────────────

    # ── UI ────────────────────────────────────────────────────────────────────
    def _construir_ui(self):
        # Cabecera
        hdr = tk.Frame(self, bg=C_MID, pady=10, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📞  Teléfono del proveedor (opcional)",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Este dato se escribirá en la columna AD del Registro Info.",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(3, 0))

        # Cuerpo
        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=16)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text="¿Desea asignar un número de teléfono al Registro Info?",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 10, "bold"),
                 wraplength=420, justify="left").pack(anchor="w", pady=(0, 4))

        tk.Label(cuerpo,
                 text="Formato: 000000000  (8-10 dígitos; los espacios se eliminan automáticamente)\n"
                      "Si deja vacío o pulsa «Sin teléfono», continuaré sin asignar teléfono.",
                 bg=C_BG, fg="#666",
                 font=FONT_LABEL,
                 justify="left").pack(anchor="w", pady=(0, 10))

        # Campo de entrada
        frm_entrada = tk.Frame(cuerpo, bg=C_BG)
        frm_entrada.pack(fill="x", pady=(0, 4))

        tk.Label(frm_entrada, text="Teléfono:",
                 bg=C_BG, fg=C_TEXT,
                 font=FONT_BOLD, width=10, anchor="w").pack(side="left")

        self._var_tel = tk.StringVar()
        self._entry = tk.Entry(frm_entrada,
                               textvariable=self._var_tel,
                               font=("Segoe UI", 11),
                               width=16, relief="solid", bd=1)
        self._entry.pack(side="left", padx=(4, 0))
        self._entry.bind("<Return>", lambda e: self._confirmar())

        # Etiqueta de error (oculta inicialmente)
        self._lbl_error = tk.Label(cuerpo, text="",
                                   bg=C_BG, fg=C_ERROR,
                                   font=("Segoe UI", 8, "italic"))
        self._lbl_error.pack(anchor="w", pady=(0, 8))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 12))

        # Botones
        tk.Button(cuerpo,
                  text="✔  Asignar teléfono y continuar",
                  command=self._confirmar,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=12, pady=10,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo,
                  text="Sin teléfono — continuar igualmente",
                  command=self._sin_telefono,
                  bg=C_MID, fg=C_TEXT_LIGHT,
                  font=FONT_BOLD, padx=12, pady=9,
                  cursor="hand2", relief="flat",
                  activebackground=C_HOVER,
                  activeforeground=C_ACCENT2).pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo, text="Cancelar",
                  command=self._cancelar,
                  font=FONT_LABEL, padx=12, pady=7,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8", activebackground="#C0C8D8").pack(fill="x")

        self._entry.focus_set()

    # ── Acciones ──────────────────────────────────────────────────────────────
    def _confirmar(self):
        valor = self._var_tel.get().strip()
        if not valor:
            # Vacío → sin teléfono
            self._sin_telefono()
            return
        # Eliminar espacios antes de validar (acepta "000 000 000" o "000000000")
        valor_limpio = valor.replace(" ", "")
        if not self._RE_TELEFONO.match(valor_limpio):
            self._lbl_error.config(
                text="⚠  Formato incorrecto. Se esperan entre 8 y 10 dígitos (los espacios se ignoran)")
            self._entry.focus_set()
            return
        self._lbl_error.config(text="")
        self.destroy()
        self._on_confirm(valor_limpio)

    def _sin_telefono(self):
        self.destroy()
        self._on_confirm(None)

    def _cancelar(self):
        self.destroy()
        self._on_confirm(False)


# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4b.1 — VENTANA: FECHAS DE VALIDEZ DEL REGISTRO INFO
# ─────────────────────────────────────────────────────────────────────────────

class VentanaFechasRINFO(tk.Toplevel):
    """
    Solicita al usuario las fechas de inicio y fin de validez de la tarifa
    (Registro Info) antes de iniciar la creación del RINFO.

    Ambas fechas son OBLIGATORIAS.

    Formatos aceptados (entrada):
        dd-mm-yyyy  |  dd.mm.yyyy  |  dd,mm,yyyy  |  ddmmyyyy
    El valor que se escribe en el archivo es siempre "dd.mm.yyyy".

    Callbacks:
        on_confirm((fecha_inicio_str, fecha_fin_str))  → fechas validadas
        on_confirm(False)                              → cancelado / cerrado
    """

    # Variantes separadoras aceptadas: guion, punto, coma, o sin separador
    _RE_FECHA = re.compile(
        r'^(\d{2})[-.,]?(\d{2})[-.,]?(\d{4})$'
    )

    def __init__(self, parent, on_confirm):
        """
        Parámetros
        ----------
        parent      : ventana padre (Toplevel o Tk).
        on_confirm  : callable(resultado) donde resultado es:
                        (str, str)  → (fecha_inicio, fecha_fin) en formato dd.mm.yyyy
                        False       → cancelado / cerrar sin crear RINFO
        """
        super().__init__(parent)
        self.title("Fechas de validez del Registro Info")
        self.configure(bg=C_BG)
        self.resizable(False, False)
        self._on_confirm = on_confirm

        cargar_icono_ventana(self)

        self._construir_ui()
        ajustar_geometria_pantalla(self, 500)
        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.grab_set()

    # ── Geometría ─────────────────────────────────────────────────────────────

    # ── UI ────────────────────────────────────────────────────────────────────
    def _construir_ui(self):
        # Cabecera
        hdr = tk.Frame(self, bg="#285FA8", pady=10, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📅  Fechas de validez de la tarifa (Reg. Info)",
                 bg="#285FA8", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Estos valores se escribirán en la hoja 3 del Registro Info.",
                 bg="#285FA8", fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(3, 0))

        # Cuerpo
        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=16)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text="Formato requerido: [dd.mm.yyyy]",
                 bg=C_BG, fg="#666",
                 font=FONT_LABEL,
                 wraplength=440, justify="left").pack(anchor="w", pady=(0, 12))

        # ── Campo Fecha Inicio ─────────────────────────────────────────────────
        tk.Label(cuerpo,
                 text="Fecha Inicio Tarifa (Inicio de validez):",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")

        frm_inicio = tk.Frame(cuerpo, bg=C_BG)
        frm_inicio.pack(fill="x", pady=(4, 0))

        tk.Label(frm_inicio, text="Inicio:",
                 bg=C_BG, fg=C_TEXT,
                 font=FONT_BOLD, width=8, anchor="w").pack(side="left")

        self._var_inicio = tk.StringVar()
        self._entry_inicio = tk.Entry(frm_inicio,
                                      textvariable=self._var_inicio,
                                      font=("Segoe UI", 11),
                                      width=14, relief="solid", bd=1)
        self._entry_inicio.pack(side="left", padx=(4, 0))
        self._entry_inicio.bind("<Return>", lambda e: self._entry_fin.focus_set())

        tk.Label(frm_inicio, text="  p.ej. 01-01-2025",
                 bg=C_BG, fg="#999", font=FONT_SMALL).pack(side="left", padx=(8, 0))

        self._lbl_error_inicio = tk.Label(cuerpo, text="",
                                          bg=C_BG, fg=C_ERROR,
                                          font=("Segoe UI", 8, "italic"))
        self._lbl_error_inicio.pack(anchor="w", pady=(2, 8))

        tk.Frame(cuerpo, bg="#E0E8F0", height=1).pack(fill="x", pady=(0, 10))

        # ── Campo Fecha Fin ────────────────────────────────────────────────────
        tk.Label(cuerpo,
                 text="Fecha Finalización Tarifa (Fin de validez):",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")

        frm_fin = tk.Frame(cuerpo, bg=C_BG)
        frm_fin.pack(fill="x", pady=(4, 0))

        tk.Label(frm_fin, text="Fin:",
                 bg=C_BG, fg=C_TEXT,
                 font=FONT_BOLD, width=8, anchor="w").pack(side="left")

        self._var_fin = tk.StringVar()
        self._entry_fin = tk.Entry(frm_fin,
                                   textvariable=self._var_fin,
                                   font=("Segoe UI", 11),
                                   width=14, relief="solid", bd=1)
        self._entry_fin.pack(side="left", padx=(4, 0))
        self._entry_fin.bind("<Return>", lambda e: self._confirmar())

        tk.Label(frm_fin, text="  p.ej. 31-12-2025",
                 bg=C_BG, fg="#999", font=FONT_SMALL).pack(side="left", padx=(8, 0))

        self._lbl_error_fin = tk.Label(cuerpo, text="",
                                       bg=C_BG, fg=C_ERROR,
                                       font=("Segoe UI", 8, "italic"))
        self._lbl_error_fin.pack(anchor="w", pady=(2, 8))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 12))

        # ── Botones ────────────────────────────────────────────────────────────
        tk.Button(cuerpo,
                  text="✔  Confirmar fechas y continuar",
                  command=self._confirmar,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=12, pady=10,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo, text="Cancelar",
                  command=self._cancelar,
                  font=FONT_LABEL, padx=12, pady=7,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8", activebackground="#C0C8D8").pack(fill="x")

        # Precargar fechas por defecto: 01.01.AAAA / 31.12.AAAA del año actual
        _anio = datetime.today().year
        self._var_inicio.set(f"01.01.{_anio}")
        self._var_fin.set(f"31.12.{_anio}")

        self._entry_inicio.focus_set()

    # ── Validación ────────────────────────────────────────────────────────────
    @classmethod
    def _parsear_fecha(cls, texto: str):
        """
        Intenta parsear el texto introducido.
        Devuelve "dd.mm.yyyy" si es válido, None si no.
        Acepta: dd-mm-yyyy | dd.mm.yyyy | dd,mm,yyyy | ddmmyyyy
        """
        texto = texto.strip()
        # Quitar separadores mixtos (guion, punto, coma, espacio)
        m = cls._RE_FECHA.match(texto)
        if not m:
            return None
        dia, mes, anio = m.group(1), m.group(2), m.group(3)
        # Validar rangos básicos
        try:
            datetime(int(anio), int(mes), int(dia))
        except ValueError:
            return None
        return f"{dia}.{mes}.{anio}"

    @classmethod
    def _a_date(cls, texto: str):
        """Devuelve un objeto date a partir de texto ya validado 'dd.mm.yyyy', o None."""
        try:
            d, m, y = texto.split(".")
            return datetime(int(y), int(m), int(d)).date()
        except Exception:
            return None

    # ── Acciones ──────────────────────────────────────────────────────────────
    def _confirmar(self):
        # Limpiar errores previos
        self._lbl_error_inicio.config(text="")
        self._lbl_error_fin.config(text="")

        fecha_inicio = self._parsear_fecha(self._var_inicio.get())
        fecha_fin    = self._parsear_fecha(self._var_fin.get())

        ok = True

        if not self._var_inicio.get().strip():
            self._lbl_error_inicio.config(
                text="⚠  Campo obligatorio. Introduzca la fecha de inicio.")
            self._entry_inicio.focus_set()
            ok = False
        elif fecha_inicio is None:
            self._lbl_error_inicio.config(
                text="⚠  Fecha inválida — compruebe día, mes y año (p.ej. 01-01-2025).")
            self._entry_inicio.focus_set()
            ok = False

        if not self._var_fin.get().strip():
            self._lbl_error_fin.config(
                text="⚠  Campo obligatorio. Introduzca la fecha de fin.")
            if ok:
                self._entry_fin.focus_set()
            ok = False
        elif fecha_fin is None:
            self._lbl_error_fin.config(
                text="⚠  Fecha inválida — compruebe día, mes y año (p.ej. 31-12-2025).")
            if ok:
                self._entry_fin.focus_set()
            ok = False

        # ── Validación cruzada: fin no puede ser anterior a inicio ────────────
        if ok and fecha_inicio and fecha_fin:
            d_inicio = self._a_date(fecha_inicio)
            d_fin    = self._a_date(fecha_fin)
            if d_inicio is not None and d_fin is not None and d_fin < d_inicio:
                self._lbl_error_fin.config(
                    text=f"⚠  La fecha de fin ({fecha_fin}) no puede ser "
                         f"anterior a la de inicio ({fecha_inicio}).")
                self._entry_fin.focus_set()
                ok = False

        if not ok:
            return

        self.destroy()
        self._on_confirm((fecha_inicio, fecha_fin))

    def _cancelar(self):
        self.destroy()
        self._on_confirm(False)


# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4b.2 — VENTANA: SELECCIÓN DE CENTROS SAP
# ─────────────────────────────────────────────────────────────────────────────

class VentanaCentrosRINFO(tk.Toplevel):
    """
    Muestra los 10 centros SAP de Canarias con casillas de verificación.
    Todos los centros aparecen marcados por defecto.
    El usuario desmarca los centros que NO desea incluir en el Registro Info.

    Callbacks:
        on_confirm(lista)   → lista de (codigo, nombre) de los centros seleccionados
                              (al menos 1 debe quedar marcado)
        on_confirm(False)   → cancelado / ventana cerrada
    """

    def __init__(self, parent, on_confirm):
        """
        Parámetros
        ----------
        parent      : ventana padre (Toplevel o Tk).
        on_confirm  : callable(resultado) donde resultado es:
                        list[(str, str)]  → centros seleccionados [(codigo, nombre), ...]
                        False             → cancelado / cerrar sin crear RINFO
        """
        super().__init__(parent)
        self.title("Selección de centros — Registro Info")
        self.configure(bg=C_BG)
        self.resizable(False, False)
        self._on_confirm = on_confirm

        cargar_icono_ventana(self)

        self._vars = {}   # codigo → BooleanVar
        self._construir_ui()
        ajustar_geometria_pantalla(self, 520)
        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.grab_set()

    # ── Geometría ─────────────────────────────────────────────────────────────

    # ── UI ────────────────────────────────────────────────────────────────────
    def _construir_ui(self):
        # Cabecera
        hdr = tk.Frame(self, bg=C_MID, pady=10, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🏨  Centros SAP — Selección para el Registro Info",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Desmarque los centros que NO se van a incluir. "
                      "Por defecto se incluyen todos.",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB,
                 wraplength=460, justify="left").pack(anchor="w", pady=(3, 0))

        # Cuerpo
        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=14)
        cuerpo.pack(fill="both", expand=True)

        # Botones rápidos Todos / Ninguno
        frm_rapido = tk.Frame(cuerpo, bg=C_BG)
        frm_rapido.pack(fill="x", pady=(0, 8))
        tk.Label(frm_rapido, text="Selección rápida:",
                 bg=C_BG, fg="#888", font=FONT_SMALL).pack(side="left")
        tk.Button(frm_rapido, text="✔ Todos",
                  command=self._seleccionar_todos,
                  bg="#E8F0FE", fg="#285FA8",
                  font=FONT_SMALL, padx=8, pady=2,
                  cursor="hand2", relief="flat",
                  activebackground="#D0E0F8").pack(side="left", padx=(8, 4))
        tk.Button(frm_rapido, text="✗ Ninguno",
                  command=self._deseleccionar_todos,
                  bg="#FEF0E8", fg="#8B3A00",
                  font=FONT_SMALL, padx=8, pady=2,
                  cursor="hand2", relief="flat",
                  activebackground="#F8D8C0").pack(side="left")

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 8))

        # Lista de centros con checkboxes
        frm_lista = tk.Frame(cuerpo, bg=C_BG)
        frm_lista.pack(fill="x")

        for i, (codigo, nombre) in enumerate(_CENTROS_LIST):
            var = tk.BooleanVar(value=True)
            self._vars[codigo] = var

            fila_bg = "#F4F7FB" if i % 2 == 0 else C_BG
            frm_fila = tk.Frame(frm_lista, bg=fila_bg, padx=6, pady=3)
            frm_fila.pack(fill="x")

            cb = tk.Checkbutton(frm_fila,
                                variable=var,
                                bg=fila_bg,
                                activebackground=fila_bg,
                                cursor="hand2",
                                relief="flat")
            cb.pack(side="left")

            tk.Label(frm_fila,
                     text=f"{codigo}",
                     bg=fila_bg, fg="#285FA8",
                     font=("Segoe UI", 9, "bold"),
                     width=6, anchor="w").pack(side="left")

            tk.Label(frm_fila,
                     text=nombre,
                     bg=fila_bg, fg=C_TEXT,
                     font=FONT_LABEL,
                     anchor="w").pack(side="left", padx=(4, 0))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(10, 0))

        # Contador de centros seleccionados (se actualiza en tiempo real)
        self._lbl_contador = tk.Label(cuerpo, text="",
                                      bg=C_BG, fg="#444",
                                      font=("Segoe UI", 8, "italic"))
        self._lbl_contador.pack(anchor="e", pady=(4, 0))
        self._actualizar_contador()

        # Trazar cambios en las variables para actualizar el contador
        for var in self._vars.values():
            var.trace_add("write", lambda *_: self._actualizar_contador())

        # Etiqueta de error
        self._lbl_error = tk.Label(cuerpo, text="",
                                   bg=C_BG, fg=C_ERROR,
                                   font=("Segoe UI", 8, "italic"))
        self._lbl_error.pack(anchor="w", pady=(2, 8))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        # Botones principales
        tk.Button(cuerpo,
                  text="✔  Confirmar selección y continuar",
                  command=self._confirmar,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=12, pady=10,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo, text="Cancelar",
                  command=self._cancelar,
                  font=FONT_LABEL, padx=12, pady=7,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8", activebackground="#C0C8D8").pack(fill="x")

    # ── Helpers UI ────────────────────────────────────────────────────────────
    def _seleccionar_todos(self):
        for var in self._vars.values():
            var.set(True)

    def _deseleccionar_todos(self):
        for var in self._vars.values():
            var.set(False)

    def _actualizar_contador(self):
        n = sum(1 for v in self._vars.values() if v.get())
        total = len(self._vars)
        color = C_ERROR if n == 0 else ("#1A7A40" if n == total else "#285FA8")
        self._lbl_contador.config(
            text=f"{n} de {total} centro(s) seleccionado(s)",
            fg=color)

    # ── Acciones ──────────────────────────────────────────────────────────────
    def _confirmar(self):
        self._lbl_error.config(text="")
        seleccionados = [
            (codigo, nombre)
            for (codigo, nombre) in _CENTROS_LIST
            if self._vars[codigo].get()
        ]
        if not seleccionados:
            self._lbl_error.config(
                text="⚠  Debe seleccionar al menos un centro para continuar.")
            return
        self.destroy()
        self._on_confirm(seleccionados)

    def _cancelar(self):
        self.destroy()
        self._on_confirm(False)


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
        self._barra("indeterminate", "Cruzando con LISTADO CODIGOS DALI-SAP...")
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
        nombre_listado = "LISTADO CODIGOS DALI - SAP.xlsx"

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

        _log(self._parent, f"LISTADO DALI-SAP localizado: {ruta_listado}", "INFO")

        # -- Cargar diccionario codigo_SAP --> (desc, unidad, grp, desc_grp) ──
        # Datos reales desde fila 8 (fila 7 = cabeceras).
        # La clave col D (indice 3) es int en el LISTADO; se normaliza a string.
        diccionario = {}
        try:
            wb_listado = openpyxl.load_workbook(ruta_listado,
                                                 read_only=True, data_only=True)
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
                 f"LISTADO DALI-SAP cargado: {len(diccionario)} entradas.", "INFO")
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
                wb_l  = openpyxl.load_workbook(ruta_listado, read_only=True, data_only=True)
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
                wb_l = openpyxl.load_workbook(ruta_listado,
                                               read_only=True, data_only=True)
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



# ─────────────────────────────────────────────────────────────────────────────
#  PASO 6 — VENTANA DE CONFIRMACIÓN FINAL
# ─────────────────────────────────────────────────────────────────────────────

class VentanaConfirmacionRINFO(tk.Toplevel):

    def __init__(self, parent, nombre_nuevo: str, ruta_nuevo: str,
                 nombre_sap: str, codigo_sap: str,
                 hojas_vaciadas: list, articulos_copiados: int):
        super().__init__(parent)
        self.title("Registro Info creado correctamente")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        self._ruta_nuevo = ruta_nuevo
        cargar_icono_ventana(self)

        self._construir_ui(nombre_nuevo, ruta_nuevo, nombre_sap, codigo_sap,
                            hojas_vaciadas, articulos_copiados)
        self._ajustar_geometria()
        self.grab_set()

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(660, sw - 40)
        # Alto: contenido real + margen; mínimo 520 para que los botones sean visibles
        h  = max(520, min(self.winfo_reqheight() + 40, sh - 60))
        x  = (sw - w) // 2
        y  = max(20, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(560, 520)

    def _construir_ui(self, nombre_nuevo, ruta_nuevo, nombre_sap, codigo_sap,
                       hojas_vaciadas, articulos_copiados):
        carpeta_guardado = os.path.dirname(ruta_nuevo)

        hdr = tk.Frame(self, bg="#1A7A40", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="✔  Registro Info creado correctamente",
                 bg="#1A7A40", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="El archivo ha sido generado y está listo para usarse en SAP.",
                 bg="#1A7A40", fg="#A8E6C0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        # ── Zona de botones — declarada ANTES del cuerpo para que pack()
        #    la reserve en la parte inferior antes de que cuerpo expanda.
        frm_pie = tk.Frame(self, bg=C_BG, padx=22, pady=12)
        frm_pie.pack(side="bottom", fill="x")

        tk.Frame(frm_pie, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(frm_pie, text="¿Desea abrir el archivo ahora?",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 8))

        tk.Button(frm_pie, text="✔  Sí, abrir ahora",
                  command=self._abrir_y_cerrar,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=14, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(frm_pie, text="No, cerrar",
                  command=self._cerrar,
                  font=FONT_LABEL, padx=14, pady=8,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(fill="x")

        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=16)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo, text="✔  Archivo generado:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  📄  {nombre_nuevo}",
                 bg="#EEF7EE", fg="#1B5E20",
                 font=("Segoe UI", 10, "bold"),
                 padx=10, pady=8, relief="flat").pack(fill="x", pady=(4, 10))

        tk.Label(cuerpo, text="Guardado en:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  📁  {carpeta_guardado}",
                 bg="#F5F5F5", fg="#333",
                 font=FONT_LABEL, padx=10, pady=6, relief="flat",
                 wraplength=560, justify="left").pack(fill="x", pady=(4, 10))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(cuerpo, text="Valores escritos en la hoja principal:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")

        frm_vals = tk.Frame(cuerpo, bg="#F0F5FB", padx=12, pady=8)
        frm_vals.pack(fill="x", pady=(4, 10))
        _fila_valor(frm_vals, "D2 — Proveedor SAP:", nombre_sap)
        _fila_valor(frm_vals, "C2 — Código SAP:",    codigo_sap or "(sin código)")
        _fila_valor(frm_vals, "Columna E (desde E2):",
                    f"{articulos_copiados} artículo(s) copiado(s) de la tarifa")

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        if hojas_vaciadas:
            tk.Label(cuerpo,
                     text=f"Hojas creadas ({len(hojas_vaciadas)}):",
                     bg=C_BG, fg=C_TEXT,
                     font=("Segoe UI", 9, "bold")).pack(anchor="w")
            for h in hojas_vaciadas:
                tk.Label(cuerpo, text=f"     •  {h}",
                         bg=C_BG, fg=C_TEXT, font=FONT_LABEL).pack(anchor="w")

        tk.Frame(cuerpo, bg=C_BG, height=6).pack()

    def _abrir_y_cerrar(self):
        try:
            os.startfile(self._ruta_nuevo)
        except AttributeError:
            subprocess.call(["xdg-open", self._ruta_nuevo])
        except Exception:
            pass
        self._cerrar()

    def _cerrar(self):
        parent = self.master
        self.destroy()
        try:
            if parent and parent.winfo_exists():
                parent.deiconify()
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────────────────────
#  HELPER UI
# ─────────────────────────────────────────────────────────────────────────────

def _fila_valor(parent_frame, etiqueta: str, valor: str):
    """Muestra una fila etiqueta: valor en el panel de confirmación."""
    frm = tk.Frame(parent_frame, bg="#F0F5FB")
    frm.pack(fill="x", pady=1)
    tk.Label(frm, text=etiqueta,
             bg="#F0F5FB", fg="#555",
             font=("Segoe UI", 8, "bold"),
             width=28, anchor="w").pack(side="left")
    tk.Label(frm, text=valor,
             bg="#F0F5FB", fg="#1A2E4A",
             font=("Segoe UI", 9),
             anchor="w").pack(side="left", padx=(4, 0))
