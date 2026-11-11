# -*- coding: utf-8 -*-
"""
comun_rinfo.py
Constantes, log y funciones auxiliares del motor RINFO (Excel, similitud de nombres,
búsqueda de tarifas, Purchasing y listados). No importa de ningún otro módulo de app.rinfo.

Separado de crear_registro_info.py en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""

import os
import re
import shutil
import unicodedata
import tkinter as tk
import datetime as _dt_mod          # módulo completo — usado en lógica de fechas de baja
from datetime import datetime
from copy import copy
import zipfile
import openpyxl
# normalizar_codigo_sap vive en sap_helpers.py (sin tkinter) para que
# también pueda importarla test_crear_registro_info.py sin arrastrar
# la UI. Ver ese fichero para la implementación y el porqué.
from app.comun.sap_helpers import (  # noqa: E402
    normalizar_codigo_sap, ultima_fila_con_datos, parsear_fecha_celda,
    formatear_fecha_dmy, clave_pir_h3, material_de_clave_pir, formatear_clave_pir,
    debe_cerrar_fin_validez)



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
