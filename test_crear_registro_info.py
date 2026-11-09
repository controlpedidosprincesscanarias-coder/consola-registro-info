# -*- coding: utf-8 -*-
"""
test_crear_registro_info.py
───────────────────────────
Tests automáticos de casos límite (edge cases) para la lógica de
crear_registro_info.py  — sin dependencias de tkinter, SAP ni ficheros reales.

Cubre:
  TC-01  Múltiples overlaps de vigencia (mismo material, varios periodos solapados)
  TC-02  Fechas abiertas (fin = None) → siempre se considera solapamiento
  TC-03  Columna J vacía (sin fecha fin) en todas las filas → sin crash
  TC-04  Varios centros SAP simultáneos → bloques correctamente generados
  TC-05  Múltiples organizaciones de compras → cada org genera su bloque
  TC-06  Múltiples proveedores mismo material → sin dedup entre proveedores
  TC-07  Deduplicación H2: mismo material + centro repetido → solo una fila
  TC-08  Deduplicación H3: misma clave completa → solo una fila
  TC-09  Gap de vigencia detectado (fecha fin / inicio siguiente con hueco)
  TC-10  Sin gap cuando los periodos son contiguos (fin + 1 día = inicio siguiente)
  TC-11  Centros sin extensión → lista de centros faltantes correcta
  TC-12  Registros omitidos Hoja1 → materiales sin extensión detectados
  TC-13  Normalización de código SAP float → str sin decimales
  TC-14  Material None en Hoja1 → fila ignorada

Central Compras Princess Canarias — (c) VAMA 2026
"""

import datetime
import unittest
from collections import defaultdict


# ─────────────────────────────────────────────────────────────────────────────
#  Helpers extraídos / replicados (sin importar el módulo completo)
# ─────────────────────────────────────────────────────────────────────────────

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


def _parse_fecha(valor):
    """Parsea dd.mm.yyyy o date/datetime. Devuelve date o None."""
    if valor is None:
        return None
    if isinstance(valor, (datetime.date, datetime.datetime)):
        return valor if isinstance(valor, datetime.date) else valor.date()
    txt = str(valor).strip()
    try:
        d, m, y = txt.split(".")
        return datetime.date(int(y), int(m), int(d))
    except Exception:
        return None


import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sap_helpers import normalizar_codigo_sap as _normalizar_codigo_sap
from sap_helpers import ultima_fila_con_datos
from sap_helpers import parsear_fecha_celda
from sap_helpers import (
    formatear_fecha_dmy, clave_pir_h3, material_de_clave_pir, formatear_clave_pir)
from sap_helpers import eliminar_filas_agrupadas
# ↑ Import real (antes había aquí una reimplementación paralela que
#   podía divergir del código de producción sin que ningún test lo
#   detectara). Ahora TC13 protege la función que usa de verdad
#   crear_registro_info.py en sus 31 puntos de uso.


def _detectar_overlaps(filas_h3):
    """
    filas_h3 : lista de dict con claves 'material', 'centro', 'fi', 'ff'
    Devuelve lista de strings descriptivos con los overlaps detectados.
    """
    periodos = defaultdict(list)
    for f in filas_h3:
        mat  = str(f["material"]).strip() if f["material"] else ""
        ctr  = str(f["centro"]).strip()   if f["centro"]   else ""
        fi   = _parse_fecha(f["fi"])
        ff   = _parse_fecha(f["ff"])
        if not mat or fi is None:
            continue
        periodos[(mat, ctr)].append((fi, ff))

    avisos = []
    for (mat, ctr), rangos in periodos.items():
        rangos_sorted = sorted(rangos, key=lambda r: r[0])
        for i in range(len(rangos_sorted) - 1):
            fi_a, ff_a = rangos_sorted[i]
            fi_b, _    = rangos_sorted[i + 1]
            if ff_a is None or fi_b <= ff_a:
                avisos.append(f"OVERLAP mat={mat} ctr={ctr} ff_a={ff_a} fi_b={fi_b}")
    return avisos


def _detectar_gaps(filas_h3):
    """Devuelve lista de strings con gaps de vigencia detectados."""
    periodos = defaultdict(list)
    for f in filas_h3:
        mat = str(f["material"]).strip() if f["material"] else ""
        ctr = str(f["centro"]).strip()   if f["centro"]   else ""
        fi  = _parse_fecha(f["fi"])
        ff  = _parse_fecha(f["ff"])
        if not mat or fi is None:
            continue
        periodos[(mat, ctr)].append((fi, ff))

    avisos = []
    for (mat, ctr), rangos in periodos.items():
        if len(rangos) < 2:
            continue
        rangos_sorted = sorted(rangos, key=lambda r: r[0])
        for i in range(len(rangos_sorted) - 1):
            _, ff_a = rangos_sorted[i]
            fi_b, _ = rangos_sorted[i + 1]
            if ff_a is None:
                continue
            un_dia = datetime.timedelta(days=1)
            try:
                if fi_b > ff_a + un_dia:
                    avisos.append(
                        f"GAP mat={mat} ctr={ctr} fin={ff_a} inicio={fi_b} "
                        f"dias={(fi_b - ff_a).days - 1}"
                    )
            except OverflowError:
                pass   # ff_a = 31.12.9999 → + 1 día desborda; no hay gap posible
    return avisos


def _dedup_h2(filas, codigo_sap, org="IC00"):
    """
    Simula la deduplicación de Hoja2.
    filas : lista de (reg_info, material, centro)
    Devuelve la lista de filas únicas (sin duplicados funcionales).
    """
    vistas = set()
    resultado = []
    prov = str(codigo_sap).strip() if codigo_sap else ""
    for reg_info, material, centro in filas:
        mat = str(material).strip() if material else ""
        ctr = str(centro).strip()   if centro   else ""
        clave = f"{prov}|{org}|{mat}|{ctr}"
        if clave in vistas:
            continue
        vistas.add(clave)
        resultado.append((reg_info, material, centro))
    return resultado


def _dedup_h3(filas, codigo_sap, org="IC00", fi="01.01.2025", ff="31.12.2025"):
    """
    Simula la deduplicación de Hoja3.
    filas : lista de (reg_info, material, centro)
    Devuelve la lista de filas únicas (sin duplicados funcionales).
    """
    vistas = set()
    resultado = []
    prov = str(codigo_sap).strip() if codigo_sap else ""
    for reg_info, material, centro in filas:
        mat = str(material).strip() if material else ""
        ctr = str(centro).strip()   if centro   else ""
        clave = f"{prov}|{org}|{mat}|{ctr}|{fi}|{ff}"
        if clave in vistas:
            continue
        vistas.add(clave)
        resultado.append((reg_info, material, centro))
    return resultado


def _centros_faltantes(filas_h2_centros):
    """
    filas_h2_centros : lista de códigos de centro presentes en Hoja2
    Devuelve el set de centros de CENTROS_SAP_CANARIAS no representados.
    """
    esperados = set(CENTROS_SAP_CANARIAS.keys())
    presentes = {str(c).strip() for c in filas_h2_centros if c}
    return esperados - presentes


def _registros_omitidos(mats_h1, mats_h2):
    """Devuelve materiales en Hoja1 que no aparecen en Hoja2."""
    return set(mats_h1) - set(mats_h2)


# ─────────────────────────────────────────────────────────────────────────────
#  Tests
# ─────────────────────────────────────────────────────────────────────────────

class TC01_MultipleOverlaps(unittest.TestCase):
    """Múltiples overlaps: tres periodos solapados para el mismo material/centro."""

    def test_detecta_dos_overlaps(self):
        filas = [
            {"material": "MAT001", "centro": "0110", "fi": "01.01.2025", "ff": "30.06.2025"},
            {"material": "MAT001", "centro": "0110", "fi": "15.06.2025", "ff": "31.12.2025"},
            {"material": "MAT001", "centro": "0110", "fi": "20.12.2025", "ff": "28.02.2026"},
        ]
        avisos = _detectar_overlaps(filas)
        # Periodo 1-2 solapa, periodo 2-3 solapa → 2 avisos
        self.assertEqual(len(avisos), 2, f"Se esperaban 2 overlaps, obtenidos: {avisos}")

    def test_sin_overlap_cuando_no_hay(self):
        filas = [
            {"material": "MAT002", "centro": "0210", "fi": "01.01.2025", "ff": "30.06.2025"},
            {"material": "MAT002", "centro": "0210", "fi": "01.07.2025", "ff": "31.12.2025"},
        ]
        avisos = _detectar_overlaps(filas)
        self.assertEqual(avisos, [], f"No se esperaban overlaps: {avisos}")


class TC02_FechasAbiertas(unittest.TestCase):
    """Fecha fin None (fecha abierta) → siempre reporta solapamiento con periodo siguiente."""

    def test_fecha_abierta_solapa(self):
        filas = [
            {"material": "MAT003", "centro": "0310", "fi": "01.01.2025", "ff": None},
            {"material": "MAT003", "centro": "0310", "fi": "01.07.2025", "ff": "31.12.2025"},
        ]
        avisos = _detectar_overlaps(filas)
        self.assertEqual(len(avisos), 1, f"Fecha abierta debe generar overlap: {avisos}")

    def test_unico_periodo_abierto_sin_overlap(self):
        filas = [
            {"material": "MAT004", "centro": "0410", "fi": "01.01.2025", "ff": None},
        ]
        avisos = _detectar_overlaps(filas)
        self.assertEqual(avisos, [], "Un único periodo abierto no genera overlap")


class TC03_ColumnaJVacia(unittest.TestCase):
    """Columna J (fecha fin) vacía en todas las filas → sin crash, sin gaps."""

    def test_sin_crash_ni_gaps(self):
        filas = [
            {"material": "MAT005", "centro": "0510", "fi": "01.01.2025", "ff": None},
            {"material": "MAT005", "centro": "0510", "fi": "01.07.2025", "ff": None},
        ]
        try:
            gaps = _detectar_gaps(filas)
        except Exception as e:
            self.fail(f"_detectar_gaps lanzó excepción inesperada: {e}")
        # Con ff=None no se puede calcular gap → lista vacía
        self.assertEqual(gaps, [], f"No deben reportarse gaps con ff=None: {gaps}")


class TC04_VariosCentros(unittest.TestCase):
    """Varios centros SAP → cada material genera una fila por centro."""

    def test_expansion_10_centros(self):
        materiales = ["MAT-A", "MAT-B", "MAT-C"]
        centros    = list(CENTROS_SAP_CANARIAS.keys())   # 10 centros
        filas_expandidas = []
        for mat in materiales:
            for c in centros:
                filas_expandidas.append(("1", mat, c))

        resultado = _dedup_h2(filas_expandidas, "PROV001")
        # Sin duplicados → 3 materiales × 10 centros = 30 filas
        self.assertEqual(len(resultado), 30)

    def test_solo_centros_activos(self):
        # Si se usan solo 3 centros, deben generarse 2 materiales × 3 centros = 6 filas
        materiales = ["MAT-X", "MAT-Y"]
        centros    = ["0110", "0210", "0310"]
        filas = [(f"R{i}", m, c) for i, (m, c) in enumerate(
            [(m, c) for m in materiales for c in centros], start=1)]
        resultado = _dedup_h2(filas, "PROV002")
        self.assertEqual(len(resultado), 6)


class TC05_MultiplesOrgsCompras(unittest.TestCase):
    """Múltiples organizaciones de compras → dedup trata cada org como independiente."""

    def test_misma_mat_distinta_org_no_dedup(self):
        # Con org_a e org_b distintas, la misma combinación mat+centro no se deduplica
        filas_org_a = [("1", "MAT100", "0110")]
        filas_org_b = [("1", "MAT100", "0110")]
        res_a = _dedup_h2(filas_org_a, "PROV003", org="IC00")
        res_b = _dedup_h2(filas_org_b, "PROV003", org="IC01")
        # Cada resultado tiene 1 fila — no se mezclan entre orgs
        self.assertEqual(len(res_a), 1)
        self.assertEqual(len(res_b), 1)


class TC06_MultiplesProveedoresMismoMaterial(unittest.TestCase):
    """Múltiples proveedores con el mismo material → no hay dedup entre proveedores."""

    def test_dos_proveedores_mismo_mat(self):
        filas = [("1", "MAT200", "0110")]
        res_prov1 = _dedup_h2(filas, "PROV_A")
        res_prov2 = _dedup_h2(filas, "PROV_B")
        # Cada proveedor tiene su propia entrada — no se eliminan mutuamente
        self.assertEqual(len(res_prov1), 1)
        self.assertEqual(len(res_prov2), 1)

    def test_mismo_proveedor_mismo_mat_sí_dedup(self):
        filas = [("1", "MAT200", "0110"), ("2", "MAT200", "0110")]
        resultado = _dedup_h2(filas, "PROV_A")
        self.assertEqual(len(resultado), 1, "Mismo proveedor+mat+centro debe deduplicarse")


class TC07_DedupH2(unittest.TestCase):
    """Deduplicación H2: mismo material + centro repetido → solo una fila."""

    def test_dedup_basico(self):
        filas = [
            ("1", "MAT300", "0110"),
            ("1", "MAT300", "0110"),  # duplicado exacto
            ("2", "MAT300", "0210"),  # distinto centro → se conserva
        ]
        resultado = _dedup_h2(filas, "PROV_X")
        self.assertEqual(len(resultado), 2)
        centros = [f[2] for f in resultado]
        self.assertIn("0110", centros)
        self.assertIn("0210", centros)

    def test_dedup_10_bloques_mismo_material(self):
        # Simula la expansión de 10 bloques con un único material
        filas = [("1", "MAT301", c) for c in CENTROS_SAP_CANARIAS.keys()] * 3
        resultado = _dedup_h2(filas, "PROV_Y")
        # 10 centros × 1 material → exactamente 10 filas únicas
        self.assertEqual(len(resultado), 10)


class TC08_DedupH3(unittest.TestCase):
    """Deduplicación H3: misma clave completa (prov|org|mat|centro|fi|ff) → una fila."""

    def test_dedup_h3_basico(self):
        filas = [
            ("1", "MAT400", "0110"),
            ("1", "MAT400", "0110"),  # duplicado
            ("2", "MAT400", "0210"),  # distinto centro
        ]
        resultado = _dedup_h3(filas, "PROV_Z")
        self.assertEqual(len(resultado), 2)

    def test_dedup_h3_fechas_distintas_no_dedup(self):
        # Mismo mat+centro pero fechas diferentes → claves distintas → no se dedup
        res1 = _dedup_h3([("1", "MAT401", "0110")], "PROV_Z",
                          fi="01.01.2024", ff="31.12.2024")
        res2 = _dedup_h3([("1", "MAT401", "0110")], "PROV_Z",
                          fi="01.01.2025", ff="31.12.2025")
        # Cada llamada independiente recibe 1 fila y debe conservarla
        self.assertEqual(len(res1), 1, "Periodo 2024 debe conservarse")
        self.assertEqual(len(res2), 1, "Periodo 2025 debe conservarse")
        # Si se juntan ambas listas en una sola llamada con fi/ff distintas,
        # la función mantiene la primera vista y deduplica la segunda.
        # Ese caso queda cubierto por TC07/TC08 básico.


class TC09_GapVigencia(unittest.TestCase):
    """Gap de vigencia detectado cuando hay hueco entre periodos consecutivos."""

    def test_gap_de_15_dias(self):
        filas = [
            {"material": "MAT500", "centro": "0610", "fi": "01.01.2025", "ff": "31.05.2025"},
            {"material": "MAT500", "centro": "0610", "fi": "15.06.2025", "ff": "31.12.2025"},
        ]
        gaps = _detectar_gaps(filas)
        self.assertEqual(len(gaps), 1, f"Se esperaba 1 gap: {gaps}")
        self.assertIn("MAT500", gaps[0])

    def test_sin_gap_cuando_contiguo(self):
        filas = [
            {"material": "MAT501", "centro": "0710", "fi": "01.01.2025", "ff": "30.06.2025"},
            {"material": "MAT501", "centro": "0710", "fi": "01.07.2025", "ff": "31.12.2025"},
        ]
        gaps = _detectar_gaps(filas)
        self.assertEqual(gaps, [], f"Periodos contiguos no deben generar gap: {gaps}")


class TC10_PeriodosContiguos(unittest.TestCase):
    """Confirma que fin+1día = inicio_siguiente NO genera ni gap ni overlap."""

    def test_sin_gap_ni_overlap(self):
        # 30.06 + 1 día = 01.07 → contiguos, sin hueco ni solape
        filas = [
            {"material": "MAT600", "centro": "0810", "fi": "01.01.2025", "ff": "30.06.2025"},
            {"material": "MAT600", "centro": "0810", "fi": "01.07.2025", "ff": "31.12.2025"},
        ]
        gaps    = _detectar_gaps(filas)
        overlaps = _detectar_overlaps(filas)
        self.assertEqual(gaps,    [], f"No debe haber gaps: {gaps}")
        self.assertEqual(overlaps, [], f"No debe haber overlaps: {overlaps}")


class TC11_CentrosSinExtension(unittest.TestCase):
    """Centros sin extensión en Hoja2 → lista de centros faltantes correcta."""

    def test_dos_centros_faltantes(self):
        # Se proporcionan 8 de los 10 centros
        centros_h2 = [c for c in CENTROS_SAP_CANARIAS.keys()
                      if c not in ("0810", "0910")]   # faltan 2
        faltantes = _centros_faltantes(centros_h2)
        self.assertEqual(faltantes, {"0810", "0910"})

    def test_todos_los_centros_presentes(self):
        centros_h2 = list(CENTROS_SAP_CANARIAS.keys())
        faltantes = _centros_faltantes(centros_h2)
        self.assertEqual(faltantes, set())

    def test_hoja2_completamente_vacia(self):
        faltantes = _centros_faltantes([])
        self.assertEqual(faltantes, set(CENTROS_SAP_CANARIAS.keys()))


class TC12_RegistrosOmitidosH1(unittest.TestCase):
    """Materiales en Hoja1 sin extensión en Hoja2 → correctamente detectados."""

    def test_detecta_omitidos(self):
        mats_h1 = {"MAT001", "MAT002", "MAT003"}
        mats_h2 = {"MAT001", "MAT003"}   # MAT002 no está en Hoja2
        omitidos = _registros_omitidos(mats_h1, mats_h2)
        self.assertEqual(omitidos, {"MAT002"})

    def test_sin_omitidos(self):
        mats_h1 = {"MAT001", "MAT002"}
        mats_h2 = {"MAT001", "MAT002", "MAT003"}  # Hoja2 tiene más, no importa
        omitidos = _registros_omitidos(mats_h1, mats_h2)
        self.assertEqual(omitidos, set())

    def test_todos_omitidos(self):
        mats_h1 = {"MAT010", "MAT011"}
        mats_h2 = set()
        omitidos = _registros_omitidos(mats_h1, mats_h2)
        self.assertEqual(omitidos, mats_h1)


class TC13_NormalizacionCodigoSAP(unittest.TestCase):
    """Normalización de código SAP float → str sin decimales."""

    def test_float_entero(self):
        self.assertEqual(_normalizar_codigo_sap(10800044.0), "10800044")

    def test_int(self):
        self.assertEqual(_normalizar_codigo_sap(10800044), "10800044")

    def test_string(self):
        self.assertEqual(_normalizar_codigo_sap("10800044"), "10800044")

    def test_string_con_espacios(self):
        self.assertEqual(_normalizar_codigo_sap("  10800044  "), "10800044")

    def test_none(self):
        self.assertIsNone(_normalizar_codigo_sap(None))

    def test_float_con_decimales_reales(self):
        # Float no entero → no se trunca → se deja como string
        resultado = _normalizar_codigo_sap(1080.5)
        self.assertEqual(resultado, "1080.5")


class TC14_MaterialNoneIgnorado(unittest.TestCase):
    """Filas con material None en Hoja1 deben ignorarse completamente."""

    def test_dedup_ignora_material_none(self):
        # Una fila con material None + una fila válida
        filas_raw = [(None, None), ("MAT999", "0110")]
        filas_validas = [(mat, ctr) for mat, ctr in filas_raw if mat is not None]
        self.assertEqual(len(filas_validas), 1)
        self.assertEqual(filas_validas[0], ("MAT999", "0110"))

    def test_overlap_detection_ignora_material_none(self):
        filas = [
            {"material": None, "centro": "0110", "fi": "01.01.2025", "ff": "30.06.2025"},
            {"material": "MAT998", "centro": "0110", "fi": "01.01.2025", "ff": "30.06.2025"},
        ]
        # None no debe generar overlap consigo mismo ni con MAT998
        avisos = _detectar_overlaps(filas)
        self.assertEqual(avisos, [])


class TC15_PeriodosInfinitos(unittest.TestCase):
    """
    Períodos 'infinitos' (fi muy antigua + ff muy lejana, p.ej. 01.01.1900 / 31.12.9999).
    Casos típicos en SAP cuando no se acota la vigencia.
    """

    def test_periodo_infinito_unico_sin_overlap(self):
        filas = [
            {"material": "MAT700", "centro": "0110",
             "fi": "01.01.1900", "ff": "31.12.9999"},
        ]
        avisos = _detectar_overlaps(filas)
        self.assertEqual(avisos, [], "Un único periodo infinito no genera overlap")

    def test_dos_periodos_infinitos_mismo_mat_detecta_overlap(self):
        filas = [
            {"material": "MAT701", "centro": "0110",
             "fi": "01.01.2020", "ff": "31.12.9999"},
            {"material": "MAT701", "centro": "0110",
             "fi": "01.01.2025", "ff": "31.12.9999"},
        ]
        avisos = _detectar_overlaps(filas)
        self.assertEqual(len(avisos), 1,
                         f"Dos periodos infinitos deben solaparse: {avisos}")

    def test_periodo_infinito_distinto_centro_no_overlap(self):
        filas = [
            {"material": "MAT702", "centro": "0110",
             "fi": "01.01.2020", "ff": "31.12.9999"},
            {"material": "MAT702", "centro": "0210",
             "fi": "01.01.2020", "ff": "31.12.9999"},
        ]
        # Centros distintos → claves distintas → no hay overlap entre ellos
        avisos = _detectar_overlaps(filas)
        self.assertEqual(avisos, [],
                         "Periodos infinitos en centros distintos no son overlap")

    def test_periodo_infinito_sin_gap_posible(self):
        filas = [
            {"material": "MAT703", "centro": "0310",
             "fi": "01.01.2020", "ff": "31.12.9999"},
            {"material": "MAT703", "centro": "0310",
             "fi": "01.01.2025", "ff": "31.12.9999"},
        ]
        # ff=31.12.9999 no es None, pero el segundo inicio está dentro del primero
        # → el detector de gaps no aplica (ya lo cubre el de overlaps)
        gaps = _detectar_gaps(filas)
        self.assertEqual(gaps, [],
                         "Dentro de un overlap no debe reportarse gap además")


class TC16_MaterialesRepetidos(unittest.TestCase):
    """
    Materiales repetidos en la tarifa (mismo código SAP en múltiples filas).
    El dedup H2/H3 debe asegurarse de que solo se inserta una fila por
    combinación funcional única.
    """

    def test_material_repetido_tres_veces_en_tarifa(self):
        # El mismo material aparece 3 veces en la tarifa (error del proveedor)
        filas = [
            ("1", "MAT800", "0110"),
            ("2", "MAT800", "0110"),   # repetición
            ("3", "MAT800", "0110"),   # repetición
        ]
        resultado = _dedup_h2(filas, "PROV_DUP")
        self.assertEqual(len(resultado), 1,
                         "Material repetido 3 veces debe deduplicarse a 1 fila")

    def test_material_repetido_distintos_centros_conserva_todos(self):
        # El mismo material se repite, pero en centros distintos → se conservan
        filas = [
            ("1", "MAT801", "0110"),
            ("2", "MAT801", "0210"),
            ("3", "MAT801", "0110"),   # repetición de centro 0110 → se elimina
        ]
        resultado = _dedup_h2(filas, "PROV_DUP2")
        self.assertEqual(len(resultado), 2,
                         "2 centros distintos → 2 filas, 3ª repetición eliminada")
        centros_res = {f[2] for f in resultado}
        self.assertEqual(centros_res, {"0110", "0210"})

    def test_material_repetido_en_h3_con_mismas_fechas(self):
        filas = [
            ("1", "MAT802", "0110"),
            ("1", "MAT802", "0110"),   # duplicado exacto
        ]
        resultado = _dedup_h3(filas, "PROV_DUP3",
                               fi="01.01.2025", ff="31.12.2025")
        self.assertEqual(len(resultado), 1,
                         "Duplicado exacto en H3 debe eliminarse")

    def test_multiples_materiales_sin_dedup_entre_distintos(self):
        # Varios materiales distintos → todos se conservan aunque haya repeticiones
        filas = [
            ("1", "MAT810", "0110"),
            ("2", "MAT811", "0110"),
            ("3", "MAT812", "0110"),
            ("1", "MAT810", "0110"),   # repetición de MAT810 → se elimina
        ]
        resultado = _dedup_h2(filas, "PROV_MULTI")
        self.assertEqual(len(resultado), 3,
                         "3 materiales distintos → 3 filas, duplicado eliminado")
        mats = {f[1] for f in resultado}
        self.assertEqual(mats, {"MAT810", "MAT811", "MAT812"})


class TC17_UltimaFilaConDatos(unittest.TestCase):
    """ultima_fila_con_datos() — detección de última fila con datos en col E."""

    def _hoja_con_valores(self, valores_col_e):
        """Crea una hoja en memoria con valores en columna E a partir de la fila 2."""
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        for i, val in enumerate(valores_col_e, start=2):
            ws.cell(row=i, column=5, value=val)
        return ws

    def test_hoja_vacia_devuelve_1(self):
        ws = self._hoja_con_valores([])
        self.assertEqual(ultima_fila_con_datos(ws), 1)

    def test_datos_contiguos(self):
        ws = self._hoja_con_valores(["A", "B", "C"])
        # filas 2, 3, 4 con datos → última fila = 4
        self.assertEqual(ultima_fila_con_datos(ws), 4)

    def test_huecos_intermedios_no_afectan_ultima(self):
        ws = self._hoja_con_valores(["A", None, "C"])
        # fila 2=A, fila 3=None, fila 4=C → última fila con datos = 4
        self.assertEqual(ultima_fila_con_datos(ws), 4)

    def test_columna_distinta(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=2, column=2, value="X")  # columna B, no E
        ws.cell(row=5, column=2, value="Y")
        self.assertEqual(ultima_fila_con_datos(ws, columna=2), 5)
        self.assertEqual(ultima_fila_con_datos(ws, columna=5), 1)  # col E vacía


class TC18_ParsearFechaCelda(unittest.TestCase):
    """parsear_fecha_celda() — normaliza valores de celda a datetime.date.

    Consolida 5 copias idénticas que había repartidas por
    crear_registro_info.py (_parse_fecha, _a_date_h3, _parse_date_bis,
    y dos variantes inline) en una sola función testeada.
    """

    def test_none_devuelve_none(self):
        self.assertIsNone(parsear_fecha_celda(None))

    def test_datetime_devuelve_date(self):
        dt = datetime.datetime(2026, 7, 21, 10, 30)
        self.assertEqual(parsear_fecha_celda(dt), datetime.date(2026, 7, 21))

    def test_date_se_devuelve_igual(self):
        d = datetime.date(2026, 7, 21)
        self.assertEqual(parsear_fecha_celda(d), d)

    def test_string_formato_punto(self):
        self.assertEqual(parsear_fecha_celda("21.07.2026"), datetime.date(2026, 7, 21))

    def test_string_formato_barra(self):
        self.assertEqual(parsear_fecha_celda("21/07/2026"), datetime.date(2026, 7, 21))

    def test_string_formato_iso(self):
        self.assertEqual(parsear_fecha_celda("2026-07-21"), datetime.date(2026, 7, 21))

    def test_string_con_espacios_se_recorta(self):
        self.assertEqual(parsear_fecha_celda("  21.07.2026  "), datetime.date(2026, 7, 21))

    def test_string_formato_no_reconocido_devuelve_none(self):
        self.assertIsNone(parsear_fecha_celda("21 de julio de 2026"))

    def test_tipo_no_soportado_devuelve_none(self):
        self.assertIsNone(parsear_fecha_celda(12345))  # int, no soportado


class TC19_HelpersH3(unittest.TestCase):
    """formatear_fecha_dmy / clave_pir_h3 / material_de_clave_pir /
    formatear_clave_pir — funciones puras extraídas de
    _actualizar_h3_heredados() (limpieza conservadora, sin trocear el
    bucle principal, que sigue teniendo demasiado estado compartido
    para dividirse con seguridad)."""

    def test_formatear_fecha_dmy(self):
        self.assertEqual(formatear_fecha_dmy(datetime.date(2026, 7, 21)), "21.07.2026")

    def _fila_h3(self, reg_info, org, centro, material):
        """Construye una fila real de Hoja 3 (celdas, no valores) con
        openpyxl, tal como la devuelve ws.iter_rows(...)."""
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        # Layout Hoja3: A=reg_info, D=org, E=centro, F=material
        ws.cell(row=1, column=1, value=reg_info)
        ws.cell(row=1, column=4, value=org)
        ws.cell(row=1, column=5, value=centro)
        ws.cell(row=1, column=6, value=material)
        return next(ws.iter_rows(min_row=1, max_row=1))

    def test_clave_pir_h3_completa(self):
        fila = self._fila_h3("5300001473", "IC00", "0110", 10101583)
        self.assertEqual(clave_pir_h3(fila), ("5300001473", "IC00", "10101583", "0110"))

    def test_clave_pir_h3_normaliza_material_float(self):
        # El material puede llegar como float (10101583.0) — debe normalizarse
        fila = self._fila_h3("5300001473", "IC00", "0110", 10101583.0)
        self.assertEqual(clave_pir_h3(fila)[2], "10101583")

    def test_clave_pir_h3_sin_material_devuelve_none(self):
        fila = self._fila_h3("5300001473", "IC00", "0110", None)
        self.assertIsNone(clave_pir_h3(fila))

    def test_clave_pir_h3_fila_corta_devuelve_none(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value="x")   # solo col A, faltan D/E/F
        fila = next(ws.iter_rows(min_row=1, max_row=1))
        self.assertIsNone(clave_pir_h3(fila))

    def test_material_de_clave_pir(self):
        clave = ("5300001473", "IC00", "10101583", "0110")
        self.assertEqual(material_de_clave_pir(clave), "10101583")

    def test_material_de_clave_pir_none(self):
        self.assertEqual(material_de_clave_pir(None), "")

    def test_formatear_clave_pir(self):
        clave = ("5300001473", "IC00", "10101583", "0110")
        self.assertEqual(formatear_clave_pir(clave), "5300001473|IC00|10101583|0110")

    def test_formatear_clave_pir_con_campos_vacios(self):
        clave = ("", "IC00", "10101583", "")
        self.assertEqual(formatear_clave_pir(clave), "IC00|10101583")


class TC20_EliminarFilasAgrupadas(unittest.TestCase):
    """eliminar_filas_agrupadas() — borrado físico de filas por rangos
    contiguos en vez de fila a fila, para evitar el coste ~O(n²) de
    llamar a ws.delete_rows() una vez por fila dispersa (fix de
    rendimiento real: ~6 min -> 16s en un caso con 4518 filas)."""

    def _construir_ws(self, n):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        for i in range(1, n + 1):
            ws.cell(row=i, column=1, value=f"v{i}")
        return wb, ws

    def _borrar_fila_a_fila(self, ws, filas):
        """Referencia: el comportamiento ORIGINAL (antes del fix), para
        comparar contra el nuevo método en cada test."""
        for f in sorted(set(filas), reverse=True):
            ws.delete_rows(f)

    def _valores(self, ws):
        return [c[0].value for c in ws.iter_rows(min_row=1, max_row=ws.max_row)]

    def test_lista_vacia_no_hace_nada(self):
        _, ws = self._construir_ws(10)
        antes = self._valores(ws)
        eliminar_filas_agrupadas(ws, [])
        self.assertEqual(self._valores(ws), antes)

    def test_una_sola_fila_intermedia(self):
        _, ws1 = self._construir_ws(20)
        self._borrar_fila_a_fila(ws1, [10])
        _, ws2 = self._construir_ws(20)
        eliminar_filas_agrupadas(ws2, [10])
        self.assertEqual(self._valores(ws1), self._valores(ws2))

    def test_bloque_consecutivo_unico(self):
        filas = list(range(10, 51))
        _, ws1 = self._construir_ws(100)
        self._borrar_fila_a_fila(ws1, filas)
        _, ws2 = self._construir_ws(100)
        eliminar_filas_agrupadas(ws2, filas)
        self.assertEqual(self._valores(ws1), self._valores(ws2))

    def test_filas_dispersas_alternas(self):
        filas = list(range(2, 101, 2))
        _, ws1 = self._construir_ws(100)
        self._borrar_fila_a_fila(ws1, filas)
        _, ws2 = self._construir_ws(100)
        eliminar_filas_agrupadas(ws2, filas)
        self.assertEqual(self._valores(ws1), self._valores(ws2))

    def test_extremos_primera_y_ultima_fila(self):
        filas = [2, 100]
        _, ws1 = self._construir_ws(100)
        self._borrar_fila_a_fila(ws1, filas)
        _, ws2 = self._construir_ws(100)
        eliminar_filas_agrupadas(ws2, filas)
        self.assertEqual(self._valores(ws1), self._valores(ws2))

    def test_indices_duplicados_en_la_entrada(self):
        # El original repetiría delete_rows sobre la misma posición dos
        # veces (bug latente: borra una fila de más la segunda vez).
        # eliminar_filas_agrupadas deduplica -- comportamiento más seguro,
        # documentado como diferencia intencional.
        _, ws = self._construir_ws(20)
        eliminar_filas_agrupadas(ws, [10, 10, 11, 11, 12])
        # Debe borrar exactamente 3 filas (10, 11, 12), no más.
        self.assertEqual(ws.max_row, 17)

    def test_multiples_bloques_no_contiguos(self):
        filas = [3, 4, 5, 20, 21, 50]
        _, ws1 = self._construir_ws(60)
        self._borrar_fila_a_fila(ws1, filas)
        _, ws2 = self._construir_ws(60)
        eliminar_filas_agrupadas(ws2, filas)
        self.assertEqual(self._valores(ws1), self._valores(ws2))

    def test_aleatorio_multiples_semillas(self):
        import random
        for seed in range(5):
            random.seed(seed)
            n = random.choice([30, 80, 150])
            candidatas = list(range(2, n + 1))
            k = random.randint(1, len(candidatas))
            filas = random.sample(candidatas, k)

            _, ws1 = self._construir_ws(n)
            self._borrar_fila_a_fila(ws1, filas)
            _, ws2 = self._construir_ws(n)
            eliminar_filas_agrupadas(ws2, filas)
            self.assertEqual(
                self._valores(ws1), self._valores(ws2),
                f"Discrepancia con seed={seed}, n={n}, k={k}")


# ─────────────────────────────────────────────────────────────────────────────
#  TC21 — Solapamiento PARCIAL con precio distinto (bug real MAHOU, 2026-09-28)
#
#  A diferencia del resto de tests de este fichero (que replican la lógica para
#  no depender de tkinter), este ejecuta el CÓDIGO REAL de GeneradorDesdeExistente
#  encadenando los mismos pasos que _ejecutar(): 6c → 9b → 10. Se salta solo si
#  el módulo no se puede importar (p.ej. entorno sin tkinter).
#
#  Escenario: el Purchasing trae un período 01.09.2025→31.08.2026 a 19,36 y la
#  tarifa nueva llega con 01.02.2026→31.01.2027 a 19,83 (solapan parcialmente).
#  Esperado: el viejo se acorta a 31.01.2026 Y se crea el nuevo a 19,83.
#  Bug: el viejo se acortaba pero el nuevo NUNCA se creaba (el material se
#  quedaba con el precio viejo y sin cobertura desde el 01.02.2026, sin error).
# ─────────────────────────────────────────────────────────────────────────────

def _importar_modulo_real():
    """Importa crear_registro_info, o None si no es posible (sin tkinter, etc.)."""
    try:
        import crear_registro_info as _cri
        return _cri
    except Exception:
        return None


_CRI = _importar_modulo_real()


@unittest.skipIf(_CRI is None, "crear_registro_info no importable (¿sin tkinter?)")
class TC21_SolapamientoParcialPrecioDistinto(unittest.TestCase):

    H1 = "Info general p.reg.info"
    H2 = "Datos org.compr.p.reg.info"
    H3 = "Det.precios p.registros info"
    REG, MAT = "5300000186", "10200147"
    CENTROS = ["0110", "0120"]

    def _ejecutar_cadena(self, periodo_heredado, precio_heredado, precio_nuevo,
                         fecha_inicio, fecha_fin, centros_con_h3=None):
        """Monta un libro mínimo y encadena los pasos 6c → 9b → 10 reales.
        Devuelve {centro: [(inicio, fin, precio), ...]} ordenado por inicio."""
        import unittest.mock as mock
        import openpyxl

        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = self.H1
        ws2 = wb.create_sheet(self.H2)
        ws3 = wb.create_sheet(self.H3)
        for ws, n in ((ws1, 44), (ws2, 41), (ws3, 27)):
            for c in range(1, n + 1):
                ws.cell(row=1, column=c, value=f"c{c}")

        ws1.cell(row=2, column=1, value=self.REG)
        ws1.cell(row=2, column=5, value=self.MAT)
        centros_h3 = self.CENTROS if centros_con_h3 is None else centros_con_h3
        r3 = 2
        for i, centro in enumerate(self.CENTROS):
            r = 2 + i
            ws2.cell(row=r, column=1, value=self.REG)
            ws2.cell(row=r, column=6, value=centro)
            ws2.cell(row=r, column=10, value=self.MAT)
            if centro not in centros_h3:
                continue   # Hoja 2 sí, Hoja 3 no (hueco del Purchasing)
            ws3.cell(row=r3, column=1, value=self.REG)
            ws3.cell(row=r3, column=4, value="IC00")
            ws3.cell(row=r3, column=5, value=centro)
            ws3.cell(row=r3, column=6, value=self.MAT)
            ws3.cell(row=r3, column=9, value=periodo_heredado[0])
            ws3.cell(row=r3, column=10, value=periodo_heredado[1])
            ws3.cell(row=r3, column=15, value=precio_heredado)
            r3 += 1

        gen = object.__new__(_CRI.GeneradorDesdeExistente)
        gen._parent = None
        gen._app_root = None
        gen._codigo_sap = "1003600"
        gen._fecha_inicio = fecha_inicio
        gen._fecha_fin = fecha_fin
        gen._centros_activos = [(c, c) for c in self.CENTROS]
        gen._audit = {
            "materiales_modificados": set(), "materiales_sin_cambios": set(),
            "n_absorbidos": 0, "n_gaps": 0, "n_reducciones": 0,
            "n_solapamientos": 0, "gaps_detalle": {}, "solap_detalle": {},
        }
        gen._fila_h1_por_codigo = {}

        with mock.patch.object(_CRI, "_log", lambda *a, **k: None):
            absorbidas = gen._actualizar_h3_heredados(
                wb, ws1, {self.MAT}, [(self.MAT, "05")], {self.MAT: precio_nuevo})
            filas = gen._generar_filas_absorbidas_con_tarifa(wb, absorbidas, {})
            gen._rellenar_h3_nuevas_filas(wb, filas, {}, {self.MAT: precio_nuevo})

        por_centro = {}
        for r in range(2, ws3.max_row + 1):
            if ws3.cell(row=r, column=6).value is None:
                continue
            centro = str(ws3.cell(row=r, column=5).value)
            por_centro.setdefault(centro, []).append((
                str(ws3.cell(row=r, column=9).value),
                str(ws3.cell(row=r, column=10).value),
                ws3.cell(row=r, column=15).value))
        for periodos in por_centro.values():
            periodos.sort(key=lambda p: (p[0][6:], p[0][3:5], p[0][:2]))
        return por_centro

    def test_solapamiento_parcial_crea_periodo_nuevo_con_precio_nuevo(self):
        por_centro = self._ejecutar_cadena(
            ("01.09.2025", "31.08.2026"), 19.36, 19.83,
            "01.02.2026", "31.01.2027")
        for centro in self.CENTROS:
            self.assertEqual(
                por_centro[centro],
                [("01.09.2025", "31.01.2026", 19.36),    # viejo, acortado
                 ("01.02.2026", "31.01.2027", 19.83)],   # nuevo, precio de tarifa
                f"centro {centro}: falta el período nuevo o está mal")

    def test_solapamiento_parcial_no_deja_huecos_ni_solapes(self):
        from datetime import datetime as _dt
        por_centro = self._ejecutar_cadena(
            ("01.09.2025", "31.08.2026"), 19.36, 19.83,
            "01.02.2026", "31.01.2027")
        for centro, periodos in por_centro.items():
            # Sin este aserto el test sería vacuo: con un solo período (el bug)
            # no hay pares que comparar y pasaría igualmente.
            self.assertGreaterEqual(
                len(periodos), 2,
                f"centro {centro}: debería haber período viejo + nuevo")
            for (i1, j1, _), (i2, j2, _) in zip(periodos, periodos[1:]):
                fin_anterior = _dt.strptime(j1, "%d.%m.%Y")
                inicio_siguiente = _dt.strptime(i2, "%d.%m.%Y")
                self.assertEqual((inicio_siguiente - fin_anterior).days, 1,
                                 f"centro {centro}: hueco o solape entre "
                                 f"{j1} y {i2}")

    def test_sin_solapamiento_sigue_actualizando_la_fila_existente(self):
        # Regresión del fix anterior (CALABACIN): sin solape no se crea fila
        # nueva, se reescribe la más reciente con el precio y fechas nuevos.
        por_centro = self._ejecutar_cadena(
            ("01.08.2026", "31.08.2026"), 0.90, 2.10,
            "01.09.2026", "30.09.2026")
        for centro in self.CENTROS:
            self.assertEqual(por_centro[centro],
                             [("01.09.2026", "30.09.2026", 2.10)])

    def test_mismo_precio_solo_extiende_fin_sin_crear_filas(self):
        por_centro = self._ejecutar_cadena(
            ("01.09.2025", "31.08.2026"), 19.36, 19.36,
            "01.02.2026", "31.01.2027")
        for centro in self.CENTROS:
            self.assertEqual(por_centro[centro],
                             [("01.09.2025", "31.01.2027", 19.36)])

    # ── Hueco por CENTRO (bug real FRAN Y CHEMI, 2026-09-28) ──────────────
    #  El Purchasing traía el material con Hoja 2 en 0110 y 0120 pero con
    #  Hoja 3 solo en 0110. La red de seguridad razonaba por material (tiene
    #  filas → no hay hueco) y el RINFO salía sin precio para 0120.

    def test_hueco_por_centro_genera_el_periodo_en_el_centro_sin_hoja3(self):
        por_centro = self._ejecutar_cadena(
            ("01.11.2025", "31.08.2026"), 2.65, 3.20,
            "01.09.2026", "30.09.2026", centros_con_h3=["0110"])
        self.assertIn("0120", por_centro,
                      "el centro sin Hoja 3 debe recibir su condición de precio")
        self.assertEqual(por_centro["0120"],
                         [("01.09.2026", "30.09.2026", 3.20)])
        # y el centro que ya tenía precio queda igual que sin el hueco
        self.assertEqual(por_centro["0110"],
                         [("01.09.2026", "30.09.2026", 3.20)])

    def test_hueco_por_centro_no_duplica_filas_donde_ya_habia_precio(self):
        por_centro = self._ejecutar_cadena(
            ("01.11.2025", "31.08.2026"), 2.65, 3.20,
            "01.09.2026", "30.09.2026", centros_con_h3=["0110"])
        self.assertEqual(len(por_centro["0110"]), 1)
        self.assertEqual(len(por_centro["0120"]), 1)


# ─────────────────────────────────────────────────────────────────────────────
#  TC22 — Layout del catálogo DALI completo (2026-09-28)
#
#  Se insertaron dos columnas ("Texto UM (expl.)" y "Precio €") entre "Unidad
#  SAP" y "Grupo productos SAP", desplazando 8 columnas. El catálogo se lee y
#  escribe SIEMPRE por nombre de cabecera para que cada valor siga en su
#  columna; estos tests fijan esa definición y el mapeo.
# ─────────────────────────────────────────────────────────────────────────────

class TC22_LayoutCatalogo(unittest.TestCase):

    CABECERAS_NUEVAS = [
        "Código DALI", "Código Proveedor", "Descripción - DALI", "Unidad - DALI",
        "IGIC - DALI", "Naturaleza - DALI", "Familia - DALI", "Subfamilia - DALI",
        "Proveedor Asignado - DALI", "Código SAP", "Descripción SAP", "Unidad SAP",
        "Texto UM (expl.)", "Precio €", "Grupo productos SAP",
        "Descripción grupo productos SAP", "Indicador impuestos SAP",
        "Categoría valoración", "Tipo producto", "Grupo compras SAP",
        "Nombre grupo compras SAP", "Activo",
    ]

    def test_definicion_canonica_coincide_con_el_excel_de_la_web(self):
        import sap_helpers as S
        self.assertEqual([h for _k, h, _a in S.CATALOGO_COLUMNAS],
                         self.CABECERAS_NUEVAS)

    def test_claves_unicas_y_las_de_origen_sap_existen(self):
        import sap_helpers as S
        claves = [k for k, _h, _a in S.CATALOGO_COLUMNAS]
        self.assertEqual(len(claves), len(set(claves)))
        for k in S.CATALOGO_CLAVES_SAP:
            self.assertIn(k, claves)
        self.assertNotIn("precio", S.CATALOGO_CLAVES_SAP)   # no viene del LISTADO

    def test_mapa_formato_nuevo(self):
        import sap_helpers as S
        m = S.mapa_columnas_catalogo(self.CABECERAS_NUEVAS)
        self.assertEqual(m["cod_sap"], 10)
        self.assertEqual(m["texto_um"], 13)
        self.assertEqual(m["precio"], 14)
        self.assertEqual(m["grp_prod"], 15)      # antes 13
        self.assertEqual(m["activo"], 22)        # antes 20

    def test_mapa_formato_viejo_de_20_columnas(self):
        import sap_helpers as S
        viejas = [h for h in self.CABECERAS_NUEVAS
                  if h not in ("Texto UM (expl.)", "Precio €")]
        m = S.mapa_columnas_catalogo(viejas)
        self.assertEqual(len(viejas), 20)
        self.assertNotIn("texto_um", m)
        self.assertNotIn("precio", m)
        self.assertEqual(m["grp_prod"], 13)
        self.assertEqual(m["activo"], 20)

    def test_mapa_sigue_a_una_columna_insertada_y_ignora_mayusculas(self):
        import sap_helpers as S
        cab = list(self.CABECERAS_NUEVAS)
        cab.insert(2, "Columna Futura")                       # desconocida: se ignora
        cab = [c.upper() if c == "Precio €" else c for c in cab]
        cab[cab.index("Texto UM (expl.)")] = "  texto  um (EXPL.) "
        m = S.mapa_columnas_catalogo(cab)
        self.assertEqual(m["clave"], 1)
        self.assertEqual(m["descripcion"], 4)                 # desplazada por la nueva
        self.assertEqual(m["texto_um"], 14)
        self.assertEqual(m["precio"], 15)

    def test_datos_sap_desde_fila_listado_texto_um_es_la_columna_F(self):
        import sap_helpers as S
        fila = (1, "AGUA", "UNIDAD", 10200458, " AGUA GARRAFA 5L VITALIA ",
                "Pieza", "PI", "020020030", "AGUAS", "3X", "3102", "BSEU",
                "200", "BEVERAGES", "SI")
        d = S.datos_sap_desde_fila_listado(fila)
        self.assertEqual(d["cod_sap"], "10200458")            # texto, sin decimales
        self.assertEqual(d["desc_sap"], "AGUA GARRAFA 5L VITALIA")   # sin espacios
        self.assertEqual(d["texto_um"], "Pieza")              # LISTADO col F
        self.assertEqual(d["unidad_sap"], "PI")               # LISTADO col G
        self.assertEqual(d["grp_prod"], "020020030")
        self.assertEqual(d["desc_grp_compras"], "BEVERAGES")
        self.assertEqual(set(d), set(S.CATALOGO_CLAVES_SAP))

    def test_datos_sap_desde_fila_corta_o_sin_sap(self):
        import sap_helpers as S
        d = S.datos_sap_desde_fila_listado((2, "AGUA VITTEL", "UNIDAD"))
        self.assertTrue(all(v is None for v in d.values()))
        d2 = S.datos_sap_desde_fila_listado(
            (3, "X", "U", None, None, "Caja", None))
        self.assertIsNone(d2["cod_sap"])
        self.assertEqual(d2["texto_um"], "Caja")

    def test_es_precio_valido(self):
        import sap_helpers as S
        for bueno in (12.5, 3, 0.01, 19.83, 1e6):
            self.assertTrue(S.es_precio_valido(bueno), bueno)
        for malo in (None, "", 0, 0.0, -1, -0.5, "abc", "1,50", True, False,
                     float("nan"), float("inf")):
            self.assertFalse(S.es_precio_valido(malo), repr(malo))


# ─────────────────────────────────────────────────────────────────────────────
#  TC23 — Estado heredado incoherente: periodos solapados con otro precio
#         (bug real FRAN Y CHEMI / ESCAROLA 10101544, 2026-10-06)
#
#  El Purchasing de ese día traía la escarola con 01.10→30.10 @1,50 y
#  31.10→31.10 @1,80 (restos de cargas repetidas del mismo RINFO) y la tarifa
#  decía 1,80 para todo octubre. La decisión "mismo precio" miraba solo el
#  período de mayor inicio (31.10, @1,80 = tarifa) y dejaba el 1,50 vigente 30
#  de los 31 días. Ejecuta el CÓDIGO REAL (6c → 9b → 10), como TC21.
# ─────────────────────────────────────────────────────────────────────────────

@unittest.skipIf(_CRI is None, "crear_registro_info no importable (¿sin tkinter?)")
class TC23_HeredadoIncoherenteSolapado(unittest.TestCase):

    H1 = "Info general p.reg.info"
    H2 = "Datos org.compr.p.reg.info"
    H3 = "Det.precios p.registros info"
    REG, MAT = "5300000627", "10101544"
    CENTROS = ["0810", "0910"]

    def _ejecutar(self, periodos_heredados, precio_nuevo, fecha_inicio, fecha_fin):
        """periodos_heredados: [(inicio, fin, precio), ...] aplicados a cada
        centro, en ESE orden físico. Devuelve {centro: [(i, f, precio), ...]}
        ordenado por inicio."""
        import unittest.mock as mock
        import openpyxl

        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = self.H1
        ws2 = wb.create_sheet(self.H2)
        ws3 = wb.create_sheet(self.H3)
        for ws, n in ((ws1, 44), (ws2, 41), (ws3, 27)):
            for c in range(1, n + 1):
                ws.cell(row=1, column=c, value=f"c{c}")

        ws1.cell(row=2, column=1, value=self.REG)
        ws1.cell(row=2, column=5, value=self.MAT)
        r3 = 2
        for i, centro in enumerate(self.CENTROS):
            ws2.cell(row=2 + i, column=1, value=self.REG)
            ws2.cell(row=2 + i, column=6, value=centro)
            ws2.cell(row=2 + i, column=10, value=self.MAT)
            for (ini, fin, precio) in periodos_heredados:
                ws3.cell(row=r3, column=1, value=self.REG)
                ws3.cell(row=r3, column=4, value="IC00")
                ws3.cell(row=r3, column=5, value=centro)
                ws3.cell(row=r3, column=6, value=self.MAT)
                ws3.cell(row=r3, column=9, value=ini)
                ws3.cell(row=r3, column=10, value=fin)
                ws3.cell(row=r3, column=15, value=precio)
                r3 += 1

        gen = object.__new__(_CRI.GeneradorDesdeExistente)
        gen._parent = None
        gen._app_root = None
        gen._codigo_sap = "1001411"
        gen._fecha_inicio = fecha_inicio
        gen._fecha_fin = fecha_fin
        gen._centros_activos = [(c, c) for c in self.CENTROS]
        gen._audit = {
            "materiales_modificados": set(), "materiales_sin_cambios": set(),
            "n_absorbidos": 0, "n_gaps": 0, "n_reducciones": 0,
            "n_solapamientos": 0, "gaps_detalle": {}, "solap_detalle": {},
        }
        gen._fila_h1_por_codigo = {}

        avisos = []
        with mock.patch.object(_CRI, "_log",
                               lambda p, m, n="INFO": avisos.append((n, m))):
            absorbidas = gen._actualizar_h3_heredados(
                wb, ws1, {self.MAT}, [(self.MAT, "05")], {self.MAT: precio_nuevo})
            filas = gen._generar_filas_absorbidas_con_tarifa(wb, absorbidas, {})
            gen._rellenar_h3_nuevas_filas(wb, filas, {}, {self.MAT: precio_nuevo})
        self._avisos = avisos

        por_centro = {}
        for r in range(2, ws3.max_row + 1):
            if ws3.cell(row=r, column=6).value is None:
                continue
            centro = str(ws3.cell(row=r, column=5).value)
            por_centro.setdefault(centro, []).append((
                str(ws3.cell(row=r, column=9).value),
                str(ws3.cell(row=r, column=10).value),
                ws3.cell(row=r, column=15).value))
        for periodos in por_centro.values():
            periodos.sort(key=lambda p: (p[0][6:], p[0][3:5], p[0][:2]))
        return por_centro

    def test_escarola_octubre_incoherente_queda_a_precio_de_tarifa(self):
        por_centro = self._ejecutar(
            [("01.08.2026", "31.08.2026", 1.4),
             ("01.09.2026", "30.09.2026", 1.5),
             ("31.10.2026", "31.10.2026", 1.8),    # resto de 1 día, @ tarifa
             ("01.10.2026", "30.10.2026", 1.5)],   # el que debía ser 1,80
            1.8, "01.10.2026", "31.10.2026")
        for centro in self.CENTROS:
            self.assertEqual(
                por_centro[centro],
                [("01.08.2026", "31.08.2026", 1.4),
                 ("01.09.2026", "30.09.2026", 1.5),   # septiembre intacto
                 ("01.10.2026", "31.10.2026", 1.8)],  # octubre completo @ tarifa
                f"centro {centro}")

    def test_estado_incoherente_genera_aviso(self):
        self._ejecutar(
            [("01.09.2026", "30.09.2026", 1.5),
             ("31.10.2026", "31.10.2026", 1.8),
             ("01.10.2026", "30.10.2026", 1.5)],
            1.8, "01.10.2026", "31.10.2026")
        avisos = [m for n, m in self._avisos if "SOLAPADOS" in m]
        self.assertEqual(len(avisos), 1)
        self.assertIn(self.MAT, avisos[0])

    def test_heredado_coherente_mismo_precio_no_se_toca(self):
        # Control: Oct ya correcto a 1,80 y la tarifa también -> sin cambios y
        # SIN aviso de incoherencia.
        por_centro = self._ejecutar(
            [("01.09.2026", "30.09.2026", 1.5),
             ("01.10.2026", "31.10.2026", 1.8)],
            1.8, "01.10.2026", "31.10.2026")
        for centro in self.CENTROS:
            self.assertEqual(por_centro[centro],
                             [("01.09.2026", "30.09.2026", 1.5),
                              ("01.10.2026", "31.10.2026", 1.8)])
        self.assertFalse([m for n, m in self._avisos if "SOLAPADOS" in m])

    def test_periodo_largo_mismo_precio_sigue_la_ruta_mismo_precio(self):
        # Control: un único período que solapa con el mismo precio NO es
        # incoherencia -> sigue extendiendo/reduciendo solo la fecha fin.
        por_centro = self._ejecutar(
            [("01.08.2026", "31.12.2026", 1.6)],
            1.6, "01.10.2026", "31.10.2026")
        for centro in self.CENTROS:
            self.assertEqual(por_centro[centro],
                             [("01.08.2026", "31.10.2026", 1.6)])
        self.assertFalse([m for n, m in self._avisos if "SOLAPADOS" in m])


# ─────────────────────────────────────────────────────────────────────────────
#  TC24 — Aviso explícito cuando un cambio de fecha/precio se pierde por un
#         conflicto con la unidad de medida (caso real FRAN Y CHEMI / COCOS
#         10101648, 2026-10-06)
#
#  Si una fila de Hoja 3 tiene a la vez Z=X (cambio de fecha/precio) y una UM
#  distinta de la del LISTADO, _sincronizar_um_hoja3 quita la Z para que la
#  carga no falle -- y SAP ignora entonces el cambio de fecha/precio. Antes solo
#  quedaba una línea de resumen fácil de pasar por alto; ahora hay un WARN
#  agregado que lo dice explícitamente. Ejecuta el CÓDIGO REAL.
# ─────────────────────────────────────────────────────────────────────────────

@unittest.skipIf(_CRI is None, "crear_registro_info no importable (¿sin tkinter?)")
class TC24_AvisoCambioPerdidoPorUM(unittest.TestCase):

    H3 = "Det.precios p.registros info"
    MAT = "10101648"

    def _ejecutar(self, filas):
        """filas: [(centro, inicio, fin, precio, R, S, Z), ...] -> (wb, avisos)"""
        import unittest.mock as mock
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = self.H3
        for c in range(1, 28):
            ws.cell(row=1, column=c, value=f"c{c}")
        for i, (centro, ini, fin, precio, r_um, s_um, z) in enumerate(filas):
            r = 2 + i
            ws.cell(row=r, column=1, value="5300000654")
            ws.cell(row=r, column=5, value=centro)
            ws.cell(row=r, column=6, value=self.MAT)
            ws.cell(row=r, column=9, value=ini)
            ws.cell(row=r, column=10, value=fin)
            ws.cell(row=r, column=15, value=precio)
            ws.cell(row=r, column=18, value=r_um)
            ws.cell(row=r, column=19, value=s_um)
            ws.cell(row=r, column=26, value=z)

        gen = object.__new__(_CRI.GeneradorDesdeExistente)
        gen._parent = None
        gen._audit = {"materiales_modificados": {self.MAT}}
        avisos = []
        with mock.patch.object(_CRI, "_log",
                               lambda p, m, n="INFO": avisos.append((n, m))):
            gen._sincronizar_um_hoja3(wb, {self.MAT: ["COCOS", "PI"]})
        return wb[self.H3], avisos

    def test_cambio_de_precio_mas_um_distinta_avisa_de_que_no_se_registra(self):
        ws, avisos = self._ejecutar([
            ("0810", "01.10.2026", "31.10.2026", 1.35, "KG", "KG", "X"),
            ("0910", "01.10.2026", "31.10.2026", 1.35, "KG", "KG", "X")])
        # comportamiento previo conservado: UM corregida y Z quitada
        self.assertEqual((ws.cell(row=2, column=18).value, ws.cell(row=2, column=19).value),
                         ("PI", "PI"))
        self.assertIsNone(ws.cell(row=2, column=26).value)
        agregados = [m for n, m in avisos if n == "WARN" and "SIN REGISTRAR" in m]
        self.assertEqual(len(agregados), 1)
        self.assertIn(self.MAT, agregados[0])
        self.assertIn("0810", agregados[0])
        self.assertIn("0910", agregados[0])
        self.assertIn("1.35", agregados[0])

    def test_solo_um_distinta_sin_z_no_genera_aviso_de_perdida(self):
        ws, avisos = self._ejecutar([
            ("0810", "01.09.2026", "30.09.2026", 1.2, "KG", "KG", None)])
        self.assertEqual(ws.cell(row=2, column=18).value, "PI")
        self.assertFalse([m for n, m in avisos if "SIN REGISTRAR" in m])

    def test_z_con_um_correcta_no_se_toca(self):
        ws, avisos = self._ejecutar([
            ("0810", "01.10.2026", "31.10.2026", 1.35, "PI", "PI", "X")])
        self.assertEqual(ws.cell(row=2, column=26).value, "X")
        self.assertFalse([m for n, m in avisos if "SIN REGISTRAR" in m])


# ─────────────────────────────────────────────────────────────────────────────
#  TC25 — Baja de un material: no dejar un período con fin anterior al inicio
#         (bug real FRAN Y CHEMI / MELOCOTON 10101624, 2026-10-06)
#
#  El Purchasing traía el melocotón con 01.08→30.09 y 01.10→30.10. Al darlo de
#  baja (por decisión del usuario) se "cerraba" la fecha fin del período más
#  reciente a 30.09 (día anterior a la nueva tarifa) SIN comprobar que ese
#  período empezaba antes: quedaba 01.10.2026 → 30.09.2026 (invertido). Las
#  pruebas con el método real usan fechas muy lejanas para no depender del día
#  en que se ejecuten los tests.
# ─────────────────────────────────────────────────────────────────────────────

@unittest.skipIf(_CRI is None, "crear_registro_info no importable (¿sin tkinter?)")
class TC25_CierreFinValidezEnBaja(unittest.TestCase):

    def test_helper_cierra_periodo_vigente_que_empieza_antes(self):
        from sap_helpers import debe_cerrar_fin_validez as f
        d = datetime.date
        self.assertTrue(f(d(2026, 8, 1), d(2026, 12, 31), d(2026, 10, 6), d(2026, 9, 30)))

    def test_helper_no_cierra_periodo_que_empieza_despues_del_cierre(self):
        from sap_helpers import debe_cerrar_fin_validez as f
        d = datetime.date
        # caso MELOCOTON: 01.10→30.10 con cierre 30.09 dejaría fin < inicio
        self.assertFalse(f(d(2026, 10, 1), d(2026, 10, 30), d(2026, 10, 6), d(2026, 9, 30)))

    def test_helper_limites(self):
        from sap_helpers import debe_cerrar_fin_validez as f
        d = datetime.date
        hoy, cierre = d(2026, 10, 6), d(2026, 9, 30)
        self.assertTrue(f(cierre, d(2026, 12, 1), hoy, cierre))        # empieza el mismo día del cierre
        self.assertFalse(f(d(2026, 8, 1), d(2026, 9, 1), hoy, cierre)) # ya vencido
        self.assertTrue(f(None, None, hoy, cierre))                    # sin fechas: comportamiento previo
        self.assertTrue(f(None, d(2026, 12, 1), hoy, cierre))

    # ── con el método real ─────────────────────────────────────────────────
    MAT, REG = "10101624", "5300006192"

    def _baja(self, periodos):
        """periodos: [(inicio, fin), ...] iguales en 0810 y 0910, en este orden
        físico. Da de baja el material (decisión simulada) y devuelve
        {centro: [(inicio, fin, Y), ...]}."""
        import unittest.mock as mock
        import openpyxl
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Info general p.reg.info"
        ws2 = wb.create_sheet("Datos org.compr.p.reg.info")
        ws3 = wb.create_sheet("Det.precios p.registros info")
        for ws, n in ((ws1, 44), (ws2, 41), (ws3, 27)):
            for c in range(1, n + 1):
                ws.cell(row=1, column=c, value=f"c{c}")
        ws1.cell(row=2, column=1, value=self.REG)
        ws1.cell(row=2, column=5, value=self.MAT)
        ws1.cell(row=2, column=6, value="MELOCOTON")
        r3 = 2
        for i, centro in enumerate(["0810", "0910"]):
            ws2.cell(row=2 + i, column=1, value=self.REG)
            ws2.cell(row=2 + i, column=6, value=centro)
            ws2.cell(row=2 + i, column=10, value=self.MAT)
            for (ini, fin) in periodos:
                ws3.cell(row=r3, column=1, value=self.REG)
                ws3.cell(row=r3, column=5, value=centro)
                ws3.cell(row=r3, column=6, value=self.MAT)
                ws3.cell(row=r3, column=9, value=ini)
                ws3.cell(row=r3, column=10, value=fin)
                ws3.cell(row=r3, column=15, value=1.6)
                r3 += 1

        class _P:
            def after(self, ms, fn): fn()
            def winfo_exists(self): return True

        gen = object.__new__(_CRI.GeneradorDesdeExistente)
        gen._parent = _P()
        gen._app_root = _P()
        gen._barra = lambda *a, **k: None
        gen._fecha_inicio = "01.10.2090"
        gen._audit = {"materiales_modificados": set(), "materiales_sin_cambios": set()}
        gen._fila_h1_por_codigo = {self.MAT: 2}   # fila de Hoja 1 (la rellena un paso previo)

        with mock.patch.object(_CRI, "_log", lambda *a, **k: None), \
             mock.patch.object(_CRI, "_ventana_sin_tarifa",
                               lambda parent, pend: [
                                   {"codigo": p["codigo"], "fila_h1": p["fila_h1"], "baja": True}
                                   for p in pend]), \
             mock.patch.object(_CRI, "_safe_after",
                               lambda w, fn, ms=0: fn()):
            gen._limpiar_marca_aj_heredados(wb, ws1, {self.MAT}, set())   # decide la baja (AJ=X)
            gen._propagar_borrado_h1_a_h2_h3(wb, ws1)                    # propaga a H2/H3 y cierra fechas

        por_centro = {}
        for r in range(2, ws3.max_row + 1):
            por_centro.setdefault(str(ws3.cell(row=r, column=5).value), []).append((
                str(ws3.cell(row=r, column=9).value),
                str(ws3.cell(row=r, column=10).value),
                ws3.cell(row=r, column=25).value))
        return por_centro

    def test_baja_no_invierte_un_periodo_que_empieza_despues_del_cierre(self):
        por_centro = self._baja([("01.08.2090", "30.09.2090"),
                                 ("01.10.2090", "30.10.2090")])
        for centro in ("0810", "0910"):
            self.assertEqual(por_centro[centro],
                             [("01.08.2090", "30.09.2090", "X"),
                              ("01.10.2090", "30.10.2090", "X")],   # fin intacto
                             f"centro {centro}")

    def test_baja_sigue_cerrando_el_periodo_vigente_que_empieza_antes(self):
        # Control: comportamiento previo conservado cuando NO hay inversión.
        por_centro = self._baja([("01.08.2090", "31.12.2090")])
        for centro in ("0810", "0910"):
            self.assertEqual(por_centro[centro],
                             [("01.08.2090", "30.09.2090", "X")],   # cerrado a día anterior
                             f"centro {centro}")


# ─────────────────────────────────────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    unittest.main(verbosity=2)
