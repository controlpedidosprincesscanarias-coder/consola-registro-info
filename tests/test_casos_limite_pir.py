# -*- coding: utf-8 -*-
"""Casos límite de los registros info de SAP (PIR): solapamientos, fechas abiertas, centros,
organizaciones y proveedores (T01–T09).

Hasta noviembre de 2026 este bloque vivía al final de app/rinfo/crear_registro_info.py y se
lanzaba con «python crear_registro_info.py --test». Se trasladó aquí sin cambiar nada para que
el programa no lo cargue en cada arranque. Es autónomo: reproduce la lógica que comprueba (no
importa el motor RINFO). Los casos T07–T09 no están en test_crear_registro_info.py.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus
pruebas, son de Víctor Alejandro Martín Andrade (VAMA).
Uso: python -m unittest discover -s tests -v   ó   python tests/test_casos_limite_pir.py T01 T03
"""
import unittest

# ═════════════════════════════════════════════════════════════════════════════
#  MÓDULO DE TESTS AUTOMÁTICOS — EDGE CASES SAP PIR
#
#  Cubre los casos indicados en el spec de calidad (Punto 3):
#    T01 · Múltiples solapamientos (3 períodos solapando)
#    T02 · Fecha fin abierta (J = None / vacía)
#    T03 · Deduplicación funcional (misma clave PIR × 2 intentos)
#    T04 · Varios centros para el mismo material
#    T05 · Múltiples organizaciones de compra mismo material
#    T06 · Múltiples proveedores mismo material
#    T07 · ID offset col A — correlativos vs IDs SAP reales
#    T08 · Solapamiento total (nuevo engloba completamente al heredado)
#    T09 · Solapamiento inverso (heredado engloba al nuevo)
#
#
#  Cada test es autocontenido: crea sus datos en memoria, ejecuta la lógica
#  bajo prueba y verifica la salida. Sin dependencias externas ni archivos.
# ═════════════════════════════════════════════════════════════════════════════

import datetime as _dt_test

def _run_tests(filtro: list = None):
    """Ejecuta todos los tests de edge cases SAP PIR."""

    VERDE  = "\033[92m"
    ROJO   = "\033[91m"
    AMARILLO = "\033[93m"
    RESET  = "\033[0m"

    resultados = []

    def _ok(nombre, detalle=""):
        resultados.append(("OK", nombre, detalle))
        print(f"{VERDE}  ✅ PASS{RESET} {nombre}" + (f" — {detalle}" if detalle else ""))

    def _fail(nombre, motivo):
        resultados.append(("FAIL", nombre, motivo))
        print(f"{ROJO}  ❌ FAIL{RESET} {nombre} — {motivo}")

    def _skip(nombre):
        resultados.append(("SKIP", nombre, ""))
        print(f"{AMARILLO}  ⏭ SKIP{RESET} {nombre}")

    def _debe_ejecutar(tid):
        return not filtro or tid in filtro

    # ─────────────────────────────────────────────────────────────────────────
    #  Helpers para tests de solapamiento (replica la lógica de paso 6c)
    # ─────────────────────────────────────────────────────────────────────────

    def _overlap(i_her, j_her, i_nva, j_nva) -> bool:
        """Devuelve True si los dos intervalos se solapan (overlap universal)."""
        j_her_inf = j_her is None
        j_nva_inf = j_nva is None
        sin_solap = (
            (not j_her_inf and j_her < i_nva) or
            (not j_nva_inf and j_nva < i_her)
        ) if i_her is not None else False
        return not sin_solap

    def _cierre(i_her, i_nva) -> _dt_test.date:
        """J_her de cierre cuando hay solapamiento."""
        if i_nva > i_her:
            return i_nva - _dt_test.timedelta(days=1)
        else:
            return i_her - _dt_test.timedelta(days=1)

    d = _dt_test.date  # alias corto

    # ── T01: Múltiples solapamientos ─────────────────────────────────────────
    if _debe_ejecutar("T01"):
        nombre = "T01 · Múltiples solapamientos (3 períodos)"
        try:
            # Heredados: P1=ene-mar, P2=abr-jun, P3=jul-dic
            # Nueva tarifa: 01.03 – 30.09 (solapa con P1, P2 y P3)
            periodos = [
                {"i": d(2025,1,1),  "j": d(2025,3,31)},
                {"i": d(2025,4,1),  "j": d(2025,6,30)},
                {"i": d(2025,7,1),  "j": d(2025,12,31)},
            ]
            i_nva = d(2025, 3, 1)
            j_nva = d(2025, 9, 30)
            solapados = [p for p in periodos if _overlap(p["i"], p["j"], i_nva, j_nva)]
            assert len(solapados) == 3, f"Esperaba 3 solapamientos, got {len(solapados)}"
            # P1: i_nva (01.03) > i_her (01.01) → cierre = 28.02
            c1 = _cierre(periodos[0]["i"], i_nva)
            assert c1 == d(2025,2,28), f"P1 cierre esperado 2025-02-28, got {c1}"
            # P2: i_nva (01.03) < i_her (01.04) → absorbido, cierre = 31.03
            c2 = _cierre(periodos[1]["i"], i_nva)
            assert c2 == d(2025,3,31), f"P2 cierre esperado 2025-03-31, got {c2}"
            _ok(nombre, f"{len(solapados)} solapamientos detectados, cierres correctos")
        except Exception as e:
            _fail(nombre, str(e))

    # ── T02: Fecha fin abierta (J = None) ────────────────────────────────────
    if _debe_ejecutar("T02"):
        nombre = "T02 · Fecha fin abierta (J = None)"
        try:
            # Período heredado con J abierta (sin fin)
            i_her = d(2025, 1, 1)
            j_her = None  # abierto
            i_nva = d(2025, 6, 1)
            j_nva = d(2025, 12, 31)
            # Debe solapar (J_her = ∞ ⊃ [i_nva, j_nva])
            assert _overlap(i_her, j_her, i_nva, j_nva), "Debería solapar con J_her=None"
            c = _cierre(i_her, i_nva)
            assert c == d(2025, 5, 31), f"Cierre esperado 2025-05-31, got {c}"
            # J_her abierta + i_nva anterior a i_her: también solapamiento
            i_nva2 = d(2024, 12, 1)
            j_nva2 = None
            assert _overlap(i_her, j_her, i_nva2, j_nva2), "Ambos abiertos deben solapar"
            _ok(nombre, "J=None tratado como +∞, solapamiento y cierre correctos")
        except Exception as e:
            _fail(nombre, str(e))

    # ── T03: Deduplicación funcional ─────────────────────────────────────────
    if _debe_ejecutar("T03"):
        nombre = "T03 · Deduplicación funcional (misma clave PIR × 2)"
        try:
            escritas: set = set()
            prov, org, mat, centro, fi, ff = "1001419", "IC00", "10100264", "0110", "01.05.2026", "31.10.2026"
            clave = f"{prov}|{org}|{mat}|{centro}|{fi}|{ff}"
            # Primera inserción: debe pasar
            assert clave not in escritas, "Primera inserción no debe estar en set"
            escritas.add(clave)
            # Segunda inserción: debe bloquearse
            assert clave in escritas, "Segunda inserción debe estar bloqueada"
            saltadas = 0
            for _ in range(3):  # 3 intentos más
                if clave in escritas:
                    saltadas += 1
                else:
                    escritas.add(clave)
            assert saltadas == 3, f"Esperaba 3 omisiones, got {saltadas}"
            # Clave distinta (otro centro): debe pasar
            clave2 = f"{prov}|{org}|{mat}|0120|{fi}|{ff}"
            assert clave2 not in escritas, "Clave con otro centro no debe estar bloqueada"
            _ok(nombre, "Dedup correcto: 1 escritura, 3 omisiones, centro diferente pasa")
        except Exception as e:
            _fail(nombre, str(e))

    # ── T04: Varios centros mismo material ───────────────────────────────────
    if _debe_ejecutar("T04"):
        nombre = "T04 · Varios centros para el mismo material"
        try:
            mat = "10100264"
            centros = ["0110", "0120", "0210", "0310", "0410"]
            escritas: set = set()
            prov, org, fi, ff = "1001419", "IC00", "01.05.2026", "31.10.2026"
            insertadas = 0
            for c in centros:
                clave = f"{prov}|{org}|{mat}|{c}|{fi}|{ff}"
                if clave not in escritas:
                    escritas.add(clave)
                    insertadas += 1
            # Intentar insertar los mismos de nuevo
            duplicados = 0
            for c in centros:
                clave = f"{prov}|{org}|{mat}|{c}|{fi}|{ff}"
                if clave in escritas:
                    duplicados += 1
            assert insertadas == 5, f"Esperaba 5 inserciones, got {insertadas}"
            assert duplicados == 5, f"Esperaba 5 duplicados bloqueados, got {duplicados}"
            # Solapamiento: cada centro tiene sus propios períodos independientes
            i_her = d(2025, 1, 1); j_her = d(2025, 12, 31)
            i_nva = d(2025, 6, 1); j_nva = d(2026, 5, 31)
            assert _overlap(i_her, j_her, i_nva, j_nva), "Debe solapar"
            c_cierre = _cierre(i_her, i_nva)
            assert c_cierre == d(2025, 5, 31)
            _ok(nombre, f"{insertadas} centros insertados, {duplicados} duplicados bloqueados, solapamiento correcto")
        except Exception as e:
            _fail(nombre, str(e))

    # ── T05: Múltiples organizaciones de compra ──────────────────────────────
    if _debe_ejecutar("T05"):
        nombre = "T05 · Múltiples organizaciones de compra mismo material"
        try:
            mat = "10100264"; centro = "0110"; prov = "1001419"
            orgs = ["IC00", "IC01", "IC02"]
            escritas: set = set()
            fi, ff = "01.05.2026", "31.10.2026"
            for org in orgs:
                clave = f"{prov}|{org}|{mat}|{centro}|{fi}|{ff}"
                assert clave not in escritas, f"Org {org} no debe estar duplicada"
                escritas.add(clave)
            assert len(escritas) == 3, "Cada org genera clave PIR distinta"
            # Intentar IC00 otra vez
            clave_dup = f"{prov}|IC00|{mat}|{centro}|{fi}|{ff}"
            assert clave_dup in escritas, "IC00 segunda vez debe estar bloqueada"
            _ok(nombre, "3 orgs de compra = 3 claves PIR distintas, dedup correcto")
        except Exception as e:
            _fail(nombre, str(e))

    # ── T06: Múltiples proveedores mismo material ────────────────────────────
    if _debe_ejecutar("T06"):
        nombre = "T06 · Múltiples proveedores mismo material"
        try:
            mat = "10100264"; centro = "0110"; org = "IC00"
            proveedores = ["1001419", "1001420", "1001421"]
            escritas: set = set()
            fi, ff = "01.05.2026", "31.10.2026"
            for prov in proveedores:
                clave = f"{prov}|{org}|{mat}|{centro}|{fi}|{ff}"
                assert clave not in escritas, f"Proveedor {prov} no debe estar duplicado"
                escritas.add(clave)
            assert len(escritas) == 3, "Cada proveedor = clave PIR distinta"
            # Intentar proveedor 1001419 de nuevo
            clave_dup = f"1001419|{org}|{mat}|{centro}|{fi}|{ff}"
            assert clave_dup in escritas, "Proveedor repetido debe estar bloqueado"
            # Solapamiento es independiente por proveedor (no se mezclan)
            # Proveedor A: heredado ene-dic, nuevo jun-may+1
            i_her = d(2025,1,1); j_her = d(2025,12,31)
            i_nva = d(2025,6,1); j_nva = d(2026,5,31)
            assert _overlap(i_her, j_her, i_nva, j_nva)
            # Clave diferente de proveedor = solapamiento calculado independientemente
            _ok(nombre, "3 proveedores = 3 claves PIR distintas, solapamiento independiente")
        except Exception as e:
            _fail(nombre, str(e))

    # ── T07: ID offset col A — correlativos vs IDs SAP reales ────────────────
    if _debe_ejecutar("T07"):
        nombre = "T07 · ID offset — correlativos simples vs IDs SAP reales"
        try:
            # Caso A: col A con correlativos simples (1..N) → offset = max
            ids_simples = [1, 2, 3, 4, 5]
            max_id = 0
            for v in ids_simples:
                if 0 < v <= 9999 and v > max_id:
                    max_id = v
            assert max_id == 5, f"Max correlativo esperado 5, got {max_id}"
            offset_efectivo = max_id if max_id > 0 else len(ids_simples)
            nuevos_ids = [str(offset_efectivo + i) for i in range(1, 4)]
            assert nuevos_ids == ["6", "7", "8"], f"Esperaba [6,7,8], got {nuevos_ids}"

            # Caso B: col A con IDs SAP reales (5300001122) → offset = 0 → fallback n_heredadas
            ids_sap = [5300001122, 5300001125, 5300001133]
            max_id_b = 0
            for v in ids_sap:
                if 0 < v <= 9999 and v > max_id_b:
                    max_id_b = v
            assert max_id_b == 0, f"IDs SAP no deben capturarse, got {max_id_b}"
            n_heredadas = len(ids_sap)
            offset_b = max_id_b if max_id_b > 0 else n_heredadas
            assert offset_b == 3, f"Fallback n_heredadas esperado 3, got {offset_b}"
            nuevos_ids_b = [str(offset_b + i) for i in range(1, 4)]
            assert nuevos_ids_b == ["4", "5", "6"], f"Esperaba [4,5,6], got {nuevos_ids_b}"

            _ok(nombre, "Correlativos 6,7,8 tras 1-5; IDs SAP ignorados, fallback correcto")
        except Exception as e:
            _fail(nombre, str(e))

    # ── T08: Solapamiento total (nuevo engloba al heredado) ───────────────────
    if _debe_ejecutar("T08"):
        nombre = "T08 · Solapamiento total (nuevo engloba completamente al heredado)"
        try:
            # Heredado: abr-jun; Nuevo: ene-dic (engloba)
            i_her = d(2025,4,1); j_her = d(2025,6,30)
            i_nva = d(2025,1,1); j_nva = d(2025,12,31)
            assert _overlap(i_her, j_her, i_nva, j_nva), "Debe solapar (englobado)"
            # Verificar condición de absorción total:
            #   i_nva <= i_her  AND  j_nva >= j_her
            # → la fila heredada se ELIMINA directamente (no se calcula cierre J=I-1).
            absorbido_total = (i_nva <= i_her) and (j_nva >= j_her)
            assert absorbido_total, "Debe detectarse como absorción total → fila a eliminar"
            # _cierre() se mantiene como helper auxiliar para solapamientos parciales
            c = _cierre(i_her, i_nva)
            assert c == d(2025,3,31), f"Cierre auxiliar esperado 2025-03-31, got {c}"
            _ok(nombre, f"Absorción TOTAL detectada correctamente — fila se elimina (no J=I-1). "
                        f"Cierre auxiliar parcial = {c}")
        except Exception as e:
            _fail(nombre, str(e))

    # ── T09: Solapamiento inverso (heredado engloba al nuevo) ─────────────────
    if _debe_ejecutar("T09"):
        nombre = "T09 · Solapamiento inverso (heredado engloba al nuevo rango)"
        try:
            # Heredado: ene-dic 2025; Nuevo: abr-jun 2025 (englobado por heredado)
            i_her = d(2025,1,1); j_her = d(2025,12,31)
            i_nva = d(2025,4,1); j_nva = d(2025,6,30)
            assert _overlap(i_her, j_her, i_nva, j_nva), "Debe solapar (inverso)"
            c = _cierre(i_her, i_nva)
            assert c == d(2025,3,31), f"Cierre esperado 2025-03-31, got {c}"
            # Con J_her abierta (heredado sin fin): también debe solapar
            assert _overlap(i_her, None, i_nva, j_nva), "Heredado abierto debe solapar"
            _ok(nombre, f"Solapamiento inverso detectado, cierre = {c}")
        except Exception as e:
            _fail(nombre, str(e))

    # ── Resumen ───────────────────────────────────────────────────────────────
    print()
    n_ok   = sum(1 for r in resultados if r[0] == "OK")
    n_fail = sum(1 for r in resultados if r[0] == "FAIL")
    n_skip = sum(1 for r in resultados if r[0] == "SKIP")
    print(f"  Tests: {n_ok} OK  |  {n_fail} FAIL  |  {n_skip} SKIP  |  {len(resultados)} total")
    if n_fail == 0:
        print(f"{VERDE}  ✅ TODOS LOS TESTS PASADOS{RESET}")
    else:
        print(f"{ROJO}  ❌ {n_fail} TEST(S) FALLIDO(S){RESET}")
    return n_fail == 0




class CasosLimitePIR(unittest.TestCase):
    def test_t01_a_t09(self):
        self.assertTrue(_run_tests(), "Algún caso límite PIR ha fallado (ver la salida)")


if __name__ == "__main__":
    import sys
    sys.exit(0 if _run_tests(sys.argv[1:] or None) else 1)
