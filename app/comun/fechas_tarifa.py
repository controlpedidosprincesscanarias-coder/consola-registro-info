# -*- coding: utf-8 -*-
"""
fechas_tarifa.py — fechas de validez del Registro Info a partir del periodo de la tarifa (13 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Motivo (registro de Víctor del 13 nov 2026): la ventana «Fechas de validez» venía siempre con 01.01 – 31.12 del año
en curso y se aceptó así para una tarifa del «4º TRIMESTRE 2026». El RINFO salió con efecto desde enero y recortó 530
vigencias que llegaban a marzo de 2027. Ahora:
  1. las fechas propuestas salen del periodo de la tarifa (celda C3, «Periodo:»); si no se entiende, las últimas
     que se usaron en este PC; y si no hay, el año en curso, como antes;
  2. al confirmar, se avisa si el inicio es anterior a hoy o si las fechas no coinciden con el periodo de la tarifa.
Sin tkinter (la ventana es VentanaFechasRINFO en app/rinfo/ventanas_rinfo.py). Pruebas en tests/test_fechas_tarifa.py.
"""
import calendar
import json
import os
import re
import unicodedata
from datetime import date

_MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
          "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
          "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7, "ago": 8, "sep": 9, "sept": 9,
          "oct": 10, "nov": 11, "dic": 12}
_ROMANOS = {"i": 1, "ii": 2, "iii": 3, "iv": 4}


def _normal(texto):
    t = unicodedata.normalize("NFKD", str(texto or ""))
    return "".join(c for c in t if not unicodedata.combining(c)).lower()


def _fin_de_mes(anio, mes):
    return date(anio, mes, calendar.monthrange(anio, mes)[1])


def _fecha(txt):
    m = re.match(r"(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$", txt)
    if not m:
        return None
    d, mes, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
    a = a + 2000 if a < 100 else a
    try:
        return date(a, mes, d)
    except ValueError:
        return None


def periodo_a_fechas(texto):
    """«4º TRIMESTRE 2026» → (date(2026,10,1), date(2026,12,31)). Entiende trimestres («4º TRIMESTRE 2026»,
    «T4 2026», «4T 2026», «IV TRIMESTRE 2026»), semestres, meses o rangos de meses («OCTUBRE 2026»,
    «OCT-DIC 2026», «OCTUBRE 2026 - MARZO 2027»), dos fechas («01/10/2026 - 31/03/2027») y el año suelto («2026»,
    «AÑO 2026»). Si no lo entiende, None."""
    t = _normal(texto).strip()
    if not t:
        return None
    # Dos fechas
    fechas = [f for f in (_fecha(x) for x in re.findall(r"\d{1,2}[./-]\d{1,2}[./-]\d{2,4}", t)) if f]
    if len(fechas) >= 2 and fechas[0] <= fechas[1]:
        return fechas[0], fechas[1]
    anios = [int(a) for a in re.findall(r"(?<!\d)(20\d{2})(?!\d)", t)]
    if not anios:
        return None
    anio = anios[0]
    # Trimestre
    m = (re.search(r"(?<!\d)([1-4])\s*(?:º|o|er|ero|do|to|°|ª)?\.?\s*(?:trimestre|trim\b|t\b)", t)
         or re.search(r"\b(?:trimestre|trim|t)\s*([1-4])\b", t))
    num = int(m.group(1)) if m else None
    if num is None:
        m = re.search(r"\b(iv|iii|ii|i)\s*(?:trimestre|trim)\b", t)
        num = _ROMANOS[m.group(1)] if m else None
    if num:
        return date(anio, 3 * num - 2, 1), _fin_de_mes(anio, 3 * num)
    # Semestre
    m = (re.search(r"(?<!\d)([12])\s*(?:º|o|er|ero|do|°|ª)?\.?\s*(?:semestre|sem\b|s\b)", t)
         or re.search(r"\b(?:semestre|sem|s)\s*([12])\b", t))
    if m:
        s = int(m.group(1))
        return date(anio, 6 * s - 5, 1), _fin_de_mes(anio, 6 * s)
    # Meses
    meses = [_MESES[p] for p in re.findall(r"[a-z]+", t) if p in _MESES]
    if meses:
        a_fin = anios[-1] if len(anios) > 1 else anio
        m_ini, m_fin = meses[0], meses[-1]
        if a_fin == anio and m_fin < m_ini:          # «NOV-FEB 2026» → cruza de año
            a_fin = anio + 1
        return date(anio, m_ini, 1), _fin_de_mes(a_fin, m_fin)
    # Año suelto («2026», «AÑO 2026», «TARIFA 2026»)
    if len(anios) == 1:
        return date(anio, 1, 1), date(anio, 12, 31)
    return None


def leer_periodo_tarifa(ruta_tarifa):
    """El texto de «Periodo:» de la tarifa final (celda C3, la escribe la Acción 4), o "" si no se puede leer."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ruta_tarifa, read_only=True, data_only=True)
        try:
            valor = wb.active.cell(3, 3).value
        finally:
            wb.close()
        return "" if valor is None else str(valor).strip()
    except Exception:  # noqa: BLE001 — sin periodo se proponen otras fechas
        return ""


# ── Últimas fechas usadas en este PC ─────────────────────────────────────────────────────────────────
def ruta_ultimas():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "ConsolaRegistroInfo", "ultimas_fechas_rinfo.json")


def leer_ultimas(ruta=None):
    try:
        with open(ruta or ruta_ultimas(), encoding="utf-8") as f:
            d = json.load(f)
        return _fecha(d["inicio"]), _fecha(d["fin"])
    except Exception:  # noqa: BLE001
        return None


def guardar_ultimas(inicio, fin, ruta=None):
    ruta = ruta or ruta_ultimas()
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump({"inicio": inicio.strftime("%d.%m.%Y"), "fin": fin.strftime("%d.%m.%Y")}, f)
    except OSError:
        pass


def propuesta(periodo, hoy=None, ultimas=None):
    """(inicio, fin, origen): del periodo de la tarifa; si no, las últimas usadas; si no, el año en curso.
    origen: "periodo" | "ultimas" | "anio"."""
    hoy = hoy or date.today()
    p = periodo_a_fechas(periodo)
    if p:
        return p[0], p[1], "periodo"
    if ultimas and all(ultimas):
        return ultimas[0], ultimas[1], "ultimas"
    return date(hoy.year, 1, 1), date(hoy.year, 12, 31), "anio"


def avisos(inicio, fin, periodo, hoy=None):
    """Lo que conviene confirmar antes de seguir (lista de textos; vacía si nada)."""
    hoy = hoy or date.today()
    f = lambda d: d.strftime("%d.%m.%Y")   # noqa: E731
    lista = []
    p = periodo_a_fechas(periodo)
    coincide = bool(p) and (inicio, fin) == p
    if inicio < hoy and not coincide:          # si coincide con la tarifa, esa retroactividad es la suya
        lista.append(f"El inicio ({f(inicio)}) es anterior a hoy ({f(hoy)}): los precios nuevos valdrán con efecto "
                     "desde esa fecha y se acortarán los periodos que ya estén vigentes en SAP.")
    if p and not coincide:
        lista.append(f"La tarifa dice «{periodo}» ({f(p[0])} – {f(p[1])}) y las fechas son {f(inicio)} – {f(fin)}.")
    return lista
