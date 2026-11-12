# -*- coding: utf-8 -*-
"""
ventana_publicar_catalogo.py — resumen de lo que se va a cambiar en Catálogo Asignaciones (etapa 6, v10.2.50).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Se abre (siempre en el hilo principal) después de que la web haya «simulado» el envío: muestra cuántos
artículos cambian de cada grupo, los que se tocaron en la web (fuera por defecto, con una casilla para
incluirlos) y los irregulares que la web no aplicará (para corregirlos). Devuelve (aplicar, incluir_editados).
"""
import tkinter as tk

from app.comun import catalogo_web as W
from app.comun.tarifa_constants import (
    C_ACCENT, C_ACCENT2, C_CARD, C_DARK, C_LOG_BG, C_LOG_FG, C_MUTED, C_TEXT_LIGHT, C_WARNING,
    FONT_BOLD, FONT_LABEL, FONT_LOG, FONT_SMALL, FONT_SUB,
)

ETIQUETA_CAMPO = {
    "nombre_articulo": "Descripción", "unidad_dali": "Unidad", "igic": "IGIC", "naturaleza": "Naturaleza",
    "familia": "Familia", "subfamilia": "Subfamilia", "proveedor": "Proveedor asignado", "activo": "Activo",
    "codigo_sap": "Código SAP", "descripcion_sap": "Descripción SAP", "unidad_sap": "Unidad SAP",
    "texto_um_sap": "Texto UM", "grupo_productos_sap": "Grupo productos SAP",
    "descripcion_grupo_sap": "Descripción grupo SAP", "indicador_impuestos_sap": "Indicador impuestos SAP",
    "categoria_valoracion": "Categoría valoración", "tipo_producto": "Tipo producto",
    "grupo_compras_sap": "Grupo compras SAP", "nombre_grupo_compras_sap": "Nombre grupo compras SAP",
    "precio": "Precio €", "codigo_proveedor": "Código proveedor",
}
MAX_LINEAS = 300


def _valor(v):
    if v is None:
        return "(vacío)"
    if isinstance(v, bool):
        return "SI" if v else "NO"
    if isinstance(v, float):
        return f"{v:.2f}".replace(".", ",")
    return str(v)


def lineas_detalle(diferencias):
    """Texto del detalle: una línea por campo cambiado (sin tkinter, se prueba aparte)."""
    lineas = []
    for grupo in W.GRUPOS:
        for d in diferencias[grupo]:
            marca = " ✎ editado en la web" if d["editado_en_web"] else ""
            if d["nuevo"] and grupo == "accion2":
                lineas.append(f"DALI {d['codigo_dali']} · ALTA «{_valor(d['campos'].get('nombre_articulo', (0, None))[1])}»")
                continue
            for campo, (web, local) in d["campos"].items():
                lineas.append(f"DALI {d['codigo_dali']} · {ETIQUETA_CAMPO.get(campo, campo)}: "
                              f"{_valor(web)} → {_valor(local)}{marca}")
    return lineas


def irregulares_utiles(simulaciones, diferencias):
    """Irregulares de la simulación, sin los «no existe» de artículos que se dan de alta en este mismo envío."""
    nuevos = {d["codigo_dali"] for d in diferencias["accion2"] if d["nuevo"]}
    salida = []
    for grupo, sim in simulaciones.items():
        for irr in sim.get("irregulares") or []:
            if irr.get("tipo") == "no_existe" and irr.get("codigo_dali") in nuevos:
                continue
            salida.append((grupo, irr))
    return salida


def pedir_confirmacion(parent, diferencias, simulaciones, motivo):
    resultado = {"aplicar": False}
    editados = {d["codigo_dali"] for difs in diferencias.values() for d in difs if d["editado_en_web"]}
    irregulares = irregulares_utiles(simulaciones, diferencias)

    win = tk.Toplevel(parent)
    win.title("Actualizar Catálogo Asignaciones")
    win.configure(bg=C_DARK)
    win.transient(parent)
    win.grab_set()
    frm = tk.Frame(win, bg=C_DARK, padx=22, pady=18)
    frm.pack(fill="both", expand=True)

    tk.Label(frm, text="Cambios para Catálogo Asignaciones", bg=C_DARK, fg=C_ACCENT2,
             font=("Segoe UI", 12, "bold")).pack(anchor="w")
    tk.Label(frm, text=f"Después de: {motivo}. Solo se envía lo que es distinto en la web.",
             bg=C_DARK, fg=C_MUTED, font=FONT_SUB).pack(anchor="w", pady=(2, 10))

    var_editados = tk.BooleanVar(value=False)
    lbl_resumen = tk.Label(frm, bg=C_DARK, fg=C_TEXT_LIGHT, font=FONT_LABEL, justify="left")
    lbl_resumen.pack(anchor="w")

    def _pintar_resumen():
        n = W.contar(diferencias, var_editados.get())
        nuevos = sum(1 for d in diferencias["accion2"] if d["nuevo"])
        texto = "\n".join(f"• {W.NOMBRE_GRUPO[g]}: {n[g]} artículo(s)" for g in W.GRUPOS)
        if nuevos:
            texto += f"\n   (de ellos, {nuevos} artículo(s) nuevo(s) que se dan de alta)"
        lbl_resumen.config(text=texto)

    _pintar_resumen()

    if editados:
        tk.Label(frm, text=f"✎ {len(editados)} artículo(s) se cambiaron en la web (o los envió otro compañero) "
                           "desde tu último envío.\nPor defecto, sus datos SAP, su precio y su código de proveedor "
                           "NO se tocan; la descripción DALI sí,\nporque solo se carga desde la Consola.",
                 bg=C_DARK, fg=C_WARNING, font=FONT_SMALL, justify="left").pack(anchor="w", pady=(10, 2))
        tk.Checkbutton(frm, text="Incluirlos también (sustituir lo de la web por lo de la Consola)",
                       variable=var_editados, command=_pintar_resumen, bg=C_DARK, fg=C_TEXT_LIGHT,
                       selectcolor=C_CARD, activebackground=C_DARK, activeforeground=C_ACCENT2,
                       font=FONT_LABEL).pack(anchor="w")

    if irregulares:
        tk.Label(frm, text=f"⚠ {len(irregulares)} artículo(s) que la web NO cambiará (revísalos y corrígelos):",
                 bg=C_DARK, fg=C_WARNING, font=FONT_BOLD).pack(anchor="w", pady=(10, 2))

    tk.Label(frm, text="Detalle:", bg=C_DARK, fg=C_MUTED, font=FONT_SMALL).pack(anchor="w", pady=(10, 0))
    caja = tk.Frame(frm, bg=C_DARK)
    caja.pack(fill="both", expand=True)
    txt = tk.Text(caja, width=100, height=16, bg=C_LOG_BG, fg=C_LOG_FG, font=FONT_LOG, wrap="none",
                  relief="flat")
    barra = tk.Scrollbar(caja, command=txt.yview)
    txt.configure(yscrollcommand=barra.set)
    barra.pack(side="right", fill="y")
    txt.pack(side="left", fill="both", expand=True)
    lineas = [f"⚠ DALI {irr.get('codigo_dali')} · {irr.get('detalle', irr.get('tipo'))}" for _g, irr in irregulares]
    lineas += lineas_detalle(diferencias)
    if len(lineas) > MAX_LINEAS:
        lineas = lineas[:MAX_LINEAS] + [f"… y {len(lineas) - MAX_LINEAS} línea(s) más (el detalle completo queda en la consola)."]
    txt.insert("1.0", "\n".join(lineas))
    txt.configure(state="disabled")

    def _aplicar():
        resultado["aplicar"] = True
        win.destroy()

    botones = tk.Frame(frm, bg=C_DARK)
    botones.pack(fill="x", pady=(14, 0))
    tk.Button(botones, text="Ahora no", command=win.destroy, font=FONT_LABEL).pack(side="right", padx=(8, 0))
    tk.Button(botones, text="Actualizar la web", command=_aplicar, font=FONT_BOLD,
              bg=C_ACCENT, fg=C_DARK).pack(side="right")
    tk.Label(botones, text="Si eliges «Ahora no», los cambios volverán a salir la próxima vez.",
             bg=C_DARK, fg=C_MUTED, font=FONT_SMALL).pack(side="left")

    win.update_idletasks()
    w, h = win.winfo_reqwidth(), win.winfo_reqheight()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    win.geometry(f"{w}x{h}+{(sw - w) // 2}+{max(40, (sh - h) // 3)}")
    win.wait_window()
    return resultado["aplicar"], bool(var_editados.get())
