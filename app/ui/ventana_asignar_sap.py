# -*- coding: utf-8 -*-
"""
ventana_asignar_sap.py — ventana de la acción «Asignar SAP pendientes» (v10.2.56).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Todos los códigos DALI activos sin SAP del catálogo, con buscador y filtros por proveedor y familia. Se elige una
fila, se escribe el código SAP abajo y con Intro se asigna y se pasa a la siguiente (en el orden que se quiera). Al
lado sale lo que dice Productos.xlsx de ese código, para confirmar que es el correcto; si no está, se avisa pero se
deja asignar (petición de Víctor). «Guardar en el catálogo» graba lo asignado; las filas guardadas desaparecen y se
puede seguir. La lógica está en app/comun/asignar_sap.py; el guardado lo hace la app (guardar_fn), en un hilo.
"""
import threading
import tkinter as tk
from tkinter import messagebox, ttk

from app.comun import asignar_sap as A
from app.comun.tarifa_constants import (
    C_ACCENT, C_ACCENT2, C_BORDER, C_DARK, C_HOVER, C_MID, C_MUTED, C_SUCCESS, C_TEXT_LIGHT, C_WARNING,
    FONT_BOLD, FONT_LABEL, FONT_SMALL, FONT_TITLE,
)
from app.comun.tarifa_core import cargar_icono_ventana
from app.ui.tarifa_dialogs import VentanaResolverSAPDuplicados

TODOS = "(todos)"
COLUMNAS = (("dali", "DALI", 70), ("descripcion", "Descripción DALI", 300), ("unidad", "Unidad", 60),
            ("proveedor", "Proveedor", 190), ("familia", "Familia", 150), ("sap", "Código SAP", 90),
            ("productos", "En Productos.xlsx", 260))


class VentanaAsignarSap(tk.Toplevel):
    def __init__(self, parent, pendientes, tabla_productos, guardar_fn, consola_fn=None):
        """
        pendientes      : lista de asignar_sap.pendientes()
        tabla_productos : sap_helpers.leer_productos_xlsx()[0], o None si no hay Productos.xlsx
        guardar_fn      : (nuevos {dali: sap}, resolver_fn) -> dict con escritos/limpiados/no_encontrados/
                          fuera_productos, o None si se canceló. Se llama en un hilo de fondo; resolver_fn
                          (conflictos) muestra VentanaResolverSAPDuplicados y devuelve el set a borrar o None.
        """
        super().__init__(parent)
        self.title("Asignar SAP pendientes — Códigos DALI activos sin SAP")
        self.configure(bg=C_MID)
        self.transient(parent)
        cargar_icono_ventana(self)
        self._todos = list(pendientes)
        self._por_dali = {p["dali"]: p for p in self._todos}
        self._productos = tabla_productos
        self._guardar_fn = guardar_fn
        self._log = consola_fn or (lambda m, n="INFO": None)
        self._asignados = {}          # {dali: sap} sin guardar
        self._guardando = False
        self.guardados = 0
        self._construir()
        self._refrescar()
        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(1180, sw - 40), min(760, sh - 80)
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(5, (sh - h) // 4)}")
        self.grab_set()

    # ── Interfaz ─────────────────────────────────────────────────────────────────────────────────────
    def _construir(self):
        hdr = tk.Frame(self, bg=C_DARK, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🧩  Asignar SAP pendientes", font=FONT_TITLE, fg=C_ACCENT, bg=C_DARK).pack(side="left",
                                                                                                     padx=14)
        self._lbl_cuenta = tk.Label(hdr, font=FONT_LABEL, fg=C_MUTED, bg=C_DARK)
        self._lbl_cuenta.pack(side="right", padx=14)

        aviso = ("Elige un artículo, escribe su código SAP y pulsa Intro: se asigna y pasa al siguiente. "
                 "Al terminar (o cuando quieras), «Guardar en el catálogo».")
        if not self._productos:
            aviso += ("\n⚠  No se ha encontrado RECURSOS/Productos.xlsx: no se puede comprobar si los códigos "
                      "existen en SAP.")
        tk.Label(self, text=aviso, font=FONT_LABEL, fg=C_TEXT_LIGHT, bg=C_MID, justify="left").pack(
            anchor="w", padx=14, pady=(8, 4))

        # Filtros
        flt = tk.Frame(self, bg=C_MID)
        flt.pack(fill="x", padx=14, pady=(2, 6))
        tk.Label(flt, text="Buscar", font=FONT_LABEL, fg=C_TEXT_LIGHT, bg=C_MID).pack(side="left")
        self._var_buscar = tk.StringVar()
        e = tk.Entry(flt, textvariable=self._var_buscar, font=FONT_LABEL, width=28)
        e.pack(side="left", padx=(6, 14), ipady=2)
        self._var_buscar.trace_add("write", lambda *_: self._refrescar())
        tk.Label(flt, text="Proveedor", font=FONT_LABEL, fg=C_TEXT_LIGHT, bg=C_MID).pack(side="left")
        self._var_prov = tk.StringVar(value=TODOS)
        self._cb_prov = ttk.Combobox(flt, textvariable=self._var_prov, state="readonly", width=28,
                                     values=[TODOS] + A.valores_filtro(self._todos, "proveedor"))
        self._cb_prov.pack(side="left", padx=(6, 14))
        self._cb_prov.bind("<<ComboboxSelected>>", lambda _e: self._refrescar())
        tk.Label(flt, text="Familia", font=FONT_LABEL, fg=C_TEXT_LIGHT, bg=C_MID).pack(side="left")
        self._var_fam = tk.StringVar(value=TODOS)
        self._cb_fam = ttk.Combobox(flt, textvariable=self._var_fam, state="readonly", width=24,
                                    values=[TODOS] + A.valores_filtro(self._todos, "familia"))
        self._cb_fam.pack(side="left", padx=(6, 14))
        self._cb_fam.bind("<<ComboboxSelected>>", lambda _e: self._refrescar())
        tk.Button(flt, text="Quitar filtros", font=FONT_SMALL, relief="flat", bg=C_BORDER, fg=C_TEXT_LIGHT,
                  activebackground=C_HOVER, command=self._quitar_filtros).pack(side="left")

        # Botones de abajo (antes que la tabla, para que siempre se vean)
        abajo = tk.Frame(self, bg=C_MID, pady=10)
        abajo.pack(fill="x", padx=14, side="bottom")
        self._btn_guardar = tk.Button(abajo, font=FONT_BOLD, fg=C_DARK, bg=C_ACCENT, activebackground=C_ACCENT2,
                                      relief="flat", cursor="hand2", command=self._guardar)
        self._btn_guardar.pack(side="left", ipadx=12, ipady=4)
        tk.Button(abajo, text="✘  Cerrar", font=FONT_BOLD, fg=C_TEXT_LIGHT, bg=C_BORDER, activebackground=C_HOVER,
                  relief="flat", cursor="hand2", command=self._cerrar).pack(side="left", padx=8, ipadx=10, ipady=4)
        tk.Label(abajo, text="Los códigos se escriben en RECURSOS/catalogo-dali-completo.xlsx y se envían a "
                             "Catálogo Asignaciones.", font=FONT_SMALL, fg=C_MUTED, bg=C_MID).pack(side="right")

        # Editor de la fila elegida
        ed = tk.Frame(self, bg=C_DARK, padx=12, pady=8)
        ed.pack(fill="x", padx=14, side="bottom")
        self._lbl_articulo = tk.Label(ed, text="Elige un artículo de la lista.", font=FONT_BOLD, fg=C_ACCENT2,
                                      bg=C_DARK, anchor="w")
        self._lbl_articulo.grid(row=0, column=0, columnspan=4, sticky="w")
        tk.Label(ed, text="Código SAP", font=FONT_LABEL, fg=C_TEXT_LIGHT, bg=C_DARK).grid(row=1, column=0, sticky="w",
                                                                                         pady=(6, 0))
        self._var_sap = tk.StringVar()
        self._ent_sap = tk.Entry(ed, textvariable=self._var_sap, font=("Segoe UI", 11), width=16)
        self._ent_sap.grid(row=1, column=1, sticky="w", padx=(8, 8), pady=(6, 0), ipady=2)
        self._var_sap.trace_add("write", lambda *_: self._comprobar_escrito())
        self._ent_sap.bind("<Return>", lambda _e: self._asignar())
        self._ent_sap.bind("<KP_Enter>", lambda _e: self._asignar())
        tk.Button(ed, text="Asignar (Intro)", font=FONT_BOLD, fg=C_DARK, bg=C_ACCENT, relief="flat",
                  command=self._asignar).grid(row=1, column=2, padx=(0, 6), pady=(6, 0))
        tk.Button(ed, text="Quitar", font=FONT_LABEL, relief="flat", bg=C_BORDER, fg=C_TEXT_LIGHT,
                  command=self._quitar).grid(row=1, column=3, pady=(6, 0))
        self._lbl_productos = tk.Label(ed, text="", font=FONT_LABEL, fg=C_MUTED, bg=C_DARK, anchor="w")
        self._lbl_productos.grid(row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))

        # Tabla
        caja = tk.Frame(self, bg=C_MID)
        caja.pack(fill="both", expand=True, padx=14)
        self._tabla = ttk.Treeview(caja, columns=[c for c, _t, _a in COLUMNAS], show="headings",
                                   selectmode="browse")
        for c, titulo, ancho in COLUMNAS:
            self._tabla.heading(c, text=titulo)
            self._tabla.column(c, width=ancho, anchor="w", stretch=c in ("descripcion", "productos"))
        self._tabla.tag_configure("asignado", foreground="#1B7A3A")
        self._tabla.tag_configure("aviso", foreground="#A15C00")
        barra = ttk.Scrollbar(caja, orient="vertical", command=self._tabla.yview)
        self._tabla.configure(yscrollcommand=barra.set)
        barra.pack(side="right", fill="y")
        self._tabla.pack(side="left", fill="both", expand=True)
        self._tabla.bind("<<TreeviewSelect>>", lambda _e: self._al_elegir())
        self._tabla.bind("<Double-1>", lambda _e: self._ent_sap.focus_set())

    # ── Datos en pantalla ────────────────────────────────────────────────────────────────────────────
    def _texto_productos(self, sap):
        if sap is None:
            return ""
        if not self._productos:
            return "(sin Productos.xlsx)"
        info = A.info_productos(sap, self._productos)
        if info is None:
            return "⚠ no está en Productos.xlsx"
        return f"✔ {info[0]}" + (f" — {info[1]}" if info[1] else "")

    def _visibles(self):
        prov = "" if self._var_prov.get() == TODOS else self._var_prov.get()
        fam = "" if self._var_fam.get() == TODOS else self._var_fam.get()
        return A.filtrar(self._todos, self._var_buscar.get(), prov, fam)

    def _refrescar(self, seleccionar=None):
        seleccionada = seleccionar if seleccionar is not None else self._dali_elegido()
        self._tabla.delete(*self._tabla.get_children())
        visibles = self._visibles()
        for p in visibles:
            sap = self._asignados.get(p["dali"])
            texto = self._texto_productos(sap)
            tag = ("aviso",) if texto.startswith("⚠") else ("asignado",) if sap is not None else ()
            self._tabla.insert("", "end", iid=str(p["dali"]), tags=tag,
                               values=(str(p["dali"]).zfill(6), p["descripcion"], p["unidad"], p["proveedor"],
                                       p["familia"], "" if sap is None else sap, texto))
        if seleccionada is not None and self._tabla.exists(str(seleccionada)):
            self._tabla.selection_set(str(seleccionada))
            self._tabla.see(str(seleccionada))
        self._actualizar_cuenta(len(visibles))

    def _actualizar_cuenta(self, n_visibles=None):
        n = len(self._todos)
        txt = f"{n} pendiente(s)"
        if n_visibles is not None and n_visibles != n:
            txt += f" · {n_visibles} con este filtro"
        txt += f" · {len(self._asignados)} asignado(s) sin guardar"
        self._lbl_cuenta.configure(text=txt)
        self._btn_guardar.configure(text=f"✔  Guardar en el catálogo ({len(self._asignados)})",
                                    state="normal" if self._asignados and not self._guardando else "disabled")

    def _quitar_filtros(self):
        self._var_buscar.set("")
        self._var_prov.set(TODOS)
        self._var_fam.set(TODOS)
        self._refrescar()

    def _dali_elegido(self):
        sel = self._tabla.selection() if hasattr(self, "_tabla") else ()
        return int(sel[0]) if sel else None

    def _al_elegir(self):
        dali = self._dali_elegido()
        if dali is None:
            return
        p = self._por_dali[dali]
        self._lbl_articulo.configure(text=f"DALI {str(dali).zfill(6)} · {p['descripcion']}  ({p['proveedor']})")
        sap = self._asignados.get(dali)
        self._var_sap.set("" if sap is None else str(sap))
        self._ent_sap.focus_set()
        self._ent_sap.icursor("end")

    def _comprobar_escrito(self):
        texto = self._var_sap.get().strip()
        if not texto:
            self._lbl_productos.configure(text="", fg=C_MUTED)
            return
        try:
            sap = A.validar_codigo(texto)
        except ValueError as e:
            self._lbl_productos.configure(text=str(e), fg=C_WARNING)
            return
        info = self._texto_productos(sap)
        color = C_WARNING if info.startswith("⚠") else C_SUCCESS if info.startswith("✔") else C_MUTED
        if info.startswith("⚠"):
            info += " — se puede asignar igualmente (luego, Acción 3 con los Excel actualizados)"
        self._lbl_productos.configure(text=info, fg=color)

    # ── Asignar / quitar ─────────────────────────────────────────────────────────────────────────────
    def _asignar(self):
        dali = self._dali_elegido()
        if dali is None:
            messagebox.showinfo("Asignar SAP", "Elige primero un artículo de la lista.", parent=self)
            return
        try:
            sap = A.validar_codigo(self._var_sap.get())
        except ValueError as e:
            messagebox.showwarning("Código SAP no válido", str(e), parent=self)
            return
        otro = next((d for d, s in self._asignados.items() if s == sap and d != dali), None)
        if otro is not None and not messagebox.askyesno(
                "Código repetido",
                f"El código SAP {sap} ya lo has asignado al DALI {str(otro).zfill(6)}.\n\n"
                "¿Asignarlo también a este? (al guardar tendrás que elegir cuál es el correcto)", parent=self):
            return
        self._asignados[dali] = sap
        siguiente = self._siguiente(dali)
        self._refrescar(seleccionar=siguiente if siguiente is not None else dali)
        self._al_elegir()

    def _siguiente(self, dali):
        hijos = list(self._tabla.get_children())
        if str(dali) not in hijos:
            return None
        i = hijos.index(str(dali))
        for iid in hijos[i + 1:]:
            if int(iid) not in self._asignados:
                return int(iid)
        return int(hijos[i + 1]) if i + 1 < len(hijos) else None

    def _quitar(self):
        dali = self._dali_elegido()
        if dali is not None and self._asignados.pop(dali, None) is not None:
            self._refrescar(seleccionar=dali)
            self._al_elegir()

    # ── Guardar ──────────────────────────────────────────────────────────────────────────────────────
    def _resolver_en_hilo_principal(self, conflictos):
        """Desde el hilo de guardado: deja la petición y espera; la ventana la atiende en el hilo principal
        (_vigilar_guardado), porque Tkinter no se puede tocar desde otro hilo."""
        peticion = {"conflictos": conflictos, "evento": threading.Event()}
        self._peticion = peticion
        peticion["evento"].wait()
        return peticion.get("valor")

    def _guardar(self, y_cerrar=False):
        if not self._asignados or self._guardando:
            return
        self._guardando = True
        self._actualizar_cuenta()
        nuevos = dict(self._asignados)
        self._peticion = None
        estado = {}

        def _tarea():
            try:
                estado["r"] = self._guardar_fn(nuevos, self._resolver_en_hilo_principal)
            except Exception as e:  # noqa: BLE001
                estado["error"] = e
            estado["fin"] = True
        threading.Thread(target=_tarea, daemon=True, name="AsignarSapGuardar").start()
        self.after(100, lambda: self._vigilar_guardado(nuevos, estado, y_cerrar))

    def _vigilar_guardado(self, nuevos, estado, y_cerrar):
        """Hilo principal: atiende la ventana de códigos repetidos y el final del guardado."""
        peticion, self._peticion = self._peticion, None
        if peticion is not None:
            dlg = VentanaResolverSAPDuplicados(self, peticion["conflictos"])
            self.wait_window(dlg)
            peticion["valor"] = dlg.resultado
            peticion["evento"].set()
        if estado.get("fin"):
            self._tras_guardar(nuevos, estado.get("r"), estado.get("error"), y_cerrar)
        else:
            self.after(100, lambda: self._vigilar_guardado(nuevos, estado, y_cerrar))

    def _tras_guardar(self, nuevos, r, error, y_cerrar):
        self._guardando = False
        if error is not None:
            if isinstance(error, PermissionError):
                messagebox.showerror("Archivo bloqueado", "No se puede escribir en RECURSOS/catalogo-dali-completo."
                                     "xlsx: está abierto en Excel.\n\nCiérralo y vuelve a pulsar «Guardar».",
                                     parent=self)
            else:
                messagebox.showerror("Error al guardar", str(error), parent=self)
            self._actualizar_cuenta()
            return
        if r is None:                          # canceló al resolver duplicados: no se ha guardado nada
            self._actualizar_cuenta()
            return
        # Todo lo enviado queda resuelto: escrito, descartado al resolver duplicados o no encontrado. Lo que se haya
        # asignado mientras se guardaba (o cambiado de código) sigue pendiente de guardar.
        for d, sap in nuevos.items():
            if self._asignados.get(d) == sap:
                self._asignados.pop(d)
        self._todos = [p for p in self._todos if p["dali"] not in r["escritos"]]
        self.guardados += len(r["escritos"])
        self._refrescar()
        msg = f"✔  {len(r['escritos'])} código(s) SAP guardados en el catálogo."
        if r["limpiados"]:
            msg += f"\n✔  Quitado el código a {len(r['limpiados'])} DALI que lo tenían repetido."
        if r["no_encontrados"]:
            msg += ("\n\n⚠  No están en el catálogo (¿se han dado de baja?): "
                    + ", ".join(str(d) for d in r["no_encontrados"]))
        if r["fuera_productos"]:
            msg += (f"\n\n⚠  {len(r['fuera_productos'])} código(s) no están en Productos.xlsx: "
                    + ", ".join(str(r["escritos"][d]) for d in r["fuera_productos"][:15])
                    + (" …" if len(r["fuera_productos"]) > 15 else "")
                    + "\nSe han guardado igualmente. Para completar sus datos SAP (descripción, unidad, grupos…):"
                      "\n  1. Descarga de SAP Productos.xlsx y Listas de productos.xlsx actualizados."
                      "\n  2. Guárdalos en la carpeta RECURSOS."
                      "\n  3. Pulsa de nuevo la Acción 3 (Actualizar listado materiales SAP).")
        if r.get("completando"):
            msg += "\n\nLos datos SAP de los que sí están en Productos.xlsx se están completando ahora (como en la Acción 3)."
        messagebox.showinfo("Asignar SAP pendientes", msg, parent=self)
        if y_cerrar and not self._asignados:
            self.destroy()

    def _cerrar(self):
        if self._guardando:
            return
        if self._asignados:
            r = messagebox.askyesnocancel(
                "Hay códigos sin guardar",
                f"Has asignado {len(self._asignados)} código(s) SAP que no se han guardado.\n\n"
                "¿Guardarlos antes de cerrar?", parent=self)
            if r is None:
                return
            if r:
                self._guardar(y_cerrar=True)
                return
        self.destroy()
