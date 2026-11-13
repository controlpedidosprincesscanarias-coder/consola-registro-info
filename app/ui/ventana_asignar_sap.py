# -*- coding: utf-8 -*-
"""
ventana_asignar_sap.py — ventana de la acción «Asignar SAP pendientes» (v10.2.50).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Todos los códigos DALI activos sin SAP del catálogo, con buscador y filtros por proveedor y familia.
(13 nov 2026, petición de Víctor: «me gusta más este sistema») La tabla es como la de «Códigos SAP faltantes» de la
Acción 4: una fila por artículo, con el código DALI, la descripción y el formato seleccionables (para copiarlos y
pegarlos en SAP al crear el material) y la columna «CÓDIGO SAP (introducir)», donde se escribe directamente; Intro
pasa a la fila siguiente. Al lado de cada código sale lo que dice Productos.xlsx (si no está, se avisa pero se deja
guardar). Lo escrito se conserva al cambiar el buscador o los filtros. «Guardar en el catálogo» graba lo escrito;
las filas guardadas desaparecen y se puede seguir. La lógica está en app/comun/asignar_sap.py; el guardado lo hace
la app (guardar_fn), en un hilo.
"""
import threading
import tkinter as tk
from tkinter import messagebox, ttk

from app.comun import asignar_sap as A
from app.comun.tarifa_constants import (
    C_ACCENT, C_ACCENT2, C_BORDER, C_CARD, C_DARK, C_HOVER, C_MID, C_MUTED, C_PANEL, C_SUCCESS, C_TEXT,
    C_TEXT_LIGHT, FONT_BOLD, FONT_LABEL, FONT_SMALL, FONT_TITLE,
)
from app.comun.tarifa_core import cargar_icono_ventana
from app.ui.tarifa_dialogs import VentanaResolverSAPDuplicados

TODOS = "(todos)"
MAX_FILAS = 250            # filas a la vez en pantalla (con miles, Tk se volvería lento): para el resto, filtrar
# (clave, título, ancho en caracteres, se estira)
COLUMNAS = (("dali", "CÓDIGO DALI", 10, False), ("descripcion", "DESCRIPCIÓN", 40, True),
            ("unidad", "FORMATO", 12, False), ("proveedor", "PROVEEDOR", 24, False),
            ("familia", "FAMILIA", 18, False), ("sap", "CÓDIGO SAP  (introducir)", 16, False),
            ("productos", "EN PRODUCTOS.XLSX", 34, True))
C_FILA_PAR, C_FILA_IMPAR = C_PANEL, "#EFF3F8"
C_OK, C_AVISO, C_MAL = "#1B7A3A", "#A15C00", "#B03A2E"


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
        self._vars = {}               # {dali: StringVar} de la columna CÓDIGO SAP (se conservan al filtrar)
        self._asignados = {}          # {dali: sap} escritos y válidos, sin guardar
        self._celdas_sap = {}         # {dali: Entry} de las filas en pantalla
        self._estados = {}            # {dali: Label «EN PRODUCTOS.XLSX»} de las filas en pantalla
        self._orden = []              # dali de las filas en pantalla, en orden
        self._dali_menu = None
        self._guardando = False
        self.guardados = 0
        self._construir()
        self._refrescar()
        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(1280, sw - 40), min(780, sh - 80)
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

        aviso = ("Escribe el código SAP de cada artículo en la columna «CÓDIGO SAP (introducir)»; Intro pasa al "
                 "siguiente. Deja en blanco los que no quieras asignar ahora.\nEl código DALI, la descripción y el "
                 "formato se pueden seleccionar y copiar (clic derecho → Copiar) para pegarlos en SAP al crear el "
                 "material.")
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
        self._lbl_info = tk.Label(abajo, text="", font=FONT_SMALL, fg=C_ACCENT2, bg=C_MID)
        self._lbl_info.pack(side="left", padx=8)
        tk.Label(abajo, text="Los códigos se escriben en RECURSOS/catalogo-dali-completo.xlsx\ny se envían a "
                             "Catálogo Asignaciones.", font=FONT_SMALL, fg=C_MUTED, bg=C_MID,
                 justify="left").pack(side="right")

        # Tabla con scroll (como la de «Códigos SAP faltantes» de la Acción 4)
        caja = tk.Frame(self, bg=C_MID)
        caja.pack(fill="both", expand=True, padx=10, pady=4)
        self._canvas = tk.Canvas(caja, bg=C_CARD, highlightthickness=0)
        barra = ttk.Scrollbar(caja, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=barra.set)
        barra.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._rejilla = tk.Frame(self._canvas, bg=C_CARD)
        ventana = self._canvas.create_window((0, 0), window=self._rejilla, anchor="nw")
        self._rejilla.bind("<Configure>", lambda _e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>", lambda e: self._canvas.itemconfig(ventana, width=e.width))
        for col, (_c, _t, ancho, estira) in enumerate(COLUMNAS):
            self._rejilla.columnconfigure(col, weight=1 if estira else 0, minsize=ancho * 7)
        # Rueda del ratón sobre la tabla
        self._canvas.bind("<Enter>", lambda _e: self.bind_all("<MouseWheel>", self._rueda))
        self._canvas.bind("<Leave>", lambda _e: self.unbind_all("<MouseWheel>"))

        self._menu = tk.Menu(self, tearoff=0)
        for clave, titulo in (("dali", "Código DALI"), ("descripcion", "Descripción"), ("unidad", "Formato"),
                              ("proveedor", "Proveedor")):
            self._menu.add_command(label=f"Copiar {titulo.lower()}",
                                   command=lambda c=clave, t=titulo: self._copiar_campo(c, t))
        self._menu.add_separator()
        self._menu.add_command(label="Copiar la fila entera", command=self._copiar_fila)

    def _rueda(self, evento):
        self._canvas.yview_scroll(int(-evento.delta / 120) or (-1 if evento.delta > 0 else 1), "units")

    # ── Datos en pantalla ────────────────────────────────────────────────────────────────────────────
    def _texto_productos(self, sap):
        if sap is None:
            return ""
        if not self._productos:
            return "(sin Productos.xlsx)"
        info = A.info_productos(sap, self._productos)
        if info is None:
            return "⚠ no está en Productos.xlsx (se guarda igualmente)"
        return f"✔ {info[0]}" + (f" — {info[1]}" if info[1] else "")

    def _visibles(self):
        prov = "" if self._var_prov.get() == TODOS else self._var_prov.get()
        fam = "" if self._var_fam.get() == TODOS else self._var_fam.get()
        return A.filtrar(self._todos, self._var_buscar.get(), prov, fam)

    def _var_de(self, dali):
        var = self._vars.get(dali)
        if var is None:
            var = tk.StringVar()
            var.trace_add("write", lambda *_a, d=dali: self._al_escribir(d))
            self._vars[dali] = var
        return var

    def _celda_lectura(self, fila, col, texto, fondo, dali, centrar=False):
        """Celda de solo lectura, seleccionable y copiable (como en la ventana de faltantes de la Acción 4)."""
        e = tk.Entry(self._rejilla, font=FONT_LABEL, fg=C_TEXT, bg=fondo, relief="flat", bd=0, width=1,
                     readonlybackground=fondo, justify="center" if centrar else "left")
        e.insert(0, texto)
        e.configure(state="readonly")
        e.grid(row=fila, column=col, padx=1, pady=1, sticky="nsew", ipady=1)
        e.bind("<Button-3>", lambda ev, d=dali: self._abrir_menu(ev, d))
        e.bind("<Control-c>", lambda _e, w=e: self._copiar_celda(w))
        e.bind("<Control-C>", lambda _e, w=e: self._copiar_celda(w))
        return e

    def _refrescar(self):
        for w in self._rejilla.winfo_children():
            w.destroy()
        self._celdas_sap, self._estados, self._orden = {}, {}, []
        for col, (_c, titulo, _a, _e) in enumerate(COLUMNAS):
            tk.Label(self._rejilla, text=titulo, font=FONT_BOLD, fg=C_TEXT, bg="#D9D9D9", anchor="center").grid(
                row=0, column=col, padx=1, pady=1, sticky="nsew")
        visibles = self._visibles()
        for i, p in enumerate(visibles[:MAX_FILAS], start=1):
            dali = p["dali"]
            fondo = C_FILA_PAR if i % 2 == 0 else C_FILA_IMPAR
            self._celda_lectura(i, 0, str(dali).zfill(6), fondo, dali, centrar=True)
            self._celda_lectura(i, 1, p["descripcion"], fondo, dali)
            self._celda_lectura(i, 2, p["unidad"], fondo, dali, centrar=True)
            self._celda_lectura(i, 3, p["proveedor"], fondo, dali)
            self._celda_lectura(i, 4, p["familia"], fondo, dali)
            ent = tk.Entry(self._rejilla, textvariable=self._var_de(dali), font=FONT_LABEL, fg=C_TEXT, bg="#FFFDF2",
                           relief="solid", bd=1, width=1)
            ent.grid(row=i, column=5, padx=1, pady=1, sticky="nsew", ipady=1)
            for tecla, paso in (("<Return>", 1), ("<KP_Enter>", 1), ("<Down>", 1), ("<Up>", -1)):
                ent.bind(tecla, lambda _e, d=dali, s=paso: self._siguiente_fila(d, s))
            ent.bind("<Button-3>", lambda ev, d=dali: self._abrir_menu(ev, d))
            lbl = tk.Label(self._rejilla, text="", font=FONT_SMALL, bg=fondo, anchor="w")
            lbl.grid(row=i, column=6, padx=1, pady=1, sticky="nsew")
            self._celdas_sap[dali], self._estados[dali] = ent, lbl
            self._orden.append(dali)
            self._pintar_estado(dali)
        if len(visibles) > MAX_FILAS:
            tk.Label(self._rejilla, text=f"… y {len(visibles) - MAX_FILAS} más: usa el buscador o los filtros para "
                                          "verlos.", font=FONT_SMALL, fg=C_ACCENT2, bg=C_CARD, anchor="w").grid(
                row=MAX_FILAS + 1, column=0, columnspan=len(COLUMNAS), sticky="w", padx=6, pady=6)
        self._canvas.yview_moveto(0)
        self._actualizar_cuenta(len(visibles))

    def _filas_visibles(self):
        """Códigos DALI de las filas en pantalla, en orden (para las pruebas y las capturas del manual)."""
        return list(self._orden)

    def _actualizar_cuenta(self, n_visibles=None):
        n = len(self._todos)
        txt = f"{n} pendiente(s)"
        if n_visibles is not None and n_visibles != n:
            txt += f" · {n_visibles} con este filtro"
        txt += f" · {len(self._asignados)} escrito(s) sin guardar"
        self._lbl_cuenta.configure(text=txt)
        self._btn_guardar.configure(text=f"✔  Guardar en el catálogo ({len(self._asignados)})",
                                    state="normal" if self._asignados and not self._guardando else "disabled")

    def _quitar_filtros(self):
        self._var_buscar.set("")
        self._var_prov.set(TODOS)
        self._var_fam.set(TODOS)
        self._refrescar()

    # ── Escribir el código SAP ───────────────────────────────────────────────────────────────────────
    def _al_escribir(self, dali):
        texto = self._vars[dali].get().strip()
        anterior = self._asignados.pop(dali, None)
        if texto:
            try:
                self._asignados[dali] = A.validar_codigo(texto)
            except ValueError:
                pass
        # el aviso de «repetido» puede cambiar también en otras filas
        afectados = {d for d, s in self._asignados.items() if s in (anterior, self._asignados.get(dali))}
        for d in afectados | {dali}:
            self._pintar_estado(d)
        self._actualizar_cuenta()

    def _pintar_estado(self, dali):
        lbl = self._estados.get(dali)
        if lbl is None:
            return
        texto = self._vars[dali].get().strip() if dali in self._vars else ""
        if not texto:
            lbl.configure(text="", fg=C_TEXT)
            return
        if dali not in self._asignados:
            try:
                A.validar_codigo(texto)
                mensaje = "✘ código no válido"
            except ValueError as e:
                mensaje = f"✘ {e}"
            lbl.configure(text=mensaje, fg=C_MAL)
            return
        sap = self._asignados[dali]
        otro = next((d for d, s in self._asignados.items() if s == sap and d != dali), None)
        info = self._texto_productos(sap)
        if otro is not None:
            info = f"⚠ repetido con el DALI {str(otro).zfill(6)} (al guardar eliges cuál) · " + info
        lbl.configure(text=info, fg=C_AVISO if info.startswith("⚠") else C_OK if info.startswith("✔") else C_TEXT)

    def _siguiente_fila(self, dali, paso=1):
        if dali in self._orden:
            i = self._orden.index(dali) + paso
            if 0 <= i < len(self._orden):
                ent = self._celdas_sap[self._orden[i]]
                ent.focus_set()
                ent.icursor("end")
                self._ver_fila(ent)
        return "break"

    def _ver_fila(self, widget):
        """Desplaza la tabla para que se vea la fila de `widget`."""
        self._rejilla.update_idletasks()
        alto = max(1, self._rejilla.winfo_height())
        y, h = widget.winfo_y(), self._canvas.winfo_height()
        arriba, abajo = self._canvas.canvasy(0), self._canvas.canvasy(h)
        if y < arriba or y + widget.winfo_height() > abajo:
            self._canvas.yview_moveto(max(0.0, (y - h / 3) / alto))

    # ── Copiar (para pegarlo en SAP al crear el material) ────────────────────────────────────────────
    def _texto_campo(self, clave, dali=None):
        dali = self._dali_menu if dali is None else dali
        if dali is None or dali not in self._por_dali:
            return None
        if clave == "dali":
            return str(dali).zfill(6)
        return self._por_dali[dali].get(clave, "")

    def _al_portapapeles(self, texto, que):
        self.clipboard_clear()
        self.clipboard_append(texto)
        self._lbl_info.configure(text=f"✔ Copiado ({que}): {texto[:70]}", fg=C_SUCCESS)

    def _copiar_celda(self, w):
        """Ctrl+C en una celda: lo seleccionado o, sin selección, la celda entera."""
        try:
            texto = w.selection_get() if w.selection_present() else w.get()
        except tk.TclError:
            texto = w.get()
        self._al_portapapeles(texto, "celda")
        return "break"

    def _copiar_campo(self, clave, titulo, dali=None):
        texto = self._texto_campo(clave, dali)
        if texto is None:
            return "break"
        self._al_portapapeles(texto, titulo.lower())
        return "break"

    def _copiar_fila(self, dali=None):
        dali = self._dali_menu if dali is None else dali
        if dali is None or dali not in self._por_dali:
            return
        p = self._por_dali[dali]
        texto = "\t".join([str(dali).zfill(6), p["descripcion"], p["unidad"], p["proveedor"], p["familia"]])
        self._al_portapapeles(texto, "fila")

    def _abrir_menu(self, evento, dali):
        self._dali_menu = dali
        self._menu.tk_popup(evento.x_root, evento.y_root)
        return "break"

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
                self._vars.pop(d, None)        # guardado o descartado: su casilla vuelve a quedar en blanco
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
