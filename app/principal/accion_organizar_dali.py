# -*- coding: utf-8 -*-
"""
accion_organizar_dali.py
Acción 1 — Organizar artículos DALI (la interfaz; la lógica está en app/acciones/organizar_dali.py).

Mixin de App (tarifa_app.py): sus métodos se usan como self.metodo(). Separado de tarifa_app.py
en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).

13 nov 2026 (petición de Víctor): al terminar ya no sale solo el mensaje de resultado, sino la ventana
app/ui/ventana_volcado_dali.py con ese resumen, la pregunta de si se quiere revisar ARTICULOS_DALI.xlsx en Excel y,
tras contestar, el botón «Actualizar Catálogo Asignaciones», que pasa el archivo al catálogo completo y lo envía a la
web (_volcar_articulos_dali_al_catalogo, lo que antes era la Acción 2). Si la validación de diccionarios falla, el
archivo no es bueno: sale el mensaje de siempre y no se ofrece actualizar.
"""

import os
# (13 nov 2026) messagebox se importa aquí, una vez. Antes se importaba dentro de proceso() solo en los bloques
# except: eso lo convertía en variable local de proceso() y los mensajes lanzados antes con self.after (faltan
# ARTICULOS1/2 y no hay ARTICULOS_DALI previo, incidencias de duplicados y validación fallida) daban NameError en
# silencio y no salían nunca. Lo vigila tests/test_accion1_mensajes.py.
from tkinter import messagebox

from app.comun import volcado_dali


class AccionOrganizarDali:
    def _organizar_dali(self):
        """
        Acción 1 — Organizar artículos DALI (versión optimizada)
        Lee ARTICULOS1.xls y ARTICULOS2.xls desde RECURSOS/, cruza por CODIGO
        con pandas.merge en memoria y escribe el resultado de golpe con ws.append().
        Los estilos se aplican por rango, no celda a celda.
        """
        # (nov 2026) Se importa aquí y no al arrancar: carga pandas (~la mitad del arranque)
        # y solo lo usa esta acción.
        from app.acciones import organizar_dali as organizar_dali_datos
        import threading

        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta or not os.path.isdir(carpeta):
            messagebox.showwarning("Sin proyecto",
                                   "Seleccione primero la carpeta del Proyecto.")
            return

        recursos_path = os.path.join(carpeta, "RECURSOS")
        file1  = os.path.join(recursos_path, "ARTICULOS1.xls")
        file2  = os.path.join(recursos_path, "ARTICULOS2.xls")
        output = os.path.join(recursos_path, "ARTICULOS_DALI.xlsx")

        def proceso():
            try:
                self.after(0, lambda: self._barra.indeterminate("Leyendo archivos..."))
                self.consola.separador("INICIO ORGANIZACIÓN DALI")

                # ── Validaciones ──────────────────────────────────────────────
                if not os.path.isdir(recursos_path):
                    raise Exception(f"Carpeta RECURSOS no encontrada en:\n{carpeta}")

                faltan_xls = not os.path.exists(file1) or not os.path.exists(file2)

                if faltan_xls:
                    # ── Comprobar si existe ARTICULOS_DALI.xlsx como alternativa ──
                    if os.path.isfile(output):
                        try:
                            import openpyxl as _oxl
                            _wb = _oxl.load_workbook(output, read_only=True, data_only=True)
                            _ws = _wb.active
                            _b3 = str(_ws["B3"].value or "").strip()
                            _b5 = str(_ws["B5"].value or "").strip()
                            _wb.close()
                            _info_fecha = f"{_b3}  {_b5}".strip() or "(fecha no disponible)"
                        except Exception:
                            _info_fecha = "(no se pudo leer la fecha)"

                        _faltantes = []
                        if not os.path.exists(file1):
                            _faltantes.append("ARTICULOS1.xls")
                        if not os.path.exists(file2):
                            _faltantes.append("ARTICULOS2.xls")
                        _msg_falta = "\n".join(f"   \u2022 {f}" for f in _faltantes)

                        _respuesta = [None]

                        def _preguntar(r=_respuesta):
                            r[0] = messagebox.askyesno(
                                "Archivos ARTICULOS no encontrados",
                                f"\u274c  No se encontraron los siguientes archivos en RECURSOS/:\n\n"
                                f"{_msg_falta}\n\n"
                                f"Para obtenerlos, acceda a DALI \u2192 opci\u00f3n 331,\n"
                                f"descargue el listado con los filtros:\n\n"
                                f"   ARTICULOS1:  C\u00f3digo, Descripci\u00f3n, Unidad,\n"
                                f"               Naturaleza, Familia, Subfamilia\n\n"
                                f"   ARTICULOS2:  C\u00f3digo, Descripci\u00f3n, Unidad,\n"
                                f"               Impuesto Compra, Proveedor por Defecto,\n"
                                f"               No permitir compra\n\n"
                                f"\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\n"
                                f"Se ha encontrado un ARTICULOS_DALI.xlsx existente:\n\n"
                                f"   {_info_fecha}\n\n"
                                f"\u00bfDesea usar este archivo para actualizar el cat\u00e1logo\n"
                                f"y Cat\u00e1logo Asignaciones?",
                                icon="warning"
                            )

                        self.after(0, _preguntar)
                        # Esperar sin límite de tiempo a que el usuario responda el diálogo
                        import time as _time
                        while _respuesta[0] is None:
                            _time.sleep(0.1)

                        if _respuesta[0] is True:
                            # Usuario quiere actualizar → la misma ventana del final (revisar y actualizar)
                            self.log_consola("Se usa el ARTICULOS_DALI.xlsx existente.", "INFO")
                            self.consola.separador("FIN ORGANIZACIÓN DALI")
                            self.after(0, lambda: self._barra.reset(""))
                            _lineas = [("Se usará el ARTICULOS_DALI.xlsx que ya está en RECURSOS:", "info"),
                                       (f"   {_info_fecha}", "info"),
                                       ("", "info"),
                                       (f"📁  {output}", "nota")]
                            self.after(0, lambda li=_lineas: self._final_accion1(
                                "DALI — Actualizar Catálogo Asignaciones", li, output))
                            return
                        else:
                            # Usuario cierra → volver a la interfaz sin hacer nada
                            self.log_consola("Operaci\u00f3n cancelada por el usuario.", "WARN")
                            self.consola.separador("FIN ORGANIZACIÓN DALI")
                            self.after(0, lambda: self._barra.reset("Cancelado"))
                            return
                    else:
                        # No hay .xls ni ARTICULOS_DALI.xlsx → informar y explicar descarga
                        _faltantes = []
                        if not os.path.exists(file1):
                            _faltantes.append("ARTICULOS1.xls")
                        if not os.path.exists(file2):
                            _faltantes.append("ARTICULOS2.xls")
                        _msg_falta = "\n".join(f"   \u2022 {f}" for f in _faltantes)
                        self.after(0, lambda mf=_msg_falta: messagebox.showerror(
                            "Archivos ARTICULOS no encontrados",
                            f"\u274c  No se encontraron en RECURSOS/:\n\n"
                            f"{mf}\n\n"
                            f"Tambi\u00e9n se verific\u00f3 que no existe un\n"
                            f"ARTICULOS_DALI.xlsx previo en RECURSOS/.\n\n"
                            f"\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\n"
                            f"\ud83d\udce5  C\u00f3mo descargar desde DALI (opci\u00f3n 131):\n\n"
                            f"  ARTICULOS1  \u2192  seleccione en este orden:\n"
                            f"     1. C\u00f3digo\n"
                            f"     2. Descripci\u00f3n\n"
                            f"     3. Unidad\n"
                            f"     4. Familia\n"
                            f"     5. Subfamilia\n"
                            f"     6. Naturaleza\n\n"
                            f"  ARTICULOS2  \u2192  seleccione en este orden:\n"
                            f"     1. C\u00f3digo\n"
                            f"     2. Descripci\u00f3n\n"
                            f"     3. Unidad\n"
                            f"     4. Impuesto Compra\n"
                            f"     5. Proveedor por Defecto\n"
                            f"     6. No permitir compra\n\n"
                            f"Guarde los archivos descargados en la carpeta:\n"
                            f"   \ud83d\udcc1  RECURSOS/\n"
                            f"con los nombres exactos  ARTICULOS1.xls  y  ARTICULOS2.xls.\n\n"
                            f"A continuaci\u00f3n vuelva a ejecutar la Acci\u00f3n 1."
                        ))
                        self.log_consola("Proceso cancelado: faltan ARTICULOS1/2.xls y no hay fichero previo.", "ERROR")
                        self.consola.separador("FIN ORGANIZACIÓN DALI")
                        self.after(0, lambda: self._barra.reset("Faltan archivos"))
                        return

                # ── Núcleo de datos (Tanda 1 de la reorganización) ─────────────
                #    Lectura, dedup de duplicados en origen, cruce por CODIGO
                #    y saneamiento — extraído a acciones/organizar_dali.py.
                #    Ver ese módulo para el detalle línea a línea (idéntico
                #    al que había aquí, solo con self.log_consola -> log).
                _resultado = organizar_dali_datos.procesar_datos_dali(
                    file1, file2, self.log_consola)

                df_final          = _resultado["df_final"]
                df1               = _resultado["df1"]
                _ultimo_cod_art1  = _resultado["ultimo_cod_art1"]
                _incidencias_dup  = _resultado["incidencias_dup"]
                output_cols       = _resultado["output_cols"]
                errores           = _resultado["errores_sin_igic"]

                # ── Volcado del log de incidencias + aviso al usuario ─────────
                #    Solo si se han detectado duplicados con datos DISTINTOS
                #    entre repeticiones (errores reales de origen, no el caso
                #    inofensivo de duplicado con datos idénticos).
                _ruta_log_incidencias = organizar_dali_datos.guardar_log_incidencias(
                    _incidencias_dup, recursos_path, self.log_consola)

                if _ruta_log_incidencias:
                    self.after(0, lambda ruta=_ruta_log_incidencias, n=len(_incidencias_dup): messagebox.showwarning(
                        "Incidencias detectadas en datos de origen",
                        f"⚠  Se han detectado {n} código(s) duplicado(s) con datos "
                        f"DISTINTOS entre repeticiones en ARTICULOS1/ARTICULOS2.\n\n"
                        f"Se ha conservado la 1ª aparición de cada uno para generar "
                        f"ARTICULOS_DALI.xlsx, pero debería revisar y corregir el "
                        f"dato correcto en DALI.\n\n"
                        f"Detalle guardado en:\n{ruta}"
                    ))

                # ── Generar Excel + limpieza + validaciones (Tanda 2) ──────────
                #    Extraído a acciones/organizar_dali.py: construcción del
                #    libro (cabecera, logo, tabla, estilos), guardado,
                #    limpieza de ARTICULOS1/2 y validación de diccionarios
                #    SAP (SUBFAMILIA/IGIC). La DECISIÓN de qué hacer si la
                #    validación falla (borrar de nuevo + avisar) se queda
                #    aquí, no en el módulo, por ser un efecto secundario serio.
                organizar_dali_datos.generar_excel_dali(
                    df_final, output_cols, output,
                    self.log_consola,
                    progreso=lambda t: self.after(0, lambda: self._barra.indeterminate(t)))

                n_filas = len(df_final)
                self.log_consola(f"Archivo generado: {os.path.basename(output)}", "OK")
                self.log_consola(f"  └ {n_filas} artículos transferidos", "OK")

                _eliminados_ok = organizar_dali_datos.eliminar_archivos_origen(
                    file1, file2, self.log_consola, motivo="limpieza")

                # ═══════════════════════════════════════════════════════════════
                #  VALIDACIÓN DE COLUMNAS E (SUBFAMILIA) y G (IGIC)
                #  Se verifica el % de coincidencia con los diccionarios SAP.
                #  Si alguna columna tiene coincidencia 0 % se lanza error y se
                #  advierte al usuario que revise los filtros de la opción 131.
                # ═══════════════════════════════════════════════════════════════
                _valido = True     # False si falla la validación: entonces no se ofrece actualizar la web
                try:
                    _val = organizar_dali_datos.validar_diccionarios_sap(
                        df_final, self.log_consola)

                    if _val["columnas_sin_match"]:
                        _valido = False
                        _cols_str = "\n".join(f"   \u2022 {c}" for c in _val["columnas_sin_match"])

                        # Eliminar ARTICULOS1/2 (proceso incorrecto)
                        organizar_dali_datos.eliminar_archivos_origen(
                            file1, file2, self.log_consola, motivo="match fallido")

                        self.after(0, lambda cs=_cols_str: messagebox.showerror(
                            "Acción 1 — Proceso NO correcto: match con diccionarios fallido",
                            "✘  El proceso de la Acción 1 NO se ha completado correctamente.\n\n"
                            "Motivo: Las siguientes columnas del archivo generado\n"
                            "no contienen valores reconocidos por los diccionarios SAP:\n\n"
                            f"{cs}\n\n"
                            "Esto indica que ARTICULOS1.xls fue descargado con un\n"
                            "orden o selección de campos incorrectos desde DALI.\n\n"
                            "────────────────────────────────────────────────\n"
                            "📥  Vuelva a descargar desde DALI (opción 131)\n"
                            "    con los campos en este orden exacto:\n\n"
                            "  ARTICULOS1  →  seleccione en este orden:\n"
                            "     1. Código\n"
                            "     2. Descripción\n"
                            "     3. Unidad\n"
                            "     4. Familia\n"
                            "     5. Subfamilia\n"
                            "     6. Naturaleza\n\n"
                            "  ARTICULOS2  →  seleccione en este orden:\n"
                            "     1. Código\n"
                            "     2. Descripción\n"
                            "     3. Unidad\n"
                            "     4. Impuesto Compra\n"
                            "     5. Proveedor por Defecto\n"
                            "     6. No permitir compra\n\n"
                            "Guarde los archivos en RECURSOS/ con los nombres\n"
                            "exactos  ARTICULOS1.xls  y  ARTICULOS2.xls\n"
                            "y vuelva a ejecutar la Acción 1.\n\n"
                            "⚠   Los archivos ARTICULOS1 y ARTICULOS2 han sido\n"
                            "   eliminados de RECURSOS/ para forzar una nueva\n"
                            "   descarga correcta desde DALI."
                        ))
                except Exception as _e_val:
                    self.log_consola(
                        f"  ⚠  No se pudo completar la validación de columnas: {_e_val}", "WARN")

                # ── Detectar códigos faltantes entre 1 y el último de ARTICULOS1 ──
                saltos, faltantes_lista = organizar_dali_datos.detectar_codigos_faltantes(
                    df_final, _ultimo_cod_art1, self.log_consola)

                if errores:
                    self.log_consola(
                        f"  └ {errores} códigos sin coincidencia en ARTICULOS2", "WARN")

                self.log_consola(f"Ruta: {output}", "INFO")
                self.consola.separador("FIN ORGANIZACIÓN DALI")
                self.after(0, lambda: self._barra.completado(
                    f"DALI — {n_filas} artículos ✔"))

                # ── Construir mensaje para el usuario ─────────────────────────
                def _mostrar_resultado(n=n_filas, s=saltos, fl=faltantes_lista,
                                       e=errores, ruta=recursos_path,
                                       techo=_ultimo_cod_art1,
                                       elim=_eliminados_ok, valido=_valido):

                    # Construir texto del mensaje
                    lineas = [
                        "✔  Proceso completado correctamente",
                        "",
                        f"   Artículos transferidos:  {n}",
                    ]
                    if e:
                        lineas.append(f"   Sin coincidencia ARTICULOS2:  {e}")

                    if fl:
                        lineas.append("")
                        _techo_str = f"{techo:06d}" if techo else "??????"
                        lineas.append(f"✘  {len(fl)} código(s) faltante(s) entre 000001 y {_techo_str}:")
                        lineas.append("   Estos códigos de ARTICULOS1 no están en el resultado.")
                        for desde, hasta, num in s[:5]:
                            if desde == hasta:
                                lineas.append(f"      {desde:06d}  ({num} ausente)")
                            else:
                                lineas.append(f"      {desde:06d} – {hasta:06d}  ({num} ausente(s))")
                        if len(s) > 5:
                            lineas.append(f"      … y {len(s)-5} grupo(s) más")

                    lineas += [
                        "",
                        f"📁  Puede acceder al listado en:",
                        f"   {ruta}",
                    ]

                    if elim:
                        lineas += [
                            "",
                            "🗑  Archivos eliminados automáticamente de RECURSOS/:",
                        ] + [f"   • {f}" for f in elim]

                    tipo  = "error" if fl else "info"
                    titulo = "DALI — Proceso completado" + (" con errores" if fl else "")

                    if valido:
                        self._final_accion1(titulo, volcado_dali.lineas_con_tipo(lineas), output)
                        return

                    msg = "\n".join(lineas)
                    if tipo == "error":
                        messagebox.showerror(titulo, msg)
                    else:
                        messagebox.showinfo(titulo, msg)

                self.after(0, _mostrar_resultado)

            except PermissionError as e:
                self.log_consola(f"ERROR: {e}", "ERROR")
                self.consola.separador("FIN ORGANIZACIÓN DALI (con error)")
                self.after(0, lambda: self._barra.reset("Error en DALI"))
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error en DALI — archivo en uso",
                    "\u2718  No se pudo guardar el archivo.\n\n"
                    "El archivo de destino está abierto en Excel (por ti o por otra "
                    "persona, ya que está en una carpeta compartida).\n\n"
                    "Cierra el archivo y vuelve a ejecutar la acción.\n\n"
                    f"Detalle técnico:\n{err}"))

            except Exception as e:
                self.log_consola(f"ERROR: {e}", "ERROR")
                self.consola.separador("FIN ORGANIZACIÓN DALI (con error)")
                self.after(0, lambda: self._barra.reset("Error en DALI"))
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error en DALI",
                    "\u2718  El proceso no se completó correctamente.\n\n"
                    f"Detalle del error:\n{err}\n\n"
                    "Revise la ventana de log para más información."))

        threading.Thread(target=proceso, daemon=True).start()

    def _final_accion1(self, titulo, lineas, ruta_articulos):
        """(13 nov 2026) Ventana del final de la Acción 1 (en el hilo principal). Si se pulsa «Actualizar Catálogo
        Asignaciones», pasa ARTICULOS_DALI.xlsx al catálogo y lo envía a la web."""
        from app.ui import ventana_volcado_dali
        if ventana_volcado_dali.preguntar(self, titulo, lineas, ruta_articulos) == "actualizar":
            self.log_consola("Actualizando el catálogo y Catálogo Asignaciones con ARTICULOS_DALI.xlsx...", "INFO")
            self._volcar_articulos_dali_al_catalogo()
        else:
            self.log_consola("Catálogo Asignaciones sin actualizar: ARTICULOS_DALI.xlsx queda en RECURSOS "
                             "(la Acción 1 ofrecerá usarlo la próxima vez).", "WARN")

