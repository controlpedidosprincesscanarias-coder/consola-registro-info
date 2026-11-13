# -*- coding: utf-8 -*-
"""
catalogo_dali.py
Catálogo DALI completo y Catálogo Asignaciones: credenciales, descarga, comprobación de
antigüedad, regeneración y código de proveedor.

Mixin de App (tarifa_app.py): sus métodos se usan como self.metodo(). Separado de tarifa_app.py
en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""

import os
import tkinter as tk
from tkinter import messagebox
import threading
import time
from app.comun.tarifa_constants import (
    C_DARK, C_ACCENT, C_ACCENT2, C_TEXT, C_TEXT_LIGHT, C_MUTED, FONT_SUB,
    FONT_LABEL, FONT_SMALL,
)
from app.comun.tarifa_core import (
    guardar_config,
    FOLDER_RECURSOS,
)
from app.comun.credenciales import (
    leer_credenciales_dali, guardar_credenciales_dali,
    borrar_credenciales_dali,
)
from app.comun.sap_helpers import (
    CATALOGO_COLUMNAS, CATALOGO_CLAVES_SAP,
    CATALOGO_FILA_CABECERA, CATALOGO_PRIMERA_FILA_DATOS,
    mapa_columnas_catalogo, datos_sap_desde_fila_listado, es_precio_valido,
)

# URL del worker de Cloudflare que hace de proxy inverso hacia el backend
# de la app web "Catálogo Asignaciones" (dali-sap-articulos-app). Usada
# solo por _descargar_catalogo_dali_web() para la descarga automática del
# catálogo completo (ver _confirmar_catalogo_actualizado()). Si esta URL
# cambia algún día (redespliegue del Worker con otro nombre), solo hay
# que actualizarla aquí.
# (Hasta nov 2026, en tarifa_app.py.)
# Desde v10.2.50 la URL está en app/comun/catalogo_web.py (la usa también el envío de cambios a la web).
from app.comun.catalogo_web import DALI_APP_API_BASE  # noqa: E402


class CatalogoDali:
    def _obtener_credenciales_dali_app(self):
        """
        Devuelve (email, password) para la app web «Catálogo Asignaciones».
        Si ya están guardadas en tarifa_config.json (email y contraseña
        CIFRADA con DPAPI de Windows, ver credenciales.py), las devuelve
        directamente sin preguntar nada. Si no, abre un diálogo modal para
        pedirlas una vez. (nov 2026: antes se guardaban en texto plano y el
        archivo llegó a subirse a GitHub; ahora nunca se escriben en claro.)

        Devuelve (None, None) si el usuario cancela el diálogo.
        """
        # (v10.2.52) Las de la sesión iniciada al abrir la Consola (sirven aunque no se hayan podido cifrar).
        if getattr(self, "_credenciales_sesion", None):
            return self._credenciales_sesion
        email_guardado, pass_guardado = leer_credenciales_dali(self.config_data)
        if email_guardado and pass_guardado:
            return email_guardado, pass_guardado

        resultado = {"email": None, "password": None}

        win = tk.Toplevel(self)
        win.title("Catálogo Asignaciones — credenciales")
        win.configure(bg=C_DARK)
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        frm = tk.Frame(win, bg=C_DARK, padx=24, pady=20)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text="Credenciales de «Catálogo Asignaciones»",
                 bg=C_DARK, fg=C_ACCENT2, font=FONT_SUB).pack(anchor="w")
        tk.Label(
            frm,
            text="Se usan para descargar el catálogo actualizado y para\n"
                 "enviar a la web los cambios de las Acciones 2, 3 y 4.\n"
                 "La contraseña se guarda\n"
                 "cifrada: solo tu usuario de Windows, en este equipo,\n"
                 "puede leerla. No se envía a ningún otro sitio.",
            bg=C_DARK, fg=C_MUTED, font=FONT_SMALL, justify="left",
        ).pack(anchor="w", pady=(4, 16))

        tk.Label(frm, text="Email", bg=C_DARK, fg=C_TEXT_LIGHT, font=FONT_LABEL).pack(anchor="w")
        entry_email = tk.Entry(frm, font=FONT_LABEL, width=34)
        entry_email.pack(fill="x", pady=(2, 10))

        tk.Label(frm, text="Contraseña", bg=C_DARK, fg=C_TEXT_LIGHT, font=FONT_LABEL).pack(anchor="w")
        entry_pass = tk.Entry(frm, font=FONT_LABEL, width=34, show="•")
        entry_pass.pack(fill="x", pady=(2, 16))

        def _guardar():
            email = entry_email.get().strip()
            pwd   = entry_pass.get().strip()
            if not email or not pwd:
                messagebox.showwarning("Datos incompletos",
                                       "Indique email y contraseña.", parent=win)
                return
            resultado["email"] = email
            resultado["password"] = pwd
            win.destroy()

        def _cancelar():
            win.destroy()

        frm_btns = tk.Frame(frm, bg=C_DARK)
        frm_btns.pack(fill="x")
        tk.Button(frm_btns, text="Cancelar", command=_cancelar,
                  font=FONT_LABEL).pack(side="right", padx=(8, 0))
        tk.Button(frm_btns, text="Guardar y continuar", command=_guardar,
                  font=FONT_LABEL, bg=C_ACCENT, fg=C_DARK).pack(side="right")

        entry_email.focus_set()
        win.bind("<Return>", lambda e: _guardar())

        # ── Centrar (mismo patrón que el resto de ventanas del proyecto) ──
        win.update_idletasks()
        _w = win.winfo_reqwidth()
        _h = win.winfo_reqheight()
        _sw = win.winfo_screenwidth()
        _sh = win.winfo_screenheight()
        win.geometry(f"{_w}x{_h}+{(_sw-_w)//2}+{max(60, (_sh-_h)//3)}")

        win.wait_window()

        if resultado["email"] and resultado["password"]:
            if not guardar_credenciales_dali(self.config_data, resultado["email"], resultado["password"]):
                self.log_consola(
                    "No se pudo cifrar la contraseña de Catálogo Asignaciones: se usa solo "
                    "en esta sesión y se volverá a pedir la próxima vez.", "WARN")
            guardar_config(self.config_data)
        return resultado["email"], resultado["password"]

    def _iniciar_sesion_consola(self):
        """
        (v10.2.52) Al abrir la Consola: inicio de sesión de Catálogo Asignaciones. Solo entra una cuenta activa,
        de administrador y con el permiso «Escribir desde la Consola Registro Info»
        (catalogo_web.comprobar_acceso); a las demás, «Usuario no autorizado.». Si hay credenciales guardadas,
        se prueban solas; si dejan de valer o la cuenta ya no está autorizada, se borran. «Salir» cierra la app.
        Desde el código (python tarifa_app.py) también se pide.
        """
        from app.comun import catalogo_web
        from app.ui import ventana_login

        email, password = leer_credenciales_dali(self.config_data)

        def _comprobar(e, p):
            return catalogo_web.comprobar_acceso(catalogo_web.ClienteCatalogo(), e, p)

        def _olvidar():
            borrar_credenciales_dali(self.config_data)
            guardar_config(self.config_data)

        try:
            r = ventana_login.pedir_inicio_sesion(self, email, password, _comprobar, al_fallar_credenciales=_olvidar)
        except tk.TclError:
            return                                # la Consola se cerró mientras tanto (p. ej., al actualizarse)
        if not r:
            try:
                self.destroy()
            except tk.TclError:
                pass
            return
        email, password, usuario = r
        self._credenciales_sesion = (email, password)
        if not guardar_credenciales_dali(self.config_data, email, password):
            self.log_consola("No se pudo cifrar la contraseña de Catálogo Asignaciones: se usa solo en esta "
                             "sesión y se volverá a pedir la próxima vez.", "WARN")
        guardar_config(self.config_data)
        self.deiconify()
        self.lift()
        self._tras_iniciar_sesion()
        self.log_consola(f"Sesión iniciada en Catálogo Asignaciones: {usuario.get('nombre') or email}.", "OK")

    def _adquirir_lock_catalogo_con_aviso(self):
        """
        Adquiere self._catalogo_lock antes de que cualquiera de los tres
        métodos que escriben catalogo-dali-completo.xlsx empiece a tocar
        el archivo (_descargar_catalogo_dali_web, _regenerar_catalogo_
        dali_completo, _actualizar_codigo_proveedor_catalogo). Si el
        lock ya está ocupado por otra de estas operaciones (p.ej. el
        usuario lanzó Acción 3 mientras el cierre de una transformación
        aún estaba escribiendo el código proveedor de otra tarifa), avisa
        en el log antes de quedarse esperando -- así la segunda operación
        se pone en cola en vez de chocar con la primera y perder datos en
        silencio. No hace falta liberar aquí: cada método libera el lock
        él mismo (try/finally) al terminar.
        """
        if not self._catalogo_lock.acquire(blocking=False):
            self.log_consola(
                "Esperando a que termine otra operación sobre el catálogo "
                "DALI completo antes de continuar...", "INFO")
            self._catalogo_lock.acquire(blocking=True)

    def _descargar_catalogo_dali_web(self, email, password):
        """
        Wrapper con bloqueo global (ver _adquirir_lock_catalogo_con_aviso):
        adquiere self._catalogo_lock antes de tocar catalogo-dali-completo.xlsx
        y lo libera siempre al terminar, delegando en _descargar_catalogo_
        dali_web_impl() para la lógica real (sin tocarla).
        """
        self._adquirir_lock_catalogo_con_aviso()
        try:
            return self._descargar_catalogo_dali_web_impl(email, password)
        finally:
            self._catalogo_lock.release()

    def _descargar_catalogo_dali_web_impl(self, email, password):
        """
        Descarga catalogo-dali-completo.xlsx directamente desde la app
        web «Catálogo Asignaciones» (POST /auth/login + GET
        /admin/export-completo, vía DALI_APP_API_BASE), sobrescribiendo
        RECURSOS/catalogo-dali-completo.xlsx. Autenticación por sesión
        (cookie), no por token -- se usa un CookieJar para mantenerla
        entre las dos peticiones.

        Recibe `email`/`password` ya resueltos (ver
        _obtener_credenciales_dali_app(), que crea una ventana Tkinter y
        por eso debe llamarse SIEMPRE desde el hilo principal, nunca
        desde aquí) -- este método hace solo peticiones de red y E/S de
        fichero, sin tocar Tkinter directamente salvo via self.after(),
        por lo que es seguro llamarlo desde un hilo de fondo (ver
        _descargar_catalogo_dali_web_bloqueante(), que es quien realmente
        lo invoca, en su propio hilo, para no congelar la interfaz
        mientras esperamos respuesta del servidor).

        Devuelve True si la descarga se completó correctamente, False
        en cualquier otro caso (credenciales incorrectas, verificación
        en dos pasos pendiente, sin conexión, permiso insuficiente...)
        -- siempre dejando un mensaje claro en el log y, si procede, una
        ventana de aviso, antes de devolver False.
        """
        import json as _json
        import urllib.request as _urlreq
        import urllib.error as _urlerr
        import http.cookiejar as _cookiejar

        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            return False

        self.log_consola("Descargando catálogo actualizado desde Catálogo Asignaciones...", "INFO")
        self.after(0, lambda: self._barra.indeterminate("Descargando catálogo actualizado..."))

        cookie_jar = _cookiejar.CookieJar()
        opener = _urlreq.build_opener(_urlreq.HTTPCookieProcessor(cookie_jar))
        # Cloudflare (WAF/Bot Fight Mode) puede rechazar con 403 el
        # User-Agent por defecto de urllib ("Python-urllib/3.x"),
        # identificable como tráfico de script -- se sustituye por uno
        # de navegador normal para que la petición pase igual que la
        # de cualquier persona usando la app desde Chrome/Edge/Firefox.
        opener.addheaders = [("User-Agent",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")]

        # ── 1. Login ────────────────────────────────────────────────
        try:
            cuerpo = _json.dumps({"email": email, "password": password}).encode("utf-8")
            peticion = _urlreq.Request(
                f"{DALI_APP_API_BASE}/auth/login", data=cuerpo, method="POST",
                headers={"Content-Type": "application/json"})
            with opener.open(peticion, timeout=20) as resp:
                datos_login = _json.loads(resp.read().decode("utf-8"))
        except _urlerr.HTTPError as e:
            try:
                detalle = _json.loads(e.read().decode("utf-8")).get("error", "")
            except Exception:
                detalle = ""
            if e.code in (401, 403):
                self.log_consola(
                    f"Login en Catálogo Asignaciones rechazado (HTTP {e.code}"
                    f"{': ' + detalle if detalle else ''}) — se han borrado las "
                    f"credenciales guardadas, vuelva a intentarlo.", "ERROR")
                borrar_credenciales_dali(self.config_data)
                guardar_config(self.config_data)
            else:
                self.log_consola(f"Error al iniciar sesión en Catálogo Asignaciones "
                                 f"(HTTP {e.code}): {detalle or e.reason}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False
        except Exception as e:
            self.log_consola(f"No se pudo conectar con Catálogo Asignaciones: {e}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False

        if datos_login.get("requiereVerificacion"):
            self.log_consola(
                "Catálogo Asignaciones pide un código de verificación (email "
                "enviado) — inicie sesión una vez desde el navegador para "
                "completarlo, y vuelva a intentar esta acción después.", "WARN")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            self.after(0, lambda: messagebox.showwarning(
                "Verificación pendiente",
                "Catálogo Asignaciones ha pedido un código de verificación "
                "por email (por inactividad prolongada de la cuenta).\n\n"
                "Inicie sesión una vez desde el navegador para completarlo, "
                "y vuelva a intentar esta acción después."))
            return False

        if not datos_login.get("ok"):
            self.log_consola(
                f"Login en Catálogo Asignaciones sin éxito: "
                f"{datos_login.get('error', '(sin detalle)')}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False

        # ── 2. Descargar el Excel completo ─────────────────────────────
        try:
            peticion = _urlreq.Request(f"{DALI_APP_API_BASE}/admin/export-completo")
            with opener.open(peticion, timeout=60) as resp:
                contenido = resp.read()
        except _urlerr.HTTPError as e:
            if e.code == 403:
                self.log_consola(
                    "Su cuenta no tiene permiso para exportar el catálogo completo "
                    "en Catálogo Asignaciones — pídaselo al administrador principal.",
                    "ERROR")
            else:
                self.log_consola(f"Error al descargar el catálogo "
                                 f"(HTTP {e.code}): {e.reason}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False
        except Exception as e:
            self.log_consola(f"No se pudo descargar el catálogo: {e}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False

        # ── 3. Guardar (escritura atómica: fichero temporal + reemplazo) ──
        # Fix (2026-09-25, error real: WinError 5 Acceso denegado en
        # os.replace): el archivo destino puede estar bloqueado un
        # instante por OneDrive sincronizando, un antivirus escaneándolo,
        # o Excel con el archivo abierto -- un bloqueo casi siempre
        # transitorio. Se reintenta unas pocas veces con una pequeña
        # espera antes de dar el guardado por fallido de verdad.
        try:
            recursos_path = os.path.join(carpeta, FOLDER_RECURSOS)
            os.makedirs(recursos_path, exist_ok=True)
            ruta_final = os.path.join(recursos_path, "catalogo-dali-completo.xlsx")
            ruta_tmp = ruta_final + ".tmp"
            with open(ruta_tmp, "wb") as f:
                f.write(contenido)

            _INTENTOS_REEMPLAZO = 5
            _ESPERA_ENTRE_INTENTOS = 1.0  # segundos
            for _intento in range(1, _INTENTOS_REEMPLAZO + 1):
                try:
                    os.replace(ruta_tmp, ruta_final)
                    break
                except PermissionError:
                    if _intento == _INTENTOS_REEMPLAZO:
                        raise
                    self.log_consola(
                        f"  ⚠ catalogo-dali-completo.xlsx está en uso "
                        f"(OneDrive, antivirus o Excel abierto) -- "
                        f"reintentando ({_intento}/{_INTENTOS_REEMPLAZO})...",
                        "WARN")
                    time.sleep(_ESPERA_ENTRE_INTENTOS)
        except Exception as e:
            self.log_consola(f"Catálogo descargado pero no se pudo guardar: {e}", "ERROR")
            self.after(0, lambda: self._barra.reset("Descarga cancelada"))
            return False

        self.log_consola(
            f"  └ catalogo-dali-completo.xlsx actualizado desde Catálogo "
            f"Asignaciones ({len(contenido):,} bytes).".replace(",", "."), "OK")
        self.after(0, lambda: self._barra.completado("Catálogo actualizado"))
        return True

    def _descargar_catalogo_dali_web_bloqueante(self, email, password):
        """
        Ejecuta _descargar_catalogo_dali_web(email, password) en un
        hilo de fondo, mostrando una ventana modal de "Descargando..."
        mientras tanto, y no devuelve el control hasta que termina --
        pero SIN congelar la interfaz mientras espera (la GUI sigue
        respondiendo: la espera se hace con wait_window(), que corre su
        propio bucle de eventos local, y el hilo se comprueba por
        sondeo con after() en vez de con join() bloqueante).

        Corrige un bug real (2026-09-20): antes, _descargar_catalogo_
        dali_web() se llamaba directamente desde el hilo principal --
        las peticiones de red (hasta 20s+60s de timeout) bloqueaban
        Tkinter entero, mostrando "(No responde)" en la barra de
        título de Windows durante toda la espera.
        """
        resultado = {"ok": False}

        def _trabajo():
            resultado["ok"] = self._descargar_catalogo_dali_web(email, password)

        win = tk.Toplevel(self)
        win.title("Descargando catálogo")
        win.configure(bg=C_DARK)
        win.resizable(False, False)
        win.transient(self)
        win.protocol("WM_DELETE_WINDOW", lambda: None)  # no cerrable a mano
        tk.Label(
            win, text="Descargando catálogo actualizado desde\n"
                      "Catálogo Asignaciones, espere un momento...",
            bg=C_DARK, fg=C_TEXT, font=FONT_LABEL, justify="center",
            padx=36,
        ).pack(pady=(28, 4))
        lbl_tiempo = tk.Label(
            win, text="Llevamos 0s...", width=24,
            bg=C_DARK, fg=C_MUTED, font=FONT_SMALL, justify="center",
        )
        lbl_tiempo.pack(pady=(0, 20))
        win.update_idletasks()
        win.grab_set()

        # ── Centrar (mismo patrón que el resto de ventanas del proyecto) ──
        _w = win.winfo_reqwidth()
        _h = win.winfo_reqheight()
        _sw = win.winfo_screenwidth()
        _sh = win.winfo_screenheight()
        win.geometry(f"{_w}x{_h}+{(_sw-_w)//2}+{max(60, (_sh-_h)//3)}")

        hilo = threading.Thread(target=_trabajo, daemon=True)
        _inicio = time.monotonic()
        hilo.start()

        def _comprobar():
            if hilo.is_alive():
                _segundos = int(time.monotonic() - _inicio)
                lbl_tiempo.config(text=f"Llevamos {_segundos}s... (suele tardar ~1 min)")
                win.after(150, _comprobar)
            else:
                win.grab_release()
                win.destroy()

        win.after(150, _comprobar)
        win.wait_window()
        return resultado["ok"]

    def _ejecutar_en_hilo_principal_y_esperar(self, funcion):
        """
        Ejecuta `funcion` (sin argumentos) de forma segura para Tkinter y
        devuelve su resultado -- tanto si quien llama es el hilo principal
        como si es un hilo de fondo (el de proceso() de cualquier Acción).

        Bug real corregido (2026-09-24, caso real: aplicación totalmente
        congelada, hubo que matarla desde el Administrador de tareas):
        Tkinter NO es thread-safe -- crear o manipular cualquier widget
        (messagebox, tk.Toplevel, wait_window...) desde un hilo que no es
        el principal puede colgar la aplicación entera, pero de forma
        INTERMITENTE (a veces "funciona" por pura suerte de sincronización
        con lo que esté haciendo el hilo principal en ese instante, hasta
        que un día no). _catalogo_reciente_o_confirmar() y _confirmar_
        catalogo_actualizado() hacían exactamente esto -- se llaman desde
        dentro de proceso() en los 6+ puntos de enganche del catálogo, y
        proceso() siempre corre en un hilo de fondo (threading.Thread).

        Si ya estamos en el hilo principal, ejecuta `funcion` directamente
        (evita un interbloqueo real: bloquear el hilo principal esperando
        un after() que solo el propio hilo principal puede procesar sería
        un punto muerto inmediato). Si no, programa `funcion` con
        self.after(0, ...) para que corra en el hilo principal, y bloquea
        (de forma segura, con threading.Event) el hilo que llama hasta que
        termine -- devolviendo su resultado, o relanzando su excepción si
        la hubo.
        """
        if threading.current_thread() is threading.main_thread():
            return funcion()

        resultado = {}
        evento = threading.Event()

        def _envoltorio():
            try:
                resultado["valor"] = funcion()
            except Exception as e:
                resultado["excepcion"] = e
            finally:
                evento.set()

        self.after(0, _envoltorio)
        evento.wait()
        if "excepcion" in resultado:
            raise resultado["excepcion"]
        return resultado.get("valor")

    def _catalogo_reciente_o_confirmar(self):
        """
        Wrapper thread-safe (ver _ejecutar_en_hilo_principal_y_esperar):
        delega en _catalogo_reciente_o_confirmar_impl() para la lógica
        real (sin tocarla), garantizando que se ejecuta en el hilo
        principal aunque quien llame sea el hilo de fondo de proceso()
        de cualquier Acción.
        """
        return self._ejecutar_en_hilo_principal_y_esperar(
            self._catalogo_reciente_o_confirmar_impl)

    def _catalogo_reciente_o_confirmar_impl(self):
        """
        Comprobación previa, específica de Acción 4 ("Generar TARIFA final
        y Procesar Tarifas Incompletas"), llamada justo antes de cada uno
        de sus 4 puntos de enganche con _actualizar_codigo_proveedor_
        catalogo() (_ejecutar_transformacion camino limpio, y los 3 caminos
        de _gestionar_faltantes_o_finalizar_inc). A diferencia del aviso de
        Acciones 2/3 (que se muestra SIEMPRE al principio), este NO
        interrumpe nada si todo está reciente -- solo aparece cuando de
        verdad puede haber un problema, para no repetir el aviso constante
        que Acción 4 dejó de mostrar a propósito (ver _confirmar_catalogo_
        actualizado, que Acción 4 ya no llama).

        Revisa la antigüedad de catalogo-dali-completo.xlsx Y de sus dos
        ficheros de origen (ARTICULOS_DALI.xlsx, LISTADO CODIGOS DALI -
        SAP.xlsx) -- si Acciones 1-3 no se han corrido hoy, estos también
        pueden estar desactualizados aunque el catálogo en sí no se haya
        tocado hace mucho.

        Umbral: 30 minutos. Por debajo, se actualiza sin más (devuelve
        True en silencio). Por encima de alguno de los tres, ofrece el
        mismo diálogo de 3 vías que Acciones 2/3 (Sí=descargar, No=
        continuar igualmente, Cancelar=no actualizar el catálogo esta vez
        -- la tarifa en sí ya quedó completada, solo se salta este paso).

        Devuelve True si se debe proceder a actualizar el catálogo con
        _actualizar_codigo_proveedor_catalogo(), False si no.
        """
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            return True

        from datetime import datetime as _dt, timedelta as _td
        UMBRAL = _td(minutes=30)
        ahora = _dt.now()

        rutas_a_revisar = [
            ("catalogo-dali-completo.xlsx",
             os.path.join(carpeta, FOLDER_RECURSOS, "catalogo-dali-completo.xlsx")),
            ("ARTICULOS_DALI.xlsx",
             os.path.join(carpeta, FOLDER_RECURSOS, "ARTICULOS_DALI.xlsx")),
            ("LISTADO CODIGOS DALI - SAP.xlsx",
             os.path.join(carpeta, "LISTADO CODIGOS DALI - SAP.xlsx")),
        ]

        if not os.path.isfile(rutas_a_revisar[0][1]):
            return True   # catálogo nunca generado: nada que comprobar

        anticuados = []
        for nombre, ruta in rutas_a_revisar:
            if not os.path.isfile(ruta):
                continue
            try:
                mtime = _dt.fromtimestamp(os.path.getmtime(ruta))
            except Exception:
                continue
            antiguedad = ahora - mtime
            if antiguedad > UMBRAL:
                minutos = int(antiguedad.total_seconds() // 60)
                txt_edad = f"{minutos} min" if minutos < 60 \
                    else f"{minutos // 60}h {minutos % 60}min"
                anticuados.append((nombre, mtime.strftime("%d/%m/%Y %H:%M"), txt_edad))

        if not anticuados:
            return True

        detalle = "\n".join(
            f"   • {n}: actualizado {f} (hace {e})" for n, f, e in anticuados)
        respuesta = messagebox.askyesnocancel(
            "Catálogo DALI completo — puede estar desactualizado",
            "Antes de actualizar el catálogo con esta tarifa, se detectó "
            "que alguno de los archivos que lo alimentan tiene más de 30 "
            "minutos de antigüedad:\n\n"
            f"{detalle}\n\n"
            "El catálogo también se mantiene desde la aplicación «Catálogo "
            "Asignaciones» — puede convenir refrescarlo antes de escribir "
            "en él.\n\n"
            "  SÍ       → descargar ahora una copia actualizada desde "
            "Catálogo Asignaciones y continuar con ella\n"
            "  NO       → actualizar igualmente con la copia actual\n"
            "  CANCELAR → no actualizar el catálogo esta vez (la tarifa "
            "ya quedó completada igualmente)")

        if respuesta is None:
            self.log_consola(
                "Actualización del catálogo omitida por decisión del "
                "usuario (catálogo/origen desactualizado).", "WARN")
            return False
        if respuesta is False:
            return True

        email, password = self._obtener_credenciales_dali_app()
        if not email or not password:
            self.log_consola("Descarga del catálogo cancelada: sin credenciales.", "WARN")
            return False
        return self._descargar_catalogo_dali_web_bloqueante(email, password)

    def _confirmar_catalogo_actualizado(self):
        """
        Wrapper thread-safe (ver _ejecutar_en_hilo_principal_y_esperar):
        delega en _confirmar_catalogo_actualizado_impl() para la lógica
        real (sin tocarla), garantizando que se ejecuta en el hilo
        principal aunque quien llame sea el hilo de fondo de proceso()
        de Acción 2 o Acción 3.
        """
        return self._ejecutar_en_hilo_principal_y_esperar(
            self._confirmar_catalogo_actualizado_impl)

    def _confirmar_catalogo_actualizado_impl(self):
        """
        Aviso previo, común a las Acciones 2 y 3 (las que acaban
        modificando RECURSOS/catalogo-dali-completo.xlsx durante la
        actualización de plantillas DALI/SAP -- Acción 4 no la llama en
        absoluto, ver nota en _ejecutar_reinscripcion_sap_desde_inc):
        antes de empezar nada, informa de la fecha de última
        actualización de ese archivo -- el catálogo también se mantiene
        desde la app web «Catálogo Asignaciones» (asignación manual de
        Código Proveedor por otras personas), así que la copia local de
        este equipo puede quedarse desactualizada sin que esta
        aplicación tenga forma de saberlo.

        Diálogo de 3 vías (Sí / No / Cancelar):
          Sí       -> descargar ahora una copia fresca desde Catálogo
                      Asignaciones (_descargar_catalogo_dali_web()) y
                      continuar con ella.
          No       -> continuar sin descargar, con la copia local tal
                      cual está.
          Cancelar -> no continuar con la acción.

        Memoria de sesión (2026-09-21, a petición del usuario): una vez
        resuelto (Sí o No, no Cancelar) dentro de esta misma ejecución
        de la aplicación, NO se vuelve a preguntar -- Acción 2 y Acción
        3 suelen correr una justo detrás de la otra como parte del
        mismo proceso de actualización de plantillas, y repetir la
        pregunta sería puro ruido. Se reinicia solo al cerrar y volver
        a abrir la aplicación (self._catalogo_verificado_sesion vive en
        la instancia, no se persiste en tarifa_config.json).

        Devuelve True si se debe continuar (con o sin descarga previa,
        o si el catálogo todavía no existe -- nada que avisar), False
        si no se debe continuar.
        """
        if getattr(self, "_catalogo_verificado_sesion", False):
            return True   # ya se preguntó y se resolvió en esta misma sesión de trabajo

        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            return True   # sin proyecto seleccionado: cada acción ya valida esto por su cuenta

        ruta_catalogo = os.path.join(carpeta, FOLDER_RECURSOS, "catalogo-dali-completo.xlsx")
        if not os.path.isfile(ruta_catalogo):
            return True   # nunca se ha generado: nada que advertir todavía

        try:
            from datetime import datetime as _dt
            _mtime = os.path.getmtime(ruta_catalogo)
            _fecha_txt = _dt.fromtimestamp(_mtime).strftime("%d/%m/%Y a las %H:%M")
        except Exception:
            _fecha_txt = "(fecha no disponible)"

        respuesta = messagebox.askyesnocancel(
            "Catálogo DALI completo — comprobar actualización",
            "catalogo-dali-completo.xlsx se actualizó por última vez:\n\n"
            f"   {_fecha_txt}\n\n"
            "Esta acción va a modificarlo ahora. El catálogo también se "
            "mantiene desde la aplicación «Catálogo Asignaciones» (otras "
            "personas pueden haber asignado códigos proveedor después de "
            "esta fecha), así que puede convenir refrescarlo antes de "
            "seguir.\n\n"
            "  SÍ       → descargar ahora una copia actualizada desde "
            "Catálogo Asignaciones y continuar con ella\n"
            "  NO       → continuar sin descargar, con la copia actual\n"
            "  CANCELAR → no continuar con esta acción")

        if respuesta is None:
            return False   # cancelar: no marcar como resuelto, para que vuelva a preguntar
        if respuesta is False:
            self._catalogo_verificado_sesion = True
            return True
        email, password = self._obtener_credenciales_dali_app()
        if not email or not password:
            self.log_consola("Descarga del catálogo cancelada: sin credenciales.", "WARN")
            return False   # sin credenciales: no marcar, para que se pueda reintentar
        exito = self._descargar_catalogo_dali_web_bloqueante(email, password)
        if exito:
            self._catalogo_verificado_sesion = True
        return exito

    def _regenerar_catalogo_dali_completo(self):
        """
        Wrapper con bloqueo global (ver _adquirir_lock_catalogo_con_aviso):
        adquiere self._catalogo_lock antes de tocar catalogo-dali-completo.xlsx
        y lo libera siempre al terminar, delegando en _regenerar_catalogo_
        dali_completo_impl() para la lógica real (sin tocarla).
        """
        self._adquirir_lock_catalogo_con_aviso()
        try:
            resultado = self._regenerar_catalogo_dali_completo_impl()
        finally:
            self._catalogo_lock.release()
        # Etapa 6 (v10.2.50): enviar a Catálogo Asignaciones lo que haya cambiado (en segundo plano).
        self._publicar_catalogo_web("regenerar el catálogo (Acción 2 o 3)")
        return resultado

    def _regenerar_catalogo_dali_completo_impl(self):
        """
        Regenera RECURSOS/catalogo-dali-completo.xlsx combinando:
          - RECURSOS/ARTICULOS_DALI.xlsx (Acción 1): código, descripción,
            unidad, familia, subfamilia, naturaleza, IGIC, proveedor
            asignado y activo (lado DALI).
          - <raíz>/LISTADO CODIGOS DALI - SAP.xlsx (Acciones 2/3): código
            SAP, descripciones/grupos SAP, categoría valoración, tipo de
            producto, grupo de compras (lado SAP).

        Llamado automáticamente al final de _actualizar_listado_dali()
        (Acción 2) y de _cruzar_y_escribir_listado_sap() (Acción 3) --
        cada vez que cualquiera de las dos actualiza el LISTADO, el
        catálogo se regenera para mantenerse siempre sincronizado. No es
        una acción independiente con botón propio: se ejecuta de forma
        síncrona dentro del hilo de fondo que ya está corriendo, así que
        no usa su propio threading.Thread ni diálogos de confirmación, y
        cualquier error aquí se registra en el log sin interrumpir el
        éxito de la acción que lo llamó (ver el try/except del llamador).

        La columna B (Código Proveedor) se conserva de la versión
        anterior del catálogo SOLO si la columna I (Proveedor Asignado -
        DALI) de ese código no ha cambiado; si cambió, se vacía B y se
        deja constancia en un .txt en RECURSOS/ (ver más abajo). Acción 4
        (_actualizar_codigo_proveedor_catalogo) escribe B por su cuenta,
        con datos más específicos (la tarifa concreta), cuando corresponda.
        """
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            return

        recursos_path        = os.path.join(carpeta, FOLDER_RECURSOS)
        ruta_articulos_dali  = os.path.join(recursos_path, "ARTICULOS_DALI.xlsx")
        ruta_listado         = os.path.join(carpeta, "LISTADO CODIGOS DALI - SAP.xlsx")
        ruta_catalogo        = os.path.join(recursos_path, "catalogo-dali-completo.xlsx")

        if not os.path.isfile(ruta_articulos_dali) or not os.path.isfile(ruta_listado):
            self.log_consola(
                "  ℹ catalogo-dali-completo.xlsx no regenerado: falta "
                "ARTICULOS_DALI.xlsx o LISTADO CODIGOS DALI - SAP.xlsx.", "INFO")
            return

        import openpyxl
        from copy import copy
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from datetime import datetime as _dt

        self.consola.separador("REGENERAR CATÁLOGO DALI COMPLETO")

        # ── 1. Leer Código Proveedor, Proveedor Asignado y Precio € del catálogo anterior ──
        #    Se localizan las columnas por NOMBRE DE CABECERA (fila 4), no por
        #    posición: el catálogo anterior puede tener el formato viejo (20
        #    columnas) o el nuevo (22, con Texto UM y Precio €) y en ambos cada
        #    valor debe salir de SU columna. A, B e I coinciden en los dos
        #    formatos, así que son el valor por defecto si falta la cabecera.
        self.after(0, lambda: self._barra.indeterminate("Actualizando catálogo DALI..."))
        _datos_previos = {}   # clave DALI -> (cod_proveedor_previo, proveedor_asignado_previo)
        _precios_previos = {}   # clave DALI -> precio previo (columna "Precio €")
        if os.path.isfile(ruta_catalogo):
            try:
                wb_prev = openpyxl.load_workbook(ruta_catalogo, read_only=True, data_only=True)
                ws_prev = wb_prev.active
                _mapa_prev = {}
                for _n_fila, row in enumerate(
                        ws_prev.iter_rows(min_row=CATALOGO_FILA_CABECERA, values_only=True),
                        start=CATALOGO_FILA_CABECERA):
                    if _n_fila == CATALOGO_FILA_CABECERA:
                        _mapa_prev = mapa_columnas_catalogo(row or ())
                        _i_clave = _mapa_prev.get("clave", 1)
                        _i_cprov = _mapa_prev.get("cod_prov", 2)
                        _i_prov  = _mapa_prev.get("proveedor", 9)
                        _i_prec  = _mapa_prev.get("precio")     # None si el catálogo viejo no la tiene
                        continue
                    if not row or len(row) < _i_clave or row[_i_clave - 1] is None:
                        continue
                    clave_prev = str(row[_i_clave - 1]).strip()
                    cod_prov_prev  = row[_i_cprov - 1] if len(row) >= _i_cprov else None
                    prov_asig_prev = row[_i_prov - 1]  if len(row) >= _i_prov  else None
                    _datos_previos[clave_prev] = (cod_prov_prev, prov_asig_prev)
                    if _i_prec and len(row) >= _i_prec and row[_i_prec - 1] not in (None, ""):
                        _precios_previos[clave_prev] = row[_i_prec - 1]
                wb_prev.close()
                _con_cod_prov_previo = sum(
                    1 for cp, _ in _datos_previos.values() if cp not in (None, ""))
                self.log_consola(
                    f"  └ {len(_datos_previos)} código(s) leídos del catálogo anterior "
                    f"({_con_cod_prov_previo} con código proveedor, "
                    f"{len(_precios_previos)} con precio)", "INFO")
            except Exception as _e_prev:
                self.log_consola(
                    f"  ⚠  No se pudo leer el catálogo anterior "
                    f"(se generará sin conservar códigos proveedor): {_e_prev}", "WARN")

        # ── 2. Leer ARTICULOS_DALI.xlsx ─────────────────────────────
        def _lim(v):
            """Quita espacios sobrantes de inicio/fin si v es texto; deja
            números/None tal cual. Evita que espacios arrastrados desde el
            origen (p.ej. Productos.xlsx) lleguen al catálogo."""
            return v.strip() if isinstance(v, str) else v

        self.log_consola("Leyendo ARTICULOS_DALI.xlsx...", "INFO")
        wb_dali = openpyxl.load_workbook(ruta_articulos_dali, read_only=True, data_only=True)
        ws_dali = wb_dali.active
        tabla_dali = {}
        for row in ws_dali.iter_rows(min_row=8, values_only=True):
            if not row or row[0] is None:
                continue
            clave = str(row[0]).strip()
            tabla_dali[clave] = {
                "descripcion": _lim(row[1]) if len(row) > 1 else None,
                "unidad":      _lim(row[2]) if len(row) > 2 else None,
                "familia":     _lim(row[3]) if len(row) > 3 else None,
                "subfamilia":  _lim(row[4]) if len(row) > 4 else None,
                "naturaleza":  _lim(row[5]) if len(row) > 5 else None,
                "igic":        _lim(row[6]) if len(row) > 6 else None,
                "proveedor":   _lim(row[7]) if len(row) > 7 else None,
                "activo":      _lim(row[8]) if len(row) > 8 else None,
            }
        wb_dali.close()
        self.log_consola(f"  └ {len(tabla_dali)} artículos DALI leídos", "OK")

        # ── 3. Leer LISTADO CODIGOS DALI - SAP.xlsx ─────────────────
        self.log_consola("Leyendo LISTADO CODIGOS DALI - SAP.xlsx...", "INFO")
        wb_lst = openpyxl.load_workbook(ruta_listado, read_only=True, data_only=True)
        ws_lst = wb_lst.active
        tabla_sap = {}
        for row in ws_lst.iter_rows(min_row=8, values_only=True):
            if not row or row[0] is None:
                continue
            clave = str(row[0]).strip()
            tabla_sap[clave] = datos_sap_desde_fila_listado(row)
        wb_lst.close()
        self.log_consola(f"  └ {len(tabla_sap)} entrada(s) del LISTADO leídas", "OK")

        # ── 4. Construir el catálogo ─────────────────────────────────
        wb_out = openpyxl.Workbook()
        ws_out = wb_out.active
        ws_out.title = "Catálogo completo"

        _codigos_ordenados = sorted(
            tabla_dali.keys(),
            key=lambda c: (0, int(c)) if c.isdigit() else (1, c))

        _c_titulo = ws_out.cell(row=1, column=1,
                    value="Listado completo de artículos DALI — importación DALI + SAP + código proveedor")
        _c_titulo.font = Font(size=14, bold=True)
        _c_subtitulo = ws_out.cell(row=2, column=1,
                    value=f"Generado el {_dt.now().strftime('%d/%m/%Y, %H:%M:%S')} · "
                          f"{len(_codigos_ordenados)} artículo(s)")
        _c_subtitulo.font = Font(size=9, italic=True, color="FF666666")

        # Layout desde la definición canónica (sap_helpers.CATALOGO_COLUMNAS):
        # cabecera, orden de columnas y anchos salen de ahí, así que este
        # método no puede volver a descuadrarse respecto a los demás.
        _claves_col = [k for k, _h, _a in CATALOGO_COLUMNAS]
        _fill_cabecera = PatternFill(fill_type="solid", fgColor="FFE5DECB")
        _lado = Side(style="thin", color="FF000000")
        _borde = Border(left=_lado, right=_lado, top=_lado, bottom=_lado)
        _centrado = Alignment(horizontal="center")
        for i, (_k, titulo, _a) in enumerate(CATALOGO_COLUMNAS, start=1):
            c = ws_out.cell(row=CATALOGO_FILA_CABECERA, column=i, value=titulo)
            c.font = Font(bold=True)
            c.fill = _fill_cabecera
            c.alignment = _centrado
            c.border = _borde

        # Estilo de las celdas de datos (centrado + borde fino, como el
        # catálogo que exporta la web). Se construye UNA vez y se copia por
        # celda: asignar font/alignment/border en cada una de las ~660.000
        # celdas sería mucho más lento.
        _proto = ws_out.cell(row=CATALOGO_PRIMERA_FILA_DATOS, column=1)
        _proto.alignment = _centrado
        _proto.border = _borde
        _estilo_datos = copy(_proto._style)

        _fila = CATALOGO_PRIMERA_FILA_DATOS
        _con_proveedor_conservado = 0
        _cambios_proveedor_asignado = []  # [(clave, prov_asig_anterior, prov_asig_nuevo, cod_prov_vaciado, precio_vaciado)]

        def _normalizar(v):
            return str(v).strip().upper() if v not in (None, "") else ""

        for clave in _codigos_ordenados:
            d = tabla_dali[clave]
            s = tabla_sap.get(clave, {})
            prov_asig_nuevo = d["proveedor"]

            cod_prov = None
            precio = None
            if clave in _datos_previos:
                cod_prov_prev, prov_asig_prev = _datos_previos[clave]
                precio_prev = _precios_previos.get(clave)
                if _normalizar(prov_asig_prev) == _normalizar(prov_asig_nuevo):
                    # Proveedor Asignado sin cambios -> conservar Código
                    # Proveedor y Precio € (ambos dependen del proveedor)
                    cod_prov = cod_prov_prev
                    precio = precio_prev
                    if cod_prov not in (None, ""):
                        _con_proveedor_conservado += 1
                elif cod_prov_prev not in (None, "") or precio_prev not in (None, ""):
                    # Proveedor Asignado cambió Y había un código proveedor y/o
                    # un precio que ya no son fiables -> vaciar ambos y dejar
                    # constancia en el .txt
                    _cambios_proveedor_asignado.append(
                        (clave, prov_asig_prev, prov_asig_nuevo, cod_prov_prev, precio_prev))

            # Cada valor va asociado a su CLAVE de columna (no a una posición):
            # el orden real lo decide CATALOGO_COLUMNAS al escribir.
            valores = {
                "clave":       int(clave) if clave.isdigit() else clave,
                "cod_prov":    cod_prov,
                "descripcion": d["descripcion"], "unidad": d["unidad"],
                "igic":        d["igic"],        "naturaleza": d["naturaleza"],
                "familia":     d["familia"],     "subfamilia": d["subfamilia"],
                "proveedor":   d["proveedor"],   "activo":     d["activo"],
                # "Precio €": igual que el código proveedor, se conserva del
                # catálogo anterior solo si el Proveedor Asignado no cambió.
                "precio":      precio,
            }
            for _k_sap in CATALOGO_CLAVES_SAP:
                valores[_k_sap] = s.get(_k_sap)
            for i, _k in enumerate(_claves_col, start=1):
                c = ws_out.cell(row=_fila, column=i, value=valores.get(_k))
                c._style = copy(_estilo_datos)
            _fila += 1

        for i, (_k, _h, ancho) in enumerate(CATALOGO_COLUMNAS, start=1):
            ws_out.column_dimensions[openpyxl.utils.get_column_letter(i)].width = ancho
        ws_out.auto_filter.ref = (
            f"A{CATALOGO_FILA_CABECERA}:"
            f"{openpyxl.utils.get_column_letter(len(CATALOGO_COLUMNAS))}"
            f"{CATALOGO_FILA_CABECERA}")

        os.makedirs(recursos_path, exist_ok=True)
        wb_out.save(ruta_catalogo)

        # ── 4b. Si hubo cambios de Proveedor Asignado, dejar constancia en .txt ──
        if _cambios_proveedor_asignado:
            _sello = _dt.now().strftime('%Y%m%d_%H%M%S')
            _ruta_txt_cambios = os.path.join(
                recursos_path, f"catalogo_cambios_proveedor_{_sello}.txt")
            with open(_ruta_txt_cambios, "w", encoding="utf-8") as _f_txt:
                _f_txt.write(
                    "CASOS IRREGULARES — CAMBIOS DE PROVEEDOR ASIGNADO DETECTADOS "
                    "AL REGENERAR EL CATÁLOGO\n"
                    f"Generado el {_dt.now().strftime('%d/%m/%Y, %H:%M:%S')}\n\n"
                    "Los siguientes códigos DALI han cambiado de \"Proveedor Asignado - DALI\"\n"
                    "respecto al catálogo anterior. Normalmente ambas bases de datos de\n"
                    "proveedores deberían coincidir exactamente en ambos listados -- esto\n"
                    "suele indicar un desfase puntual (p.ej. el nombre se actualizó en una\n"
                    "base de datos antes que en la otra). Por precaución, se han vaciado su\n"
                    "Código Proveedor (columna B) y su Precio € -- revíselos y complételos de\n"
                    "nuevo si procede.\n\n"
                )
                _f_txt.write(
                    f"{'Código DALI':<12} | {'Proveedor anterior':<35} | "
                    f"{'Proveedor nuevo':<35} | {'Código Proveedor vaciado':<25} | Precio € vaciado\n")
                _f_txt.write("-" * 130 + "\n")
                for _clave_c, _prov_ant, _prov_new, _cod_vaciado, _precio_vaciado in _cambios_proveedor_asignado:
                    _f_txt.write(
                        f"{_clave_c:<12} | {str(_prov_ant or ''):<35} | "
                        f"{str(_prov_new or ''):<35} | {str(_cod_vaciado or ''):<25} | "
                        f"{'' if _precio_vaciado in (None, '') else _precio_vaciado}\n")
            self.log_consola(
                f"  ⚠  CASO IRREGULAR: {len(_cambios_proveedor_asignado)} código(s) DALI "
                f"cambiaron de Proveedor Asignado — código proveedor y precio vaciados, "
                f"detalle en {os.path.basename(_ruta_txt_cambios)}", "WARN")
            for _clave_c, _prov_ant, _prov_new, _cod_vaciado, _precio_vaciado in _cambios_proveedor_asignado[:30]:
                _vaciado = []
                if _cod_vaciado not in (None, ""):
                    _vaciado.append(f"código {_cod_vaciado}")
                if _precio_vaciado not in (None, ""):
                    _vaciado.append(f"precio {_precio_vaciado}")
                self.log_consola(
                    f"      · DALI {_clave_c}: «{_prov_ant or '(vacío)'}» → "
                    f"«{_prov_new or '(vacío)'}» ({' y '.join(_vaciado)} vaciado)", "WARN")
            if len(_cambios_proveedor_asignado) > 30:
                self.log_consola(
                    f"      … y {len(_cambios_proveedor_asignado) - 30} más "
                    f"(omitidos del log, ver {os.path.basename(_ruta_txt_cambios)})", "WARN")

            def _avisar_cambios_proveedor_asignado(cambios=_cambios_proveedor_asignado,
                                                    nombre_txt=os.path.basename(_ruta_txt_cambios)):
                lineas = [
                    "⚠  CASO IRREGULAR — Proveedor Asignado cambió",
                    "",
                    f"Los siguientes {len(cambios)} código(s) DALI han cambiado de",
                    "\"Proveedor Asignado - DALI\" respecto al catálogo anterior.",
                    "Normalmente ambas bases de datos de proveedores deberían",
                    "coincidir exactamente -- revise estos códigos, puede tratarse",
                    "de un desfase puntual entre listados:",
                    "",
                ]
                for _clave_c, _prov_ant, _prov_new, _cod_vaciado, _precio_vaciado in cambios[:20]:
                    lineas.append(
                        f"   • DALI {_clave_c}: «{_prov_ant or '(vacío)'}» → "
                        f"«{_prov_new or '(vacío)'}»")
                if len(cambios) > 20:
                    lineas.append(f"   … y {len(cambios) - 20} más (ver log de consola).")
                lineas += [
                    "",
                    "Su Código Proveedor y su Precio € se han vaciado en catalogo-dali-completo.xlsx.",
                    f"Detalle completo en RECURSOS/{nombre_txt}",
                ]
                messagebox.showwarning(
                    "Caso irregular — Proveedor Asignado cambió", "\n".join(lineas))
            self.after(0, _avisar_cambios_proveedor_asignado)

        _n_filas = _fila - 5
        self.log_consola(f"  └ catalogo-dali-completo.xlsx regenerado: {_n_filas} fila(s), "
                          f"{_con_proveedor_conservado} código(s) proveedor conservados", "OK")
        self.after(0, lambda: self._barra.completado("Catálogo DALI actualizado"))
        self.consola.separador("FIN REGENERAR CATÁLOGO DALI COMPLETO")

    def _refrescar_sap_catalogo_para_dali(self, codigos_dali):
        """
        Actualiza SOLO las columnas J-S (datos de origen SAP: código SAP,
        descripción, unidad, grupo de producto, IGIC, categoría, tipo,
        grupo de compras) de las filas de catalogo-dali-completo.xlsx
        correspondientes a los códigos DALI indicados, leyendo su estado
        actual desde LISTADO CODIGOS DALI - SAP.xlsx. No toca el resto
        de columnas (proveedor, datos DALI) ni añade filas nuevas.

        Fix (2026-09-25, hallazgo real del usuario): cuando Ventana
        SAPFaltantes resuelve un código SAP duplicado y borra el SAP de
        un DALI que NO pertenece a la tarifa en curso (ver
        VentanaResolverSAPDuplicados), _actualizar_codigo_proveedor_
        catalogo() nunca lo refresca -- solo toca los códigos DALI de
        la propia tarifa. El catálogo se queda con el SAP viejo hasta
        la siguiente Acción 2/3. Este método cubre ese hueco con una
        actualización puntual, ligera, sin regenerar el catálogo entero.

        codigos_dali: iterable de códigos DALI (int o str) a refrescar.
        No hace nada si el catálogo o el LISTADO no existen todavía.
        """
        if not codigos_dali:
            return
        self._adquirir_lock_catalogo_con_aviso()
        try:
            import openpyxl

            carpeta       = self._carpeta_proyecto.get().strip()
            recursos_path = os.path.join(carpeta, FOLDER_RECURSOS)
            ruta_catalogo = os.path.join(recursos_path, "catalogo-dali-completo.xlsx")
            ruta_listado  = os.path.join(carpeta, "LISTADO CODIGOS DALI - SAP.xlsx")
            if not os.path.isfile(ruta_catalogo) or not os.path.isfile(ruta_listado):
                return

            claves = set()
            for c in codigos_dali:
                try:
                    claves.add(str(int(float(str(c).strip()))))
                except (ValueError, TypeError):
                    continue
            if not claves:
                return

            wb_l = openpyxl.load_workbook(ruta_listado, read_only=True, data_only=True)
            ws_l = wb_l.active
            _datos_sap = {}
            for row in ws_l.iter_rows(min_row=8, values_only=True):
                if not row or row[0] is None:
                    continue
                clave = str(row[0]).strip()
                if clave not in claves:
                    continue
                _datos_sap[clave] = datos_sap_desde_fila_listado(row)
            wb_l.close()

            wb_cat = openpyxl.load_workbook(ruta_catalogo)
            ws_cat = wb_cat.active
            # Columnas localizadas por cabecera: valen tanto el formato viejo
            # (20 columnas) como el nuevo (con Texto UM y Precio €).
            _mapa = mapa_columnas_catalogo(
                [c.value for c in ws_cat[CATALOGO_FILA_CABECERA]])
            _col_clave = _mapa.get("clave", 1)
            _fila_por_clave = {}
            for row in ws_cat.iter_rows(min_row=CATALOGO_PRIMERA_FILA_DATOS):
                if not row or len(row) < _col_clave or row[_col_clave - 1].value is None:
                    continue
                _fila_por_clave[str(row[_col_clave - 1].value).strip()] = row[0].row

            actualizadas = 0
            for clave in claves:
                fila = _fila_por_clave.get(clave)
                if fila is None:
                    continue
                s = _datos_sap.get(clave, {})
                # Solo se tocan las columnas de origen SAP; "Precio €", el
                # proveedor y los datos DALI no se modifican aquí.
                for _k in CATALOGO_CLAVES_SAP:
                    _col = _mapa.get(_k)
                    if _col:
                        # ws_cat.cell(..., value=v) NO borra si v es None (bug
                        # conocido de openpyxl, ya visto antes en este mismo
                        # proyecto) -- hace falta asignar .value en dos pasos.
                        ws_cat.cell(row=fila, column=_col).value = s.get(_k)
                actualizadas += 1

            wb_cat.save(ruta_catalogo)
            wb_cat.close()
            if actualizadas:
                self.log_consola(
                    f"  └ Catálogo: datos SAP refrescados para {actualizadas} "
                    f"código(s) DALI tras resolución de duplicados.", "OK")
        except Exception as _e_ref:
            self.log_consola(
                f"  ⚠  No se pudo refrescar el catálogo tras resolver "
                f"duplicados: {_e_ref}", "WARN")
        finally:
            self._catalogo_lock.release()
        # Etapa 6 (v10.2.50): enviar a Catálogo Asignaciones los datos SAP cambiados.
        self._publicar_catalogo_web("resolver un código SAP duplicado")

    def _actualizar_codigo_proveedor_catalogo(self, ruta_tarifa):
        """
        Wrapper con bloqueo global (ver _adquirir_lock_catalogo_con_aviso):
        adquiere self._catalogo_lock antes de tocar catalogo-dali-completo.xlsx
        y lo libera siempre al terminar, delegando en _actualizar_codigo_
        proveedor_catalogo_impl() para la lógica real (sin tocarla).
        """
        self._adquirir_lock_catalogo_con_aviso()
        try:
            resultado = self._actualizar_codigo_proveedor_catalogo_impl(ruta_tarifa)
        finally:
            self._catalogo_lock.release()
        # Etapa 6 (v10.2.50): enviar a Catálogo Asignaciones el precio y el código de proveedor cambiados.
        self._publicar_catalogo_web("completar la tarifa (Acción 4)", os.path.basename(str(ruta_tarifa or "")) or None)
        return resultado

    def _actualizar_codigo_proveedor_catalogo_impl(self, ruta_tarifa):
        """
        Llamada desde _ejecutar_reinscripcion_sap_desde_inc() (Acción 4)
        justo después de que una tarifa queda lista para crear/modificar
        el Registro Info (renombrada, sin [INCOMPLETO] ni prefijo INC).

        Actualiza en RECURSOS/catalogo-dali-completo.xlsx, ÚNICAMENTE las
        filas de los códigos DALI presentes en esta tarifa (col B, desde
        fila 8) CUYO Proveedor Asignado (ARTICULOS_DALI.xlsx, el mismo
        valor que alimenta la col I del catálogo) coincide con el
        proveedor de esta tarifa (celda C2). Cuando hay match, refresca
        las 20 columnas con los valores actuales de:
          - RECURSOS/ARTICULOS_DALI.xlsx (Acción 1)         → cols A, C-I
          - <raíz>/LISTADO CODIGOS DALI - SAP.xlsx (Acc. 2/3) → cols J-T
          - la propia tarifa (Acción 4, col C)               → col B (Código Proveedor)

        Si el Proveedor Asignado NO coincide con el proveedor de la
        tarifa, esa fila NO se toca en absoluto (ninguna de las 20
        columnas) -- se avisa en el log y con una ventana de advertencia
        al terminar, para que se pueda revisar la asignación DALI.

        Si un código DALI de la tarifa no existe todavía como fila en el
        catálogo, se añade al final (solo si el proveedor coincide). No
        se toca ninguna fila de códigos DALI ajenos a esta tarifa.

        Si el catálogo todavía no existe (no se ha ejecutado la Acción 8
        nunca), no hace nada salvo avisar en el log -- no es un error.
        """
        try:
            import openpyxl

            carpeta       = self._carpeta_proyecto.get().strip()
            recursos_path = os.path.join(carpeta, FOLDER_RECURSOS)
            ruta_catalogo = os.path.join(recursos_path, "catalogo-dali-completo.xlsx")
            if not os.path.isfile(ruta_catalogo):
                self.log_consola(
                    "  ℹ catalogo-dali-completo.xlsx no existe todavía "
                    "(genérelo con la Acción 8) — catálogo no actualizado.",
                    "INFO")
                return

            # ── 1. Proveedor (C2), códigos DALI + código proveedor de la tarifa ──
            wb_tar = openpyxl.load_workbook(ruta_tarifa, read_only=True, data_only=True)
            ws_tar = wb_tar.active
            _nombre_prov_tarifa = ws_tar["C2"].value
            _nombre_prov_tarifa = str(_nombre_prov_tarifa).strip() if _nombre_prov_tarifa else ""
            _cod_prov_tarifa = {}     # clave DALI (str) -> código proveedor (tarifa col C)
            _precio_tarifa = {}       # clave DALI (str) -> PRECIO de la tarifa final (tarifa col F)
            _fila_tarifa_por_clave = {}   # clave DALI (str) -> nº de fila en la tarifa (para completar col C luego)
            _fila_actual_tar = 7
            for row in ws_tar.iter_rows(min_row=8, values_only=True):
                _fila_actual_tar += 1
                if not row or len(row) < 2:
                    continue
                v_dali = row[1]   # col B
                if v_dali is None:
                    continue
                try:
                    clave = str(int(float(str(v_dali).strip())))
                except (ValueError, TypeError):
                    continue
                v_prov = row[2] if len(row) > 2 else None   # col C
                _cod_prov_tarifa[clave] = v_prov
                _precio_tarifa[clave] = row[5] if len(row) > 5 else None   # col F: PRECIO
                _fila_tarifa_por_clave[clave] = _fila_actual_tar
            wb_tar.close()

            if not _cod_prov_tarifa:
                self.log_consola(
                    "  ⚠  No se pudieron leer códigos DALI de la tarifa "
                    "— catálogo no actualizado.", "WARN")
                return

            if not _nombre_prov_tarifa:
                self.log_consola(
                    "  ⚠  No se pudo leer el proveedor de la tarifa (celda C2) "
                    "— ninguna fila se actualizará en el catálogo hasta corregirlo.",
                    "WARN")

            # ── 2. Datos DALI (Acción 1), solo los códigos que necesitamos ──
            def _lim(v):
                """Quita espacios sobrantes de inicio/fin si v es texto."""
                return v.strip() if isinstance(v, str) else v

            ruta_articulos_dali = os.path.join(recursos_path, "ARTICULOS_DALI.xlsx")
            _datos_dali = {}
            if os.path.isfile(ruta_articulos_dali):
                wb_a = openpyxl.load_workbook(ruta_articulos_dali, read_only=True, data_only=True)
                ws_a = wb_a.active
                for row in ws_a.iter_rows(min_row=8, values_only=True):
                    if not row or row[0] is None:
                        continue
                    clave = str(row[0]).strip()
                    if clave not in _cod_prov_tarifa:
                        continue
                    _datos_dali[clave] = {
                        "descripcion": _lim(row[1]) if len(row) > 1 else None,
                        "unidad":      _lim(row[2]) if len(row) > 2 else None,
                        "familia":     _lim(row[3]) if len(row) > 3 else None,
                        "subfamilia":  _lim(row[4]) if len(row) > 4 else None,
                        "naturaleza":  _lim(row[5]) if len(row) > 5 else None,
                        "igic":        _lim(row[6]) if len(row) > 6 else None,
                        "proveedor":   _lim(row[7]) if len(row) > 7 else None,
                        "activo":      _lim(row[8]) if len(row) > 8 else None,
                    }
                wb_a.close()
            else:
                self.log_consola(
                    "  ⚠  ARTICULOS_DALI.xlsx no encontrado — "
                    "catálogo actualizado solo con los datos de la tarifa.", "WARN")

            # ── 3. Datos SAP (Acciones 2/3), solo los códigos que necesitamos ──
            ruta_listado = os.path.join(carpeta, "LISTADO CODIGOS DALI - SAP.xlsx")
            _datos_sap = {}
            if os.path.isfile(ruta_listado):
                wb_l = openpyxl.load_workbook(ruta_listado, read_only=True, data_only=True)
                ws_l = wb_l.active
                for row in ws_l.iter_rows(min_row=8, values_only=True):
                    if not row or row[0] is None:
                        continue
                    clave = str(row[0]).strip()
                    if clave not in _cod_prov_tarifa:
                        continue
                    _datos_sap[clave] = datos_sap_desde_fila_listado(row)
                wb_l.close()
            else:
                self.log_consola(
                    "  ⚠  LISTADO CODIGOS DALI-SAP no encontrado — "
                    "catálogo actualizado solo parcialmente.", "WARN")

            # ── 4. Abrir catálogo, localizar filas existentes por código DALI ──
            wb_cat = openpyxl.load_workbook(ruta_catalogo)
            ws_cat = wb_cat.active
            # Columnas por CABECERA (no por posición): cada valor cae en su
            # columna tanto en el formato viejo (20 cols) como en el nuevo
            # (22, con Texto UM y Precio €). Las columnas que el fichero no
            # tenga se saltan en lugar de escribir en otra.
            _mapa = mapa_columnas_catalogo(
                [c.value for c in ws_cat[CATALOGO_FILA_CABECERA]])
            _col_clave = _mapa.get("clave", 1)
            _col_cprov = _mapa.get("cod_prov", 2)
            _fila_por_clave = {}
            for row in ws_cat.iter_rows(min_row=CATALOGO_PRIMERA_FILA_DATOS):
                if not row or len(row) < _col_clave or row[_col_clave - 1].value is None:
                    continue
                _fila_por_clave[str(row[_col_clave - 1].value).strip()] = row[0].row

            def _normalizar(v):
                return str(v).strip().upper() if v not in (None, "") else ""

            _nombre_norm = _normalizar(_nombre_prov_tarifa)
            _actualizadas   = 0
            _anadidas       = 0
            _omitidas_proveedor = []   # [(clave, proveedor_asignado)]
            _completados_desde_catalogo = []   # [(clave, cod_prov_recuperado)]
            _fila_siguiente = ws_cat.max_row + 1

            # "Precio €" funciona como "Código Proveedor": se pega el valor de
            # la tarifa (col F, precio final ya preparado para el paso RINFO)
            # y, si el catálogo ya tenía uno, se SUSTITUYE. Solo se acepta un
            # número mayor que 0: una celda vacía, un 0 o un texto no deben
            # borrar un precio bueno del catálogo (en ese caso se conserva el
            # que ya hubiera y se avisa).
            _precios_escritos = 0
            _precios_sustituidos = 0    # sustituyen un valor anterior DISTINTO
            _precios_invalidos = []     # DALI sin precio válido en la tarifa
            _col_precio = _mapa.get("precio")

            for clave, cod_prov in _cod_prov_tarifa.items():
                d = _datos_dali.get(clave, {})
                s = _datos_sap.get(clave, {})

                _proveedor_asignado = d.get("proveedor")
                if _normalizar(_proveedor_asignado) != _nombre_norm:
                    _omitidas_proveedor.append((clave, _proveedor_asignado))
                    continue

                if clave in _fila_por_clave:
                    fila = _fila_por_clave[clave]
                    _actualizadas += 1
                    # La tarifa trae el código proveedor vacío para este
                    # material -- si el catálogo ya tenía uno guardado de
                    # antes, se conserva (no se borra) y además se recupera
                    # para completar la propia tarifa (ver más abajo).
                    if cod_prov in (None, ""):
                        _cod_prov_existente = ws_cat.cell(row=fila, column=_col_cprov).value
                        if _cod_prov_existente not in (None, ""):
                            cod_prov = _cod_prov_existente
                            _completados_desde_catalogo.append((clave, cod_prov))
                else:
                    fila = _fila_siguiente
                    _fila_siguiente += 1
                    _anadidas += 1

                valores = {
                    "clave":       int(clave) if clave.isdigit() else clave,
                    "cod_prov":    cod_prov,
                    "descripcion": d.get("descripcion"), "unidad": d.get("unidad"),
                    "igic":        d.get("igic"),        "naturaleza": d.get("naturaleza"),
                    "familia":     d.get("familia"),     "subfamilia": d.get("subfamilia"),
                    "proveedor":   d.get("proveedor"),   "activo":     d.get("activo"),
                }
                for _k_sap in CATALOGO_CLAVES_SAP:
                    valores[_k_sap] = s.get(_k_sap)
                _precio = _precio_tarifa.get(clave)
                if es_precio_valido(_precio):
                    valores["precio"] = _precio
                    _precios_escritos += 1
                    if _col_precio:
                        _prev = ws_cat.cell(row=fila, column=_col_precio).value
                        if _prev not in (None, "") and _prev != _precio:
                            _precios_sustituidos += 1
                else:
                    _precios_invalidos.append(clave)
                # (cell(..., value=None) no borra: si un dato no viene, el que
                # ya hubiera en el catálogo se conserva -- comportamiento
                # previo que se mantiene, y es lo que protege el precio cuando
                # la tarifa no trae uno válido.)
                for _k, _v in valores.items():
                    _col = _mapa.get(_k)
                    if _col:
                        ws_cat.cell(row=fila, column=_col, value=_v)

            wb_cat.save(ruta_catalogo)
            wb_cat.close()

            self.log_consola(
                f"  └ Catálogo DALI actualizado desde la tarifa: "
                f"{_actualizadas} fila(s) modificadas, {_anadidas} fila(s) nuevas añadidas",
                "OK")
            if not _col_precio:
                self.log_consola(
                    "  ⚠  El catálogo no tiene la columna «Precio €» (formato antiguo): "
                    "los precios de la tarifa no se han guardado; se incorporarán "
                    "cuando se regenere el catálogo (Acción 2/3).", "WARN")
            else:
                self.log_consola(
                    f"  └ Precio €: {_precios_escritos} precio(s) pegado(s) desde la "
                    f"tarifa ({_precios_sustituidos} sustituyen un valor anterior "
                    f"distinto)", "OK")
                if _precios_invalidos:
                    _lista_inv = ", ".join(_precios_invalidos[:10]) + (
                        f" … y {len(_precios_invalidos) - 10} más"
                        if len(_precios_invalidos) > 10 else "")
                    self.log_consola(
                        f"  ⚠  {len(_precios_invalidos)} código(s) DALI sin precio "
                        f"válido en la tarifa (columna F vacía, 0 o no numérica) — se "
                        f"conserva el que ya tuviera el catálogo: DALI {_lista_inv}",
                        "WARN")

            # ── 5. Completar en la propia tarifa los códigos proveedor que ──
            #      faltaban y se han recuperado del catálogo, para que el
            #      Registro Info que se cree a partir de ella salga más
            #      completo.
            if _completados_desde_catalogo:
                try:
                    wb_tar2 = openpyxl.load_workbook(ruta_tarifa)
                    ws_tar2 = wb_tar2.active
                    for _clave_c, _cod_c in _completados_desde_catalogo:
                        _fila_c = _fila_tarifa_por_clave.get(_clave_c)
                        if _fila_c is not None:
                            ws_tar2.cell(row=_fila_c, column=3, value=_cod_c)
                    wb_tar2.save(ruta_tarifa)
                    wb_tar2.close()
                    self.log_consola(
                        f"  └ {len(_completados_desde_catalogo)} código(s) proveedor "
                        f"completados en la tarifa desde el catálogo DALI:", "OK")
                    for _clave_c, _cod_c in _completados_desde_catalogo[:30]:
                        self.log_consola(
                            f"      · DALI {_clave_c} → «{_cod_c}»", "OK")
                    if len(_completados_desde_catalogo) > 30:
                        self.log_consola(
                            f"      … y {len(_completados_desde_catalogo) - 30} más "
                            f"(omitidos del log)", "OK")

                    def _avisar_completados_desde_catalogo(
                            completados=_completados_desde_catalogo):
                        lineas = [
                            f"✔  {len(completados)} código(s) proveedor completados",
                            "",
                            "Estos artículos venían sin Código Proveedor en la tarifa,",
                            "pero el catálogo DALI completo ya tenía uno guardado de",
                            "antes -- se ha recuperado y escrito también en la tarifa,",
                            "para que el Registro Info quede más completo:",
                            "",
                        ]
                        for _clave_c, _cod_c in completados[:20]:
                            lineas.append(f"   • DALI {_clave_c} → «{_cod_c}»")
                        if len(completados) > 20:
                            lineas.append(f"   … y {len(completados) - 20} más (ver log de consola).")
                        messagebox.showinfo(
                            "Códigos proveedor completados desde el catálogo",
                            "\n".join(lineas))
                    self.after(0, _avisar_completados_desde_catalogo)
                except Exception as _e_completar:
                    self.log_consola(
                        f"  ⚠  No se pudieron escribir los códigos proveedor "
                        f"recuperados en la tarifa: {_e_completar}", "WARN")

            if _omitidas_proveedor:
                self.log_consola(
                    f"  ⚠  CASO IRREGULAR: {len(_omitidas_proveedor)} código(s) DALI de la "
                    f"tarifa NO tienen «{_nombre_prov_tarifa}» como Proveedor Asignado en "
                    f"ARTICULOS_DALI.xlsx — fila(s) NO modificada(s):",
                    "WARN")
                for _clave_om, _prov_om in _omitidas_proveedor[:30]:
                    self.log_consola(
                        f"      · DALI {_clave_om}: Proveedor Asignado real = "
                        f"«{_prov_om or '(vacío)'}»", "WARN")
                if len(_omitidas_proveedor) > 30:
                    self.log_consola(
                        f"      … y {len(_omitidas_proveedor) - 30} más (omitidos del log)", "WARN")

                def _avisar_proveedor_no_coincide(nombre=_nombre_prov_tarifa,
                                                   omitidas=_omitidas_proveedor):
                    lineas = [
                        f"⚠  CASO IRREGULAR — Proveedor Asignado no coincide",
                        "",
                        f"Esta tarifa es de «{nombre}», pero en ARTICULOS_DALI.xlsx los",
                        f"siguientes {len(omitidas)} código(s) DALI tienen asignado un",
                        "proveedor distinto. Normalmente ambas bases de datos de",
                        "proveedores deberían coincidir exactamente -- revise estos",
                        "códigos, puede tratarse de un desfase puntual entre listados:",
                        "",
                    ]
                    for _clave_om, _prov_om in omitidas[:20]:
                        lineas.append(f"   • DALI {_clave_om}: proveedor real = «{_prov_om or '(vacío)'}»")
                    if len(omitidas) > 20:
                        lineas.append(f"   … y {len(omitidas) - 20} más (ver log de consola).")
                    lineas += ["", "No se ha modificado su fila en catalogo-dali-completo.xlsx."]
                    messagebox.showwarning(
                        "Caso irregular — Proveedor Asignado no coincide", "\n".join(lineas))
                self.after(0, _avisar_proveedor_no_coincide)

        except Exception as e_cat:
            self.log_consola(
                f"  ⚠  No se pudo actualizar catalogo-dali-completo.xlsx: {e_cat}", "WARN")
