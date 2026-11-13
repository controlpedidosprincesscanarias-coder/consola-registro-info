# -*- coding: utf-8 -*-
"""
ventana_login.py — inicio de sesión de Catálogo Asignaciones al abrir la Consola (v10.2.52).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Decisión de Víctor (13 nov 2026): solo entra quien tiene en Catálogo Asignaciones una cuenta activa, de
administrador y con el permiso «Escribir desde la Consola Registro Info»; a cualquier otra, «Usuario no
autorizado.» (sin anotar nada). La comprobación está en app/comun/catalogo_web.comprobar_acceso (con pruebas);
aquí solo van la ventana y el hilo (la petición a la web no congela la ventana).

pedir_inicio_sesion(root, email, password, comprobar) abre la ventana (sin transient: la principal está oculta) y
devuelve (email, password, usuario) si entra, o None si se pulsa «Salir» o se cierra. Si llegan credenciales
guardadas, las prueba solas al abrirse. `comprobar(email, password)` devuelve (estado, mensaje, usuario).
"""
import threading
import tkinter as tk

from app.comun.tarifa_constants import (
    C_ACCENT, C_ACCENT2, C_DARK, C_ERROR, C_MUTED, C_TEXT_LIGHT, FONT_BOLD, FONT_LABEL, FONT_SMALL,
)
from app.comun.tarifa_core import cargar_icono_ventana


def pedir_inicio_sesion(root, email_guardado, password_guardada, comprobar, al_fallar_credenciales=None):
    resultado = {"valor": None}
    win = tk.Toplevel(root)
    win.title("Consola Registro Info — Iniciar sesión")
    win.configure(bg=C_DARK)
    win.resizable(False, False)
    cargar_icono_ventana(win)

    frm = tk.Frame(win, bg=C_DARK, padx=28, pady=22)
    frm.pack(fill="both", expand=True)
    tk.Frame(frm, bg=C_ACCENT, height=3).pack(fill="x", pady=(0, 14))
    tk.Label(frm, text="CREACIÓN  REGISTRO  INFO  SAP", bg=C_DARK, fg=C_ACCENT,
             font=("Segoe UI", 13, "bold")).pack(anchor="w")
    tk.Label(frm, text="Inicia sesión con tu usuario de Catálogo Asignaciones.",
             bg=C_DARK, fg=C_MUTED, font=FONT_SMALL).pack(anchor="w", pady=(2, 16))

    tk.Label(frm, text="Email", bg=C_DARK, fg=C_TEXT_LIGHT, font=FONT_LABEL).pack(anchor="w")
    e_email = tk.Entry(frm, font=FONT_LABEL, width=36)
    e_email.pack(fill="x", pady=(2, 10))
    tk.Label(frm, text="Contraseña", bg=C_DARK, fg=C_TEXT_LIGHT, font=FONT_LABEL).pack(anchor="w")
    e_pass = tk.Entry(frm, font=FONT_LABEL, width=36, show="•")
    e_pass.pack(fill="x", pady=(2, 8))
    if email_guardado:
        e_email.insert(0, email_guardado)

    lbl_estado = tk.Label(frm, text="", bg=C_DARK, fg=C_ERROR, font=FONT_BOLD, justify="left", wraplength=320)
    lbl_estado.pack(anchor="w", pady=(4, 10))

    botones = tk.Frame(frm, bg=C_DARK)
    botones.pack(fill="x")
    btn_salir = tk.Button(botones, text="Salir", font=FONT_LABEL, command=lambda: _cerrar(None))
    btn_salir.pack(side="right", padx=(8, 0))
    btn_entrar = tk.Button(botones, text="Entrar", font=FONT_BOLD, bg=C_ACCENT, fg=C_DARK)
    btn_entrar.pack(side="right")
    tk.Label(frm, text="La contraseña se guarda cifrada en este equipo (solo tu usuario de Windows puede leerla).",
             bg=C_DARK, fg=C_MUTED, font=("Segoe UI", 7), wraplength=320, justify="left").pack(anchor="w", pady=(12, 0))

    def _cerrar(valor):
        resultado["valor"] = valor
        win.destroy()

    def _ocupado(si, texto=""):
        estado = "disabled" if si else "normal"
        for w in (e_email, e_pass, btn_entrar):
            w.configure(state=estado)
        lbl_estado.configure(text=texto, fg=C_ACCENT2 if si else C_ERROR)

    def _entrar(email=None, password=None):
        email = (email if email is not None else e_email.get()).strip()
        password = password if password is not None else e_pass.get()
        _ocupado(True, "Comprobando acceso…")
        caja = {}
        hilo = threading.Thread(target=lambda: caja.update(r=comprobar(email, password)), daemon=True)
        hilo.start()

        def _esperar():
            if hilo.is_alive():
                win.after(120, _esperar)
                return
            estado, mensaje, usuario = caja.get("r") or ("error", "Error inesperado al comprobar el acceso.", None)
            if estado == "ok":
                _cerrar((email, password, usuario))
                return
            if estado in ("credenciales", "no_autorizado") and al_fallar_credenciales:
                al_fallar_credenciales()          # que no se vuelvan a probar solas
            _ocupado(False, mensaje)
            e_pass.delete(0, "end")
            (e_pass if e_email.get().strip() else e_email).focus_set()

        win.after(120, _esperar)

    btn_entrar.configure(command=_entrar)
    win.bind("<Return>", lambda _e: _entrar() if str(btn_entrar["state"]) == "normal" else None)
    win.protocol("WM_DELETE_WINDOW", lambda: _cerrar(None))

    win.update_idletasks()
    w, h = win.winfo_reqwidth(), win.winfo_reqheight()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    win.geometry(f"{w}x{h}+{(sw - w) // 2}+{max(60, (sh - h) // 3)}")
    win.lift()
    win.focus_force()
    (e_pass if email_guardado else e_email).focus_set()
    if email_guardado and password_guardada:
        win.after(200, lambda: _entrar(email_guardado, password_guardada))
    win.wait_window()
    return resultado["valor"]
