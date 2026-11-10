# -*- coding: utf-8 -*-
"""
credenciales.py
Usuario y contraseña de «Catálogo Asignaciones» guardados CIFRADOS en tarifa_config.json con
DPAPI de Windows (CryptProtectData): solo el mismo usuario de Windows, en el mismo equipo, puede
descifrarlos. No hace falta ninguna clave guardada junto a los datos ni librerías extra (ctypes).
Es lo mismo que hace el Organizador Princess con la contraseña de Control de Pedidos.

Antes (hasta noviembre de 2026) la contraseña se guardaba en claro en tarifa_config.json, y ese
archivo llegó a subirse a GitHub. Ahora:
  - la contraseña se guarda en «dali_app_password_cifrada» con el prefijo «dpapi1:»;
  - si se encuentra una contraseña antigua en claro («dali_app_password»), se cifra y se borra
    la copia en claro (migrar_credenciales_dali);
  - si no se puede cifrar (no es Windows o DPAPI falla), la contraseña NO se guarda: se usa solo
    durante la sesión y se vuelve a pedir la próxima vez. Nunca se escribe en claro.

Sin tkinter: se puede probar con test_credenciales.py en cualquier sistema.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import base64
import sys

PREFIJO = "dpapi1:"
CLAVE_EMAIL = "dali_app_email"
CLAVE_CIFRADA = "dali_app_password_cifrada"
CLAVE_ANTIGUA = "dali_app_password"          # en claro: solo se lee para migrarla y se borra


def dpapi_disponible() -> bool:
    return sys.platform.startswith("win")


def _crypt(datos: bytes, cifrar: bool) -> bytes:
    import ctypes
    from ctypes import wintypes

    class _BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    fn = crypt32.CryptProtectData if cifrar else crypt32.CryptUnprotectData
    fn.argtypes = [ctypes.POINTER(_BLOB), ctypes.c_void_p, ctypes.POINTER(_BLOB),
                   ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_BLOB)]
    fn.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p

    buf = ctypes.create_string_buffer(datos, len(datos))
    entrada = _BLOB(len(datos), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    salida = _BLOB()
    if not fn(ctypes.byref(entrada), None, None, None, None, 0, ctypes.byref(salida)):
        raise OSError(ctypes.get_last_error(), "DPAPI falló")
    try:
        return ctypes.string_at(salida.pbData, salida.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(salida.pbData, ctypes.c_void_p))


def proteger(texto: str):
    """Texto cifrado con el prefijo «dpapi1:», o None si no se puede cifrar."""
    if not texto or not dpapi_disponible():
        return None
    try:
        return PREFIJO + base64.b64encode(_crypt(texto.encode("utf-8"), True)).decode("ascii")
    except Exception:  # noqa: BLE001
        return None


def desproteger(valor: str):
    """Texto descifrado, o None si no se puede (otro equipo u otro usuario de Windows)."""
    if not valor or not str(valor).startswith(PREFIJO) or not dpapi_disponible():
        return None
    try:
        return _crypt(base64.b64decode(valor[len(PREFIJO):]), False).decode("utf-8")
    except Exception:  # noqa: BLE001
        return None


def leer_credenciales_dali(config: dict):
    """(email, contraseña) guardados, o (None, None) si no hay o no se pueden descifrar."""
    email = (config or {}).get(CLAVE_EMAIL)
    clave = desproteger((config or {}).get(CLAVE_CIFRADA))
    if email and clave:
        return email, clave
    return None, None


def guardar_credenciales_dali(config: dict, email: str, clave: str) -> bool:
    """Guarda el email y la contraseña CIFRADA en `config` (hay que llamar después a
    guardar_config). Devuelve False si no se pudo cifrar: entonces no se guarda la contraseña."""
    config.pop(CLAVE_ANTIGUA, None)
    cifrada = proteger(clave)
    if not cifrada:
        config.pop(CLAVE_CIFRADA, None)
        return False
    config[CLAVE_EMAIL] = email
    config[CLAVE_CIFRADA] = cifrada
    return True


def borrar_credenciales_dali(config: dict):
    for k in (CLAVE_EMAIL, CLAVE_CIFRADA, CLAVE_ANTIGUA):
        config.pop(k, None)


def migrar_credenciales_dali(config: dict) -> bool:
    """Si hay una contraseña antigua en claro, la cifra (si se puede) y borra la copia en claro.
    Devuelve True si `config` ha cambiado y hay que guardarlo."""
    antigua = (config or {}).get(CLAVE_ANTIGUA)
    if antigua is None:
        return False
    email = config.get(CLAVE_EMAIL)
    if not (email and antigua and guardar_credenciales_dali(config, email, antigua)):
        borrar_credenciales_dali(config)
    return True
