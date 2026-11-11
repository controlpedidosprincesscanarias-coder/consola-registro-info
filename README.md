# Consola Registro Info SAP (TARIFA Finder)

App de escritorio de la **Central de Compras Princess Canarias** para preparar las tarifas de los proveedores y generar los **Registros de Información de compras de SAP (RINFO)**. También mantiene los listados de artículos DALI y de materiales SAP.

> La idea y el desarrollo de la aplicación, incluidas sus pruebas, son de **Víctor Alejandro Martín Andrade**, para el Departamento de Central de Compras de Princess Hotels en Canarias. Ver `AUTORIA.md`.

**Versión actual:** v 10.2.36 (`APP_VERSION` en `tarifa_app.py`).

## Uso
Se distribuye solo `TARIFA_Finder.exe`, que no necesita Python instalado. Hace falta Excel para algunas acciones.

1. Elige la **carpeta del proyecto** (con `TARIFAS/` y `RECURSOS/`). Se recuerda en `tarifa_config.json`, junto al `.exe`.
2. Usa las **acciones** del panel izquierdo:

| Nº | Acción |
|---|---|
| 1 | Organizar artículos DALI y generar `ARTICULOS_DALI` |
| 2 | Actualizar el listado de artículos DALI |
| 3 | Actualizar el listado de materiales SAP |
| 4 | Generar la TARIFA final y procesar las tarifas incompletas |
| 5 | Abrir la última TARIFA procesada |
| 6 | Actualizar a mano los archivos [INCOMPLETOS] |
| 7 | Listado de tarifa de proveedor vacía (en construcción) |

Desde la tarifa preparada se crea el **Registro Info SAP**, a partir del «Purchasing Info Records» descargado de SAP.

La **consola** de la derecha muestra cada paso con su hora y un color por nivel: `SISTEMA`, `INFO`, `OK`, `WARN` y `ERROR`.

**Catálogo Asignaciones:** la primera vez que hace falta descargar su catálogo, la app pide tu usuario y tu contraseña. La contraseña se guarda **cifrada**: solo tu usuario de Windows, en ese equipo, puede leerla. Si la contraseña deja de valer, se borra y se vuelve a pedir.

## Compilar (Windows)
```bash
pip install pyinstaller pillow pandas openpyxl xlrd pywin32 psutil
```
Después ejecuta `COMPILAR_TARIFA.bat`: pregunta la versión, la escribe en `tarifa_app.py` y llama a `build_exe.py`. El ejecutable queda en `dist/TARIFA_Finder.exe`.

## Pruebas
```bash
python -m unittest discover -s tests -v
```

## Archivos
| Archivo | Contenido |
|---|---|
| `tarifa_app.py` | Punto de arranque y clase de la ventana principal (`APP_VERSION`) |
| `app/principal/` | Métodos de la ventana principal por temas: interfaz, Acciones 1, 2, 3, 4/6 y catálogo DALI |
| `app/rinfo/` | Motor de generación de RINFO: `generador_existente.py`, `creador_rinfo.py`, `ventanas_rinfo.py`, `comun_rinfo.py` y `crear_registro_info.py` (entrada) |
| `app/tarifas/transformar_tarifa.py` | Transformación de la tarifa del proveedor |
| `app/acciones/organizar_dali.py` | Lógica de la Acción 1 (sin interfaz) |
| `app/comun/` | Constantes, configuración, consola, `sap_helpers.py` y `credenciales.py` |
| `app/ui/tarifa_dialogs.py` | Ventanas modales |
| `assets/` | Logos, icono e imágenes de ayuda (van dentro del `.exe`) |
| `tests/` | Pruebas automáticas |
| `build_exe.py`, `COMPILAR_TARIFA.bat`, `_bat_helper.py` | Compilación |
| `HISTORIAL_REORGANIZACION_RINFO.md` | Historial de cambios y próximos pasos |
