# CLAUDE.md — Consola Registro Info SAP (TARIFA Finder)

> **Autoría:** la idea y el desarrollo de la aplicación, incluidas sus pruebas, son de **Víctor Alejandro Martín Andrade**, para el Departamento de Central de Compras de Princess Hotels en Canarias. Ver `AUTORIA.md`.

Instrucciones para Claude Code en este repositorio. Responde siempre en **español**.

## Qué es
App de escritorio para Windows (Python + Tkinter) de la Central de Compras: prepara las tarifas de los proveedores y genera los **Registros de Información de compras de SAP (RINFO)** a partir de la tarifa y del «Purchasing Info Records» descargado de SAP. Mantiene también los listados de artículos DALI y de materiales SAP. Ventana «CREACIÓN REGISTRO INFO SAP»; ejecutable `TARIFA_Finder.exe`.
- **Versión:** `APP_VERSION` en `tarifa_app.py` (formato `"v X.Y.Z"`). La pide y la escribe `COMPILAR_TARIFA.bat` al compilar, con `_bat_helper.py`.
- **Compilación** (en Windows): `COMPILAR_TARIFA.bat` → `build_exe.py` → PyInstaller `--onefile --windowed`, que genera `dist/TARIFA_Finder.exe`. El `.spec` lo genera PyInstaller y no se sube.
- **Datos:** carpeta de proyecto en la unidad de red (`TARIFAS/`, `RECURSOS/`), elegida en la app y guardada en `tarifa_config.json`, junto al `.exe` o al script. Ese archivo **no se sube nunca** (está en `.gitignore`).
- **Catálogo Asignaciones** (repo `dali-sap-articulos-app`): antes de algunas acciones se descarga su catálogo a través del Worker `DALI_APP_API_BASE` (`tarifa_app.py`), con el usuario y la contraseña de cada persona.

## Estructura
- `tarifa_app.py` — ventana principal (`App`) y las acciones del panel:
  1. Organizar artículos DALI;
  2. Actualizar listado DALI;
  3. Actualizar listado SAP;
  4. Generar TARIFA final;
  5. Abrir la última tarifa;
  6. Archivos [INCOMPLETOS];
  7. Tarifa vacía (en construcción).
  
  También contiene la integración con Catálogo Asignaciones y el catálogo DALI completo.
- `crear_registro_info.py` — motor de RINFO: `GeneradorDesdeExistente` (desde un Purchasing existente) y `CreadorRINFO`, con sus ventanas.
- `transformar_tarifa.py` — transforma la tarifa del proveedor al formato de trabajo.
- `acciones/organizar_dali.py` — lógica de la Acción 1, sin tkinter.
- `sap_helpers.py` — funciones de negocio puras (códigos SAP, fechas, filas), sin tkinter, compartidas con las pruebas.
- `credenciales.py` — contraseña de Catálogo Asignaciones cifrada con DPAPI de Windows, sin tkinter.
- `tarifa_core.py` (rutas, configuración, consola de log, barra de progreso), `tarifa_dialogs.py` (ventanas modales), `tarifa_constants.py` (colores, fuentes, tablas fijas), `assets/` (logos e icono).
- **Dependencias** (sin importaciones circulares): `tarifa_constants` → `tarifa_core` → `tarifa_dialogs` / `crear_registro_info` / `transformar_tarifa` → `tarifa_app`. Los módulos sin tkinter (`sap_helpers`, `credenciales`, `acciones/`) no importan de la interfaz. Nada importa de `tarifa_app.py`.

## Seguridad
- **Nunca** escribir contraseñas, tokens, claves ni correos de cuentas en el código, en los commits ni en los documentos. El 9 nov 2026 se subió por error `tarifa_config.json` con la contraseña de Catálogo Asignaciones en claro; desde entonces está en `.gitignore` y la contraseña va cifrada (`credenciales.py`).
- La contraseña de Catálogo Asignaciones solo se guarda cifrada con DPAPI. Si no se puede cifrar, no se guarda. No reintroducir el guardado en claro.
- Datos reales (tarifas, Purchasing, listados `.xls`/`.xlsx`, logs) **nunca** al repositorio.

## Reglas de trabajo
1. Leer el código real antes de proponer cambios y buscar todos los usos de una función antes de modificarla.
2. Trabajar en rama y con PR; **no hacer push a `main` sin confirmación de Víctor**. Commits pequeños con mensaje claro.
3. Los errores que se encuentren se reparan en un **commit propio**, aunque ya estuvieran, sin mezclarlos con código que solo se mueve.
4. **`HISTORIAL_REORGANIZACION_RINFO.md` es el historial de esta app**: se le añade una sección numerada por cada tanda de trabajo, y se mantiene al día «Próximos pasos». Antes de empezar, leer sus secciones «Patrón de trabajo» y «Próximos pasos».
5. Al trocear un método largo, auditar **todas** las variables del bloque extraído (si se usan después). Ya hubo dos fallos así. Muchos procesos corren en hilos sin captura: si la app «se queda colgada», sospechar primero de una excepción silenciosa en el hilo.
6. Sin Windows ni datos reales no se prueba de verdad: compilar, pasar las pruebas y probar la lógica aislada (`object.__new__(Clase)`, ver el historial). Después, decir claramente qué queda por probar con datos reales en Windows.
7. Conservar siempre la línea de autoría en documentos y cabeceras, y añadirla en los nuevos.

## Pruebas
```bash
python -m unittest test_crear_registro_info test_credenciales -v
```
- 103 pruebas. 18 de ellas importan `crear_registro_info` (que usa tkinter): sin tkinter se saltan. En Linux hay que lanzarlas con un Python que tenga tkinter y con `xvfb-run`.
- Dependencias para probar: `pandas openpyxl xlrd pillow psutil`. Para compilar, además `pyinstaller pywin32`.
- `test_crear_registro_info.py` prueba el código real a través de `sap_helpers.py` y de `crear_registro_info`; no se deben reimplementar funciones dentro de la prueba.
