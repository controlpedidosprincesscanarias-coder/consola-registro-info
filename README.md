# Consola Registro Info SAP (TARIFA Finder)

App de escritorio de la **Central de Compras Princess Canarias** para preparar las tarifas de los proveedores y generar los **Registros de Información de compras de SAP (RINFO)**. También mantiene los listados de artículos DALI y de materiales SAP.

> La idea y el desarrollo de la aplicación, incluidas sus pruebas, son de **Víctor Alejandro Martín Andrade**, para el Departamento de Central de Compras de Princess Hotels en Canarias. Ver `AUTORIA.md`.

**Versión:** la pone sola cada compilación en GitHub (`compilacion/version.py`). La de cada `.exe` es la etiqueta de su Release: la última publicada está en la página **Releases**. La v10.2.38 (borrador) es la validada con datos reales; la primera que se actualiza sola es la v10.2.40.

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
| 8 | Asignar SAP pendientes: todos los códigos DALI activos sin SAP, para grabarlos en el orden que quieras (desde v10.2.50) |

Desde la tarifa preparada se crea el **Registro Info SAP**, a partir del «Purchasing Info Records» descargado de SAP.

La **consola** de la derecha muestra cada paso con su hora y un color por nivel: `SISTEMA`, `INFO`, `OK`, `WARN` y `ERROR`.

**Manual de usuario:** `docs/Manual_Usuario_ConsolaRegistroInfo.pdf`. Dentro de la app, el acceso rápido **Manual de usuario** lo abre y la barra **🔎 Consultar el manual** (encima de la consola) busca la duda que escribas en sus apartados, sin Internet.

**Dónde dejar el `.exe`:** no se instala. Déjalo siempre en la misma carpeta y con su nombre (`TARIFA_Finder.exe`): las actualizaciones lo sustituyen en ese mismo sitio, así que los accesos directos (también el del Organizador) no hay que cambiarlos.

**Asignar SAP pendientes (Acción 8, desde v10.2.50):** abre la lista de todos los códigos DALI activos que no tienen código SAP, con buscador y filtros por proveedor y familia. Eliges uno, escribes su código SAP y pulsas Intro: se asigna y pasa al siguiente. Al lado se ve lo que dice `Productos.xlsx` de ese código. «Guardar en el catálogo» lo graba (si un código ya lo tiene otro artículo, te pregunta cuál es el correcto) y completa sus datos SAP como la Acción 3; los cambios se envían a Catálogo Asignaciones. Un código que no esté en `Productos.xlsx` se guarda igualmente: para completar sus datos, descarga de SAP `Productos.xlsx` y `Listas de productos.xlsx` actualizados, guárdalos en RECURSOS y pulsa de nuevo la Acción 3.

**Un solo archivo de artículos (desde v10.2.50):** la Consola ya no usa `LISTADO CODIGOS DALI - SAP.xlsx`: todo está en `RECURSOS/catalogo-dali-completo.xlsx`. La Acción 2 actualiza en él los datos DALI, la Acción 3 los datos SAP y la ventana de códigos SAP faltantes escribe ahí el código SAP; las tarifas y los Registros Info leen de él. Todo se empareja por código DALI. La primera vez que se abre esta versión, el LISTADO que haya se pasa al catálogo (con un informe de diferencias en RECURSOS) y se renombra a «LISTADO CODIGOS DALI - SAP (antiguo, ya no se usa).xlsx»; no se borra.

**Inicio de sesión (desde v10.2.50):** al abrir la Consola se pide el usuario y la contraseña de **Catálogo Asignaciones**. Solo entra una cuenta activa, de administrador y con el permiso «Escribir desde la Consola Registro Info» (el administrador principal también); con cualquier otra sale «Usuario no autorizado.». Si las credenciales están guardadas en el equipo, se comprueban solas al abrir. El aviso de versión nueva sale igual aunque no se haya podido entrar.

**Catálogo Asignaciones:** la primera vez que hace falta descargar su catálogo, la app pide tu usuario y tu contraseña. La contraseña se guarda **cifrada**: solo tu usuario de Windows, en ese equipo, puede leerla. Si la contraseña deja de valer, se borra y se vuelve a pedir.

**Cambios en Catálogo Asignaciones (desde v10.2.50):** después de las Acciones 2, 3 y 4 (y de resolver un código SAP duplicado), la app envía a la web **solo lo que ha cambiado**: altas y datos DALI, datos SAP, y precio y código de proveedor de la tarifa. Antes de cambiar nada muestra un resumen; lo que se haya cambiado en la web desde tu último envío queda fuera salvo que marques la casilla, y lo que la web no puede aplicar (otro proveedor asignado, otro código de proveedor) sale para corregirlo. Si eliges «Ahora no» o no hay conexión, los cambios vuelven a salir la próxima vez. Hace falta el permiso «Escribir desde la Consola Registro Info» en Catálogo Asignaciones. El código DALI y su descripción solo se cargan desde aquí. El Excel `catalogo-dali-completo.xlsx` se sigue escribiendo igual mientras se comprueba que los dos coinciden (fase 1).

## Actualizaciones
Al abrirse, la app comprueba si hay una versión nueva publicada en GitHub. Si la hay, muestra las novedades y ofrece **Actualizar ahora**: descarga el `.exe` nuevo, comprueba que es el publicado (SHA-256), lo pone en el sitio del actual y se vuelve a abrir sola. La carpeta del proyecto y la contraseña guardada no se tocan. La versión anterior se borra sola en el siguiente arranque, y la consola lo indica.

## Compilar
**En GitHub (lo normal):**
1. Ve a la pestaña **Actions** → **Compilar Consola Registro Info** → **Run workflow**.
2. Elige el destino:
   - **solo-descarga**: el `.exe` queda en la propia ejecución, en «Artifacts».
   - **borrador**: además crea una Release en borrador, que se publica a mano.
3. La versión sale sola: la última más 2 (10.2.36 → 10.2.38). Se puede forzar otra en el campo «version».

Antes de compilar se pasan todas las pruebas.

**En el propio equipo (Windows):**
```bash
pip install -r compilacion/requisitos.txt
```
Después ejecuta `COMPILAR_TARIFA.bat`. El ejecutable queda en `dist/TARIFA_Finder.exe`.

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
| `app/comun/` | Constantes, configuración, consola, `sap_helpers.py`, `credenciales.py`, `actualizaciones.py` y `manual.py` (abrir y consultar el manual) |
| `app/ui/` | Ventanas modales, aviso de actualización y «Consultar el manual» |
| `assets/` | Logos, icono, imágenes de ayuda y el manual con su índice (van dentro del `.exe`) |
| `docs/` | Manual de usuario (`.docx` y `.pdf`) y su fuente (`docs/manuales_fuente/`, ver su README) |
| `tests/` | Pruebas automáticas |
| `app/comun/asignar_sap.py`, `app/ui/ventana_asignar_sap.py`, `app/principal/accion_asignar_sap.py` | Acción 8: Asignar SAP pendientes (v10.2.50) |
| `app/comun/catalogo_local.py` | El catálogo completo como único archivo de artículos (v10.2.50) |
| `app/comun/catalogo_web.py`, `app/principal/publicar_catalogo_web.py`, `app/ui/ventana_publicar_catalogo.py` | Envío de los cambios a Catálogo Asignaciones (v10.2.50) |
| `build_exe.py`, `compilacion/`, `.github/workflows/compilar.yml` | Compilación (en GitHub, con versión automática) |
| `COMPILAR_TARIFA.bat`, `_bat_helper.py` | Compilación local en Windows |
| `HISTORIAL_REORGANIZACION_RINFO.md` | Historial de cambios y próximos pasos |
