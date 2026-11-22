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
| 1 | Organizar artículos DALI, generar `ARTICULOS_DALI` y, al final, actualizar el catálogo y Catálogo Asignaciones (desde v10.2.64 incluye lo que hacía la antigua Acción 2; antes pregunta si quieres revisar el Excel). Desde v10.2.74 vale `ARTICULOS1.xls` + `ARTICULOS2.xls` o `ARTICULOS.htm` |
| 2 | Actualizar el listado de materiales SAP (desde v10.2.74, antes comprueba si hace falta descargar de nuevo los listados de SAP) |
| 3 | Generar la TARIFA final y procesar las tarifas incompletas |
| 4 | Actualizar a mano los archivos [INCOMPLETOS] (sombreado mientras no haya ninguno en TARIFAS) |
| 5 | Asignar SAP pendientes: todos los códigos DALI activos sin SAP, para grabarlos en el orden que quieras (desde v10.2.50) |
| 6 | Generar plantilla Tarifa desde RINFO: del RINFO o Purchasing, la tarifa que se envía al proveedor para la revisión de precios (desde v10.2.94) |

**Números nuevos (desde v10.2.86, petición de Víctor):** los botones van del 1 al 6 sin huecos. Antes eran 1, 3, 4, 5, 6, 7 y 8: la 3 es ahora la 2, la 4 la 3, la 6 la 4, la 8 la 5 y la 7 la 6; la 5 («Abrir la última TARIFA procesada») se quitó porque no se usaba. La Acción 4 queda sombreada mientras no haya tarifas [INCOMPLETO] en TARIFAS (lo comprueba cada 10 s, `app/comun/incompletos.py`). En este README las acciones ya llevan el número nuevo.

Desde la tarifa preparada se crea el **Registro Info SAP**, a partir del «Purchasing Info Records» descargado de SAP.

La **consola** de la derecha muestra cada paso con su hora y un color por nivel: `SISTEMA`, `INFO`, `OK`, `WARN` y `ERROR`.

**Manual de usuario:** `docs/Manual_Usuario_ConsolaRegistroInfo.pdf`. Dentro de la app, el acceso rápido **Manual de usuario** lo abre y la barra **🔎 Consultar el manual** (encima de la consola) busca la duda que escribas en sus apartados, sin Internet.

**Dónde dejar el `.exe`:** no se instala. Déjalo siempre en la misma carpeta y con su nombre (`TARIFA_Finder.exe`): las actualizaciones lo sustituyen en ese mismo sitio, así que los accesos directos (también el del Organizador) no hay que cambiarlos. También vale una carpeta de red (G:): ahí Windows no deja cambiar el programa mientras está abierto, así que la Consola se cierra, se actualiza y se vuelve a abrir sola en unos segundos.

**Asignar SAP pendientes (Acción 5, desde v10.2.50; tabla editable desde v10.2.56):** abre la tabla de todos los códigos DALI activos que no tienen código SAP, igual que la de «Códigos SAP faltantes» de la Acción 3, con buscador y filtros por proveedor y familia. Escribes el código SAP de cada artículo en su fila (columna «CÓDIGO SAP (introducir)») e Intro pasa a la siguiente; lo escrito se conserva al filtrar. El código DALI, la descripción y el formato se pueden seleccionar y copiar (Ctrl+C o clic derecho), para pegarlos en SAP al crear el material. Al lado se ve lo que dice `Productos.xlsx` de cada código. «Guardar en el catálogo» lo graba (si un código ya lo tiene otro artículo, te pregunta cuál es el correcto) y completa sus datos SAP como la Acción 2; los cambios se envían a Catálogo Asignaciones.

**Faltan los listados de SAP actualizados (desde v10.2.56):** si guardas códigos SAP (Acción 5 o códigos SAP faltantes de la Acción 3) que todavía no están en `Productos.xlsx`, una ventana explica con imágenes cómo descargar de SAP `Productos.xlsx` («Gestionar datos maestros de producto») y `Listas de productos.xlsx` («Lista de productos»): botón ⌄ de exportar → «Exportar como…» → Excel, y guardarlos en RECURSOS. «Ya están en RECURSOS» lanza sola la Acción 2; «Más tarde» avisa de que la web queda parcialmente actualizada hasta entonces. Las imágenes se generan con `compilacion/generar_guia_sap.py`. Un listado cuenta como «descargado de nuevo» cuando su fecha cambia desde que salió el aviso (desde v10.2.76 no se compara con la hora del PC, que en la unidad de red puede no coincidir con la del servidor).

**Acción 1 con `ARTICULOS.htm` (desde v10.2.74):** además de `ARTICULOS1.xls` y `ARTICULOS2.xls`, la Acción 1 acepta `ARTICULOS.htm`: el listado de artículos de DALI exportado como informe, con las 9 columnas en un solo archivo (código, descripción, unidad, familia, subfamilia, naturaleza, impuesto compra, proveedor por defecto y no permitir compra). El resultado es el mismo; además, las descripciones largas salen sin el espacio que DALI mete al exportar a Excel donde partía la línea («PLATANO- MANZANA»). La consola dice cuántos artículos ha leído y si coinciden con el «Total Líneas» del informe. Si en RECURSOS están las dos descargas, se usa **la más reciente** y la consola dice cuál; después se borran las tres, como hasta ahora los `.xls`, para no usar por error una descarga vieja.

**Acción 6 — Generar plantilla Tarifa desde RINFO (desde v10.2.94, petición de Víctor, 21-22 nov 2026):** el paso al revés de «Crear / Actualizar Registro Info SAP». Se arrastra a su ventana (o se elige con «Examinar…») uno o varios RINFO o Purchasing Info Records, y la Consola genera en TARIFAS, sin más preguntas, la tarifa para que el proveedor rellene la nueva revisión de precios: **una por proveedor**, con sus artículos, el precio vigente hoy, el código de proveedor, el teléfono y los hoteles; si los precios cambian según el hotel, **una por cada grupo de hoteles con los mismos precios** (el grupo va en el nombre: «PROVEEDOR (JN-FV) [dd-mm-aaaa].xlsx»). El código DALI, la descripción y el formato salen del catálogo por código SAP. No salen los registros o centros marcados para borrar ni los precios caducados; los escalados llevan el precio base y se avisan. El «Periodo:» y las columnas **PRECIO NUEVO** y **FECHA INICIO** van en amarillo para el proveedor. Cuando la devuelve, se guarda en TARIFAS y se usa tal cual en «Crear / Actualizar Registro Info SAP»: manda el PRECIO NUEVO donde esté relleno y la FECHA INICIO se propone en la ventana de fechas. Nunca se pisa un archivo. El arrastrar y soltar solo funciona en Windows.

**Teléfono del proveedor y hoteles en la TARIFA (desde v10.2.94, petición de Víctor, 22 nov 2026):** la TARIFA final lleva «Teléfono:» en la fila 4 y, a la derecha de la cabecera (H1:L5), el recuadro **HOTELES** con las siglas GY IT LP TA ST MG GC MT JN FV y una casilla para marcar con ✔ (o X) los hoteles a los que se aplica. Salen en blanco. Si están rellenos, «Crear / Actualizar Registro Info SAP» no pregunta el teléfono ni los centros (lo dice en la consola); si falta alguno, sale su ventana de siempre. Las fechas se siguen preguntando y las tarifas antiguas funcionan igual.

**Cambio de nombre del proveedor en DALI (desde v10.2.78, con Catálogo Asignaciones v1.20.72):** si **todos** los artículos de un proveedor pasan a la vez a un nombre que no existía (por ejemplo, «B REFRESHING COMPANY S.L. ( BEBIDAS REFRESCANTES CANARIAS SL )» → «B REFRESHING COMPANY S.L.»), es el mismo proveedor: la Consola conserva su código de proveedor y su precio, y la web renombra el proveedor en vez de crear otro (se conservan precios, códigos, imágenes y fichas). Lo dicen la consola («↻ Proveedor renombrado en DALI…») y la ventana de envío. Si solo cambian algunos artículos, o el nombre nuevo ya existía, es un cambio de proveedor de verdad: el código y el precio del anterior se vacían, como hasta ahora.

**Código de proveedor que la web no acepta (desde v10.2.78):** si la web ya tiene guardado otro código de proveedor para un artículo, no lo cambia; ahora la Consola se queda también con el de la web y lo dice una vez (antes el aviso salía en cada envío). Si el bueno es el tuyo, cámbialo en la ficha del artículo en Catálogo Asignaciones.

**Sin ventanas de resumen repetidas (desde v10.2.78):** al terminar una acción que envía cambios a Catálogo Asignaciones (Acción 1, Acción 2, códigos SAP faltantes de la Acción 3 y Acción 5) ya no sale la ventana «Actualización completada» ni la de «Caso irregular — Proveedor Asignado cambió»: lo que va a cambiar se ve en la ventana **«Cambios para Catálogo Asignaciones»**, y el resumen queda en el log (y los cambios de proveedor, también en el `.txt` de RECURSOS). Solo sale una ventana cuando hay algo que revisar (códigos sin coincidencia en la Acción 2, DALI no encontrados, códigos que no están en `Productos.xlsx`…). Desde v10.2.80, lo mismo al completar una tarifa en la Acción 3 («Reinscripción completada» y «Tarifa completada»): si todo va bien, el resumen y el nombre del archivo van al log; con duplicados o pendientes, sigue saliendo el aviso. Desde v10.2.88 (petición de Víctor, 20 nov 2026) tampoco salen «Procesado correctamente», «Tarifa completada» tras los SAP faltantes, «Actualización SAP completada con advertencias», «Categoría valoración ambigua», «Caso irregular — Proveedor Asignado no coincide», «Códigos proveedor completados desde el catálogo», «Incidencias detectadas en datos de origen», «Sin cambios», «Catálogo Asignaciones sin actualizar», «Web parcialmente actualizada» ni «Catálogo único»: todo va a la consola. Quedan las ventanas que piden decidir algo o avisan de algo que impide seguir.

**Crear Registro Info más rápido y con menos avisos amarillos (v10.2.82):** al generar un RINFO desde un Purchasing existente, las filas absorbidas de la Hoja 3 se borran de una sola pasada (con PRO A PRO, 690 filas salteadas tardaban 94 s; ahora menos de un segundo, con el mismo resultado). Los casos que no piden hacer nada («actualización J omitida: periodo ya cubierto con el mismo precio» y «absorción total») salen en una sola línea informativa con la lista de materiales, en vez de una línea amarilla por material. Siguen en amarillo los que hay que revisar (periodos solapados con otro precio, inicio retrasado, casos irregulares…).

**Acción 2 guiada (desde v10.2.74):** al pulsar la Acción 2, la Consola trae lo cambiado en la web y comprueba sola si hace falta descargar de nuevo `Productos.xlsx` y `Listas de productos.xlsx`: busca los artículos con código SAP y sin sus datos SAP (descripción, unidad, grupo o categoría de valoración), mira si esos códigos ya están en los listados de RECURSOS y avisa si algún listado tiene más de 7 días. Una ventana muestra la fecha de cada listado y el resultado. Si falta algo, o un listado es antiguo, sale la guía de descarga con «¿Necesita descargar de nuevo…?» **Sí / No**: con «Sí» se abre RECURSOS y una línea dice, cada 2 segundos, si cada listado se ha descargado de nuevo y si ya trae los códigos que faltaban. El botón **«Actualizar Catálogo Asignaciones»** está oculto hasta contestar (como al final de la Acción 1); si no hace falta preguntar, sale directamente. Al pulsarlo se hace la Acción 2 de siempre y el envío a la web.

**Un solo archivo de artículos (desde v10.2.50):** la Consola ya no usa `LISTADO CODIGOS DALI - SAP.xlsx`: todo está en `RECURSOS/catalogo-dali-completo.xlsx`. El final de la Acción 1 actualiza en él los datos DALI (antes, la Acción 2), la Acción 2 los datos SAP y la ventana de códigos SAP faltantes escribe ahí el código SAP; las tarifas y los Registros Info leen de él. Todo se empareja por código DALI. La primera vez que se abre esta versión, el LISTADO que haya se pasa al catálogo (con un informe de diferencias en RECURSOS) y se renombra a «LISTADO CODIGOS DALI - SAP (antiguo, ya no se usa).xlsx»; no se borra.

**Inicio de sesión (desde v10.2.50):** al abrir la Consola se pide el usuario y la contraseña de **Catálogo Asignaciones**. Solo entra una cuenta activa, de administrador y con el permiso «Escribir desde la Consola Registro Info» (el administrador principal también); con cualquier otra sale «Usuario no autorizado.». Si las credenciales están guardadas en el equipo, se comprueban solas al abrir. El aviso de versión nueva sale igual aunque no se haya podido entrar.

**Catálogo Asignaciones:** la primera vez que hace falta descargar su catálogo, la app pide tu usuario y tu contraseña. La contraseña se guarda **cifrada**: solo tu usuario de Windows, en ese equipo, puede leerla. Si la contraseña deja de valer, se borra y se vuelve a pedir.

**Cambios en Catálogo Asignaciones (desde v10.2.50):** al final de la Acción 1, después de las Acciones 2 y 3 (y de resolver un código SAP duplicado), la app envía a la web **solo lo que ha cambiado**: altas y datos DALI, datos SAP, y precio y código de proveedor de la tarifa. Antes de cambiar nada muestra un resumen, y lo que la web no puede aplicar (otro proveedor asignado, otro código de proveedor) sale para corregirlo. Si eliges «Ahora no» o no hay conexión, los cambios vuelven a salir la próxima vez. Hace falta el permiso «Escribir desde la Consola Registro Info» en Catálogo Asignaciones. El código DALI y su descripción solo se cargan desde aquí. El Excel `catalogo-dali-completo.xlsx` se sigue escribiendo igual mientras se comprueba que los dos coinciden (fase 1).

**Lo cambiado en la web llega al Excel (desde v10.2.64):** el código SAP, su descripción y el código de proveedor que se cambien en Catálogo Asignaciones bajan solos a `catalogo-dali-completo.xlsx` antes de cada envío y al empezar las Acciones 2, 3 y 5. En cada dato (datos SAP, precio y código de proveedor) gana la última modificación, en la web o en la Consola; los datos DALI siempre los manda la Consola. Con el Excel abierto no se puede copiar: se avisa, y esos datos tampoco se envían, para no pisar los de la web.

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
| `app/principal/` | Métodos de la ventana principal por temas: interfaz, Acciones 1 (con el paso al catálogo, `accion_listado_dali.py`), 3, 4/6 y catálogo DALI |
| `app/rinfo/` | Motor de generación de RINFO: `generador_existente.py`, `creador_rinfo.py`, `ventanas_rinfo.py`, `comun_rinfo.py` y `crear_registro_info.py` (entrada) |
| `app/tarifas/transformar_tarifa.py` | Transformación de la tarifa del proveedor |
| `app/acciones/organizar_dali.py` | Lógica de la Acción 1 (sin interfaz) |
| `app/comun/` | Constantes, configuración, consola, `sap_helpers.py`, `credenciales.py`, `actualizaciones.py` y `manual.py` (abrir y consultar el manual) |
| `app/ui/` | Ventanas modales, aviso de actualización y «Consultar el manual» |
| `assets/` | Logos, icono, imágenes de ayuda y el manual con su índice (van dentro del `.exe`) |
| `docs/` | Manual de usuario (`.docx` y `.pdf`) y su fuente (`docs/manuales_fuente/`, ver su README) |
| `tests/` | Pruebas automáticas |
| `app/comun/asignar_sap.py`, `app/ui/ventana_asignar_sap.py`, `app/principal/accion_asignar_sap.py` | Acción 5: Asignar SAP pendientes (v10.2.50) |
| `app/comun/volcado_dali.py`, `app/ui/ventana_volcado_dali.py` | Final de la Acción 1: revisar `ARTICULOS_DALI.xlsx` y actualizar Catálogo Asignaciones (v10.2.64) |
| `app/comun/catalogo_local.py` | El catálogo completo como único archivo de artículos (v10.2.50) |
| `app/comun/catalogo_web.py`, `app/principal/publicar_catalogo_web.py`, `app/ui/ventana_publicar_catalogo.py` | Envío de los cambios a Catálogo Asignaciones (v10.2.50) |
| `build_exe.py`, `compilacion/`, `.github/workflows/compilar.yml` | Compilación (en GitHub, con versión automática) |
| `COMPILAR_TARIFA.bat`, `_bat_helper.py` | Compilación local en Windows |
| `HISTORIAL_REORGANIZACION_RINFO.md` | Historial de cambios y próximos pasos |
