// Manual de usuario de la Consola Registro Info SAP (TARIFA Finder).
// La idea y el desarrollo de la aplicación, incluidas sus pruebas, son de Víctor Alejandro Martín Andrade,
// para el Departamento de Central de Compras de Princess Hotels en Canarias.
const V = process.env.VERSION || "v 10.2.52";
const L = require("./lib.js");
const { p, vineta, pasos, h1, h2, imagen, leyenda, recuadro, tabla, flujo, portada, indice, generar } = L;
const fecha = new Date().toLocaleDateString("es-ES", { day: "numeric", month: "long", year: "numeric" });
const s = [];

s.push(...portada({ rol: "Compradores", subtitulo: "Tarifas de proveedor, catálogo DALI / SAP y Registros Info", version: V, fecha,
  publico: "Para el personal de la Central de Compras que prepara las tarifas y los Registros Info de SAP" }));
s.push(...indice([
  ["1. Qué es la Consola Registro Info SAP", "Para qué sirve y cómo es el trabajo de principio a fin"],
  ["2. Ponerla en marcha", "Dónde se guarda, inicio de sesión y actualizaciones"],
  ["3. La ventana principal", "Panel de acciones, consola, consulta del manual y accesos rápidos"],
  ["4. Preparar el proyecto", "Carpetas TARIFAS y RECURSOS y qué archivos van en cada una"],
  ["5. El catálogo DALI / SAP (Acciones 1, 2, 3 y 8)", "ARTICULOS_DALI, catálogo completo y códigos SAP pendientes"],
  ["6. Generar la TARIFA (Acciones 4, 5 y 6)", "De la tarifa del proveedor a la TARIFA final; códigos SAP faltantes y duplicados"],
  ["7. Crear el Registro Info (RINFO)", "Con un Purchasing Info Records existente o desde cero, y cargarlo en SAP"],
  ["8. Cambios en Catálogo Asignaciones", "Qué se envía a la web y cómo revisarlo antes"],
  ["9. Consultar el manual y actualizaciones", "La ayuda dentro de la app y las versiones nuevas"],
  ["10. Consulta rápida y problemas frecuentes", "Tablas de referencia"],
]));

// ── 1 ──────────────────────────────────────────────────────────────────────
s.push(h1("1. Qué es la Consola Registro Info SAP"));
s.push(p("La **Consola Registro Info SAP** (el programa **TARIFA_Finder.exe**) prepara las tarifas que envían los proveedores y genera con ellas los **Registros de Información de compras de SAP (RINFO)**, listos para importarlos en SAP. También mantiene al día el **catálogo de artículos**, que relaciona cada artículo de **DALI** con su material de **SAP**, y envía sus cambios a la web **Catálogo Asignaciones**."));
s.push(p("Todo se hace dentro de una **carpeta de proyecto** (en la unidad de red), con dos subcarpetas: **TARIFAS** (tarifas y Registros Info) y **RECURSOS** (el catálogo y los listados exportados de DALI y de SAP). La consola de la derecha va contando cada paso, con su hora y un color según el resultado."));
s.push(h2("El trabajo de principio a fin"));
s.push(...flujo([["Catálogo al día", "Acciones 1, 2, 3 y 8"], ["Tarifa del proveedor (.xls)", "a TARIFAS/"],
  ["TARIFA final", "Acción 4"], ["Registro Info", "Crear / Actualizar"], ["Importar en SAP", "en SAP S/4HANA"]]));
s.push(...tabla(["Acción", "Para qué sirve", "Cuándo usarla"], [
  ["1. Organizar artículos DALI", "Junta ARTICULOS1.xls y ARTICULOS2.xls (exportados de DALI) en ARTICULOS_DALI.xlsx", "Cuando hay artículos nuevos en DALI"],
  ["2. Actualizar listado artículos DALI", "Pasa los datos de ARTICULOS_DALI.xlsx al catálogo completo, código a código", "Después de la Acción 1"],
  ["3. Actualizar listado materiales SAP", "Completa los datos SAP del catálogo con Productos.xlsx y Listas de productos.xlsx", "Cuando hay materiales nuevos en SAP"],
  ["4. Generar TARIFA final", "Convierte la tarifa del proveedor en la TARIFA final con los códigos SAP", "Con cada tarifa nueva"],
  ["5. Abrir última TARIFA", "Abre en Excel la última TARIFA procesada", "Para revisarla"],
  ["6. Archivos [INCOMPLETOS]", "Abre en Excel una tarifa incompleta para corregirla a mano", "Si una tarifa queda [INCOMPLETO]"],
  ["7. Tarifa de proveedor vacía", "En preparación («PRÓXIMO»)", "—"],
  ["8. Asignar SAP pendientes", "Lista todos los artículos DALI activos sin código SAP para escribir sus códigos", "Cuando haya artículos sin código SAP"],
  ["Crear / Actualizar Registro Info", "Genera el RINFO de una TARIFA final (acceso rápido)", "Con la TARIFA ya completa"],
], [3000, 4300, 2338]));

// ── 2 ──────────────────────────────────────────────────────────────────────
s.push(h1("2. Ponerla en marcha"));
s.push(p("La consola **no se instala**: es un único archivo, **TARIFA_Finder.exe**, que funciona sin instalar Python ni nada más. Cada persona tiene su propia copia en su PC."));
s.push(...pasos(["Copia **TARIFA_Finder.exe** en una carpeta de tu PC (por ejemplo, el Escritorio o «Documentos\\Consola»). Mejor siempre la misma: ahí se guarda también tu configuración.",
  "Ábrelo con doble clic. La primera vez, Windows puede mostrar «Windows protegió su PC» porque el programa no está firmado: pulsa **Más información** → **Ejecutar de todas formas**.",
  "**Inicia sesión** con tu usuario de Catálogo Asignaciones (ver abajo).",
  "Elige la **carpeta del proyecto** (o créala, capítulo 4). La consola la recuerda para las próximas veces.",
  "Para tenerlo a mano, crea un acceso directo: clic derecho sobre el .exe → **Mostrar más opciones** → **Enviar a** → **Escritorio (crear acceso directo)**."]));
s.push(h2("Inicio de sesión"));
s.push(p("Al abrirse, la consola pide tu **usuario y contraseña de Catálogo Asignaciones** (los mismos de la web). Hasta que no entras, la ventana principal no aparece."));
s.push(...imagen("c_login", 8, "Inicio de sesión (datos de ejemplo)."));
s.push(...leyenda([[1, "Tu **email** de Catálogo Asignaciones."], [2, "Tu **contraseña**."], [3, "**Entrar**. Si todo está bien, se abre la ventana principal."], [4, "**Salir** cierra la consola."]]));
s.push(...recuadro("importante", "Quién puede entrar", ["Solo entra una cuenta de Catálogo Asignaciones **activa**, de **administrador** y con el permiso **«Escribir desde la Consola Registro Info»** (el administrador principal de la web también). Con cualquier otra cuenta sale **«Usuario no autorizado.»**: pide el permiso al administrador principal de Catálogo Asignaciones.",
  "Si el email o la contraseña no son correctos, sale «Email o contraseña incorrectos.». Si la web pide un código de verificación (pasa tras mucho tiempo sin entrar), inicia sesión una vez desde el navegador y vuelve a intentarlo."]));
s.push(p("La contraseña se guarda **cifrada** en tu PC: las siguientes veces la consola entra sola, sin preguntar. Si deja de valer (o tu cuenta pierde el permiso), se borra y te la vuelve a pedir. Una vez dentro, la descarga del catálogo y el envío de cambios a la web usan esta misma sesión."));
s.push(...recuadro("nota", "Qué se guarda junto al programa", ["Junto a TARIFA_Finder.exe se crea **tarifa_config.json** con la carpeta del proyecto y tu contraseña de Catálogo Asignaciones **cifrada** (solo tu usuario de Windows, en ese PC, puede leerla). No lo borres ni lo compartas: si lo borras, la consola vuelve a preguntar la carpeta y la contraseña.",
  "Si mueves el .exe a otra carpeta, mueve también tarifa_config.json."]));
s.push(...recuadro("consejo", "Accesos directos que no hay que cambiar", ["Deja el .exe **siempre en la misma carpeta y con su nombre, TARIFA_Finder.exe**. Las actualizaciones ponen la versión nueva en ese mismo sitio y con ese mismo nombre, así que el acceso directo del Escritorio, el de la barra de tareas o el del **Organizador Princess** siguen valiendo sin tocarlos.",
  "Solo habría que cambiarlos si descargas una copia nueva a mano en otra carpeta o le cambias el nombre."]));
s.push(h2("Actualizaciones"));
s.push(p("Al abrirse, la consola comprueba si hay una **versión nueva** publicada. Si la hay, te enseña las novedades y te ofrece **Actualizar ahora**: descarga la versión nueva, comprueba que es la oficial, la pone en el sitio de la actual y se vuelve a abrir sola. El aviso sale aunque todavía no hayas iniciado sesión. Lo verás en el capítulo 9."));

// ── 3 ──────────────────────────────────────────────────────────────────────
s.push(h1("3. La ventana principal"));
s.push(...imagen("c_principal", 17, "Ventana principal de la consola (datos de ejemplo)."));
s.push(...leyenda([[1, "**Carpeta del proyecto**: la ruta y el botón 📁 para elegirla."],
  [2, "**TARIFAS ✓ / RECURSOS ✓**: indican si existen las dos subcarpetas. **RAÍZ** abre la carpeta del proyecto en el Explorador."],
  [3, "**Crear estructura PROYECTO**: crea las carpetas si empiezas de cero (capítulo 4)."],
  [4, "**Estado de TARIFA.xls**: si hay una tarifa de proveedor esperando en TARIFAS. **Buscar** la localiza."],
  [5, "**Progreso** de la acción en curso."],
  [6, "**Acciones 1 a 8** (capítulos 5 y 6)."],
  [7, "**Consultar el manual**: escribe tu duda y pulsa Intro (capítulo 9)."],
  [8, "**Consola**: cada paso con su hora. Colores: SISTEMA (dorado), INFO (azul), OK (verde), WARN (amarillo, revisar) y ERROR (rojo)."],
  [9, "**guardar log** guarda la consola en un .txt; **limpiar** la vacía."],
  [10, "**Accesos rápidos**: portal de **SAP S/4HANA**, **Crear / Actualizar Registro Info SAP** (capítulo 7) y **Manual de usuario** (abre este PDF)."],
  [11, "**Versión** instalada. Púlsala para ver las **novedades** de esa versión."]]));
s.push(...recuadro("consejo", "Lee la consola", ["Las líneas **WARN** (amarillas) no son errores, pero piden que revises algo: códigos SAP que faltan, duplicados, proveedores distintos… Si algo «se queda parado», mira la última línea de la consola antes de volver a pulsar."]));

// ── 4 ──────────────────────────────────────────────────────────────────────
s.push(h1("4. Preparar el proyecto"));
s.push(p("Si es la primera vez, pulsa **Crear estructura PROYECTO**: crea la carpeta PROYECTO con TARIFAS y RECURSOS. Si ya existe, no se toca nada de lo que haya dentro."));
s.push(...imagen("c_crear_proyecto", 10, "Crear la estructura de carpetas."));
s.push(...leyenda([[1, "**Crear carpetas**. Después, la consola la deja elegida como carpeta del proyecto."]]));
s.push(h2("Qué archivos van en cada carpeta"));
s.push(...tabla(["Archivo", "Carpeta", "De dónde sale"], [
  ["ARTICULOS1.xls y ARTICULOS2.xls", "RECURSOS", "Exportados de DALI (los usa la Acción 1 y luego se borran)"],
  ["ARTICULOS_DALI.xlsx", "RECURSOS", "Lo genera la Acción 1"],
  ["Productos.xlsx", "RECURSOS", "Exportado de SAP (lista de materiales, transacción ME3N o ME2M)"],
  ["Listas de productos.xlsx", "RECURSOS", "Exportado de SAP (categoría de valoración por centro)"],
  ["Proveedores.xlsx", "RECURSOS", "Proveedores de SAP con su código (para el Registro Info)"],
  ["catalogo-dali-completo.xlsx", "RECURSOS", "El catálogo de artículos DALI / SAP. Lo mantienen las Acciones 2, 3 y 8 y la ventana de códigos SAP faltantes (se puede descargar de Catálogo Asignaciones)"],
  ["Tarifa del proveedor (.xls)", "TARIFAS", "La que envía el proveedor"],
  ["TARIFA final «PROVEEDOR [dd-mm-aaaa].xlsx»", "TARIFAS", "La genera la Acción 4"],
  ["Purchasing Info Records.xlsx", "TARIFAS", "Descargado de SAP, si el proveedor ya tiene Registros Info"],
  ["RINFO «PROVEEDOR [dd-mm-aaaa].xlsx»", "TARIFAS", "El Registro Info que genera la consola"],
], [3600, 1700, 4338]));
s.push(...recuadro("importante", "Archivos abiertos en Excel", ["Cierra en Excel los archivos que vaya a escribir la consola (el catálogo, la tarifa…) antes de lanzar una acción. Si un archivo está abierto, Windows no deja guardarlo y la consola lo indicará en rojo."]));

// ── 5 ──────────────────────────────────────────────────────────────────────
s.push(h1("5. El catálogo DALI / SAP (Acciones 1, 2, 3 y 8)"));
s.push(p("El **catálogo completo** (**RECURSOS/catalogo-dali-completo.xlsx**) relaciona cada código DALI con su material SAP (descripción, unidad, IGIC, grupo de compras…), su proveedor, su código de proveedor y su precio. Es la base de la TARIFA y del Registro Info, así que conviene tenerlo al día. Todo se empareja siempre **por código DALI**."));
s.push(...recuadro("nota", "Adiós al LISTADO CODIGOS DALI - SAP", ["Hasta noviembre de 2026 había además un **LISTADO CODIGOS DALI - SAP.xlsx** con casi los mismos datos. Ya no se usa: todo está en el catálogo.",
  "La primera vez que se abre la versión nueva, la consola pasa al catálogo los datos SAP del LISTADO, deja un informe con las diferencias que hubiera (**RECURSOS/paso_listado_a_catalogo_….txt**) y le cambia el nombre a «LISTADO CODIGOS DALI - SAP (antiguo, ya no se usa).xlsx». No se borra: puedes guardarlo como copia."]));
s.push(h2("Acción 1 — Organizar artículos DALI"));
s.push(...pasos(["Exporta de DALI los artículos en **ARTICULOS1.xls** y **ARTICULOS2.xls** y déjalos en **RECURSOS**.",
  "Pulsa **1 · Organizar artículos DALI**. La consola los cruza por código, quita duplicados y genera **ARTICULOS_DALI.xlsx**.",
  "Si un código aparece repetido con datos distintos, te avisa y guarda el detalle en RECURSOS (INCIDENCIAS_DUPLICADOS_DALI…). Los .xls de origen se borran al terminar."]));
s.push(h2("Acción 2 — Actualizar listado artículos DALI"));
s.push(p("Pasa al catálogo los datos de DALI de **ARTICULOS_DALI.xlsx** (descripción, unidad, familia, proveedor asignado, activo…), código a código. Los artículos nuevos de DALI se añaden. El **código SAP**, el **código de proveedor** y el **precio** de cada artículo se conservan; si un artículo cambia de proveedor asignado, su código de proveedor y su precio se vacían y la consola lo avisa como «CASO IRREGULAR»."));
s.push(h2("Acción 3 — Actualizar listado materiales SAP"));
s.push(p("Lee **Productos.xlsx** y **Listas de productos.xlsx** de RECURSOS y vuelve a rellenar los datos SAP del catálogo (descripción, unidad, grupo de productos, categoría de valoración, IGIC, grupo de compras…) de cada artículo que tenga código SAP. Si un material tiene categorías distintas según el centro, toma la más frecuente y lo dice en la consola. Si un mismo código SAP está en dos artículos, te pregunta cuál es el correcto."));
s.push(h2("Descargar el catálogo de Catálogo Asignaciones"));
s.push(p("Antes de las Acciones 2 y 3, y al cerrar una TARIFA, la consola puede preguntarte si quieres **descargar** antes el catálogo actualizado desde la web (por si otros compañeros han cambiado algo). Usa la sesión con la que entraste. Si en tu PC hay códigos SAP que todavía no se han enviado a la web, se conservan."));
s.push(h2("Acción 8 — Asignar SAP pendientes"));
s.push(p("Abre la lista de **todos los artículos DALI activos que no tienen código SAP**, para que vayas escribiendo sus códigos en el orden que quieras."));
s.push(...imagen("c_asignar_sap", 17, "Asignar SAP pendientes (datos de ejemplo)."));
s.push(...leyenda([[1, "**Buscar** (código, descripción, proveedor o familia, sin importar mayúsculas ni acentos) y filtros por **Proveedor** y **Familia**. **Quitar filtros** vuelve a la lista completa."],
  [2, "Los artículos pendientes. Los que ya tienen código asignado salen en **verde** si el código está en Productos.xlsx y en **naranja** si no."],
  [3, "Elige un artículo, escribe aquí su **código SAP** y pulsa **Intro** (o **Asignar**): se asigna y pasa al siguiente sin código. **Quitar** deshace la asignación del elegido."],
  [4, "Lo que dice **Productos.xlsx** del código que escribes, para que compruebes que es el correcto."],
  [5, "**Guardar en el catálogo** lo asignado. Las filas guardadas desaparecen y puedes seguir con las demás."]]));
s.push(...pasos(["Pulsa **8 · Asignar SAP pendientes**.",
  "Busca o filtra el artículo, escribe su código SAP y pulsa **Intro**. Repite con los que quieras.",
  "Pulsa **Guardar en el catálogo**. Si un código ya lo tiene otro artículo, te pregunta cuál es el correcto (la misma ventana que en el capítulo 6).",
  "La consola completa sola los datos SAP de esos artículos con Productos.xlsx y Listas de productos.xlsx (como la Acción 3) y envía los cambios a Catálogo Asignaciones (capítulo 8)."]));
s.push(...recuadro("importante", "Si un código no está en Productos.xlsx", ["Se guarda igualmente (sale en naranja). Para completar sus datos SAP:",
  "1) descarga de SAP **Productos.xlsx** y **Listas de productos.xlsx** actualizados; 2) guárdalos en **RECURSOS**; 3) pulsa de nuevo la **Acción 3**."]));
s.push(...recuadro("consejo", "Sin prisa", ["No hace falta completarlos todos de una vez: guarda cuando quieras y cierra. La próxima vez solo saldrán los que sigan sin código. Si cierras con códigos sin guardar, la consola te pregunta."]));

// ── 6 ──────────────────────────────────────────────────────────────────────
s.push(h1("6. Generar la TARIFA (Acciones 4, 5 y 6)"));
s.push(h2("Acción 4 — Generar TARIFA final"));
s.push(...pasos(["Deja la tarifa del proveedor (el **.xls** que envía) en **TARIFAS** y pulsa **Buscar** (o la Acción 4).",
  "Si no se llama TARIFA.xls, la consola te deja elegirla de la lista y la renombra.",
  "Pulsa **4 · Generar TARIFA final**. La consola copia los artículos, cruza cada código DALI con el catálogo para poner su código SAP, calcula los precios con los descuentos y marca duplicados.",
  "Resultado en TARIFAS: «**PROVEEDOR [dd-mm-aaaa].xlsx**» si está completa, o «**PROVEEDOR [INCOMPLETO] [dd-mm-aaaa].xlsx**» si faltan códigos SAP o hay duplicados. El .xls de origen se borra."]));
s.push(...imagen("c_seleccion_xls", 11, "Elegir la tarifa del proveedor."));
s.push(...leyenda([[1, "Los **.xls** que hay en TARIFAS."], [2, "**Usar este archivo**: lo renombra a TARIFA.xls y sigue."]]));
s.push(h2("Códigos SAP faltantes"));
s.push(p("Si algún código DALI de la tarifa no tiene SAP en el catálogo, se abre esta ventana para escribirlo. Lo que guardes se escribe a la vez en el catálogo y en la tarifa, y se envía a Catálogo Asignaciones. Para los artículos sin SAP que no son de esta tarifa, usa la **Acción 8** (capítulo 5)."));
s.push(...imagen("c_sap_faltantes", 16, "Códigos DALI sin SAP asignado (datos de ejemplo)."));
s.push(...leyenda([[1, "Código DALI, descripción y formato de cada artículo sin SAP."], [2, "Escribe aquí su **código SAP**. Deja en blanco los que no quieras completar ahora."], [3, "**Guardar en el catálogo**."]]));
s.push(h2("Códigos SAP duplicados"));
s.push(p("Si el código SAP que escribes ya está asignado a otro DALI (o lo has escrito dos veces), la consola te pide cuál es el correcto."));
s.push(...imagen("c_sap_duplicados", 15, "Resolver un código SAP duplicado."));
s.push(...leyenda([[1, "Marca el **DALI incorrecto**: se le borrará ese SAP. El que quede sin marcar lo conserva. En cada grupo debe quedar al menos uno sin marcar."], [2, "**Confirmar y continuar guardado**."], [3, "**Cancelar todo el guardado** si prefieres revisarlo antes."]]));
s.push(h2("Terminar una tarifa [INCOMPLETO]"));
s.push(p("Cuando ya estén los códigos (por ejemplo, después de dar de alta los materiales en SAP y repetir la **Acción 3**), pulsa otra vez la **Acción 4**: si no hay TARIFA.xls, te ofrece las tarifas [INCOMPLETO] para volver a cruzarlas con el catálogo."));
s.push(...imagen("c_incompletos_a4", 12, "Acción 4 con tarifas [INCOMPLETO]."));
s.push(...leyenda([[1, "Tarifas [INCOMPLETO] de TARIFAS."], [2, "**Seleccionar**: aplica los SAP y las descripciones del catálogo. Si queda completa, se quita [INCOMPLETO] del nombre."]]));
s.push(h2("Acción 6 — Corregir a mano un archivo [INCOMPLETO]"));
s.push(...imagen("c_incompletos_a6", 12, "Acción 6: tarifas incompletas para edición manual."));
s.push(...leyenda([[1, "Los pasos a seguir."], [2, "Tarifas [INCOMPLETO]."], [3, "**Abrir en Excel**: corrige (líneas duplicadas, códigos pendientes), guarda, cierra Excel y vuelve a pasar la **Acción 4**."]]));
s.push(p("La **Acción 5** abre en Excel la última TARIFA procesada para revisarla."));
s.push(...recuadro("nota", "Catálogo DALI completo", ["Al completar una TARIFA, la consola pasa al catálogo DALI completo los códigos y precios del proveedor. Si un código DALI de la tarifa tiene asignado **otro proveedor** en DALI, no lo toca y lo avisa como «CASO IRREGULAR»: revísalo en DALI. Después, los precios y códigos de proveedor cambiados se envían a Catálogo Asignaciones (capítulo 8)."]));

// ── 7 ──────────────────────────────────────────────────────────────────────
s.push(h1("7. Crear el Registro Info (RINFO)"));
s.push(p("Con la TARIFA ya completa (sin [INCOMPLETO]), pulsa el acceso rápido **Crear / Actualizar Registro Info SAP**."));
s.push(...imagen("c_rinfo_tarifa", 12, "Elegir la TARIFA."));
s.push(...leyenda([[1, "TARIFAS finales disponibles (no salen las [INCOMPLETO] ni los RINFO)."], [2, "**Continuar**: la consola busca el proveedor en Proveedores.xlsx y un Purchasing compatible en TARIFAS."]]));
s.push(h2("Si el proveedor ya tiene Registros Info en SAP (Purchasing encontrado)"));
s.push(p("Si en TARIFAS hay un **Purchasing Info Records.xlsx** de ese proveedor (descargado de SAP), la consola lo usa como base: conserva los registros que ya existen, actualiza precios y fechas y añade los materiales nuevos de la tarifa."));
s.push(...imagen("c_rinfo_purchasing", 12, "Purchasing Info Records encontrado."));
s.push(...leyenda([[1, "**Proveedor SAP** (nombre y código) encontrado en Proveedores.xlsx."], [2, "El **Purchasing** encontrado y su parecido con el proveedor."],
  [3, "**Continuar usando este Purchasing** (lo normal)."], [4, "**Comenzar desde cero**: un Registro Info nuevo sin tener en cuenta el Purchasing."]]));
s.push(h2("Si no hay Purchasing"));
s.push(...imagen("c_rinfo_sin_purchasing", 10, "Proveedor sin Purchasing en TARIFAS."));
s.push(...leyenda([[1, "**Descargar el Purchasing desde SAP primero**: muestra los 7 pasos en SAP. Guárdalo en TARIFAS y vuelve a empezar."], [2, "**Crear nuevo Registro Info desde cero**, si el proveedor no tiene ninguno en SAP."]]));
s.push(...imagen("c_rinfo_instrucciones", 11, "Cómo descargar el Purchasing Info Records desde SAP."));
s.push(...leyenda([[1, "Acceso directo a SAP S/4HANA. Sigue los pasos y guarda el archivo descargado en TARIFAS."]]));
s.push(h2("Teléfono, fechas y centros"));
s.push(...imagen("c_rinfo_telefono", 9, "Teléfono del proveedor (opcional)."));
s.push(...leyenda([[1, "Teléfono (8 a 10 cifras; los espacios se quitan solos)."], [2, "**Asignar teléfono y continuar**."], [3, "**Sin teléfono**, si no lo tienes."]]));
s.push(...imagen("c_rinfo_fechas", 9, "Fechas de validez de la tarifa."));
s.push(...leyenda([[1, "**Inicio** de validez (dd.mm.aaaa)."], [2, "**Fin** de validez."], [3, "**Confirmar fechas y continuar**."]]));
s.push(...imagen("c_rinfo_centros", 9, "Centros SAP del Registro Info."));
s.push(...leyenda([[1, "**Todos / Ninguno** para marcar o desmarcar de golpe."], [2, "Centros (hoteles). Desmarca los que no vayan a comprar esta tarifa."], [3, "**Confirmar selección y continuar**."]]));
s.push(h2("Avisos durante la generación"));
s.push(p("Con un Purchasing existente, puede preguntarte qué hacer con materiales que ya están en SAP pero **no vienen en la tarifa nueva**:"));
s.push(...imagen("c_rinfo_sin_tarifa", 13, "Materiales sin tarifa activa."));
s.push(...leyenda([[1, "Materiales del Purchasing que no están en la tarifa, con su fin de validez. Marca los que quieras **dar de baja**."], [2, "**MANTENER**: no cambia nada."], [3, "**DAR BAJA**: los marca para borrar en SAP."]]));
s.push(p("Y si algún material sigue en la tarifa pero su Registro Info está **marcado para borrar** en SAP, te avisa de que hay que reactivarlo a mano en SAP, con los centros afectados:"));
s.push(...imagen("c_rinfo_reactivar", 13, "Materiales a reactivar manualmente en SAP."));
s.push(...leyenda([[1, "Material, Registro Info y centros afectados."], [2, "Los pasos en SAP («Marcar reg.info de compra para borrar»), con una captura de referencia."]]));
s.push(h2("Resultado"));
s.push(...imagen("c_rinfo_generado", 10, "Registro Info generado desde un Purchasing."));
s.push(...leyenda([[1, "Resumen: materiales que ya existían y nuevos añadidos."], [2, "**Sí, abrir ahora** para revisarlo en Excel."]]));
s.push(...imagen("c_rinfo_creado", 10, "Registro Info creado desde cero."));
s.push(...leyenda([[1, "Valores escritos: proveedor SAP, código y artículos copiados."], [2, "**Sí, abrir ahora**."]]));
s.push(...recuadro("importante", "Antes de importarlo en SAP", ["Mira el **resumen de validación** de la consola: si dice «Sin incidencias detectadas», el archivo puede importarse.",
  "Importa cada RINFO **una sola vez**. Si necesitas repetir la carga, vuelve a descargar el Purchasing de SAP y genera el RINFO otra vez: un Purchasing antiguo ya no refleja lo que hay en SAP."]));

// ── 8 ──────────────────────────────────────────────────────────────────────
s.push(h1("8. Cambios en Catálogo Asignaciones"));
s.push(p("Después de la **Acción 2**, la **Acción 3**, la **Acción 8**, al completar una **TARIFA** (Acción 4) o al resolver un código SAP duplicado, la consola envía a la web **Catálogo Asignaciones** **solo lo que ha cambiado**: altas y datos DALI, datos SAP, y precio y código de proveedor. Lo hace en segundo plano: puedes seguir trabajando mientras tanto."));
s.push(p("La primera vez en cada PC descarga una copia del catálogo de la web (tarda un poco); después solo pregunta lo que ha cambiado. Antes de tocar nada, la web comprueba el envío y la consola te enseña un resumen:"));
s.push(...imagen("c_publicar", 13, "Resumen de los cambios para Catálogo Asignaciones (datos de ejemplo)."));
s.push(...leyenda([[1, "Cuántos artículos cambian de cada tipo: **datos DALI** (con las altas), **datos SAP** y **precio y código de proveedor**."],
  [2, "Artículos que se **cambiaron en la web** (o que envió otro compañero) desde tu último envío. Por defecto, sus datos SAP, su precio y su código de proveedor **no se tocan**; marca la casilla si quieres sustituirlos por los de la consola. La descripción DALI siempre se envía: solo se carga desde la consola."],
  [3, "Artículos que la web **no cambiará**: el proveedor asignado no es el de la tarifa, o el código de proveedor guardado es otro. Revísalos y corrígelos."],
  [4, "**Detalle**: cada cambio, con el valor de la web y el nuevo."],
  [5, "**Actualizar la web**: aplica los cambios."], [6, "**Ahora no**: no envía nada; los cambios volverán a salir la próxima vez."]]));
s.push(...recuadro("nota", "Al terminar", ["La consola dice si la web y tu catálogo coinciden. Si quedan diferencias (cambios de la web que dejaste fuera o artículos irregulares), te dice cuántas.",
  "En la web nunca se borra nada desde la consola, y un precio o un código de proveedor vacío en tu catálogo no borra el de la web."]));
s.push(...recuadro("importante", "Permiso necesario", ["Para enviar cambios (y para abrir la consola) tu cuenta de Catálogo Asignaciones necesita el permiso **«Escribir desde la Consola Registro Info»**. Si falta, la consola lo dice en rojo."]));

// ── 9 ──────────────────────────────────────────────────────────────────────
s.push(h1("9. Consultar el manual y actualizaciones"));

s.push(h2("Consultar el manual"));
s.push(p("En la barra **🔎 Consultar el manual**, encima de la consola, escribe tu duda con tus palabras (por ejemplo, «códigos SAP faltantes» o «contraseña catálogo») y pulsa **Intro**. La consola busca en este manual y te enseña los apartados que mejor encajan, con un fragmento y la página. No necesita Internet."));
s.push(...imagen("c_consulta_manual", 13, "Resultados de una consulta."));
s.push(...leyenda([[1, "Tu consulta. Puedes cambiarla y pulsar **Buscar** otra vez."], [2, "Apartados del manual que mejor responden, con su capítulo, página y un fragmento."],
  [3, "**Abrir el manual** completo (también con el acceso rápido **Manual de usuario**)."], [4, "**Guardar una copia** del manual donde quieras."]]));
s.push(h2("Actualizaciones"));
s.push(...imagen("c_actualizacion", 11, "Aviso de versión nueva (ejemplo)."));
s.push(...leyenda([[1, "Las **novedades** de la versión nueva."], [2, "**Actualizar ahora**: descarga, comprueba que es la versión oficial (SHA-256), sustituye el programa y se vuelve a abrir solo."], [3, "**Más tarde**: te lo volverá a ofrecer la próxima vez que la abras."]]));
s.push(...recuadro("nota", "Qué pasa con la versión anterior", ["Al abrirse la versión nueva, borra sola la anterior (TARIFA_Finder.anterior.exe) y lo dice en la consola: «Versión anterior borrada». Tu carpeta del proyecto, tu contraseña guardada y tu copia del catálogo de la web no se tocan."]));

// ── 10 ─────────────────────────────────────────────────────────────────────
s.push(h1("10. Consulta rápida y problemas frecuentes"));
s.push(...tabla(["Si ves…", "Qué significa", "Qué hacer"], [
  ["«TARIFA.xls NO ENCONTRADA»", "No hay ninguna tarifa .xls en TARIFAS", "Copia el .xls del proveedor en TARIFAS y pulsa Buscar"],
  ["La tarifa sale como [INCOMPLETO]", "Faltan códigos SAP o hay duplicados", "Completa los SAP (ventana que se abre) o usa la Acción 6, y repite la Acción 4"],
  ["«DALI sin DESCRIPCION/UNIDAD SAP»", "El material tiene código SAP pero el catálogo no tiene sus datos de SAP", "Exporta Productos.xlsx de SAP a RECURSOS y pasa la Acción 3"],
  ["Artículos DALI activos sin código SAP", "Aún no tienen material SAP asignado", "Asígnalos con la Acción 8"],
  ["«CASO IRREGULAR»", "El DALI tiene asignado otro proveedor en DALI", "Revisa en DALI el proveedor asignado de esos códigos"],
  ["«No se encontró Purchasing Info Records compatible»", "En TARIFAS no hay un Purchasing de ese proveedor", "Descárgalo de SAP (7 pasos) o crea el RINFO desde cero"],
  ["Error al guardar un archivo", "Seguramente está abierto en Excel", "Ciérralo en Excel y repite la acción"],
  ["«Usuario no autorizado.» al abrir", "Tu cuenta no es de administrador, no está activa o no tiene el permiso «Escribir desde la Consola Registro Info»", "Pide el permiso al administrador principal de Catálogo Asignaciones"],
  ["«Email o contraseña incorrectos.»", "Los datos no coinciden con los de la web", "Escríbelos de nuevo (son los de Catálogo Asignaciones)"],
  ["Ventana «Cambios para Catálogo Asignaciones»", "La consola va a enviar a la web lo que ha cambiado", "Revisa el resumen y pulsa Actualizar la web (capítulo 8)"],
  ["paso_listado_a_catalogo_….txt en RECURSOS", "Diferencias encontradas al pasar el antiguo LISTADO al catálogo (solo una vez)", "Revísalo por si alguna diferencia era un error"],
  ["«Listado no encontrado» en otro PC", "Ese PC tiene una consola antigua, que aún busca el LISTADO", "Actualízala (Actualizar ahora al abrirla)"],
  ["«No se pudo comprobar si hay versión nueva»", "Sin conexión con GitHub en ese momento", "Nada: se comprobará la próxima vez"],
  ["La app parece «colgada»", "Está trabajando o espera una ventana", "Mira la consola y si hay una ventana detrás; guarda el log si hay un error"],
], [3000, 3300, 3338]));
s.push(...recuadro("consejo", "Para pedir ayuda", ["Pulsa **guardar log** y envía el .txt junto con el nombre del archivo con el que estabas trabajando. Con eso se puede ver exactamente qué pasó."]));

generar(process.argv[2] || "salida/Manual_Usuario_ConsolaRegistroInfo.docx", "Consola Registro Info SAP — Manual de usuario", "de usuario", s);
