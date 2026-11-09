# -*- coding: utf-8 -*-

"""
tarifa_constants.py
-------------------
Constantes de diseño, tipografía y utilidades compartidas entre
los módulos de TARIFA Finder.

Importar en cada módulo:
    from tarifa_constants import *
o de forma explícita:
    from tarifa_constants import C_MID, FONT_TITLE, resource_path, ...

Central Compras Princess Canarias — (c) VAMA 2026
"""

import os
import sys


# ---------------------------------------------------------
#  RUTA DE RECURSOS  (compatible .py y .exe PyInstaller)
# ---------------------------------------------------------
def resource_path(relative_path: str) -> str:
    try:
        base = sys._MEIPASS
    except AttributeError:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


# ---------------------------------------------------------
#  PALETA DE COLORES — Princess 2026
# ---------------------------------------------------------
# ── Fondo y estructura ─────────────────────────────────
C_DARK       = "#060D18"   # Navbar más oscuro
C_MID        = "#0F2540"   # Azul marino medio
C_SIDEBAR    = "#0A1628"   # Sidebar principal
C_CARD       = "#0D1F35"   # Tarjetas / cards
C_BORDER     = "#162840"   # Bordes sutiles
C_HOVER      = "#142840"   # Estado hover en sidebar
# ── Acento dorado Princess ─────────────────────────────
C_ACCENT     = "#C8A84B"   # Dorado principal
C_ACCENT2    = "#E8C96B"   # Dorado claro
C_ACCENT_DIM = "#6A5820"   # Dorado apagado (disabled)
# ── Superficies claras (ventanas popup) ────────────────
C_BG         = "#F4F6F9"
C_PANEL      = "#FFFFFF"
C_TEXT       = "#1A1A2E"   # Texto oscuro sobre fondo claro
C_TEXT_LIGHT = "#DCE8F4"   # Texto claro sobre fondo oscuro
C_MUTED      = "#4A6A8D"   # Texto atenuado sobre oscuro
# ── Estados ────────────────────────────────────────────
C_SUCCESS    = "#27AE60"
C_ERROR      = "#E74C3C"
C_WARNING    = "#F0A500"
# ── Consola ────────────────────────────────────────────
C_LOG_BG     = "#070F1C"
C_LOG_FG     = "#A8C8A0"
C_LOG_INFO   = "#5B9BD5"
C_LOG_WARN   = "#F0C040"
C_LOG_ERR    = "#E07070"
C_LOG_OK     = "#60D060"


# ---------------------------------------------------------
#  TIPOGRAFÍA — Segoe UI en toda la app
# ---------------------------------------------------------
FONT_TITLE   = ("Segoe UI", 13, "bold")
FONT_SUB     = ("Segoe UI",  8, "italic")
FONT_LABEL   = ("Segoe UI",  9)
FONT_BOLD    = ("Segoe UI",  9, "bold")
FONT_LOG     = ("Consolas",  8)   # sobreescrito en runtime por ConsolaLog
FONT_SMALL   = ("Segoe UI",  8)
FONT_TINY    = ("Segoe UI",  7)
FONT_SECTION = ("Segoe UI",  7, "bold")


# ---------------------------------------------------------
#  LOGOS EMBEBIDOS — base64
# ---------------------------------------------------------
# Logo SAP S/4HANA (120×44 px) — reutilizado en header y acceso rápido
SAP_LOGO_B64 = (
    "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"
)

# ---------------------------------------------------------
#  DICCIONARIOS SAP — Acción 1 (Organizar DALI)
#  Catálogos de valores válidos para validar las columnas
#  SUBFAMILIA e IGIC del ARTICULOS_DALI.xlsx generado.
#  (Movidos aquí desde tarifa_app.py::_organizar_dali)
# ---------------------------------------------------------
DALI_CLAVES_SUBFAMILIA = {
    "ACCIDENTES Y DESASTRES","ACCESORIOS IMAGEN Y SONIDO BAZAR",
    "ACEITES","AZUCARES Y SIROPES","HUEVOS","AGUAS",
    "ALARMAS Y DETECTORES","ALIMENTACION DE AGUA",
    "ALIMENTACION VARIOS 5ª GAMA","ALIMENTACION VARIOS DIREC / PROPIEDAD",
    "ANISADOS","APERITIVOS","APARATOS SANITARIOS","ARTICULOS BAZAR",
    "ARTICULOS DE COSTURA","ARTICULOS PARA COCTELERIA",
    "ARTICULOS PLAYA BAZAR","ARTICULOS Y COMP. OFICINA","ARROCES",
    "ATENCIONES CLIENTES","BEBIDAS BAZAR","BLOQUES Y LADRILLOS",
    "BOLLERIA CONGELADA","BOLLERIA FRESCA","BOMBILLAS",
    "BRANDYS Y COÑACS","CAFE, CACAOS, SOLUBLES E INFUSIONES",
    "CAJAS FUERTES","CALEFACCION SIN INSTALACION",
    "CANDADOS Y ACCESORIOS","CARNE CONGELADA AVES Y LEPORIDOS",
    "CARNE CONGELADA CAZA","CARNE CONGELADA OVINO",
    "CARNE CONGELADA PORCINO","CARNE CONGELADA VACUNO",
    "CARNE REFRIGERADA AVES Y LEPORIDOS","CARNE REFRIGERADA OVINO",
    "CARNE REFRIGERADA PORCINO","CARNE REFRIGERADA VACUNO",
    "CARNES AHUMADAS","CAVAS, ESPUMOSOS Y CHAMPAGNE",
    "CAVAS SIN ALCOHOL","CELULOSA","CERRADURAS","CERVEZAS",
    "CERVEZAS SIN ALCOHOL","CEREALES","CLASIFICACION Y ARCHIVO",
    "COMBINADOS","COMIDA BAZAR","COMPLEMENTOS","COMPLEMENTOS BAÑO",
    "COMPLEMENTOS COCINA - ECONOMATO","PORCELANA - CERAMICA - BARRO",
    "CUBERTERIA","GAMMA BLANCA","PLASTICOS, MELAMINA Y SILICONA",
    "COMPLEMENTOS COMEDOR - BARES","COMPLEMENTOS DECORACION",
    "COMPLEMENTOS HABITACIONES","COMPLEMENTOS RECEPCION",
    "COMPLEMENTOS TELEFONO","COMPLEMENTOS Y ADORNOS PASTELERIA",
    "COMPLEMENTOS ZONAS COMUNES","COMPUTO Y SONIDO",
    "CONSERVAS DE PESCADO","CONSERVAS EN ACEITE","CONSERVAS EN ALMIBAR",
    "CONSERVAS VEGETALES","CONDIMENTOS Y ESPECIAS",
    "CONSUMIBLES IMPRESION","COSMETICA Y PELUQUERIA","CORTINAS",
    "CREMAS SOLARES BAZAR","CRISTALERIA","CUADROS","HARDWARE",
    "TELEFONOS","DECORACION BUFFET","DECORACION JARDIN","DIGESTIVOS",
    "ELECTRODOMESTICO PEQUEÑO","ELABORADOS CARNICOS","ELABORADOS FRESCOS",
    "EMBUTIDOS, FIAMBRES Y PATES","ENCURTIDOS Y ACEITUNAS",
    "ENVASES ALIMENTACION","ENVASES BARRIL","ENVASES BEBIDAS",
    "ENVASES GAS CARBONICO (CO2)","ENVASES LIMPIEZA Y QUIMICOS",
    "EQUIPOS DE SONIDO","EQUIPOS IMAGEN","EQUIPOS TRATAMIENTO DE AGUAS",
    "ALMACENAJE","ESCRITURA Y CORRECCION","ESPEJOS","ESCALERAS",
    "ETIQUETAS VARIAS","FERRETERIA DE FIJACION","FERRETERIA PARA MUEBLES",
    "FLAN, GELATINA, MOUSE Y CREMOSOS","AIRE ACONDICIONADO",
    "OTROS JARDINERIA","OTROS MANO DE OBRA EXTERNO",
    "MTO DE PISCINAS EXTERNO","OTROS PINTURAS","PANADERIA CONGELADA",
    "PANADERIA FRESCA","PANTALLAS Y BASES","PAPEL Y BLOCS","PASTAS",
    "CONFITADOS Y MERMELADAS (CONSERVAS DE FRUTAS Y DULCES)",
    "PASTELERIA CONGELADA","PASTELERIA FRESCA","PESCADO CONGELADO",
    "PESCADO FRESCO","PESCADOS AHUMADOS","PILAS","PINTURAS EXTERIOR",
    "PINTURAS INTERIOR","PRECOCINADO PASTA CONGELADA","PRECOCINADOS CARNE",
    "PRECOCINADOS DE QUESOS","PRECOCINADOS PASTA Y ARROZ",
    "PRECOCINADOS PESCADO Y MARISCO","PRECOCINADOS VERDURAS",
    "PROD. ACOGIDA","PROD. LIMPIEZA PISOS Y ZONAS COMUNES",
    "PROD. LIMPIEZAS COCINAS","PROD. LIMPIEZA ROPA","PROD. PISCINAS",
    "PRODUCTOS LIBRERIA BAZAR","PRODUCTOS PARA CELIACOS",
    "PRODUCTOS Y MATERIAL ASEO BAZAR","PROTECCION CONTRA DESASTRES",
    "PUROS","QUESOS","RECAMBIO MAQUINARIA GIMNASIO",
    "RECEPCIONISTA CABALLERO","REFRESCOS","REPARACIONES EXTERNAS",
    "RON","ROPA BAÑO","ROPA CAMA","ROPA MESA","ROPA Y COMPLEMENTOS BAZAR",
    "ROPA ZONAS COMUNES","SALSA","SALSAS","SANGRIAS","SEGURIDAD LABORAL",
    "SIDRAS","SISTEMAS DE VIGILANCIA","SNACKS Y FRUTOS SECOS",
    "SOBRES, ENVIOS Y ETIQUETAS","SOPAS, CREMAS, PURES Y CALDOS",
    "SOUVENIRS BAZAR","SUCEDANEOS","SUELOS PARA EXTERIOR",
    "TELEFONIA MOVIL","TELEFONILLOS Y TIMBRES","TEQUILAS Y AGUARDIENTES",
    "TEXTIL LIMPIEZA","UN SOLO USO ALIMENTACION - VARIOS",
    "UN SOLO USO BEBIDAS - VARIOS","MATERIAL MERCHANDASING Y CATALOGOS",
    "UNIFORMIDAD","UTENSILIOS LIMPIEZA","VAJILLA","VARIOS ELECTRICIDAD",
    "VARIOS FERRETERIA","VARIOS FONTANERIA","VARIOS JARDINERIA",
    "VARIOS MANO DE OBRA EXTERNO","VARIOS PINTURAS","VENTILADORES",
    "VERDURA, PAPAS Y HORTALIZAS CONGELADAS","VERDURA Y HORTALIZAS FRESCAS",
    "VINAGRES","VINOS BLANCOS","VINOS GENEROSOS","VINOS ROSADOS",
    "VINOS TINTOS","VODKA","WHISKYS","YOGURES","ZUMOS",
}
DALI_CLAVES_IGIC = {
    "0 EXENTO","03 REDUCIDO","05 REDUCIDO",
    "15 INCREMENTADO","07 GENERAL","20 ESPECIAL",
}

# SUBFAMILIA DALI -> (codigo grupo compras SAP, nombre grupo compras SAP)
# Usada en tarifa_app.py::_actualizar_listado_sap (Accion 3). Extraida de
# ahi por ser una tabla de datos estatica sin logica, siguiendo el mismo
# patron que DALI_CLAVES_SUBFAMILIA/DALI_CLAVES_IGIC mas arriba.
TABLA_SUBFAMILIA_GRUPO_COMPRAS = {
    "ACCIDENTES Y DESASTRES":                      ("900", "S. PERSONAL"),
    "ACCESORIOS IMAGEN Y SONIDO BAZAR":            ("301", "COMPU. EQUIPMENT"),
    "ACEITES":                                     ("100", "FOOD"),
    "AZUCARES Y SIROPES":                          ("100", "FOOD"),
    "HUEVOS":                                      ("100", "FOOD"),
    "AGUAS":                                       ("200", "BEVERAGES"),
    "ALARMAS Y DETECTORES":                        ("900", "S. PERSONAL"),
    "ALIMENTACION DE AGUA":                        ("400", "MAINTENANCE"),
    "ALIMENTACION VARIOS 5ª GAMA":                 ("100", "FOOD"),
    "ALIMENTACION VARIOS DIREC / PROPIEDAD":       ("100", "FOOD"),
    "ANISADOS":                                    ("200", "BEVERAGES"),
    "APERITIVOS":                                  ("200", "BEVERAGES"),
    "APARATOS SANITARIOS":                         ("400", "MAINTENANCE"),
    "ARTICULOS BAZAR":                             ("500", "SUPPLIES"),
    "ARTICULOS DE COSTURA":                        ("500", "SUPPLIES"),
    "ARTICULOS PARA COCTELERIA":                   ("500", "SUPPLIES"),
    "ARTICULOS PLAYA BAZAR":                       ("500", "SUPPLIES"),
    "ARTICULOS Y COMP. OFICINA":                   ("500", "SUPPLIES"),
    "ARROCES":                                     ("100", "FOOD"),
    "ATENCIONES CLIENTES":                         ("500", "SUPPLIES"),
    "BEBIDAS BAZAR":                               ("200", "BEVERAGES"),
    "BLOQUES Y LADRILLOS":                         ("400", "MAINTENANCE"),
    "BOLLERIA CONGELADA":                          ("100", "FOOD"),
    "BOLLERIA FRESCA":                             ("100", "FOOD"),
    "BOMBILLAS":                                   ("400", "MAINTENANCE"),
    "BRANDYS Y COÑACS":                            ("200", "BEVERAGES"),
    "CAFE, CACAOS, SOLUBLES E INFUSIONES":         ("100", "FOOD"),
    "CAJAS FUERTES":                               ("400", "MAINTENANCE"),
    "CALEFACCION SIN INSTALACION":                 ("400", "MAINTENANCE"),
    "CANDADOS Y ACCESORIOS":                       ("400", "MAINTENANCE"),
    "CARNE CONGELADA AVES Y LEPORIDOS":            ("100", "FOOD"),
    "CARNE CONGELADA CAZA":                        ("100", "FOOD"),
    "CARNE CONGELADA OVINO":                       ("100", "FOOD"),
    "CARNE CONGELADA PORCINO":                     ("100", "FOOD"),
    "CARNE CONGELADA VACUNO":                      ("100", "FOOD"),
    "CARNE REFRIGERADA AVES Y LEPORIDOS":          ("100", "FOOD"),
    "CARNE REFRIGERADA OVINO":                     ("100", "FOOD"),
    "CARNE REFRIGERADA PORCINO":                   ("100", "FOOD"),
    "CARNE REFRIGERADA VACUNO":                    ("100", "FOOD"),
    "CARNES AHUMADAS":                             ("100", "FOOD"),
    "CAVAS, ESPUMOSOS Y CHAMPAGNE":                ("200", "BEVERAGES"),
    "CAVAS SIN ALCOHOL":                           ("200", "BEVERAGES"),
    "CELULOSA":                                    ("500", "SUPPLIES"),
    "CERRADURAS":                                  ("400", "MAINTENANCE"),
    "CERVEZAS":                                    ("200", "BEVERAGES"),
    "CERVEZAS SIN ALCOHOL":                        ("200", "BEVERAGES"),
    "CEREALES":                                    ("100", "FOOD"),
    "CLASIFICACION Y ARCHIVO":                     ("501", "PAPER SUPPLIES"),
    "COMBINADOS":                                  ("200", "BEVERAGES"),
    "COMIDA BAZAR":                                ("500", "SUPPLIES"),
    "COMPLEMENTOS":                                ("500", "SUPPLIES"),
    "COMPLEMENTOS BAÑO":                           ("500", "SUPPLIES"),
    "COMPLEMENTOS COCINA - ECONOMATO":             ("300", "OPE. EQUIPMENT"),
    "PORCELANA - CERAMICA - BARRO":                ("300", "OPE. EQUIPMENT"),
    "CUBERTERIA":                                  ("300", "OPE. EQUIPMENT"),
    "GAMMA BLANCA":                                ("300", "OPE. EQUIPMENT"),
    "PLASTICOS, MELAMINA Y SILICONA":              ("300", "OPE. EQUIPMENT"),
    "COMPLEMENTOS COMEDOR - BARES":                ("500", "SUPPLIES"),
    "COMPLEMENTOS DECORACION":                     ("500", "SUPPLIES"),
    "COMPLEMENTOS HABITACIONES":                   ("500", "SUPPLIES"),
    "COMPLEMENTOS RECEPCION":                      ("500", "SUPPLIES"),
    "COMPLEMENTOS TELEFONO":                       ("301", "COMPU. EQUIPMENT"),
    "COMPLEMENTOS Y ADORNOS PASTELERIA":           ("100", "FOOD"),
    "COMPLEMENTOS ZONAS COMUNES":                  ("400", "MAINTENANCE"),
    "COMPUTO Y SONIDO":                            ("301", "COMPU. EQUIPMENT"),
    "CONSERVAS DE PESCADO":                        ("100", "FOOD"),
    "CONSERVAS EN ACEITE":                         ("100", "FOOD"),
    "CONSERVAS EN ALMIBAR":                        ("100", "FOOD"),
    "CONSERVAS VEGETALES":                         ("100", "FOOD"),
    "CONDIMENTOS Y ESPECIAS":                      ("100", "FOOD"),
    "CONSUMIBLES IMPRESION":                       ("501", "PAPER SUPPLIES"),
    "COSMETICA Y PELUQUERIA":                      ("500", "SUPPLIES"),
    "CORTINAS":                                    ("700", "UNIFORMS"),
    "CREMAS SOLARES BAZAR":                        ("500", "SUPPLIES"),
    "CRISTALERIA":                                 ("300", "OPE. EQUIPMENT"),
    "CUADROS":                                     ("500", "SUPPLIES"),
    "HARDWARE":                                    ("301", "COMPU. EQUIPMENT"),
    "TELEFONOS":                                   ("301", "COMPU. EQUIPMENT"),
    "DECORACION BUFFET":                           ("500", "SUPPLIES"),
    "DECORACION JARDIN":                           ("500", "SUPPLIES"),
    "DIGESTIVOS":                                  ("200", "BEVERAGES"),
    "ELECTRODOMESTICO PEQUEÑO":                    ("300", "OPE. EQUIPMENT"),
    "ELABORADOS CARNICOS":                         ("100", "FOOD"),
    "ELABORADOS FRESCOS":                          ("100", "FOOD"),
    "EMBUTIDOS, FIAMBRES Y PATES":                 ("100", "FOOD"),
    "ENCURTIDOS Y ACEITUNAS":                      ("100", "FOOD"),
    "ENVASES ALIMENTACION":                        ("500", "SUPPLIES"),
    "ENVASES BARRIL":                              ("500", "SUPPLIES"),
    "ENVASES BEBIDAS":                             ("500", "SUPPLIES"),
    "ENVASES GAS CARBONICO (CO2)":                 ("500", "SUPPLIES"),
    "ENVASES LIMPIEZA Y QUIMICOS":                 ("800", "CHEMICALS"),
    "EQUIPOS DE SONIDO":                           ("301", "COMPU. EQUIPMENT"),
    "EQUIPOS IMAGEN":                              ("301", "COMPU. EQUIPMENT"),
    "EQUIPOS TRATAMIENTO DE AGUAS":                ("400", "MAINTENANCE"),
    "ALMACENAJE":                                  ("400", "MAINTENANCE"),
    "ESCRITURA Y CORRECCION":                      ("501", "PAPER SUPPLIES"),
    "ESPEJOS":                                     ("400", "MAINTENANCE"),
    "ESCALERAS":                                   ("400", "MAINTENANCE"),
    "ETIQUETAS VARIAS":                            ("501", "PAPER SUPPLIES"),
    "FERRETERIA DE FIJACION":                      ("400", "MAINTENANCE"),
    "FERRETERIA PARA MUEBLES":                     ("400", "MAINTENANCE"),
    "FLAN, GELATINA, MOUSE Y CREMOSOS":            ("100", "FOOD"),
    "FLORES, PLANTAS Y MACETAS":                   ("500", "SUPPLIES"),
    "FRUTA CONGELADA":                             ("100", "FOOD"),
    "FRUTA FRESCA":                                ("100", "FOOD"),
    "GALLETAS":                                    ("100", "FOOD"),
    "GAS":                                         ("400", "MAINTENANCE"),
    "GAS CARBONICO (CO2)":                         ("200", "BEVERAGES"),
    "GINEBRAS":                                    ("200", "BEVERAGES"),
    "GOLOSINAS":                                   ("100", "FOOD"),
    "GOLOSINAS Y BOLLERIA BAZAR":                  ("100", "FOOD"),
    "GRANIZADOS Y HORCHATAS":                      ("200", "BEVERAGES"),
    "GRIFERIA Y VALVULAS":                         ("400", "MAINTENANCE"),
    "HARINAS, FECULAS Y LEVADURAS":                ("100", "FOOD"),
    "HELADOS BAZAR":                               ("100", "FOOD"),
    "HELADOS GRANEL":                              ("100", "FOOD"),
    "HELADOS IMPULSO":                             ("100", "FOOD"),
    "HERRAMIENTAS JARDINERIA":                     ("400", "MAINTENANCE"),
    "HIDROSANITARIOS":                             ("400", "MAINTENANCE"),
    "HORTALIZAS":                                  ("100", "FOOD"),
    "HORTALIZAS Y VERDURAS FRESCAS":               ("100", "FOOD"),
    "IBERICOS":                                    ("100", "FOOD"),
    "IDENTIFICACIÓN PERSONAL":                     ("900", "S. PERSONAL"),
    "MASAJE":                                      ("900", "S. PERSONAL"),
    "EQUIPO CONTRA INCENDIO":                      ("900", "S. PERSONAL"),
    "ILUMINACION EXTERIOR":                        ("400", "MAINTENANCE"),
    "ILUMINACION INTERIOR":                        ("400", "MAINTENANCE"),
    "JUEGOS DE MESA":                              ("500", "SUPPLIES"),
    "JUGUETERIA Y DEPORTE BAZAR":                  ("500", "SUPPLIES"),
    "LAMPARAS":                                    ("400", "MAINTENANCE"),
    "LECHE POLVO":                                 ("100", "FOOD"),
    "LECHE Y BATIDOS":                             ("100", "FOOD"),
    "LEGUMBRES SECAS":                             ("100", "FOOD"),
    "LIBROS - DICCIONARIOS - LIBRERIA":            ("500", "SUPPLIES"),
    "LICORES":                                     ("200", "BEVERAGES"),
    "CIGARRILLOS":                                 ("200", "BEVERAGES"),
    "LICORES COCTELERIA":                          ("200", "BEVERAGES"),
    "LICORES SIN ALCOHOL":                         ("200", "BEVERAGES"),
    "MANTEQUILLAS Y MARGARINAS":                   ("100", "FOOD"),
    "MANUALIDADES":                                ("500", "SUPPLIES"),
    "MAQUINARIA BARES":                            ("300", "OPE. EQUIPMENT"),
    "ELECTRODOMESTIC O PEQUEÑO":                   ("300", "OPE. EQUIPMENT"),
    "MAQUINARIA COCINA":                           ("300", "OPE. EQUIPMENT"),
    "MAQUINARIA COMEDOR":                          ("300", "OPE. EQUIPMENT"),
    "MAQUINARIA DE CONSTRUCCION":                  ("300", "OPE. EQUIPMENT"),
    "MAQUINARIA JARDIN":                           ("300", "OPE. EQUIPMENT"),
    "MAQUINARIA LAVANDERIA":                       ("300", "OPE. EQUIPMENT"),
    "MAQUINARIA PISCINAS":                         ("300", "OPE. EQUIPMENT"),
    "MAQUINARIA PISOS Y ZONAS COMUNES":            ("300", "OPE. EQUIPMENT"),
    "MARCAS Y SUCEDANEOS":                         ("100", "FOOD"),
    "MARISCO CONGELADO":                           ("100", "FOOD"),
    "MARISCO REFRIGERADO":                         ("100", "FOOD"),
    "MATERIAL 1 SOLO USO":                         ("500", "SUPPLIES"),
    "MATERIAL DE DISFRACES":                       ("500", "SUPPLIES"),
    "PUBLICIDAD PROMOCIONAL":                      ("500", "SUPPLIES"),
    "MATERIAL DE IMPRENTA":                        ("600", "PRINTED PAPER"),
    "MATERIAL DEPORTIVO":                          ("500", "SUPPLIES"),
    "MATERIAL MERCHANDISING Y CATALOGOS":          ("500", "SUPPLIES"),
    "MATERIAL PISCINA":                            ("400", "MAINTENANCE"),
    "MECANISMOS ELECTRICOS":                       ("400", "MAINTENANCE"),
    "MENAJE COCINA":                               ("300", "OPE. EQUIPMENT"),
    "MINIATURAS MINIBARES":                        ("200", "BEVERAGES"),
    "MOBILIARIO ANIMACION":                        ("400", "MAINTENANCE"),
    "MOBILIARIO BAÑOS":                            ("400", "MAINTENANCE"),
    "MOBILIARIO COCINA - ECONOMATO":               ("400", "MAINTENANCE"),
    "MOBILIARIO COMEDOR":                          ("400", "MAINTENANCE"),
    "MOBILIARIO EXTERIOR":                         ("400", "MAINTENANCE"),
    "MOBILIARIO HABITACION":                       ("400", "MAINTENANCE"),
    "MOBILIARIO OFICINA":                          ("400", "MAINTENANCE"),
    "MOBILIARIO ZONAS COMUNES":                    ("400", "MAINTENANCE"),
    "MOUSE Y CREMOSOS":                            ("100", "FOOD"),
    "MULTIMEDIA (ANTENAS Y ACCESORIOS)":           ("301", "COMPU. EQUIPMENT"),
    "NATAS":                                       ("100", "FOOD"),
    "OTROS":                                       ("500", "SUPPLIES"),
    "OTROS ANIMACION":                             ("500", "SUPPLIES"),
    "OTROS BAÑO":                                  ("500", "SUPPLIES"),
    "OTROS ELECTRICIDAD":                          ("400", "MAINTENANCE"),
    "CABLES":                                      ("400", "MAINTENANCE"),
    "OTROS ILUMINACION":                           ("400", "MAINTENANCE"),
    "AIRE ACONDICIONADO":                          ("400", "MAINTENANCE"),
    "OTROS JARDINERIA":                            ("400", "MAINTENANCE"),
    "OTROS MANO DE OBRA EXTERNO":                  ("901", "S. MORAL PERSON"),
    "MTO DE PISCINAS EXTERNO":                     ("901", "S. MORAL PERSON"),
    "OTROS PINTURAS":                              ("400", "MAINTENANCE"),
    "PANADERIA CONGELADA":                         ("100", "FOOD"),
    "PANADERIA FRESCA":                            ("100", "FOOD"),
    "PANTALLAS Y BASES":                           ("301", "COMPU. EQUIPMENT"),
    "PAPEL Y BLOCS":                               ("501", "PAPER SUPPLIES"),
    "PASTAS":                                      ("100", "FOOD"),
    "CONFITADOS Y MERMELADAS (CONSERVAS DE FRUTAS Y DULCES)": ("100", "FOOD"),
    "PASTELERIA CONGELADA":                        ("100", "FOOD"),
    "PASTELERIA FRESCA":                           ("100", "FOOD"),
    "PESCADO CONGELADO":                           ("100", "FOOD"),
    "PESCADO FRESCO":                              ("100", "FOOD"),
    "PESCADOS AHUMADOS":                           ("100", "FOOD"),
    "PILAS":                                       ("500", "SUPPLIES"),
    "PINTURAS EXTERIOR":                           ("400", "MAINTENANCE"),
    "PINTURAS INTERIOR":                           ("400", "MAINTENANCE"),
    "PRECOCINADO PASTA CONGELADA":                 ("100", "FOOD"),
    "PRECOCINADOS CARNE":                          ("100", "FOOD"),
    "PRECOCINADOS DE QUESOS":                      ("100", "FOOD"),
    "PRECOCINADOS PASTA Y ARROZ":                  ("100", "FOOD"),
    "PRECOCINADOS PESCADO Y MARISCO":              ("100", "FOOD"),
    "PRECOCINADOS VERDURAS":                       ("100", "FOOD"),
    "PROD. ACOGIDA":                               ("500", "SUPPLIES"),
    "PROD. LIMPIEZA PISOS Y ZONAS COMUNES":        ("800", "CHEMICALS"),
    "PROD. LIMPIEZAS COCINAS":                     ("800", "CHEMICALS"),
    "PROD. LIMPIEZA ROPA":                         ("800", "CHEMICALS"),
    "PROD. PISCINAS":                              ("400", "MAINTENANCE"),
    "PRODUCTOS LIBRERIA BAZAR":                    ("500", "SUPPLIES"),
    "PRODUCTOS PARA CELIACOS":                     ("100", "FOOD"),
    "PRODUCTOS Y MATERIAL ASEO BAZAR":             ("500", "SUPPLIES"),
    "PROTECCION CONTRA DESASTRES":                 ("900", "S. PERSONAL"),
    "PUROS":                                       ("200", "BEVERAGES"),
    "QUESOS":                                      ("100", "FOOD"),
    "RECAMBIO MAQUINARIA GIMNASIO":                ("300", "OPE. EQUIPMENT"),
    "RECEPCIONISTA CABALLERO":                     ("900", "S. PERSONAL"),
    "REFRESCOS":                                   ("200", "BEVERAGES"),
    "REPARACIONES EXTERNAS":                       ("901", "S. MORAL PERSON"),
    "RON":                                         ("200", "BEVERAGES"),
    "ROPA BAÑO":                                   ("700", "UNIFORMS"),
    "ROPA CAMA":                                   ("700", "UNIFORMS"),
    "ROPA MESA":                                   ("700", "UNIFORMS"),
    "ROPA Y COMPLEMENTOS BAZAR":                   ("500", "SUPPLIES"),
    "ROPA ZONAS COMUNES":                          ("700", "UNIFORMS"),
    "SALSA":                                       ("100", "FOOD"),
    "SALSAS":                                      ("100", "FOOD"),
    "SANGRIAS":                                    ("200", "BEVERAGES"),
    "SEGURIDAD LABORAL":                           ("900", "S. PERSONAL"),
    "SIDRAS":                                      ("200", "BEVERAGES"),
    "SISTEMAS DE VIGILANCIA":                      ("900", "S. PERSONAL"),
    "SNACKS Y FRUTOS SECOS":                       ("100", "FOOD"),
    "SOBRES, ENVIOS Y ETIQUETAS":                  ("501", "PAPER SUPPLIES"),
    "SOPAS, CREMAS, PURES Y CALDOS":               ("100", "FOOD"),
    "SOUVENIRS BAZAR":                             ("500", "SUPPLIES"),
    "SUCEDANEOS":                                  ("100", "FOOD"),
    "SUELOS PARA EXTERIOR":                        ("400", "MAINTENANCE"),
    "TELEFONIA MOVIL":                             ("301", "COMPU. EQUIPMENT"),
    "TELEFONILLOS Y TIMBRES":                      ("400", "MAINTENANCE"),
    "TEQUILAS Y AGUARDIENTES":                     ("200", "BEVERAGES"),
    "TEXTIL LIMPIEZA":                             ("700", "UNIFORMS"),
    "UN SOLO USO ALIMENTACION - VARIOS":           ("500", "SUPPLIES"),
    "UN SOLO USO BEBIDAS - VARIOS":                ("500", "SUPPLIES"),
    "MATERIAL MERCHANDASING Y CATALOGOS":          ("500", "SUPPLIES"),
    "UNIFORMIDAD":                                 ("700", "UNIFORMS"),
    "UTENSILIOS LIMPIEZA":                         ("800", "CHEMICALS"),
    "VAJILLA":                                     ("300", "OPE. EQUIPMENT"),
    "VARIOS ELECTRICIDAD":                         ("400", "MAINTENANCE"),
    "VARIOS FERRETERIA":                           ("400", "MAINTENANCE"),
    "VARIOS FONTANERIA":                           ("400", "MAINTENANCE"),
    "VARIOS JARDINERIA":                           ("400", "MAINTENANCE"),
    "VARIOS MANO DE OBRA EXTERNO":                 ("901", "S. MORAL PERSON"),
    "VARIOS PINTURAS":                             ("400", "MAINTENANCE"),
    "VENTILADORES":                                ("400", "MAINTENANCE"),
    "VERDURA, PAPAS Y HORTALIZAS CONGELADAS":      ("100", "FOOD"),
    "VERDURA Y HORTALIZAS FRESCAS":                ("100", "FOOD"),
    "VINAGRES":                                    ("100", "FOOD"),
    "VINOS BLANCOS":                               ("200", "BEVERAGES"),
    "VINOS GENEROSOS":                             ("200", "BEVERAGES"),
    "VINOS ROSADOS":                               ("200", "BEVERAGES"),
    "VINOS TINTOS":                                ("200", "BEVERAGES"),
    "VODKA":                                       ("200", "BEVERAGES"),
    "WHISKYS":                                     ("200", "BEVERAGES"),
    "YOGURES":                                     ("100", "FOOD"),
    "ZUMOS":                                       ("200", "BEVERAGES"),
}

# IGIC (texto tal como aparece en Productos.xlsx) -> (codigo condicion
# fiscal SAP, descripcion). Usada en tarifa_app.py::_actualizar_listado_sap.
TABLA_IGIC_GRUPO_COMPRAS = {
    "0 EXENTO":        ("3X", "IGIC soportado 0%"),
    "03 REDUCIDO":     ("AB", "IGIC soportado 3%"),
    "05 REDUCIDO":     ("EA", "IGIC soportado 5%"),
    "15 INCREMENTADO": ("ON", "IGIC soportado 15%"),
    "07 GENERAL":      ("Q1", "IVA soportado IGIC 7%"),
    "20 ESPECIAL":     ("S5", "IVA soportado IGIC 20%"),
}
