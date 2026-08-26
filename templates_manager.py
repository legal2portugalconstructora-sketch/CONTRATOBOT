# -*- coding: utf-8 -*-
"""
templates_manager.py — "Plantillas por proyecto": guarda y recupera, por
proyecto, el Excel MODELO y los 4 documentos Word (CONTRATO, REGLAMENTO,
MEMBRESIA, DERECHOHABIENTES) para no tener que volver a subirlos cada vez.

Es solo una capa de manejo de archivos sobre disco -- no toca ninguna
lógica de negocio de excel_pipeline.py, word_pipeline.py ni
word_com_bridge.py.
"""

import os

# Carpeta base de plantillas, relativa a la UBICACIÓN DE ESTE ARCHIVO (no al
# directorio desde el que se ejecute "streamlit run"), para que funcione
# igual sin importar desde dónde se lance la app. El nombre "PLANTILLAS"
# (mayúsculas) coincide con la carpeta que el usuario ya arma a mano.
CARPETA_PLANTILLAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "PLANTILLAS")


def obtener_proyectos():
    """Lista los proyectos que aparecen en el selector de la interfaz,
    escaneando EN VIVO las subcarpetas de PLANTILLAS/ en disco (una
    carpeta = un proyecto) -- así, agregar una carpeta de proyecto nueva
    (con o sin plantillas adentro todavía) alcanza para que aparezca en
    el selector, sin tener que reiniciar el servidor ni tocar código.
    Ignora archivos sueltos y carpetas ocultas/de sistema (que empiecen
    con '.' o '$', ej. el '~$...' que deja Word abierto). Devuelve los
    nombres ordenados alfabéticamente."""
    if not os.path.isdir(CARPETA_PLANTILLAS):
        return []
    nombres = [
        entrada.name
        for entrada in os.scandir(CARPETA_PLANTILLAS)
        if entrada.is_dir() and not entrada.name.startswith((".", "$", "~"))
    ]
    return sorted(nombres)

# Nombres de archivo fijos esperados dentro de la carpeta de cada proyecto
# (deben coincidir EXACTO con lo que arma el usuario a mano en disco).
NOMBRE_ARCHIVO_MODELO_EXCEL = "MODELO.xlsx"
NOMBRE_ARCHIVO_CONTRATO = "CONTRATO.docx"
NOMBRE_ARCHIVO_REGLAMENTO = "REGLAMENTO.docx"
NOMBRE_ARCHIVO_MEMBRESIA = "MEMBRESIA.docx"
NOMBRE_ARCHIVO_DERECHOHABIENTES = "DERECHOHABIENTES.docx"

ARCHIVOS_PLANTILLA = [
    NOMBRE_ARCHIVO_MODELO_EXCEL,
    NOMBRE_ARCHIVO_CONTRATO,
    NOMBRE_ARCHIVO_REGLAMENTO,
    NOMBRE_ARCHIVO_MEMBRESIA,
    NOMBRE_ARCHIVO_DERECHOHABIENTES,
]


def nombre_carpeta_proyecto(proyecto):
    """Convierte el nombre de un proyecto (ej. 'Costa del Valle') al
    nombre de carpeta usado dentro de PLANTILLAS/ -- se mantiene el
    nombre del proyecto tal cual (mayúsculas, CON espacios, ej. 'COSTA DEL
    VALLE'), igual a como el usuario arma las carpetas a mano en disco."""
    return proyecto.strip().upper()


def _ruta_carpeta_proyecto(proyecto):
    return os.path.join(CARPETA_PLANTILLAS, nombre_carpeta_proyecto(proyecto))


def obtener_plantillas_proyecto(proyecto):
    """Devuelve un diccionario {nombre_archivo: {"existe": bool, "ruta":
    str o None}} para cada uno de los 5 archivos de plantilla esperados
    en la carpeta del proyecto. Si la carpeta de plantillas o la del
    proyecto todavía no existen, simplemente devuelve que ningún archivo
    existe (no lanza error)."""
    ruta_carpeta = _ruta_carpeta_proyecto(proyecto)
    resultado = {}
    for nombre_archivo in ARCHIVOS_PLANTILLA:
        ruta_archivo = os.path.join(ruta_carpeta, nombre_archivo)
        existe = os.path.isfile(ruta_archivo)
        resultado[nombre_archivo] = {
            "existe": existe,
            "ruta": ruta_archivo if existe else None,
        }
    return resultado


# Nombres de archivo aceptados para el fondo de un proyecto (opcional -- si
# ninguno existe en la carpeta del proyecto, simplemente no se aplica fondo).
NOMBRES_ARCHIVO_FONDO = ["fondo.jpeg", "fondo.jpg", "fondo.png"]


def obtener_ruta_fondo_proyecto(proyecto):
    """Devuelve la ruta al archivo de fondo del proyecto (el primero que
    exista de NOMBRES_ARCHIVO_FONDO), o None si no tiene ninguno guardado."""
    ruta_carpeta = _ruta_carpeta_proyecto(proyecto)
    for nombre_archivo in NOMBRES_ARCHIVO_FONDO:
        ruta = os.path.join(ruta_carpeta, nombre_archivo)
        if os.path.isfile(ruta):
            return ruta
    return None


def guardar_plantilla_proyecto(proyecto, nombre_archivo, bytes_contenido):
    """Guarda 'bytes_contenido' como 'nombre_archivo' dentro de la carpeta
    del proyecto, creando la carpeta de plantillas y/o la del proyecto si
    todavía no existen. Sobrescribe el archivo si ya existía."""
    ruta_carpeta = _ruta_carpeta_proyecto(proyecto)
    os.makedirs(ruta_carpeta, exist_ok=True)
    ruta_archivo = os.path.join(ruta_carpeta, nombre_archivo)
    with open(ruta_archivo, "wb") as f:
        f.write(bytes_contenido)
    return ruta_archivo


# =========================================================================
# CELDAS DE "Ax.A" POR PROYECTO — el Paso 4 del Excel (excel_pipeline.py,
# procesar_paso4_axa) y la lectura de UUUUUUU en Word (word_pipeline.py,
# leer_datos_para_contrato) leen/escriben celdas fijas en la hoja "Ax.A"
# del MODELO. El layout de esa hoja puede variar de un proyecto a otro,
# así que este mapeo permite tener coordenadas distintas por proyecto sin
# tocar la lógica de excel_pipeline.py ni word_pipeline.py. Si un proyecto
# no aparece en CELDAS_AXA_POR_PROYECTO, se usan las de
# CELDAS_AXA_POR_DEFECTO (las de COSTA DEL VALLE / COSTA AZUL).
CELDAS_AXA_POR_DEFECTO = {
    "e51_e54": ["E51", "E52", "E53", "E54"],  # Ax.A MODELO -- valores manuales
    "g115": "G115",                            # Ax.A MODELO -- desde Cotizador!I23
    "j119_origen": "J119",                      # Ax.A ORIGEN -- fuente de G120
    "g120": "G120",                             # Ax.A MODELO -- desde Ax.A!J119
    "g124": "G124",                             # Ax.A MODELO -- sede (Arequipa/Ilo)
    "o132": "O132",                             # Ax.A -- fuente de UUUUUUU (Word)
    # Fórmula que se escribe en "g124" según la sede elegida. Referencia
    # celdas propias de la hoja "Ax.A" de cada proyecto (ver "g124" arriba),
    # así que también varía por proyecto igual que el resto de este mapeo.
    # NOTA: se usa "UPPER" (nombre en inglés) y NO "MAYUSC" (nombre en
    # español). Dentro del .xlsx, Excel siempre guarda las fórmulas con los
    # nombres de función en inglés, sin importar el idioma de la interfaz;
    # los traduce automáticamente al mostrarlos (en Excel en español se verá
    # como "=MAYUSC(O127)"). Si se escribe "MAYUSC" directamente en el
    # archivo, Excel no lo reconoce como función válida y antepone "@"
    # (operador de intersección implícita), dejando la fórmula rota.
    "sedes_validas": {
        "Arequipa": "=UPPER(O127)",
        "Ilo": "=UPPER(O138)",
    },
    # Rangos de "Ax.A" que se pegan como imagen en el CONTRATO (ver
    # word_com_bridge.py y pages/2_🤖CONTRATO_Word.py). Partidos en 2 partes
    # para saltear una fila "hueca" de la hoja (sin alto explícito, por la
    # propiedad "zeroHeight" de Excel -- ver el aviso de la fila 89/98 más
    # abajo) que no debe aparecer en la imagen pegada.
    "rango_imagen_parte1": "B3:N97",
    "pagina_imagen_parte1": 11,
    "rango_imagen_parte2": "B99:N190",
    "pagina_imagen_parte2": 12,
}

# COSTA AZUL: desde "PRECIO DE VENTA" en adelante, la hoja "Ax.A" está
# corrida 21 filas hacia arriba respecto de COSTA DEL VALLE (confirmado
# comparando 5 textos ancla distintos: "PRECIO VALOR TOTAL DEL BIEN",
# "FECHA INICIO DE PAGO", "DEPOSITO DE INICIAL", la firma de "GIRALDO
# PORTUGAL MOINA" y "CÓNYUGE / COPROPIETARIA" -- los 5 con el mismo
# desplazamiento exacto de -21). "e51_e54" NO cambia (misma fila en ambos).
CELDAS_AXA_POR_PROYECTO = {
    "COSTA AZUL": {
        "g115": "G94",
        "j119_origen": "J98",
        "g120": "G99",
        "g124": "G103",
        "o132": "O111",
        "sedes_validas": {
            "Arequipa": "=UPPER(O106)",
            "Ilo": "=UPPER(O117)",
        },
        # La fila 89 de "Ax.A" no tiene alto explícito (zeroHeight la deja en
        # 0) y no debe aparecer en la imagen pegada en el CONTRATO -- se
        # salta partiendo el rango en B3:N88 / B90:N169 (equivalente a como
        # COSTA DEL VALLE ya salta su fila 98 con B3:N97 / B99:N190).
        "rango_imagen_parte1": "B3:N88",
        "rango_imagen_parte2": "B90:N169",
    },
}


def obtener_celdas_axa(proyecto):
    """Devuelve el mapeo de celdas de 'Ax.A' a usar para 'proyecto': las
    de CELDAS_AXA_POR_DEFECTO, sobrescritas por las que ese proyecto tenga
    en CELDAS_AXA_POR_PROYECTO (si tiene alguna). Los valores que son a su
    vez diccionarios (ej. "sedes_validas") se combinan un nivel, para que
    un proyecto pueda sobrescribir solo una sede sin repetir la otra."""
    especificas = CELDAS_AXA_POR_PROYECTO.get(nombre_carpeta_proyecto(proyecto), {})
    resultado = dict(CELDAS_AXA_POR_DEFECTO)
    for clave, valor in especificas.items():
        if isinstance(valor, dict) and isinstance(resultado.get(clave), dict):
            resultado[clave] = {**resultado[clave], **valor}
        else:
            resultado[clave] = valor
    return resultado
