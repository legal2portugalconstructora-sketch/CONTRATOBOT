# -*- coding: utf-8 -*-
"""ui_utils.py — Utilidades de interfaz compartidas entre las páginas."""

import io
import os

import streamlit as st

from templates_manager import (
    obtener_proyectos,
    obtener_plantillas_proyecto,
    obtener_ruta_fondo_proyecto,
    NOMBRE_ARCHIVO_MODELO_EXCEL,
    NOMBRE_ARCHIVO_CONTRATO,
    NOMBRE_ARCHIVO_REGLAMENTO,
    NOMBRE_ARCHIVO_MEMBRESIA,
    NOMBRE_ARCHIVO_DERECHOHABIENTES,
)

# Ruta absoluta (no depende del directorio desde el que se lance
# "streamlit run") al logo de la app, usado como ícono de pestaña y en la
# barra lateral de todas las páginas.
RUTA_LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "CONTRATOBOT.png")

# Guía paso a paso en PDF para pages/4_🤖CONTRATO_Ayuda.py: se agrega a
# mano en la carpeta principal de la app con este nombre exacto (mismo
# patrón que las plantillas de PLANTILLAS/<PROYECTO>/ -- no se sube desde
# la propia app). Si el archivo no existe, la página de Ayuda simplemente
# no muestra esa sección.
RUTA_GUIA_PDF = os.path.join(os.path.dirname(os.path.abspath(__file__)), "GUIA_PASO_A_PASO.pdf")

# Opción del selector de proyecto para cuando el proyecto no tiene carpeta
# en templates_manager.obtener_proyectos() (o el usuario no quiere usar el
# sistema de plantillas para esta corrida).
OPCION_PROYECTO_OTRO = "Otro / no listado"

# Etiqueta corta a mostrar en el resumen de plantillas, por cada nombre de
# archivo fijo de templates_manager.py.
ETIQUETAS_ARCHIVOS_PLANTILLA = [
    ("MODELO EXCEL", NOMBRE_ARCHIVO_MODELO_EXCEL),
    ("CONTRATO", NOMBRE_ARCHIVO_CONTRATO),
    ("REGLAMENTO", NOMBRE_ARCHIVO_REGLAMENTO),
    ("MEMBRESIA", NOMBRE_ARCHIVO_MEMBRESIA),
    ("DERECHOHABIENTES", NOMBRE_ARCHIVO_DERECHOHABIENTES),
]

# Los 3 módulos de trabajo de la app. Fuente ÚNICA para Inicio.py (tarjetas
# de módulo) y pages/4_🤖CONTRATO_Ayuda.py (guía de uso) -- así ninguna de
# las dos se puede desactualizar respecto de la otra: si el día de mañana
# se agrega/saca un módulo, se edita acá una sola vez.
MODULOS = [
    {
        "titulo": "Contrato Excel",
        "descripcion": "Pasos 1 a 6: cotizador, ficha, medidas de planos y Ax.B, sobre el Excel del cliente.",
        "icono": "📊",
        "color": "#16A34A",
        "pagina": "pages/1_🤖CONTRATO_Excel.py",
    },
    {
        "titulo": "Contrato Word",
        "descripcion": "Generá CONTRATO, REGLAMENTO, MEMBRESIA y DERECHOHABIENTES desde el Excel ya procesado.",
        "icono": "📝",
        "color": "#2563EB",
        "pagina": "pages/2_🤖CONTRATO_Word.py",
    },
    {
        "titulo": "Contrato PDF",
        "descripcion": "Convertí a PDF hasta 4 Word ya revisados a mano, sin cambiarles el nombre.",
        "icono": "🧾",
        "color": "#D97706",
        "pagina": "pages/3_🤖CONTRATO_PDF.py",
    },
]

ACCIONES_RAPIDAS = [
    {"label": "Nuevo Excel", "icono": "📊", "pagina": "pages/1_🤖CONTRATO_Excel.py"},
    {"label": "Nuevo Word", "icono": "📝", "pagina": "pages/2_🤖CONTRATO_Word.py"},
    {"label": "Convertir a PDF", "icono": "🧾", "pagina": "pages/3_🤖CONTRATO_PDF.py"},
]

# Preguntas frecuentes. Fuente ÚNICA para pages/4_🤖CONTRATO_Ayuda.py (las
# muestra como expanders) y para el buscador de Inicio.py (busca en
# "pregunta" y "respuesta") -- mismo motivo que MODULOS/ACCIONES_RAPIDAS:
# que ninguna de las dos se pueda desactualizar respecto de la otra.
PREGUNTAS_FRECUENTES = [
    {
        "pregunta": "El procesamiento del Excel/Word se queda 'cargando' mucho tiempo",
        "respuesta": (
            "Puede pasar si tenés otra ventana de Excel o Word ya abierta en la misma "
            "computadora (la app usa Excel/Word reales por detrás). Cerrá cualquier "
            "Excel/Word abierto manualmente y probá de nuevo."
        ),
    },
    {
        "pregunta": "'Uno de los archivos no es un .xlsx válido'",
        "respuesta": (
            "El archivo que subiste no es un Excel (.xlsx) real, o está corrupto. "
            "Confirmá que sea el archivo correcto y que no se haya guardado en otro "
            "formato (ej. .xls viejo)."
        ),
    },
    {
        "pregunta": "'No hay plantilla de X guardada para <proyecto>'",
        "respuesta": (
            "Falta agregar ese archivo a mano en PLANTILLAS/<PROYECTO>/ con el nombre "
            "exacto que pide el aviso. Mirá el apartado \"Proyectos configurados\" de "
            "la página de Ayuda para ver qué le falta a cada proyecto."
        ),
    },
    {
        "pregunta": "¿Por qué en Word tarda tanto el botón 'Procesar CONTRATO'?",
        "respuesta": (
            "Antes de generar los documentos, recalcula TODAS las fórmulas del Excel "
            "con Excel real (necesario porque Excel de datos complejos no se recalcula "
            "solo al guardarse con otras herramientas). Es el paso más lento de toda "
            "la app."
        ),
    },
]


def aplicar_tema_base():
    """Emite el <style> con las variables de tema (--cb-*, claro y oscuro
    según prefers-color-scheme) y las reglas de layout base (fondo, sidebar,
    color de texto, tarjetas con borde, alertas) que las 3 páginas
    (Excel/Word/PDF) usaban repetidas de forma casi idéntica. Hay que
    llamarla al principio del bloque de estilo de cada página, ANTES de
    cualquier <style> con reglas propias de esa página (ej. tratamiento del
    h1, stepper, colores de botones) -- así, si una regla propia de la
    página pisa algo de acá con el mismo selector, gana por orden en el
    documento (misma especificidad + !important), no por casualidad.

    Los colores del botón primario/de descarga y del expander SÍ se dejaron
    fuera a propósito: son distintos entre Excel (acento morado) y
    Word/PDF (verde), así que centralizarlos cambiaría cómo se ven."""
    st.markdown(
        """
        <style>
        :root {
            --cb-bg: #ffffff;
            --cb-bg-elevated: #f6f7f9;
            --cb-card: #ffffff;
            --cb-border: #e3e5ea;
            --cb-text: #16181d;
            --cb-text-muted: #6b7280;
            --cb-accent: #6d6bf7;
            --cb-accent-hover: #5654e0;
            --cb-success: #16a34a;
            --cb-warning: #d97706;
            --cb-overlay: rgba(255, 255, 255, 0.45);
        }

        @media (prefers-color-scheme: dark) {
            :root {
                --cb-bg: #0c0e13;
                --cb-bg-elevated: #12151d;
                --cb-card: #171b26;
                --cb-border: #262c3b;
                --cb-text: #e8eaee;
                --cb-text-muted: #8b93a6;
                --cb-accent: #6d6bf7;
                --cb-accent-hover: #5654e0;
                --cb-success: #22c55e;
                --cb-warning: #f59e0b;
                --cb-overlay: rgba(8, 9, 12, 0.72);
            }
        }

        [data-testid="stAppViewContainer"], .stApp { background: var(--cb-bg) !important; }
        [data-testid="stHeader"] { background: transparent !important; }
        [data-testid="stSidebar"] {
            background: var(--cb-bg-elevated) !important;
            border-right: 1px solid var(--cb-border);
        }
        [data-testid="stMain"] h1, [data-testid="stMain"] h2, [data-testid="stMain"] h3,
        [data-testid="stMain"] h4, [data-testid="stMain"] h5, [data-testid="stMain"] p,
        [data-testid="stMain"] span, [data-testid="stMain"] label, [data-testid="stSidebar"] * {
            color: var(--cb-text) !important;
        }
        [data-testid="stCaptionContainer"] { color: var(--cb-text-muted) !important; }

        [data-testid="stVerticalBlockBorderWrapper"] {
            background: var(--cb-card) !important;
            border: 1px solid var(--cb-border) !important;
            border-radius: 14px !important;
        }
        [data-testid="stVerticalBlockBorderWrapper"] > div { gap: 0.9rem; }

        [data-testid="stAlert"] {
            background: var(--cb-card) !important;
            border: 1px solid var(--cb-border) !important;
            border-radius: 10px !important;
        }
        [data-testid="stAlert"] p, [data-testid="stAlert"] span {
            color: var(--cb-text) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def mostrar_logo():
    """Muestra el logo de la app (CONTRATOBOT.png) en la barra lateral.
    Hay que llamarla en cada página (Inicio.py y cada archivo de
    pages/), porque Streamlit no lo persiste solo entre páginas."""
    st.logo(RUTA_LOGO)


def selector_proyecto():
    """Muestra el selector "📁 Proyecto" (persistente en session_state bajo
    la key 'proyecto_seleccionado', así que se recuerda al navegar entre
    páginas) más un resumen de qué plantillas (Excel MODELO + los 4 Word)
    ya tiene guardadas ese proyecto y cuáles faltan. Devuelve el nombre
    del proyecto elegido, o OPCION_PROYECTO_OTRO si se eligió esa opción."""
    opciones = obtener_proyectos() + [OPCION_PROYECTO_OTRO]
    proyecto = st.selectbox("📁 Proyecto", options=opciones, key="proyecto_seleccionado")
    # Espejo en una key "plana" (no ligada a un widget): Streamlit limpia el
    # valor de un widget de session_state si pasan 2+ corridas sin que se
    # vuelva a instanciar (ej. un wizard por pasos que deja de mostrar este
    # selectbox varios pasos) -- una key plana no tiene ese problema.
    st.session_state["proyecto_actual"] = proyecto

    if proyecto != OPCION_PROYECTO_OTRO:
        plantillas = obtener_plantillas_proyecto(proyecto)
        resumen = " · ".join(
            f"{etiqueta} {'✅' if plantillas[archivo]['existe'] else '❌'}"
            for etiqueta, archivo in ETIQUETAS_ARCHIVOS_PLANTILLA
        )
        st.caption(resumen)

    return proyecto


def aplicar_fondo_proyecto(proyecto):
    """Si 'proyecto' tiene un archivo de fondo guardado (ver
    templates_manager.obtener_ruta_fondo_proyecto -- fondo.jpeg/jpg/png en su
    carpeta de PLANTILLAS), lo aplica como imagen de fondo de toda la página,
    con un velo oscuro semitransparente encima para que las tarjetas y el
    texto sigan siendo legibles. Si el proyecto no tiene fondo guardado (o
    no se eligió ninguno todavía), deja explícitamente el fondo liso normal.

    SIEMPRE emite el bloque <style> (con la imagen, o con 'none' si no hay
    imagen) en vez de omitirlo condicionalmente: si un proyecto CON fondo se
    cambia por uno SIN fondo, Streamlit no borra solo el <style> de la
    corrida anterior en esa misma posición del script -- queda "pegado" el
    fondo del proyecto anterior hasta que se lo pisa explícitamente.

    Hay que llamarla DESPUÉS del <style> principal de la página (para que
    esta regla, al aparecer más abajo en el HTML, le gane en la cascada de
    CSS a la del fondo liso -- ambas usan !important, así que gana la
    última en el documento)."""
    ruta = None
    if proyecto and proyecto != OPCION_PROYECTO_OTRO:
        ruta = obtener_ruta_fondo_proyecto(proyecto)

    if ruta is None:
        st.markdown(
            """
            <style>
            [data-testid="stAppViewContainer"] { background-image: none !important; }
            </style>
            """,
            unsafe_allow_html=True,
        )
        return

    import base64
    with open(ruta, "rb") as f:
        datos_b64 = base64.b64encode(f.read()).decode("ascii")
    extension = os.path.splitext(ruta)[1].lower().lstrip(".")
    mime = "jpeg" if extension in ("jpg", "jpeg") else extension

    st.markdown(
        f"""
        <style>
        [data-testid="stAppViewContainer"] {{
            background-image:
                linear-gradient(var(--cb-overlay), var(--cb-overlay)),
                url("data:image/{mime};base64,{datos_b64}") !important;
            background-size: cover !important;
            background-position: center top !important;
            background-attachment: fixed !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def cargar_plantilla_o_aviso(proyecto, nombre_archivo_plantilla, etiqueta_documento):
    """Carga la plantilla guardada del proyecto DESDE DISCO -- no ofrece
    subirla ni guardarla a través de la página (esa función quedó
    deshabilitada a pedido del usuario; las plantillas se agregan a mano
    en la carpeta PLANTILLAS/<PROYECTO>/ de templates_manager.py).

    - Si el proyecto tiene 'nombre_archivo_plantilla' guardado: lo carga
      y muestra "✅ Usando plantilla de {proyecto}".
    - Si no (o el proyecto es OPCION_PROYECTO_OTRO): muestra un aviso de
      que falta agregar ese archivo a mano en la carpeta PLANTILLAS, y
      devuelve None.

    Devuelve un objeto tipo-archivo (BytesIO) listo para usar, o None.
    """
    if proyecto == OPCION_PROYECTO_OTRO:
        st.info(
            f"ℹ️ Elegí un proyecto de la lista para poder usar su plantilla de {etiqueta_documento}."
        )
        return None

    plantilla_info = obtener_plantillas_proyecto(proyecto)[nombre_archivo_plantilla]
    if not plantilla_info["existe"]:
        st.warning(
            f"⚠️ No hay plantilla de {etiqueta_documento} guardada para '{proyecto}'. "
            f"Agregá el archivo '{nombre_archivo_plantilla}' a mano en la carpeta "
            f"PLANTILLAS/{proyecto}/."
        )
        return None

    with open(plantilla_info["ruta"], "rb") as f:
        archivo = io.BytesIO(f.read())
    st.success(f"✅ Usando plantilla de {proyecto}")
    return archivo


def _renderizar_alertas(alertas):
    """Muestra las alertas con letra más grande, en vez del st.error()
    normal (que se ve chico). Escapa HTML para que el contenido de las
    alertas no rompa el render ni permita inyectar código."""
    import html as _html_mod

    if not alertas:
        st.success("✅ ¡Datos perfectos! No se encontraron inconsistencias.")
        return

    st.markdown(
        "<div style='font-size:1.5rem; font-weight:700; color:#842029; margin:0.75rem 0 0.5rem 0;'>"
        "⚠️ Alertas de Auditoría Detectadas</div>",
        unsafe_allow_html=True,
    )
    for al in alertas:
        st.markdown(
            "<div style='font-size:1.15rem; line-height:1.5; background-color:#f8d7da; "
            "color:#58151c; border:1px solid #f1aeb5; border-radius:8px; "
            "padding:14px 18px; margin-bottom:10px;'>"
            f"{_html_mod.escape(al)}</div>",
            unsafe_allow_html=True,
        )
