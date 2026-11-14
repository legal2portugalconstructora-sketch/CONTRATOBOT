# -*- coding: utf-8 -*-
"""Página "🤖CONTRATO – Word": generación/edición de documentos Word.

Rediseño visual (wizard por pasos, tarjetas oscuras, mejor jerarquía),
igual patrón que pages/1_🤖CONTRATO_Excel.py. La lógica de negocio y el
flujo de procesamiento NO cambiaron -- esto es solo la capa de interfaz.
"""

import io
import os
import time

import streamlit as st

from auth_utils import requerir_login, mostrar_boton_logout
from historial_contratos import (
    extraer_lote_y_nombre,
    guardar_copia,
    nombre_archivo_descarga,
    registrar_numero_membresia,
    obtener_ultimo_numero_membresia,
    obtener_historial_membresias,
)
from stats_tracker import registrar_generacion, modo_pruebas_activo, establecer_modo_pruebas
from n8n_integracion import registrar_contrato_en_n8n
from excel_pipeline import HOJA_AXA, HOJA_AXB
from word_pipeline import (
    OPCIONES_PROPIEDAD,
    procesar_paso1_unico_propietario,
    procesar_paso2_casados_mismo_sitio,
    procesar_paso3_copropietarios_no_casados_mismo_sitio,
    procesar_paso4_copropietarios_no_casados_sitios_diferentes,
    procesar_reglamento,
    procesar_membresia,
    procesar_derechohabientes,
    leer_datos_conyuge_paso2,
    leer_datos_extra_paso4,
    MEMBRESIA_REGLAMENTO_VALORES,
)
from word_com_bridge import (
    RANGO_LISTAX_IMAGEN,
    HOJA_LISTAX_IMAGEN,
    PAGINA_CONTRATO_IMAGEN,
    RANGO_AXB_PARTE1_IMAGEN,
    PAGINA_AXB_PARTE1_IMAGEN,
    RANGO_AXB_PARTE2_IMAGEN,
    PAGINA_AXB_PARTE2_IMAGEN,
    HOJA_AXE_IMAGEN,
    RANGO_AXE_IMAGEN,
    PAGINA_AXE_IMAGEN,
    agregar_imagenes_a_contrato,
    recalcular_excel_con_com,
    recortar_contrato_a_maximo_paginas,
    MAX_PAGINAS_CONTRATO,
)
from templates_manager import (
    NOMBRE_ARCHIVO_CONTRATO,
    NOMBRE_ARCHIVO_REGLAMENTO,
    NOMBRE_ARCHIVO_MEMBRESIA,
    NOMBRE_ARCHIVO_DERECHOHABIENTES,
    obtener_celdas_axa,
)
from ui_utils import (
    aplicar_tema_base,
    mostrar_logo,
    RUTA_LOGO,
    selector_proyecto,
    cargar_plantilla_o_aviso,
    aplicar_fondo_proyecto,
)

# Qué función de word_pipeline.py procesa el CONTRATO según la opción de
# "Tipo de propiedad" elegida (ver OPCIONES_PROPIEDAD). Las opciones que
# todavía no están aquí muestran una advertencia en vez de procesarse.
FUNCIONES_PASO_CONTRATO = {
    "UNICO PROPIETARIO": procesar_paso1_unico_propietario,
    "CASADOS QUE VIVEN EN EL MISMO SITIO": procesar_paso2_casados_mismo_sitio,
    "COPROPIETARIOS NO CASADOS QUE VIVEN EN EL MISMO SITIO": procesar_paso3_copropietarios_no_casados_mismo_sitio,
    "COPROPIETARIOS NO CASADOS QUE VIVEN EN SITIOS DIFERENTES": procesar_paso4_copropietarios_no_casados_sitios_diferentes,
}

# Qué función de word_pipeline.py da los VALORES EXTRA (además de los 7
# placeholders de siempre) a aplicar en REGLAMENTO/MEMBRESIA/DERECHOHABIENTES
# según el Paso elegido (el CONTRATO ya resuelve esto internamente en su
# propia función de FUNCIONES_PASO_CONTRATO). Los Pasos que no están aquí
# no agregan valores extra.
FUNCIONES_VALORES_EXTRA_POR_PASO = {
    "CASADOS QUE VIVEN EN EL MISMO SITIO": leer_datos_conyuge_paso2,
    "COPROPIETARIOS NO CASADOS QUE VIVEN EN EL MISMO SITIO": leer_datos_conyuge_paso2,
    "COPROPIETARIOS NO CASADOS QUE VIVEN EN SITIOS DIFERENTES": leer_datos_extra_paso4,
}

st.set_page_config(page_title="🤖CONTRATO – Word", page_icon=RUTA_LOGO, layout="wide")

authenticator = requerir_login()
mostrar_boton_logout(authenticator)

mostrar_logo()

with st.sidebar:
    st.divider()
    _modo_pruebas_guardado = modo_pruebas_activo()
    _modo_pruebas_nuevo = st.checkbox(
        "🧪 Modo de pruebas", value=_modo_pruebas_guardado, key="sb_modo_pruebas_word",
        help=(
            "Mientras esté activo, lo que proceses acá no se cuenta en las estadísticas "
            "ni se archiva en CONTRATOS\\ ni en el historial de membresías -- para nadie, "
            "hasta que lo desactives."
        ),
    )
    if _modo_pruebas_nuevo != _modo_pruebas_guardado:
        establecer_modo_pruebas(_modo_pruebas_nuevo)

if modo_pruebas_activo():
    st.warning("🧪 **Modo de pruebas activo** — lo que proceses acá no se cuenta en las estadísticas.")

# =========================================================================
# ESTILO — variables de tema y layout base vienen de aplicar_tema_base()
# (compartida con las otras 2 páginas); acá solo quedan las reglas propias
# de ESTA página (encabezado, expander, botones verdes de descarga,
# stepper, inputs). Solo afecta esta página (Streamlit vuelve a montar el
# DOM en cada navegación de página). No se tocó ninguna regla de color de
# los botones DESCARGAR (siguen verdes) -- solo se sumaron reglas nuevas.
# =========================================================================
aplicar_tema_base()
st.markdown(
    """
    <style>
    [data-testid="stExpander"] {
        background: var(--cb-bg-elevated) !important;
        border: 1px solid var(--cb-border) !important;
        border-radius: 12px !important;
        overflow: hidden;
    }

    /* Encabezado de página: como queda directo sobre la imagen de fondo
       por proyecto (no dentro de ninguna tarjeta), en modo claro se le da
       fondo propio -- en vez de depender del velo general. En modo oscuro
       se revierte más abajo a como estaba (transparente, sin tarjeta). */
    .cb-page-header {
        display: flex;
        flex-direction: column;
        gap: 0.2rem;
        margin-bottom: 1.2rem;
        background: rgba(255, 255, 255, 0.88) !important;
        padding: 0.9rem 1.2rem;
        border-radius: 14px;
        border: 1px solid var(--cb-border);
    }
    .cb-page-header .cb-kicker {
        text-transform: uppercase;
        letter-spacing: 0.08em;
        font-size: 0.72rem;
        font-weight: 700;
        color: var(--cb-accent) !important;
    }
    .cb-page-header .cb-title {
        font-size: 1.9rem;
        font-weight: 700;
        margin: 0;
    }
    .cb-page-header .cb-subtitle {
        color: var(--cb-text-muted) !important;
        font-size: 0.98rem;
        max-width: 62ch;
    }
    /* Revierte el fondo del encabezado en modo oscuro -- queda EXACTAMENTE
       como estaba antes (transparente, sin tarjeta). Va DESPUÉS de la
       regla de arriba a propósito, para ganarle en la cascada (misma
       especificidad + !important en ambas, gana la última en el documento). */
    @media (prefers-color-scheme: dark) {
        .cb-page-header {
            background: transparent !important;
            padding: 0;
            border-radius: 0;
            border: none;
        }
    }

    /* Etiquetas de subsección dentro de una tarjeta */
    .cb-subsection-label {
        font-size: 0.78rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: var(--cb-text-muted) !important;
        margin: 0.2rem 0 0.1rem;
    }

    /* Botones */
    .stButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        border: 1px solid var(--cb-border) !important;
    }
    .stButton > button[kind="primary"] {
        background: var(--cb-accent) !important;
        border: none !important;
        color: white !important;
        font-size: 1.15rem !important;
        padding: 0.75rem 1.5rem !important;
    }
    .stButton > button[kind="primary"] * { color: white !important; }
    .stButton > button[kind="primary"]:hover:not(:disabled) { background: var(--cb-accent-hover) !important; }
    .stButton > button[kind="secondary"] * { color: var(--cb-text) !important; }

    /* Botones de descarga grandes y verdes, para los 4 documentos (igual que antes) */
    div[data-testid="stDownloadButton"] button {
        font-size: 1.4rem;
        font-weight: 700;
        padding: 0.9rem 2rem;
        background-color: #1e8e3e;
        color: white;
        border: none;
        border-radius: 10px;
    }
    div[data-testid="stDownloadButton"] button * { color: white !important; }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #17702f;
        color: white;
    }
    div[data-testid="stDownloadButton"] button:active {
        background-color: #145c27;
        color: white;
    }

    /* Inputs */
    [data-baseweb="select"] > div, .stTextInput input, [data-testid="stFileUploaderDropzone"] {
        background: var(--cb-bg-elevated) !important;
        border-color: var(--cb-border) !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        border: 2px dashed var(--cb-border) !important;
        border-radius: 12px !important;
    }

    /* Casillas de verificación más grandes */
    [data-testid="stCheckbox"] {
        transform: scale(1.35);
        transform-origin: left top;
        margin: 0.3rem 0 1.1rem 0;
    }

    /* Stepper superior (igual que en la página Excel) */
    .cb-stepper { display: flex; align-items: center; margin: 0.25rem 0 2rem; }
    .cb-step { display: flex; align-items: center; gap: 0.55rem; }
    .cb-step-circle {
        width: 30px; height: 30px; border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        font-size: 0.8rem; font-weight: 700;
        border: 2px solid var(--cb-border); color: var(--cb-text-muted);
        background: var(--cb-card); flex-shrink: 0;
    }
    .cb-step.active .cb-step-circle { border-color: var(--cb-accent); color: white; background: var(--cb-accent); }
    .cb-step.done .cb-step-circle { border-color: var(--cb-success); color: white; background: var(--cb-success); }
    .cb-step-label { font-size: 0.8rem; color: var(--cb-text-muted); white-space: nowrap; }
    .cb-step.active .cb-step-label { color: var(--cb-text); font-weight: 700; }
    .cb-step.done .cb-step-label { color: var(--cb-text-muted); }
    .cb-step-line { flex: 1; height: 2px; background: var(--cb-border); margin: 0 0.6rem; }
    .cb-step-line.done { background: var(--cb-success); }

    @media (max-width: 800px) {
        .cb-step-label { display: none; }
        .cb-step-line { margin: 0 0.3rem; }
    }

    .cb-card {
        background: var(--cb-card);
        border: 1px solid var(--cb-border);
        border-radius: 14px;
        padding: 1.1rem 1.3rem;
        margin-bottom: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="cb-page-header">
        <span class="cb-kicker">Área legal · Generación de documentos</span>
        <h1 class="cb-title">🤖CONTRATO – Word</h1>
        <p class="cb-subtitle">
            Generá el CONTRATO, REGLAMENTO, MEMBRESIA y DERECHOHABIENTES de un cliente a partir
            de un solo Excel de datos, paso a paso.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)


def sv(clave, defecto=None):
    """Lee 'clave' de session_state con un valor por defecto, sin importar
    si el widget que la puebla se renderizó en esta corrida (necesario
    porque el wizard solo muestra los widgets del paso actual)."""
    return st.session_state.get(clave, defecto)


# =========================================================================
# ESTADO PERSISTENTE
# =========================================================================
if "word_wizard_step" not in st.session_state:
    st.session_state["word_wizard_step"] = 0
if "word_resultado_bytes" not in st.session_state:
    st.session_state["word_resultado_bytes"] = None
    st.session_state["word_info"] = None
if "reglamento_resultado_bytes" not in st.session_state:
    st.session_state["reglamento_resultado_bytes"] = None
    st.session_state["reglamento_info"] = None
if "membresia_resultado_bytes" not in st.session_state:
    st.session_state["membresia_resultado_bytes"] = None
    st.session_state["membresia_info"] = None
if "derechohabientes_resultado_bytes" not in st.session_state:
    st.session_state["derechohabientes_resultado_bytes"] = None
    st.session_state["derechohabientes_info"] = None

# =========================================================================
# STEPPER
# =========================================================================
NOMBRES_PASOS_WIZARD_WORD = ["📁 Proyecto", "🧾 Datos", "⚙️ Configuración", "📋 Resumen", "✅ Resultado"]
paso_actual = st.session_state["word_wizard_step"]

html_stepper = ['<div class="cb-stepper">']
for i, nombre in enumerate(NOMBRES_PASOS_WIZARD_WORD):
    estado = "done" if i < paso_actual else ("active" if i == paso_actual else "")
    icono = "✓" if estado == "done" else str(i + 1)
    html_stepper.append(f'<div class="cb-step {estado}"><div class="cb-step-circle">{icono}</div>'
                         f'<div class="cb-step-label">{nombre}</div></div>')
    if i < len(NOMBRES_PASOS_WIZARD_WORD) - 1:
        linea_done = "done" if i < paso_actual else ""
        html_stepper.append(f'<div class="cb-step-line {linea_done}"></div>')
html_stepper.append("</div>")
st.markdown("".join(html_stepper), unsafe_allow_html=True)


def _ir_a_paso(nuevo_paso):
    st.session_state["word_wizard_step"] = nuevo_paso
    st.rerun()


def _alertar_placeholders_sin_valor(info):
    for placeholder in info.get("placeholders_sin_valor", []):
        st.warning(
            f"⚠️ No se pudo completar el dato de '{placeholder}': la celda del Excel vino "
            f"vacía o sin valor (None)."
        )


# =========================================================================
# PASO 0 — Proyecto y plantillas
# =========================================================================
if paso_actual == 0:
    with st.container(border=True):
        st.subheader("📁 Proyecto y plantillas")
        st.caption("Elegí el proyecto para cargar automáticamente sus documentos base guardados.")
        proyecto_seleccionado = selector_proyecto()

        st.markdown('<p class="cb-subsection-label">Documentos base</p>', unsafe_allow_html=True)
        colw1, colw2, colw3, colw4 = st.columns(4)
        with colw1:
            st.markdown("**CONTRATO**")
            doc_contrato_upload = cargar_plantilla_o_aviso(proyecto_seleccionado, NOMBRE_ARCHIVO_CONTRATO, "CONTRATO")
        with colw2:
            st.markdown("**REGLAMENTO**")
            cargar_plantilla_o_aviso(proyecto_seleccionado, NOMBRE_ARCHIVO_REGLAMENTO, "REGLAMENTO")
        with colw3:
            st.markdown("**MEMBRESIA**")
            cargar_plantilla_o_aviso(proyecto_seleccionado, NOMBRE_ARCHIVO_MEMBRESIA, "MEMBRESIA")
        with colw4:
            st.markdown("**DERECHOHABIENTES**")
            cargar_plantilla_o_aviso(proyecto_seleccionado, NOMBRE_ARCHIVO_DERECHOHABIENTES, "DERECHOHABIENTES")

    _, col_next = st.columns([3, 1])
    with col_next:
        if st.button("Siguiente →", type="primary", disabled=doc_contrato_upload is None, width="stretch"):
            _ir_a_paso(1)

# =========================================================================
# PASO 1 — Datos de este cliente/cotización
# =========================================================================
elif paso_actual == 1:
    with st.container(border=True):
        st.subheader("🧾 Datos de este cliente/cotización")
        excel_word_upload = st.file_uploader(
            "📥 Sube el Excel con los datos (hoja 'DATOS')", type=["xlsx"], key="word_excel"
        )
        st.caption(
            "⚠️ Al presionar 'Procesar', este Excel se recalcula automáticamente abriéndolo con "
            "Excel real (requiere Windows + Excel instalados en esta máquina) antes de leer los "
            "datos — así fórmulas como DATOS!H26:H28 o Ax.A!O132 quedan con su valor actualizado, "
            "sin que tengas que abrirlo y guardarlo vos mismo."
        )
        # Espejo en key plana (Streamlit limpia el valor de un widget de
        # session_state si pasan 2+ corridas sin volver a instanciarlo, y
        # este archivo se necesita recién en el Paso 3 del wizard).
        st.session_state["m_excel_word"] = excel_word_upload

    col_back, _, col_next = st.columns([1, 2, 1])
    with col_back:
        if st.button("← Atrás", width="stretch"):
            _ir_a_paso(0)
    with col_next:
        if st.button("Siguiente →", type="primary", disabled=excel_word_upload is None, width="stretch"):
            _ir_a_paso(2)

# =========================================================================
# PASO 2 — Configuración del contrato
# =========================================================================
elif paso_actual == 2:
    with st.container(border=True):
        st.subheader("⚙️ Configuración del contrato")

        st.markdown('<p class="cb-subsection-label">🏠 Tipo de propiedad</p>', unsafe_allow_html=True)
        opcion_propiedad = st.radio(
            "Selecciona una opción", options=OPCIONES_PROPIEDAD, key="radio_tipo_propiedad"
        )

        st.divider()

        st.markdown('<p class="cb-subsection-label">📕 Membresía (CONTRATO y REGLAMENTO)</p>', unsafe_allow_html=True)
        opcion_membresia_reglamento = st.radio(
            "¿El cliente tiene membresía?",
            options=list(MEMBRESIA_REGLAMENTO_VALORES.keys()),
            horizontal=True,
            key="radio_membresia_reglamento",
            help="Se aplica al CONTRATO y al REGLAMENTO, para los 4 tipos de propiedad. En "
                 "el documento, reemplaza el texto 'JDIP' por 'ANEXO F' (con membresía) o "
                 "'ANEXO E' (sin membresía).",
        )

        st.markdown('<p class="cb-subsection-label">🪪 MEMBRESIA</p>', unsafe_allow_html=True)
        numero_membresia = st.text_input(
            "Número de Membresía",
            key="input_numero_membresia",
            help="Reemplaza el texto 'GTAA' dentro del documento de MEMBRESIA.",
        )
        st.caption("Opcional — dejalo en blanco si el cliente no tiene membresía.")

        _ultima_membresia = obtener_ultimo_numero_membresia()
        if _ultima_membresia:
            _fecha_ultima = _ultima_membresia["timestamp"].split("T")[0]
            st.caption(
                f'📜 Último usado: **{_ultima_membresia["numero_membresia"]}** '
                f'({_ultima_membresia["nombre_cliente"]}, {_fecha_ultima})'
            )
            with st.expander("Ver historial completo de números de membresía"):
                for _item in obtener_historial_membresias(limite=20):
                    st.write(
                        f'{_item["timestamp"].split("T")[0]} — **{_item["numero_membresia"]}** — '
                        f'{_item["nombre_cliente"]} ({_item["proyecto"]})'
                    )
        else:
            st.caption("📜 Todavía no hay ningún número de membresía registrado.")

        st.session_state["m_opcion_propiedad"] = opcion_propiedad
        st.session_state["m_membresia_reglamento"] = opcion_membresia_reglamento
        st.session_state["m_numero_membresia"] = numero_membresia

    with st.container(border=True):
        st.subheader("📎 Imágenes de Excel")
        pegar_imagenes_excel = st.checkbox(
            "📎 Pegar las imágenes de Excel (List.Ax, Ax.A, Ax.B, Ax.E) en el CONTRATO",
            value=False,
            key="chk_pegar_imagenes",
        )
        if pegar_imagenes_excel:
            st.caption(
                "⚠️ Esto usa Excel y Word REALES vía pywin32, así que solo funciona si esta app "
                "corre en Windows con Excel y Word instalados en esta misma máquina — no funcionará "
                "en Streamlit Cloud ni en un servidor Linux."
            )
        st.session_state["m_pegar_imagenes"] = pegar_imagenes_excel

    col_back, _, col_next = st.columns([1, 2, 1])
    with col_back:
        if st.button("← Atrás", width="stretch"):
            _ir_a_paso(1)
    with col_next:
        if st.button("Siguiente →", type="primary", width="stretch"):
            _ir_a_paso(3)

# =========================================================================
# PASO 3 — Resumen y Procesar
# =========================================================================
elif paso_actual == 3:
    proyecto_seleccionado = sv("proyecto_actual")
    doc_contrato_upload = cargar_plantilla_o_aviso(proyecto_seleccionado, NOMBRE_ARCHIVO_CONTRATO, "CONTRATO")
    doc_reglamento_upload = cargar_plantilla_o_aviso(proyecto_seleccionado, NOMBRE_ARCHIVO_REGLAMENTO, "REGLAMENTO")
    doc_membresia_upload = cargar_plantilla_o_aviso(proyecto_seleccionado, NOMBRE_ARCHIVO_MEMBRESIA, "MEMBRESIA")
    doc_derechohabientes_upload = cargar_plantilla_o_aviso(
        proyecto_seleccionado, NOMBRE_ARCHIVO_DERECHOHABIENTES, "DERECHOHABIENTES"
    )
    excel_word_upload = sv("m_excel_word")

    opcion_propiedad = sv("m_opcion_propiedad", OPCIONES_PROPIEDAD[0])
    opcion_membresia_reglamento = sv("m_membresia_reglamento", list(MEMBRESIA_REGLAMENTO_VALORES.keys())[0])
    numero_membresia = sv("m_numero_membresia", "")
    pegar_imagenes_excel = sv("m_pegar_imagenes", False)

    with st.container(border=True):
        st.subheader("📋 Resumen antes de procesar")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(
                f'<div class="cb-card"><b>Proyecto:</b> {proyecto_seleccionado}<br>'
                f'<b>Excel del cliente:</b> {excel_word_upload.name if excel_word_upload else "❌ falta"}<br>'
                f'<b>CONTRATO:</b> {"✅" if doc_contrato_upload else "❌ falta"} · '
                f'<b>REGLAMENTO:</b> {"✅" if doc_reglamento_upload else "—"} · '
                f'<b>MEMBRESIA:</b> {"✅" if doc_membresia_upload else "—"} · '
                f'<b>DERECHOHABIENTES:</b> {"✅" if doc_derechohabientes_upload else "—"}'
                f'</div>',
                unsafe_allow_html=True,
            )
        with c2:
            st.markdown(
                f'<div class="cb-card"><b>Tipo de propiedad:</b> {opcion_propiedad}<br>'
                f'<b>Membresía (CONTRATO y REGLAMENTO):</b> {opcion_membresia_reglamento}<br>'
                f'<b>Número de Membresía:</b> {numero_membresia or "(sin indicar)"}<br>'
                f'<b>Imágenes de Excel en el CONTRATO:</b> {"Sí" if pegar_imagenes_excel else "No"}'
                f'</div>',
                unsafe_allow_html=True,
            )

    col_back, _, col_procesar = st.columns([1, 2, 1])
    with col_back:
        if st.button("← Atrás", width="stretch"):
            _ir_a_paso(2)
    with col_procesar:
        procesar_click = st.button(
            "🚀 Procesar CONTRATO", type="primary", width="stretch",
            disabled=not (excel_word_upload and doc_contrato_upload),
        )

    if procesar_click:
        _inicio_generacion = time.time()
        with st.spinner("Recalculando fórmulas del Excel con Excel real..."):
            try:
                excel_word_upload.seek(0)
                excel_bytes_recalculado = recalcular_excel_con_com(excel_word_upload.read())
            except Exception as e:
                st.error(f"❌ No se pudo recalcular las fórmulas del Excel con Excel real: {e}")
                st.stop()

        with st.spinner("Procesando..."):
            # Valores extra (además de los 7 placeholders de siempre) que
            # aplican según el Paso elegido, ej. KKKKKK/JJJJJJJ (datos del
            # cónyuge) en el Paso 2 -- se agregan a REGLAMENTO, MEMBRESIA
            # y DERECHOHABIENTES (el CONTRATO ya los resuelve por su cuenta
            # dentro de su propia función de FUNCIONES_PASO_CONTRATO).
            funcion_valores_extra = FUNCIONES_VALORES_EXTRA_POR_PASO.get(opcion_propiedad)
            valores_extra = (
                funcion_valores_extra(io.BytesIO(excel_bytes_recalculado))
                if funcion_valores_extra is not None
                else None
            )
            celdas_axa = obtener_celdas_axa(proyecto_seleccionado)

            # CONTRATO y REGLAMENTO se procesan en bloques try/except
            # INDEPENDIENTES a propósito: si uno de los dos falla, el otro
            # igual debe guardarse y mostrarse (antes estaban en el mismo
            # try, y un error en REGLAMENTO cortaba con st.stop() antes de
            # guardar un CONTRATO ya procesado correctamente).
            contrato_resultado_bytes = None
            contrato_info = None
            contrato_error = None
            funcion_paso_contrato = FUNCIONES_PASO_CONTRATO.get(opcion_propiedad)
            if funcion_paso_contrato is not None:
                try:
                    contrato_resultado_bytes, contrato_info = funcion_paso_contrato(
                        io.BytesIO(excel_bytes_recalculado), doc_contrato_upload, opcion_membresia_reglamento,
                        celdas_axa=celdas_axa,
                    )
                    especificaciones_imagenes = []
                    if pegar_imagenes_excel:
                        especificaciones_imagenes.append(
                            (HOJA_LISTAX_IMAGEN, RANGO_LISTAX_IMAGEN, PAGINA_CONTRATO_IMAGEN)
                        )
                        especificaciones_imagenes.append(
                            (HOJA_AXA, celdas_axa["rango_imagen_parte1"], celdas_axa["pagina_imagen_parte1"])
                        )
                        especificaciones_imagenes.append(
                            (HOJA_AXA, celdas_axa["rango_imagen_parte2"], celdas_axa["pagina_imagen_parte2"])
                        )
                        especificaciones_imagenes.append(
                            (HOJA_AXB, RANGO_AXB_PARTE1_IMAGEN, PAGINA_AXB_PARTE1_IMAGEN)
                        )
                        especificaciones_imagenes.append(
                            (HOJA_AXB, RANGO_AXB_PARTE2_IMAGEN, PAGINA_AXB_PARTE2_IMAGEN)
                        )
                        especificaciones_imagenes.append(
                            (HOJA_AXE_IMAGEN, RANGO_AXE_IMAGEN, PAGINA_AXE_IMAGEN)
                        )
                    # El CONTRATO está diseñado para ocupar exactamente 15
                    # páginas -- un placeholder con un nombre/texto más largo
                    # de lo normal a veces hace que se pase a una página 16
                    # extra al final (con muy poco o nada de contenido real).
                    # Se recorta SIEMPRE (haya o no imágenes pegadas), para
                    # que nunca salga un CONTRATO de más de 15 páginas.
                    #
                    # Cuando SÍ hay imágenes, el recorte se hace en la MISMA
                    # sesión de Word que ya se abre para pegarlas (pasando
                    # max_paginas a agregar_imagenes_a_contrato) -- así se
                    # evita abrir Word una segunda vez solo para el recorte
                    # (ahorra un ciclo completo de abrir/cerrar Word).
                    if especificaciones_imagenes:
                        contrato_resultado_bytes = agregar_imagenes_a_contrato(
                            excel_bytes_recalculado, contrato_resultado_bytes, especificaciones_imagenes,
                            max_paginas=MAX_PAGINAS_CONTRATO,
                        )
                        contrato_info["imagenes_pegadas"] = especificaciones_imagenes
                    else:
                        contrato_resultado_bytes = recortar_contrato_a_maximo_paginas(contrato_resultado_bytes)
                except Exception as e:
                    contrato_resultado_bytes = None
                    contrato_error = str(e)
                    st.error(f"❌ Ocurrió un error al procesar el CONTRATO: {e}")
            else:
                contrato_error = (
                    f"Todavía no tengo las instrucciones para '{opcion_propiedad}'. "
                    f"Por ahora solo puedo procesar: {', '.join(FUNCIONES_PASO_CONTRATO.keys())}."
                )
                st.warning(contrato_error)

            # REGLAMENTO: aplica para los 4 Pasos/tipos de propiedad por igual.
            reglamento_resultado_bytes = None
            reglamento_info = None
            reglamento_error = None
            if doc_reglamento_upload is not None:
                try:
                    reglamento_resultado_bytes, reglamento_info = procesar_reglamento(
                        io.BytesIO(excel_bytes_recalculado), doc_reglamento_upload, opcion_membresia_reglamento,
                        valores_extra, celdas_axa=celdas_axa,
                    )
                except Exception as e:
                    reglamento_resultado_bytes = None
                    reglamento_error = str(e)
                    st.error(f"❌ Ocurrió un error al procesar el REGLAMENTO: {e}")

            # MEMBRESIA: mismos valores de placeholders que el CONTRATO,
            # aplica para los 4 Pasos/tipos de propiedad por igual.
            membresia_resultado_bytes = None
            membresia_info = None
            membresia_error = None
            if doc_membresia_upload is not None:
                try:
                    membresia_resultado_bytes, membresia_info = procesar_membresia(
                        io.BytesIO(excel_bytes_recalculado), doc_membresia_upload, numero_membresia,
                        valores_extra, celdas_axa=celdas_axa,
                    )
                except Exception as e:
                    membresia_resultado_bytes = None
                    membresia_error = str(e)
                    st.error(f"❌ Ocurrió un error al procesar el MEMBRESIA: {e}")

            # DERECHOHABIENTES: mismos valores de placeholders que el
            # CONTRATO, aplica para los 4 Pasos/tipos de propiedad por igual.
            derechohabientes_resultado_bytes = None
            derechohabientes_info = None
            derechohabientes_error = None
            if doc_derechohabientes_upload is not None:
                try:
                    derechohabientes_resultado_bytes, derechohabientes_info = procesar_derechohabientes(
                        io.BytesIO(excel_bytes_recalculado), doc_derechohabientes_upload, valores_extra,
                        celdas_axa=celdas_axa,
                    )
                except Exception as e:
                    derechohabientes_resultado_bytes = None
                    derechohabientes_error = str(e)
                    st.error(f"❌ Ocurrió un error al procesar el DERECHOHABIENTES: {e}")

        # Los motivos de error se guardan SIEMPRE (haya o no resultado nuevo),
        # para poder explicar en el Paso 4 por qué falta un documento -- el
        # st.error(...) de arriba se muestra durante ESTE mismo procesamiento,
        # pero _ir_a_paso(4) hace un st.rerun() inmediatamente después, que
        # borra ese mensaje de la pantalla antes de que se alcance a leer (el
        # Paso 4 no vuelve a mostrar nada del CONTRATO si no hay bytes, así
        # que sin esto se ve "como si no se hubiera procesado", sin explicar
        # el motivo).
        st.session_state["contrato_error"] = contrato_error
        st.session_state["reglamento_error"] = reglamento_error
        st.session_state["membresia_error"] = membresia_error
        st.session_state["derechohabientes_error"] = derechohabientes_error

        if contrato_resultado_bytes is not None:
            st.session_state["word_resultado_bytes"] = contrato_resultado_bytes
            st.session_state["word_info"] = contrato_info
        if reglamento_resultado_bytes is not None:
            st.session_state["reglamento_resultado_bytes"] = reglamento_resultado_bytes
            st.session_state["reglamento_info"] = reglamento_info
        if membresia_resultado_bytes is not None:
            st.session_state["membresia_resultado_bytes"] = membresia_resultado_bytes
            st.session_state["membresia_info"] = membresia_info
        if derechohabientes_resultado_bytes is not None:
            st.session_state["derechohabientes_resultado_bytes"] = derechohabientes_resultado_bytes
            st.session_state["derechohabientes_info"] = derechohabientes_info

        # Estadísticas reales para la página de inicio (ver stats_tracker.py):
        # "contrato generado" = el CONTRATO (no REGLAMENTO/MEMBRESIA/
        # DERECHOHABIENTES) se procesó sin error.
        registrar_generacion(
            exito=contrato_resultado_bytes is not None,
            segundos=time.time() - _inicio_generacion,
            categoria="word",
        )

        # Archivo histórico (ver historial_contratos.py): una copia de
        # cada documento generado en CONTRATOS\<proyecto>\<lote nombre>\,
        # y los nombres de descarga del Paso 4 se arman igual: "<TIPO> -
        # <lote nombre>". También registra el número de membresía usado
        # (para el aviso "Último usado" del Paso 2). No corta el
        # procesamiento con el Modo de pruebas activo, sin lote/nombre, o
        # si la carpeta no es accesible en esta máquina -- pero el motivo
        # queda guardado para avisar en el Paso 4 (en vez de fallar en
        # silencio sin que nadie se entere).
        _lote, _nombre_cliente = extraer_lote_y_nombre(excel_bytes_recalculado)
        st.session_state["archivo_nombre_contrato"] = nombre_archivo_descarga(
            "CONTRATO", _lote, _nombre_cliente, "docx",
        )
        st.session_state["archivo_nombre_reglamento"] = nombre_archivo_descarga(
            "REGLAMENTO", _lote, _nombre_cliente, "docx",
        )
        st.session_state["archivo_nombre_membresia"] = nombre_archivo_descarga(
            "MEMBRESIA", _lote, _nombre_cliente, "docx",
        )
        st.session_state["archivo_nombre_derechohabientes"] = nombre_archivo_descarga(
            "DERECHOHABIENTES", _lote, _nombre_cliente, "docx",
        )
        _rutas_archivadas = []
        _motivo_no_archivado = None
        if contrato_resultado_bytes is not None:
            _r, _motivo_no_archivado = guardar_copia(
                proyecto_seleccionado, _lote, _nombre_cliente, "CONTRATO.docx", contrato_resultado_bytes,
            )
            if _r:
                _rutas_archivadas.append(_r)

            # Aviso a n8n (registro maestro externo, ver n8n_integracion.py):
            # un POST por CONTRATO generado, nunca por REGLAMENTO/MEMBRESIA/
            # DERECHOHABIENTES. Si falta lote/nombre, no hay datos suficientes
            # para armar el registro -- se omite (mismo criterio que
            # guardar_copia más arriba).
            if _lote and _nombre_cliente:
                (
                    st.session_state["n8n_estado"],
                    st.session_state["n8n_advertencias"],
                    st.session_state["n8n_id_contrato"],
                ) = registrar_contrato_en_n8n(
                    excel_bytes_recalculado, proyecto_seleccionado, _lote, _nombre_cliente,
                    numero_membresia, opcion_propiedad,
                )
            else:
                st.session_state["n8n_estado"] = None
                st.session_state["n8n_advertencias"] = []
                st.session_state["n8n_id_contrato"] = None
        if reglamento_resultado_bytes is not None:
            _r, _motivo_reglamento = guardar_copia(
                proyecto_seleccionado, _lote, _nombre_cliente, "REGLAMENTO.docx", reglamento_resultado_bytes,
            )
            if _r:
                _rutas_archivadas.append(_r)
            _motivo_no_archivado = _motivo_no_archivado or _motivo_reglamento
        if membresia_resultado_bytes is not None:
            _r, _motivo_membresia = guardar_copia(
                proyecto_seleccionado, _lote, _nombre_cliente, "MEMBRESIA.docx", membresia_resultado_bytes,
            )
            if _r:
                _rutas_archivadas.append(_r)
            _motivo_no_archivado = _motivo_no_archivado or _motivo_membresia
            registrar_numero_membresia(proyecto_seleccionado, _nombre_cliente, numero_membresia)
        if derechohabientes_resultado_bytes is not None:
            _r, _motivo_derecho = guardar_copia(
                proyecto_seleccionado, _lote, _nombre_cliente, "DERECHOHABIENTES.docx", derechohabientes_resultado_bytes,
            )
            if _r:
                _rutas_archivadas.append(_r)
            _motivo_no_archivado = _motivo_no_archivado or _motivo_derecho

        st.session_state["archivo_carpeta_archivada"] = (
            os.path.dirname(_rutas_archivadas[0]) if _rutas_archivadas else None
        )
        st.session_state["archivo_motivo_no_archivado"] = None if _rutas_archivadas else _motivo_no_archivado

        _ir_a_paso(4)

# =========================================================================
# PASO 4 — Resultado
# =========================================================================
elif paso_actual == 4:
    st.subheader("✅ Resultado")

    if st.session_state.get("archivo_carpeta_archivada"):
        st.caption(f'📁 Archivado en: `{st.session_state["archivo_carpeta_archivada"]}`')
    elif st.session_state.get("archivo_motivo_no_archivado"):
        st.caption(f'⚠️ No se archivó en CONTRATOS\\: {st.session_state["archivo_motivo_no_archivado"]}.')

    _n8n_estado = st.session_state.get("n8n_estado")
    _n8n_id = st.session_state.get("n8n_id_contrato")
    _n8n_sufijo_id = f" (ID: {_n8n_id})" if _n8n_id else ""
    if _n8n_estado == "Registrado en el registro maestro":
        st.caption(f'📡 {_n8n_estado}{_n8n_sufijo_id}')
    elif _n8n_estado and _n8n_estado.startswith("Falta el código del proyecto"):
        st.error(f'❌ {_n8n_estado}')
    elif _n8n_estado and _n8n_estado.startswith("No se puede asignar ID_CONTRATO"):
        st.error(f'❌ {_n8n_estado}')
    elif _n8n_estado:
        st.caption(f'⏳ {_n8n_estado}{_n8n_sufijo_id}')
    for _advertencia_n8n in st.session_state.get("n8n_advertencias") or []:
        st.warning(f'⚠️ {_advertencia_n8n}')

    if st.session_state["word_resultado_bytes"] is not None:
        info = st.session_state["word_info"]
        with st.container(border=True):
            st.success("¡CONTRATO procesado!")
            _alertar_placeholders_sin_valor(info)

            st.download_button(
                label="DESCARGAR",
                data=st.session_state["word_resultado_bytes"],
                file_name=st.session_state.get("archivo_nombre_contrato", "CONTRATO.docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch",
            )

            with st.expander("📋 Ver resumen del CONTRATO", expanded=False):
                st.write(f"**Párrafos eliminados:** {info['parrafos_eliminados']} (de 3 esperados)")
                st.write(f"**'JDIP' reemplazado por:** {info['jdip_reemplazado_por']}")
                st.write("**Valores usados para reemplazar los placeholders:**")
                for placeholder, valor in info["valores_usados"].items():
                    st.write(f"- {placeholder} → {valor}")
                for hoja_img, rango_img, pagina_img in info.get("imagenes_pegadas", []):
                    st.write(f"**Imagen:** {hoja_img}!{rango_img} pegada en la página {pagina_img} ✅")
    elif st.session_state.get("contrato_error"):
        st.error(f"❌ El CONTRATO no se generó: {st.session_state['contrato_error']}")

    if st.session_state["reglamento_resultado_bytes"] is not None:
        info_reglamento = st.session_state["reglamento_info"]
        with st.container(border=True):
            st.success("¡REGLAMENTO procesado!")
            _alertar_placeholders_sin_valor(info_reglamento)

            st.download_button(
                label="DESCARGAR",
                data=st.session_state["reglamento_resultado_bytes"],
                file_name=st.session_state.get("archivo_nombre_reglamento", "REGLAMENTO.docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch",
            )

            with st.expander("📋 Ver resumen del REGLAMENTO", expanded=False):
                st.write(f"**'JDIP' reemplazado por:** {info_reglamento['jdip_reemplazado_por']}")
                st.write("**Valores usados para reemplazar los placeholders:**")
                for placeholder, valor in info_reglamento["valores_usados"].items():
                    st.write(f"- {placeholder} → {valor}")
    elif st.session_state.get("reglamento_error"):
        st.error(f"❌ El REGLAMENTO no se generó: {st.session_state['reglamento_error']}")

    if st.session_state["membresia_resultado_bytes"] is not None:
        info_membresia = st.session_state["membresia_info"]
        with st.container(border=True):
            st.success("¡MEMBRESIA procesado!")
            _alertar_placeholders_sin_valor(info_membresia)

            st.download_button(
                label="DESCARGAR",
                data=st.session_state["membresia_resultado_bytes"],
                file_name=st.session_state.get("archivo_nombre_membresia", "MEMBRESIA.docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch",
            )

            with st.expander("📋 Ver resumen de MEMBRESIA", expanded=False):
                st.write(f"**'GTAA' reemplazado por:** {info_membresia['numero_membresia']}")
                st.write("**Valores usados para reemplazar los placeholders:**")
                for placeholder, valor in info_membresia["valores_usados"].items():
                    st.write(f"- {placeholder} → {valor}")
    elif st.session_state.get("membresia_error"):
        st.error(f"❌ El MEMBRESIA no se generó: {st.session_state['membresia_error']}")

    if st.session_state["derechohabientes_resultado_bytes"] is not None:
        info_derechohabientes = st.session_state["derechohabientes_info"]
        with st.container(border=True):
            st.success("¡DERECHOHABIENTES procesado!")
            _alertar_placeholders_sin_valor(info_derechohabientes)

            st.download_button(
                label="DESCARGAR",
                data=st.session_state["derechohabientes_resultado_bytes"],
                file_name=st.session_state.get("archivo_nombre_derechohabientes", "DERECHOHABIENTES.docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch",
            )

            with st.expander("📋 Ver resumen de DERECHOHABIENTES", expanded=False):
                st.write("**Valores usados para reemplazar los placeholders:**")
                for placeholder, valor in info_derechohabientes["valores_usados"].items():
                    st.write(f"- {placeholder} → {valor}")
    elif st.session_state.get("derechohabientes_error"):
        st.error(f"❌ El DERECHOHABIENTES no se generó: {st.session_state['derechohabientes_error']}")

    if st.button("🔄 Nuevo procesamiento", type="primary"):
        st.session_state["word_resultado_bytes"] = None
        st.session_state["reglamento_resultado_bytes"] = None
        st.session_state["membresia_resultado_bytes"] = None
        st.session_state["derechohabientes_resultado_bytes"] = None
        st.session_state["contrato_error"] = None
        st.session_state["reglamento_error"] = None
        st.session_state["membresia_error"] = None
        st.session_state["derechohabientes_error"] = None
        st.session_state["n8n_estado"] = None
        st.session_state["n8n_advertencias"] = []
        st.session_state["n8n_id_contrato"] = None
        _ir_a_paso(0)

# Fondo de imagen por proyecto (si el proyecto ya elegido tiene un archivo
# "fondo.jpeg/jpg/png" guardado). Va al FINAL del script (no arriba, junto al
# <style> principal): "proyecto_actual" recién queda actualizado para esta
# corrida después de que el Paso 0 (si está activo) llama a
# selector_proyecto() -- llamar esto antes leería el valor de la corrida
# anterior. La posición en el HTML no afecta la cascada de CSS (lo que
# importa es que este <style> aparezca DESPUÉS del principal en el
# documento, y así es, sin importar en qué parte del script se emita).
aplicar_fondo_proyecto(st.session_state.get("proyecto_actual"))
