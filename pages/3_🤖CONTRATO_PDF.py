# -*- coding: utf-8 -*-
"""Página "🤖CONTRATO – PDF": convertir Word (.docx) ya revisados a mano a
PDF, hasta 4 a la vez, sin cambiarles el nombre.

Pensada para el flujo: generás el CONTRATO/REGLAMENTO/MEMBRESIA/
DERECHOHABIENTES en la página Word, los revisás/corregís a mano en Word, y
después subís acá esos mismos archivos ya corregidos para bajarlos en PDF.
No depende de ningún wizard ni de ningún dato de cliente -- es un conversor
suelto, independiente del resto del flujo."""

import os

import streamlit as st

from auth_utils import requerir_login, mostrar_boton_logout
from historial_contratos import encontrar_carpeta_por_prefijo, guardar_copia_en_carpeta
from ui_utils import aplicar_tema_base, mostrar_logo, RUTA_LOGO, aplicar_fondo_proyecto
from word_com_bridge import convertir_docx_a_pdf_con_com

MAX_ARCHIVOS = 4

st.set_page_config(page_title="🤖CONTRATO – PDF", page_icon=RUTA_LOGO, layout="wide")

authenticator = requerir_login()
mostrar_boton_logout(authenticator)

mostrar_logo()

# =========================================================================
# ESTILO — variables de tema y layout base vienen de aplicar_tema_base()
# (compartida con las otras 2 páginas); acá solo quedan las reglas propias
# de ESTA página (encabezado, botones verdes de descarga, dropzone).
# =========================================================================
aplicar_tema_base()
st.markdown(
    """
    <style>
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
    @media (prefers-color-scheme: dark) {
        .cb-page-header {
            background: transparent !important;
            padding: 0;
            border-radius: 0;
            border: none;
        }
    }

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

    [data-testid="stFileUploaderDropzone"] {
        background: var(--cb-bg-elevated) !important;
        border: 2px dashed var(--cb-border) !important;
        border-radius: 12px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="cb-page-header">
        <span class="cb-kicker">Área legal · Conversión de documentos</span>
        <h1 class="cb-title">🤖CONTRATO – PDF</h1>
        <p class="cb-subtitle">
            Subí hasta 4 archivos Word (.docx) ya revisados/corregidos a mano y convertilos
            a PDF. Cada PDF se descarga con el mismo nombre con el que subiste el .docx.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

if "pdf_resultados" not in st.session_state:
    st.session_state["pdf_resultados"] = []  # lista de (nombre_pdf, bytes_pdf)

with st.container(border=True):
    archivos_subidos = st.file_uploader(
        "Archivos Word (.docx)",
        type=["docx"],
        accept_multiple_files=True,
        key="pdf_uploader",
        help=f"Hasta {MAX_ARCHIVOS} archivos a la vez.",
    )

    demasiados_archivos = archivos_subidos and len(archivos_subidos) > MAX_ARCHIVOS
    if demasiados_archivos:
        st.error(
            f"Subiste {len(archivos_subidos)} archivos; el máximo es {MAX_ARCHIVOS}. "
            f"Sacá alguno para poder convertir."
        )

    if st.button(
        "Convertir a PDF",
        type="primary",
        disabled=not archivos_subidos or demasiados_archivos,
    ):
        resultados = []
        with st.spinner("Convirtiendo a PDF con Word..."):
            for archivo in archivos_subidos:
                nombre_pdf = os.path.splitext(archivo.name)[0] + ".pdf"
                try:
                    pdf_bytes = convertir_docx_a_pdf_con_com(archivo.getvalue())
                    resultados.append((nombre_pdf, pdf_bytes, None))

                    # Archivo histórico (ver historial_contratos.py): esta
                    # página no sabe a qué proyecto pertenece un .docx --
                    # si el nombre subido ya tiene el formato "<TIPO> -
                    # <lote nombre>" (porque vino de una descarga del
                    # módulo Word), busca esa carpeta entre todos los
                    # proyectos y guarda el PDF ahí también. Si no la
                    # encuentra, no hace nada -- la descarga normal de acá
                    # abajo sigue funcionando igual.
                    nombre_sin_ext = os.path.splitext(archivo.name)[0]
                    if " - " in nombre_sin_ext:
                        prefijo_candidato = nombre_sin_ext.split(" - ", 1)[1]
                        carpeta = encontrar_carpeta_por_prefijo(prefijo_candidato)
                        if carpeta:
                            guardar_copia_en_carpeta(carpeta, nombre_pdf, pdf_bytes)
                except Exception as exc:
                    resultados.append((nombre_pdf, None, str(exc)))
        st.session_state["pdf_resultados"] = resultados

for nombre_pdf, pdf_bytes, error in st.session_state["pdf_resultados"]:
    with st.container(border=True):
        if error is not None:
            st.error(f"{nombre_pdf}: no se pudo convertir — {error}")
        else:
            st.success(f"{nombre_pdf} convertido")
            st.download_button(
                label=f"DESCARGAR {nombre_pdf}",
                data=pdf_bytes,
                file_name=nombre_pdf,
                mime="application/pdf",
                width="stretch",
            )

# Fondo de imagen por proyecto (ver la misma nota en pages/2_🤖CONTRATO_Word.py):
# va al final para leer "proyecto_actual" ya actualizado en esta corrida.
aplicar_fondo_proyecto(st.session_state.get("proyecto_actual"))
