# -*- coding: utf-8 -*-
"""Inicio.py — Punto de entrada de la app multipágina: panel de inicio.

Estilo propio de esta página (paleta "--home-*", tipografía Inter), distinto
a propósito del de las 3 páginas de trabajo (Excel/Word/PDF, que usan su
propio sistema "--cb-*" vía ui_utils.aplicar_tema_base()): esta es la
portada/dashboard, esas son el wizard de trabajo -- no comparten tema para
no mezclar dos lenguajes visuales distintos a mitad de camino.

Las estadísticas (contratos generados/de hoy/tiempo promedio/errores) son
datos REALES leídos de stats_tracker.py (un .json en disco que se va
completando cada vez que se procesa un CONTRATO en la página Word) -- no
son números de ejemplo. Si todavía no se generó ningún contrato, se
muestran en cero de forma explícita, no se inventa nada.
"""

import unicodedata
from datetime import datetime

import streamlit as st
from dotenv import load_dotenv
from rapidfuzz import fuzz

from auth_utils import requerir_login, mostrar_boton_logout
from ui_utils import mostrar_logo, RUTA_LOGO, MODULOS, ACCIONES_RAPIDAS, PREGUNTAS_FRECUENTES
from stats_tracker import obtener_resumen, modo_pruebas_activo, establecer_modo_pruebas
from n8n_integracion import (
    reintentar_pendientes_n8n,
    obtener_codigos_proyecto,
    obtener_correlativo_actual,
    fijar_correlativo_real,
)

load_dotenv()

st.set_page_config(page_title="🤖CONTRATO", page_icon=RUTA_LOGO, layout="wide")

authenticator = requerir_login()
mostrar_boton_logout(authenticator)

mostrar_logo()

# Reintenta avisos a n8n que hayan quedado pendientes (ver
# n8n_integracion.py) -- UNA sola vez por sesión de navegador, no en cada
# interacción/rerun de Streamlit (que pasa muy seguido).
if not st.session_state.get("n8n_reintento_hecho"):
    reintentar_pendientes_n8n()
    st.session_state["n8n_reintento_hecho"] = True

nombre_usuario = st.session_state.get("name", "")

# =========================================================================
# ESTILO — paleta corporativa propia de esta página (azul oscuro/azul,
# fondo claro, tarjetas blancas con sombra), tipografía Inter, sidebar
# modernizado (hover + indicador de página activa vía aria-current, el
# atributo real que usa Streamlit -- ver nota abajo).
# =========================================================================
# st.html() (no st.markdown+unsafe_allow_html) a propósito: mezclar <link>
# y <style> en un mismo st.markdown hace que el parser de Markdown de
# Streamlit corte el bloque de HTML "crudo" en la primera línea en blanco
# que encuentra ahí adentro, y todo lo que sigue después del corte se
# muestra como texto plano en vez de aplicarse como CSS (bug real, visto
# en pantalla). st.html() inserta el HTML tal cual, sin pasar por ese
# intérprete de Markdown, así que no tiene ese problema.
st.html(
    """
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
    :root {
        --home-primary: #1E3A8A;
        --home-primary-dark: #16296b;
        --home-accent: #2563EB;
        --home-accent-hover: #1d4ed8;
        --home-bg: #F8FAFC;
        --home-surface: #FFFFFF;
        --home-border: #E2E8F0;
        --home-text: #0F172A;
        --home-text-muted: #64748B;
        --home-success: #16A34A;
        --home-warning: #D97706;
        --home-danger: #DC2626;
        --home-shadow: 0 1px 2px rgba(15,23,42,.06), 0 6px 16px rgba(15,23,42,.07);
        --home-shadow-hover: 0 4px 10px rgba(15,23,42,.08), 0 14px 28px rgba(15,23,42,.12);
    }
    @media (prefers-color-scheme: dark) {
        :root {
            --home-bg: #0B1220;
            --home-surface: #131B2E;
            --home-border: #23304A;
            --home-text: #E5E9F2;
            --home-text-muted: #8E9BB5;
            --home-accent: #3B82F6;
            --home-accent-hover: #60A5FA;
            --home-shadow: 0 1px 2px rgba(0,0,0,.35), 0 6px 18px rgba(0,0,0,.35);
            --home-shadow-hover: 0 6px 14px rgba(0,0,0,.4), 0 16px 32px rgba(0,0,0,.45);
        }
    }

    html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"],
    [data-testid="stSidebar"], .stButton button, .stTextInput input {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif !important;
    }

    [data-testid="stAppViewContainer"], .stApp { background: var(--home-bg) !important; }
    [data-testid="stHeader"] { background: transparent !important; }
    [data-testid="stMain"] h1, [data-testid="stMain"] h2, [data-testid="stMain"] h3,
    [data-testid="stMain"] h4, [data-testid="stMain"] p, [data-testid="stMain"] span,
    [data-testid="stMain"] label { color: var(--home-text) !important; }
    [data-testid="stCaptionContainer"] { color: var(--home-text-muted) !important; }

    /* --- Sidebar modernizado --- */
    [data-testid="stSidebar"] {
        background: var(--home-surface) !important;
        border-right: 1px solid var(--home-border);
    }
    [data-testid="stSidebar"] * { color: var(--home-text) !important; }
    [data-testid="stSidebarNav"] { padding-top: 0.5rem; }
    /* Cada link de página del menú lateral. Streamlit le pone
       aria-current="page" SOLO al link de la página activa -- es el
       atributo real que usa (lo confirmamos en el bundle de Streamlit
       1.58), no una convención inventada. */
    [data-testid="stSidebarNavLink"] {
        border-radius: 10px !important;
        margin: 2px 8px;
        transition: background .15s ease, color .15s ease;
    }
    [data-testid="stSidebarNavLink"]:hover {
        background: rgba(37, 99, 235, 0.08) !important;
    }
    [data-testid="stSidebarNavLink"][aria-current="page"] {
        background: rgba(37, 99, 235, 0.12) !important;
        border-left: 3px solid var(--home-accent);
        font-weight: 700 !important;
    }
    [data-testid="stSidebarNavLink"][aria-current="page"] span {
        color: var(--home-accent) !important;
        font-weight: 700 !important;
    }

    /* --- Barra superior --- */
    .home-topbar-wrap [data-testid="stTextInput"] input {
        background: var(--home-surface) !important;
        border: 1px solid var(--home-border) !important;
        border-radius: 10px !important;
        color: var(--home-text) !important;
    }
    .home-topbar-wrap [data-testid="stTextInput"] input:focus {
        border-color: var(--home-accent) !important;
        box-shadow: 0 0 0 3px rgba(37,99,235,.15) !important;
    }
    /* Botones "ícono" (🔔/⚙️) del popover: circulares, discretos */
    div[data-testid="stPopover"] > div > button {
        border-radius: 50% !important;
        width: 38px; height: 38px;
        padding: 0 !important;
        border: 1px solid var(--home-border) !important;
        background: var(--home-surface) !important;
        box-shadow: var(--home-shadow);
    }
    div[data-testid="stPopover"] > div > button:hover {
        border-color: var(--home-accent) !important;
    }

    /* --- Hero de bienvenida --- */
    .home-hero {
        display: flex; align-items: center; justify-content: space-between;
        gap: 2rem; background: linear-gradient(135deg, var(--home-primary) 0%, var(--home-accent) 100%);
        border-radius: 20px; padding: 2.4rem 2.6rem; margin: 0.5rem 0 1.8rem 0;
        box-shadow: var(--home-shadow);
    }
    .home-hero-kicker {
        color: rgba(255,255,255,.85) !important; font-weight: 600; font-size: .92rem; margin: 0 0 .3rem 0;
    }
    .home-hero-title {
        color: #fff !important; font-size: 2.4rem; font-weight: 800; margin: 0; letter-spacing: -0.02em;
    }
    .home-hero-subtitle {
        color: rgba(255,255,255,.92) !important; font-size: 1.15rem; font-weight: 600; margin: .3rem 0 .6rem 0;
    }
    .home-hero-desc {
        color: rgba(255,255,255,.82) !important; font-size: .98rem; max-width: 46ch; margin: 0; line-height: 1.5;
    }
    .home-hero-illustration { flex-shrink: 0; opacity: .95; }
    .home-hero-cta [data-testid="stPageLink"] { margin-top: 1.2rem; }
    .home-hero-cta [data-testid="stPageLink"] p {
        color: var(--home-primary) !important; font-weight: 700 !important; font-size: 1.02rem !important;
    }
    .home-hero-cta [data-testid="stPageLink"] {
        background: #fff !important; border-radius: 10px !important;
        padding: .7rem 1.3rem !important; width: fit-content !important;
        box-shadow: 0 4px 14px rgba(0,0,0,.18);
        transition: transform .15s ease, box-shadow .15s ease;
    }
    .home-hero-cta [data-testid="stPageLink"]:hover {
        transform: translateY(-2px); box-shadow: 0 8px 20px rgba(0,0,0,.22);
    }

    @media (max-width: 900px) {
        .home-hero { flex-direction: column; align-items: flex-start; }
        .home-hero-illustration svg { width: 140px; height: 140px; }
    }

    /* --- Tarjetas (módulos) con elevación al pasar el mouse --- */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: var(--home-surface) !important;
        border: 1px solid var(--home-border) !important;
        border-radius: 16px !important;
        box-shadow: var(--home-shadow);
        transition: transform .18s ease, box-shadow .18s ease;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:hover {
        transform: translateY(-3px);
        box-shadow: var(--home-shadow-hover);
    }
    [data-testid="stVerticalBlockBorderWrapper"] > div { gap: .6rem; }

    .home-card-icon {
        width: 46px; height: 46px; border-radius: 12px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.4rem; margin-bottom: .3rem;
    }
    .home-card-title { font-size: 1.08rem; font-weight: 700; color: var(--home-text) !important; margin: 0; }
    .home-card-desc { font-size: .88rem; color: var(--home-text-muted) !important; line-height: 1.45; min-height: 3.2em; }

    /* Botones/links de módulo y acciones rápidas: estilo "secundario" prolijo */
    [data-testid="stPageLink"] {
        border: 1px solid var(--home-border) !important;
        border-radius: 10px !important;
        background: var(--home-bg) !important;
        padding: .5rem .9rem !important;
        transition: border-color .15s ease, background .15s ease;
    }
    [data-testid="stPageLink"]:hover {
        border-color: var(--home-accent) !important;
        background: rgba(37,99,235,.06) !important;
    }
    [data-testid="stPageLink"] p { color: var(--home-text) !important; font-weight: 600 !important; font-size: .9rem !important; }

    /* --- Estadísticas --- */
    .home-stat-card {
        background: var(--home-surface); border: 1px solid var(--home-border);
        border-radius: 16px; padding: 1.1rem 1.3rem; box-shadow: var(--home-shadow);
        display: flex; align-items: center; gap: .9rem;
    }
    .home-stat-icon {
        width: 42px; height: 42px; border-radius: 11px; flex-shrink: 0;
        display: flex; align-items: center; justify-content: center; font-size: 1.25rem;
    }
    .home-stat-value { font-size: 1.55rem; font-weight: 800; color: var(--home-text); line-height: 1.1; }
    .home-stat-label { font-size: .8rem; color: var(--home-text-muted); font-weight: 600; }

    .home-section-title {
        font-size: 1.05rem; font-weight: 700; color: var(--home-text) !important;
        margin: 1.6rem 0 .8rem 0;
    }

    /* Botones normales (secundarios) que no son page_link, ej. logout del popover */
    .stButton > button {
        border-radius: 10px !important; font-weight: 600 !important;
        border: 1px solid var(--home-border) !important;
    }
    </style>
    """
)

# =========================================================================
# BARRA SUPERIOR — búsqueda (filtra módulos/acciones más abajo), campana de
# notificaciones (errores reales registrados) y configuración (Modo de
# pruebas). Ya no hay login, así que no hay perfil ni "cerrar sesión".
# =========================================================================
resumen = obtener_resumen(categoria="word")
resumen_excel = obtener_resumen(categoria="excel")
errores_totales = resumen["errores"] + resumen_excel["errores"]

st.markdown('<div class="home-topbar-wrap">', unsafe_allow_html=True)
col_busqueda, col_notif, col_config = st.columns([7, 1, 1])
with col_busqueda:
    busqueda = st.text_input(
        "Buscar",
        placeholder="🔍 Buscar y presioná Enter (ej. 'membresía', 'plantilla', 'pdf')...",
        key="home_busqueda", label_visibility="collapsed",
        help="Busca en los módulos, las acciones rápidas y las preguntas frecuentes de Ayuda.",
    )
with col_notif:
    with st.popover("🔔", width="content"):
        st.markdown("**Notificaciones**")
        if errores_totales > 0:
            st.warning(f"⚠️ {errores_totales} error(es) registrados (Excel + Word) en las últimas corridas.")
        else:
            st.caption("Sin notificaciones nuevas.")
with col_config:
    with st.popover("⚙️", width="content"):
        st.markdown("**Configuración**")
        st.caption("El tema (claro/oscuro) sigue la preferencia de tu sistema/navegador.")
        st.divider()
        _modo_pruebas_guardado = modo_pruebas_activo()
        _modo_pruebas_nuevo = st.checkbox(
            "🧪 Modo de pruebas",
            value=_modo_pruebas_guardado,
            key="chk_modo_pruebas",
            help=(
                "Mientras esté activo, los contratos y Excel que se procesen NO se suman "
                "a 'Contratos generados' / 'Excel procesados' de las estadísticas -- ni acá "
                "ni para nadie más que use la app, hasta que lo desactives (queda guardado "
                "en disco, no se apaga solo al cerrar la página ni al reiniciar el servidor)."
            ),
        )
        if _modo_pruebas_nuevo != _modo_pruebas_guardado:
            establecer_modo_pruebas(_modo_pruebas_nuevo)
        if _modo_pruebas_nuevo:
            st.caption("🧪 Activo: las estadísticas no se están actualizando ahora mismo.")

        st.divider()
        with st.expander("🔢 Corregir correlativo de ID_CONTRATO (uso excepcional)"):
            st.caption(
                "Solo si asignaste un ID_CONTRATO a mano fuera de la app (ej. en la "
                "matriz de contratos) -- subí aquí el correlativo para que el siguiente "
                "ID automático de ese proyecto no repita un número que ya usaste."
            )
            _codigos_disponibles = obtener_codigos_proyecto()
            if not _codigos_disponibles:
                st.caption("No hay proyectos con código configurado (ver 'codigos_proyecto' en config.yaml).")
            else:
                _proyecto_corr = st.selectbox(
                    "Proyecto", sorted(_codigos_disponibles.keys()), key="sel_proyecto_correlativo",
                )
                _codigo_corr = _codigos_disponibles[_proyecto_corr]
                _anio_corr = st.number_input(
                    "Año", min_value=2000, max_value=2100,
                    value=datetime.now().year, step=1, key="num_anio_correlativo",
                )
                _correlativo_actual = obtener_correlativo_actual(_codigo_corr, int(_anio_corr))
                st.caption(
                    f"Correlativo actual de {_codigo_corr}-{int(_anio_corr)}: "
                    f"**{_correlativo_actual:04d}** (el próximo automático sería "
                    f"{_codigo_corr}-{int(_anio_corr)}-{_correlativo_actual + 1:04d})."
                )
                _nuevo_correlativo = st.number_input(
                    "Subir correlativo a", min_value=0, max_value=9999,
                    value=_correlativo_actual, step=1, key="num_nuevo_correlativo",
                )
                if st.button("Guardar correlativo", key="btn_guardar_correlativo"):
                    fijar_correlativo_real(_codigo_corr, int(_anio_corr), int(_nuevo_correlativo))
                    st.success(
                        f"Listo. El próximo CONTRATO de {_codigo_corr}-{int(_anio_corr)} va a "
                        f"ser {_codigo_corr}-{int(_anio_corr)}-{int(_nuevo_correlativo) + 1:04d}."
                    )
st.markdown('</div>', unsafe_allow_html=True)

if _modo_pruebas_nuevo:
    st.warning(
        "🧪 **Modo de pruebas activo** — los contratos y Excel que se procesen ahora "
        "no se están contando en las estadísticas. Desactivalo en ⚙️ Configuración cuando "
        "termines de probar."
    )

# =========================================================================
# HERO — bienvenida + CTA principal + ilustración (SVG propio, sin
# depender de ninguna imagen externa).
# =========================================================================
saludo = f", {nombre_usuario}" if nombre_usuario else ""
st.markdown(
    f"""
    <div class="home-hero">
        <div>
            <p class="home-hero-kicker">Bienvenido{saludo} 👋</p>
            <h1 class="home-hero-title">CONTRATOBOT</h1>
            <p class="home-hero-subtitle">Sistema Inteligente de Automatización de Contratos</p>
            <p class="home-hero-desc">
                Generá, revisá y exportá contratos, reglamentos y anexos en minutos,
                sin copiar datos a mano entre Excel y Word.
            </p>
        </div>
        <div class="home-hero-illustration">
            <svg width="180" height="180" viewBox="0 0 220 220" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Automatización documental">
                <defs>
                    <linearGradient id="homeGrad1" x1="0" y1="0" x2="1" y2="1">
                        <stop offset="0%" stop-color="#FFFFFF"/>
                        <stop offset="100%" stop-color="#DBEAFE"/>
                    </linearGradient>
                </defs>
                <rect x="28" y="66" width="120" height="150" rx="14" fill="rgba(255,255,255,.25)"/>
                <rect x="50" y="46" width="120" height="150" rx="14" fill="url(#homeGrad1)"/>
                <rect x="68" y="72" width="84" height="8" rx="4" fill="#93A5C4"/>
                <rect x="68" y="92" width="84" height="8" rx="4" fill="#C3D3EA"/>
                <rect x="68" y="112" width="60" height="8" rx="4" fill="#C3D3EA"/>
                <rect x="68" y="142" width="84" height="8" rx="4" fill="#C3D3EA"/>
                <rect x="68" y="162" width="50" height="8" rx="4" fill="#C3D3EA"/>
                <circle cx="162" cy="160" r="34" fill="#FFFFFF"/>
                <path d="M160 142 L144 168 H158 L154 186 L178 156 H162 Z" fill="#2563EB"/>
            </svg>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown('<div class="home-hero-cta">', unsafe_allow_html=True)
st.page_link(
    "pages/2_🤖CONTRATO_Word.py",
    label="✏️  Crear nuevo contrato",
    icon=None,
)
st.markdown('</div>', unsafe_allow_html=True)

# =========================================================================
# BÚSQUEDA aplicada a módulos, acciones rápidas y preguntas frecuentes de
# Ayuda (ver ui_utils.PREGUNTAS_FRECUENTES -- misma lista que usa
# pages/4_🤖CONTRATO_Ayuda.py, así que si mañana se agrega una pregunta
# nueva ahí, también aparece acá sin tocar este archivo).
# =========================================================================
# Umbral de coincidencia difusa (0 a 100, escala de rapidfuzz.fuzz): con
# esto alcanza para tolerar errores de tipeo chicos (ej. "membressia") sin
# empezar a traer resultados que no tienen nada que ver con lo buscado.
UMBRAL_COINCIDENCIA_DIFUSA = 70


def _normalizar_busqueda(texto):
    """Minúsculas y sin tildes -- para que "membresia" (sin tilde, muy común
    al escribir rápido) encuentre igual "membresía", y separado de la
    coincidencia difusa en sí (que ya de por sí tolera errores de tipeo,
    pero no está pensada para la falta sistemática de tildes del español)."""
    texto = unicodedata.normalize("NFKD", texto or "")
    return "".join(c for c in texto if not unicodedata.combining(c)).lower()


termino = _normalizar_busqueda(busqueda)


def _coincide(texto_a, texto_b=""):
    if not termino:
        return True
    combinado = _normalizar_busqueda(f"{texto_a} {texto_b}")
    if termino in combinado:
        return True
    # Antes de esto, la búsqueda pedía coincidencia EXACTA de substring: un
    # solo error de tipeo (ej. "membresia" -> "menbresia") no encontraba
    # nada. fuzz.partial_ratio compara el término contra la mejor porción
    # del texto que se le parezca, así que tolera esos errores chicos.
    return fuzz.partial_ratio(termino, combinado) >= UMBRAL_COINCIDENCIA_DIFUSA


modulos_filtrados = (
    MODULOS if not termino
    else [m for m in MODULOS if _coincide(m["titulo"], m["descripcion"])]
)
acciones_filtradas = (
    ACCIONES_RAPIDAS if not termino
    else [a for a in ACCIONES_RAPIDAS if _coincide(a["label"])]
)
preguntas_filtradas = (
    [] if not termino
    else [p for p in PREGUNTAS_FRECUENTES if _coincide(p["pregunta"], p["respuesta"])]
)

if termino and preguntas_filtradas:
    st.markdown('<p class="home-section-title">De la guía de Ayuda</p>', unsafe_allow_html=True)
    for pregunta in preguntas_filtradas:
        with st.container(border=True):
            st.markdown(f'**{pregunta["pregunta"]}**')
            st.write(pregunta["respuesta"])
    st.page_link("pages/4_🤖CONTRATO_Ayuda.py", label="❓ Ver toda la guía de Ayuda →")

if termino and not modulos_filtrados and not acciones_filtradas and not preguntas_filtradas:
    st.info(f"Sin resultados para '{busqueda}'. Probá con otra palabra (ej. 'excel', 'membresía', 'plantilla').")

# =========================================================================
# MÓDULOS PRINCIPALES
# =========================================================================
if modulos_filtrados:
    st.markdown('<p class="home-section-title">Módulos</p>', unsafe_allow_html=True)
    columnas = st.columns(len(modulos_filtrados))
    for columna, modulo in zip(columnas, modulos_filtrados):
        with columna:
            with st.container(border=True):
                st.markdown(
                    f'<div class="home-card-icon" style="background:{modulo["color"]}22; color:{modulo["color"]};">'
                    f'{modulo["icono"]}</div>',
                    unsafe_allow_html=True,
                )
                st.markdown(f'<p class="home-card-title">{modulo["titulo"]}</p>', unsafe_allow_html=True)
                st.markdown(f'<p class="home-card-desc">{modulo["descripcion"]}</p>', unsafe_allow_html=True)
                st.page_link(modulo["pagina"], label="Abrir →")

# =========================================================================
# ACCIONES RÁPIDAS
# =========================================================================
if acciones_filtradas:
    st.markdown('<p class="home-section-title">Acciones rápidas</p>', unsafe_allow_html=True)
    columnas_acciones = st.columns(len(acciones_filtradas))
    for columna, accion in zip(columnas_acciones, acciones_filtradas):
        with columna:
            st.page_link(accion["pagina"], label=f'{accion["icono"]}  {accion["label"]}')

# =========================================================================
# ESTADÍSTICAS — datos reales de stats_tracker.py (en cero si todavía no
# se procesó nada desde acá). Un apartado por página de trabajo (Word y
# Excel tienen sus propios eventos, separados por "categoria").
# =========================================================================
def _mostrar_fila_estadisticas(resumen_categoria, etiqueta_items, etiqueta_hoy, etiqueta_vacio):
    if not resumen_categoria["hay_datos"]:
        st.caption(etiqueta_vacio)

    tiles = [
        ("📄", "#2563EB", str(resumen_categoria["total_generados"]), etiqueta_items),
        ("📅", "#16A34A", str(resumen_categoria["generados_hoy"]), etiqueta_hoy),
        (
            "⏱️", "#D97706",
            f'{resumen_categoria["tiempo_promedio_seg"]:g}s' if resumen_categoria["hay_datos"] else "—",
            "Tiempo promedio",
        ),
        ("⚠️", "#DC2626", str(resumen_categoria["errores"]), "Errores"),
    ]
    cols_stats = st.columns(4)
    for columna, (icono, color, valor, etiqueta) in zip(cols_stats, tiles):
        with columna:
            st.markdown(
                f'<div class="home-stat-card">'
                f'<div class="home-stat-icon" style="background:{color}22; color:{color};">{icono}</div>'
                f'<div><div class="home-stat-value">{valor}</div>'
                f'<div class="home-stat-label">{etiqueta}</div></div>'
                f'</div>',
                unsafe_allow_html=True,
            )


st.markdown('<p class="home-section-title">Estadísticas — Contrato Word</p>', unsafe_allow_html=True)
_mostrar_fila_estadisticas(
    resumen, "Contratos generados", "Contratos de hoy",
    "Todavía no se generó ningún contrato desde acá — las estadísticas van a ir apareciendo solas.",
)

st.markdown('<p class="home-section-title">Estadísticas — Contrato Excel</p>', unsafe_allow_html=True)
_mostrar_fila_estadisticas(
    resumen_excel, "Excel procesados", "Excel de hoy",
    "Todavía no se procesó ningún Excel desde acá — las estadísticas van a ir apareciendo solas.",
)
