# -*- coding: utf-8 -*-
"""Página "🤖CONTRATO – Excel": Pasos 1 a 6 sobre los archivos Excel.

Rediseño visual (wizard por pasos, tarjetas oscuras, log de actividad).
La lógica de negocio (excel_pipeline.procesar_todo y sus parámetros) y el
flujo de procesamiento NO cambiaron -- esto es solo la capa de interfaz.
"""

import time

import streamlit as st
from openpyxl.utils.exceptions import InvalidFileException

from auth_utils import requerir_login, mostrar_boton_logout
from excel_pipeline import (
    cargar_memoria,
    guardar_en_memoria,
    borrar_memoria,
    procesar_todo,
    FILA_AUTOFILL_ORIGEN,
    FILA_AUTOFILL_AXB,
)
from historial_contratos import extraer_lote_y_nombre, guardar_copia
from stats_tracker import registrar_generacion, modo_pruebas_activo, establecer_modo_pruebas
from templates_manager import NOMBRE_ARCHIVO_MODELO_EXCEL, obtener_celdas_axa
from ui_utils import (
    _renderizar_alertas,
    aplicar_tema_base,
    mostrar_logo,
    RUTA_LOGO,
    selector_proyecto,
    cargar_plantilla_o_aviso,
    aplicar_fondo_proyecto,
)

st.set_page_config(page_title="🤖CONTRATO – Excel", page_icon=RUTA_LOGO, layout="wide")

authenticator = requerir_login()
mostrar_boton_logout(authenticator)

mostrar_logo()

with st.sidebar:
    st.divider()
    _modo_pruebas_guardado = modo_pruebas_activo()
    _modo_pruebas_nuevo = st.checkbox(
        "🧪 Modo de pruebas", value=_modo_pruebas_guardado, key="sb_modo_pruebas_excel",
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
# de ESTA página (título "nube", expander, botones morados, dropzone,
# métricas, inputs, stepper). Solo afecta esta página (Streamlit vuelve a
# montar el DOM en cada navegación de página).
# =========================================================================
aplicar_tema_base()
st.markdown(
    """
    <style>
    /* Título principal: como queda directo sobre la imagen de fondo por
       proyecto (no dentro de ninguna tarjeta), en modo claro se le da una
       "nube" de color propia -- fondo blanco redondeado y letra en el
       color de acento -- en vez de depender del velo general. En modo
       oscuro se revierte más abajo a como estaba (sin nube, texto plano). */
    [data-testid="stMain"] h1 {
        display: inline-block;
        background: rgba(255, 255, 255, 0.88) !important;
        padding: 0.35rem 1.1rem;
        border-radius: 14px;
        border: 1px solid var(--cb-border);
        color: var(--cb-accent) !important;
    }
    /* Revierte la "nube" de arriba en modo oscuro -- queda EXACTAMENTE
       como estaba antes (sin nube, texto plano). Va DESPUÉS de la regla de
       arriba a propósito, para ganarle en la cascada (misma especificidad
       + !important en ambas, así que gana la última en el documento). */
    @media (prefers-color-scheme: dark) {
        [data-testid="stMain"] h1 {
            display: block;
            background: transparent !important;
            padding: 0;
            border-radius: 0;
            border: none;
            color: var(--cb-text) !important;
        }
    }

    /* Tarjetas: expanders */
    [data-testid="stExpander"] {
        background: var(--cb-card) !important;
        border: 1px solid var(--cb-border) !important;
        border-radius: 14px !important;
        overflow: hidden;
    }

    /* Botones */
    .stButton > button, .stDownloadButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        border: 1px solid var(--cb-border) !important;
        transition: all .15s ease;
    }
    .stButton > button[kind="primary"], .stDownloadButton > button {
        background: var(--cb-accent) !important;
        border: none !important;
        color: white !important;
    }
    .stButton > button[kind="primary"] *, .stDownloadButton > button * {
        color: white !important;
    }
    .stButton > button[kind="primary"]:hover, .stDownloadButton > button:hover {
        background: var(--cb-accent-hover) !important;
    }
    .stButton > button[kind="secondary"] {
        background: var(--cb-card) !important;
        color: var(--cb-text) !important;
    }
    .stButton > button[kind="secondary"] * {
        color: var(--cb-text) !important;
    }
    .stButton > button[kind="secondary"]:hover:not(:disabled) {
        border-color: var(--cb-accent) !important;
        color: var(--cb-accent) !important;
    }
    .stButton > button[kind="secondary"]:hover:not(:disabled) * {
        color: var(--cb-accent) !important;
    }
    .stDownloadButton > button {
        font-size: 1.35rem !important;
        padding: 0.85rem 2rem !important;
    }

    /* Zona de arrastrar y soltar */
    [data-testid="stFileUploaderDropzone"] {
        background: var(--cb-bg-elevated) !important;
        border: 2px dashed var(--cb-border) !important;
        border-radius: 14px !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover { border-color: var(--cb-accent) !important; }

    /* Métricas */
    [data-testid="stMetric"] {
        background: var(--cb-card);
        border: 1px solid var(--cb-border);
        border-radius: 12px;
        padding: 0.75rem 1rem;
    }

    /* Inputs */
    [data-baseweb="select"] > div, .stTextInput input, .stDateInput input {
        background: var(--cb-bg-elevated) !important;
        border-color: var(--cb-border) !important;
        color: var(--cb-text) !important;
    }

    /* Stepper superior */
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
    .cb-card-muted { color: var(--cb-text-muted); font-size: 0.92rem; }
    </style>
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
if "wizard_step" not in st.session_state:
    st.session_state["wizard_step"] = 0
if "resultado_bytes" not in st.session_state:
    st.session_state["resultado_bytes"] = None
    st.session_state["alertas_detectadas"] = None
    st.session_state["resumen"] = None
    st.session_state["pasos_ejecutados"] = {}
    st.session_state["log_actividad"] = []

st.title("🤖CONTRATO – Excel")

memoria = cargar_memoria()
with st.sidebar:
    st.header("🧠 Memoria de Aprendizaje")
    if memoria["reglas_adicionales"]:
        for r in memoria["reglas_adicionales"]:
            st.caption(f"• {r}")
        if st.button("🗑️ Borrar todas las reglas"):
            borrar_memoria()
            st.rerun()
    else:
        st.info("Aún no me has enseñado reglas nuevas.")

# =========================================================================
# STEPPER
# =========================================================================
NOMBRES_PASOS_WIZARD = ["📁 Proyecto", "🧾 Datos", "⚙️ Configuración", "📋 Resumen", "✅ Resultado"]
paso_actual = st.session_state["wizard_step"]

html_stepper = ['<div class="cb-stepper">']
for i, nombre in enumerate(NOMBRES_PASOS_WIZARD):
    estado = "done" if i < paso_actual else ("active" if i == paso_actual else "")
    icono = "✓" if estado == "done" else str(i + 1)
    html_stepper.append(f'<div class="cb-step {estado}"><div class="cb-step-circle">{icono}</div>'
                         f'<div class="cb-step-label">{nombre}</div></div>')
    if i < len(NOMBRES_PASOS_WIZARD) - 1:
        linea_done = "done" if i < paso_actual else ""
        html_stepper.append(f'<div class="cb-step-line {linea_done}"></div>')
html_stepper.append("</div>")
st.markdown("".join(html_stepper), unsafe_allow_html=True)


def _ir_a_paso(nuevo_paso):
    st.session_state["wizard_step"] = nuevo_paso
    st.rerun()


# =========================================================================
# PASO 0 — Proyecto y plantilla
# =========================================================================
if paso_actual == 0:
    with st.container(border=True):
        st.subheader("📁 Proyecto y plantilla")
        proyecto_seleccionado = selector_proyecto()
        archivo_modelo_upload = cargar_plantilla_o_aviso(
            proyecto_seleccionado, NOMBRE_ARCHIVO_MODELO_EXCEL, "Excel MODELO"
        )

    _, col_next = st.columns([3, 1])
    with col_next:
        if st.button("Siguiente →", type="primary", disabled=archivo_modelo_upload is None, width="stretch"):
            _ir_a_paso(1)

# =========================================================================
# PASO 1 — Datos del cliente (Excel ORIGEN, siempre manual)
# =========================================================================
elif paso_actual == 1:
    with st.container(border=True):
        st.subheader("🧾 Datos de este cliente/cotización")
        archivo_origen_upload = st.file_uploader(
            "📥 Sube el Excel de ORIGEN (Con Datos)", type=["xlsx"], key="excel_origen_uploader"
        )
        # Espejo en key plana (ver nota en ui_utils.selector_proyecto): sin
        # esto, el archivo subido en el Paso 1 se "olvidaría" al llegar al
        # Paso 3 (2+ pasos después sin volver a mostrar este uploader).
        st.session_state["m_origen"] = archivo_origen_upload
        if archivo_origen_upload is not None:
            st.success(f"✅ '{archivo_origen_upload.name}' cargado correctamente.")
        else:
            st.warning("⚠️ Todavía falta subir el Excel de ORIGEN.")

    col_back, _, col_next = st.columns([1, 2, 1])
    with col_back:
        if st.button("← Atrás", width="stretch"):
            _ir_a_paso(0)
    with col_next:
        if st.button("Siguiente →", type="primary", disabled=archivo_origen_upload is None, width="stretch"):
            _ir_a_paso(2)

# =========================================================================
# PASO 2 — Configuración de Pasos
# =========================================================================
elif paso_actual == 2:
    with st.container(border=True):
        st.subheader("⚙️ Configuración de Pasos")
        with st.expander("⚙️ Opciones avanzadas", expanded=False):
            st.caption("Normalmente no hace falta tocar esto — son los Pasos internos del "
                       "pipeline de Excel, todos vienen activados.")
            ejecutar_paso2 = st.checkbox("Paso 2 — Cotizador", value=sv("chk_paso2", True), key="chk_paso2")
            ejecutar_paso3 = st.checkbox(
                "Paso 3 — FICHA (inserción de filas + AutoFill emulado)", value=sv("chk_paso3", True), key="chk_paso3"
            )
            ejecutar_paso4 = st.checkbox(
                "Paso 4 — PONER MEDIDAS PLANOS", value=sv("chk_paso4", True), key="chk_paso4"
            )
            ejecutar_paso5 = st.checkbox(
                "Paso 5 — Ax.B (inserción de filas + AutoFill emulado)", value=sv("chk_paso5", True), key="chk_paso5"
            )
            ejecutar_paso6 = st.checkbox(
                "Paso 6 — MARCAR SI TIENE MEMBRESÍA", value=sv("chk_paso6", True), key="chk_paso6"
            )
            # Espejo en keys planas (ver nota junto al file_uploader del Paso
            # 1): estas se leen de nuevo recién en el Paso 3, 1 paso después.
            st.session_state["m_paso2"] = ejecutar_paso2
            st.session_state["m_paso3"] = ejecutar_paso3
            st.session_state["m_paso4"] = ejecutar_paso4
            st.session_state["m_paso5"] = ejecutar_paso5
            st.session_state["m_paso6"] = ejecutar_paso6

        if ejecutar_paso2:
            with st.container(border=True):
                st.markdown("**Paso 2 — Cotizador**")
                usar_fecha_manual = st.checkbox(
                    "Usar una fecha específica para la cotización",
                    value=sv("chk_fecha_manual", False), key="chk_fecha_manual",
                    help="Por defecto la celda I16 del Cotizador usa la fecha de hoy (fórmula "
                         "=HOY()). Activá esto para escribir una fecha fija en su lugar.",
                )
                st.session_state["m_fecha_manual_chk"] = usar_fecha_manual
                if usar_fecha_manual:
                    fecha_manual_input = st.date_input("Fecha a usar en I16", key="input_fecha_manual")
                    st.session_state["m_fecha_manual_val"] = fecha_manual_input

        if ejecutar_paso4:
            with st.container(border=True):
                st.markdown("**Paso 4 — Medidas del plano**")
                colE1, colE2, colE3, colE4 = st.columns(4)
                with colE1:
                    e51 = st.text_input("Adelante (m)", key="input_e51",
                                         help="Se escribe en la celda E51 de la hoja Ax.A.")
                with colE2:
                    e52 = st.text_input("Derecha (m)", key="input_e52",
                                         help="Se escribe en la celda E52 de la hoja Ax.A.")
                with colE3:
                    e53 = st.text_input("Izquierda (m)", key="input_e53",
                                         help="Se escribe en la celda E53 de la hoja Ax.A.")
                with colE4:
                    e54 = st.text_input("Fondo (m)", key="input_e54",
                                         help="Se escribe en la celda E54 de la hoja Ax.A.")
                st.caption("Opcional — dejá en blanco la(s) medida(s) que no correspondan.")
                sede = st.segmented_control(
                    "Sede",
                    options=["Arequipa", "Ilo"], default=sv("seg_sede", "Arequipa"),
                    required=True, key="seg_sede",
                    help="Define qué fórmula se escribe en la celda G124 de la hoja Ax.A.",
                )
                st.session_state["m_e51"] = e51
                st.session_state["m_e52"] = e52
                st.session_state["m_e53"] = e53
                st.session_state["m_e54"] = e54
                st.session_state["m_sede"] = sede

        if ejecutar_paso6:
            with st.container(border=True):
                st.markdown("**Paso 6 — Membresía**")
                membresia = st.segmented_control(
                    "Membresía",
                    options=["CON MEMBRESÍA", "SIN MEMBRESIA"], default=sv("seg_membresia", "CON MEMBRESÍA"),
                    required=True, key="seg_membresia",
                )
                st.session_state["m_membresia"] = membresia

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
    # Todo se lee de las keys "espejo" (ver notas en los Pasos 0-2), no de
    # las keys de los widgets -- Streamlit limpia el valor de un widget de
    # session_state si pasan 2+ corridas sin volver a instanciarlo, y para
    # cuando llegamos al Paso 3 ya pasó eso con los widgets del Paso 0/1/2.
    proyecto_seleccionado = sv("proyecto_actual")
    archivo_modelo_upload = cargar_plantilla_o_aviso(proyecto_seleccionado, NOMBRE_ARCHIVO_MODELO_EXCEL, "Excel MODELO")
    archivo_origen_upload = sv("m_origen")

    ejecutar_paso2 = sv("m_paso2", True)
    ejecutar_paso3 = sv("m_paso3", True)
    ejecutar_paso4 = sv("m_paso4", True)
    ejecutar_paso5 = sv("m_paso5", True)
    ejecutar_paso6 = sv("m_paso6", True)

    fecha_manual_valor = sv("m_fecha_manual_val") if sv("m_fecha_manual_chk", False) else None
    valor_e51, valor_e52 = sv("m_e51", ""), sv("m_e52", "")
    valor_e53, valor_e54 = sv("m_e53", ""), sv("m_e54", "")
    sede_seleccionada = sv("m_sede", "Arequipa") if ejecutar_paso4 else None
    membresia_seleccionada = sv("m_membresia", "CON MEMBRESÍA") if ejecutar_paso6 else None

    with st.container(border=True):
        st.subheader("📋 Resumen antes de procesar")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(
                f'<div class="cb-card"><b>Proyecto:</b> {proyecto_seleccionado}<br>'
                f'<b>Excel MODELO:</b> {"✅ plantilla cargada" if archivo_modelo_upload else "❌ falta"}<br>'
                f'<b>Excel ORIGEN:</b> {archivo_origen_upload.name if archivo_origen_upload else "❌ falta"}'
                f'</div>',
                unsafe_allow_html=True,
            )
        with c2:
            lineas_pasos = []
            if ejecutar_paso2:
                lineas_pasos.append(
                    f"Paso 2 — Cotizador ({'fecha manual: ' + str(fecha_manual_valor) if fecha_manual_valor else '=HOY()'})"
                )
            if ejecutar_paso3:
                lineas_pasos.append("Paso 3 — FICHA")
            if ejecutar_paso4:
                lineas_pasos.append(f"Paso 4 — PONER MEDIDAS PLANOS (Sede: {sede_seleccionada})")
            if ejecutar_paso5:
                lineas_pasos.append("Paso 5 — Ax.B")
            if ejecutar_paso6:
                lineas_pasos.append(f"Paso 6 — Membresía: {membresia_seleccionada}")
            if not lineas_pasos:
                lineas_pasos.append("Ningún Paso 2-6 activado (solo se procesará el Paso 1, hoja DATOS).")
            st.markdown(
                '<div class="cb-card"><b>Pasos que se ejecutarán:</b><br>' + "<br>".join(lineas_pasos) + "</div>",
                unsafe_allow_html=True,
            )

    col_back, _, col_procesar = st.columns([1, 2, 1])
    with col_back:
        if st.button("← Atrás", width="stretch"):
            _ir_a_paso(2)
    with col_procesar:
        procesar_click = st.button(
            "🚀 Procesar", type="primary", width="stretch",
            disabled=archivo_modelo_upload is None or archivo_origen_upload is None,
        )

    if procesar_click:
        _inicio_generacion = time.time()
        with st.status("Procesando documentos...", expanded=True) as status:
            status.write("Leyendo el Excel de ORIGEN y el Excel MODELO...")
            archivo_origen_upload.seek(0)
            try:
                resultado_bytes, alertas_detectadas, resumen = procesar_todo(
                    archivo_origen_upload,
                    archivo_modelo_upload,
                    memoria["reglas_adicionales"],
                    ejecutar_paso2=ejecutar_paso2,
                    fecha_manual=fecha_manual_valor,
                    ejecutar_paso3=ejecutar_paso3,
                    ejecutar_paso4=ejecutar_paso4,
                    valores_e51_e54=[valor_e51, valor_e52, valor_e53, valor_e54],
                    sede=sede_seleccionada,
                    ejecutar_paso5=ejecutar_paso5,
                    ejecutar_paso6=ejecutar_paso6,
                    membresia=membresia_seleccionada,
                    celdas_axa=obtener_celdas_axa(proyecto_seleccionado),
                )
            except InvalidFileException:
                status.update(label="Error al procesar", state="error")
                st.error("❌ Uno de los archivos no es un .xlsx válido (puede estar corrupto o ser otro formato).")
                registrar_generacion(
                    exito=False, segundos=time.time() - _inicio_generacion, categoria="excel",
                )
                st.stop()
            except Exception as e:
                status.update(label="Error al procesar", state="error")
                st.error(f"❌ Ocurrió un error inesperado al procesar los archivos: {e}")
                registrar_generacion(
                    exito=False, segundos=time.time() - _inicio_generacion, categoria="excel",
                )
                st.stop()
            status.write("Aplicando los Pasos seleccionados...")

            pasos_ejecutados = {
                "paso2": ejecutar_paso2, "paso3": ejecutar_paso3, "paso4": ejecutar_paso4,
                "paso5": ejecutar_paso5, "paso6": ejecutar_paso6,
            }

            log_actividad = [
                f"✅ Paso 1 (DATOS): {resumen['filas_procesadas']} fila(s) procesada(s), "
                f"{resumen['filas_con_guion_ignoradas']} con '-' ignorada(s) de validación."
            ]
            if pasos_ejecutados.get("paso2"):
                ok = resumen.get("cotizador_rangos_finales_copiados")
                log_actividad.append(
                    f"{'✅' if ok else '⚠️'} Paso 2 (Cotizador): rangos finales E35:I35/E37:I37 "
                    f"{'copiados' if ok else 'NO copiados (hubo alertas previas en el Paso 2)'}."
                )
            if pasos_ejecutados.get("paso3"):
                log_actividad.append(
                    f"✅ Paso 3 (FICHA): {resumen.get('ficha_filas_insertadas', 0)} fila(s) insertada(s) "
                    f"debajo de la fila {FILA_AUTOFILL_ORIGEN}."
                )
            if pasos_ejecutados.get("paso4"):
                log_actividad.append(
                    f"✅ Paso 4 (Ax.A): G115={resumen.get('axa_g115', '(sin cambio)')} | "
                    f"G120={resumen.get('axa_g120', '(sin cambio)')} | G124={resumen.get('axa_g124', '(sin cambio)')}."
                )
            if pasos_ejecutados.get("paso5"):
                log_actividad.append(
                    f"✅ Paso 5 (Ax.B): {resumen.get('axb_filas_insertadas', 0)} fila(s) insertada(s) "
                    f"debajo de la fila {FILA_AUTOFILL_AXB}."
                )
            if pasos_ejecutados.get("paso6"):
                if resumen.get("listax_membresia") == "CON MEMBRESÍA":
                    log_actividad.append("✅ Paso 6 (List.Ax): CON MEMBRESÍA → G17 con relleno blanco.")
                elif resumen.get("listax_membresia") == "SIN MEMBRESIA":
                    log_actividad.append(
                        f"✅ Paso 6 (List.Ax): SIN MEMBRESIA → G19 ('{resumen.get('listax_g19_movido_a_g17')}') "
                        f"movido a G17, y C19:F19 borradas."
                    )
            if alertas_detectadas:
                log_actividad.append(f"⚠️ Se detectaron {len(alertas_detectadas)} alerta(s) de auditoría.")
            else:
                log_actividad.append("✅ Sin alertas de auditoría.")

            status.update(label="¡Procesamiento completado!", state="complete")

        st.session_state["resultado_bytes"] = resultado_bytes
        st.session_state["alertas_detectadas"] = alertas_detectadas
        st.session_state["resumen"] = resumen
        st.session_state["pasos_ejecutados"] = pasos_ejecutados
        st.session_state["log_actividad"] = log_actividad
        registrar_generacion(
            exito=True, segundos=time.time() - _inicio_generacion, categoria="excel",
        )

        # Archivo histórico (ver historial_contratos.py): copia del
        # COTIZADOR en CONTRATOS\<proyecto>\<lote nombre>\, y el prefijo
        # se guarda para nombrar el botón de descarga del Paso 4 igual.
        # No corta el procesamiento con el Modo de pruebas activo, sin
        # lote/nombre, o si la carpeta no es accesible en esta máquina --
        # pero el motivo queda guardado para avisar en el Paso 4 (en vez
        # de fallar en silencio sin que nadie se entere).
        _lote, _nombre_cliente = extraer_lote_y_nombre(resultado_bytes)
        _ruta_archivada, _motivo_no_archivado = guardar_copia(
            proyecto_seleccionado, _lote, _nombre_cliente, "COTIZADOR.xlsx", resultado_bytes,
        )
        st.session_state["archivo_prefijo"] = (
            f"{_lote} {_nombre_cliente} - " if _lote and _nombre_cliente else ""
        )
        st.session_state["archivo_ruta_archivada"] = _ruta_archivada
        st.session_state["archivo_motivo_no_archivado"] = _motivo_no_archivado

        _ir_a_paso(4)

# =========================================================================
# PASO 4 — Resultado
# =========================================================================
elif paso_actual == 4:
    resumen = st.session_state["resumen"]
    pasos = st.session_state["pasos_ejecutados"]

    with st.container(border=True):
        st.subheader("✅ Resultado")
        st.success("¡Procesamiento completado!")

        _renderizar_alertas(st.session_state["alertas_detectadas"])

        st.download_button(
            label="DESCARGAR",
            data=st.session_state["resultado_bytes"],
            file_name=f'{st.session_state.get("archivo_prefijo", "")}COTIZADOR.xlsx',
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
        )

        if st.session_state.get("archivo_ruta_archivada"):
            st.caption(f'📁 Archivado en: `{st.session_state["archivo_ruta_archivada"]}`')
        elif st.session_state.get("archivo_motivo_no_archivado"):
            st.caption(f'⚠️ No se archivó en CONTRATOS\\: {st.session_state["archivo_motivo_no_archivado"]}.')

        with st.expander("🗒️ Registro de actividad", expanded=True):
            for linea in st.session_state["log_actividad"]:
                st.write(linea)

    with st.expander("📋 Ver resumen detallado", expanded=False):
        c1, c2 = st.columns(2)
        c1.metric("Filas procesadas (DATOS)", resumen["filas_procesadas"])
        c2.metric("Celdas con '-' (ignoradas de validación)", resumen["filas_con_guion_ignoradas"])

        if pasos.get("paso2"):
            st.write(
                f"**Paso 2:** rangos finales E35:I35 y E37:I37 "
                f"{'copiados ✅' if resumen.get('cotizador_rangos_finales_copiados') else 'NO copiados ⚠️'}"
            )
        if pasos.get("paso3"):
            st.write(
                f"**Paso 3:** {resumen.get('ficha_filas_insertadas', 0)} fila(s) insertada(s) en 'FICHA' "
                f"debajo de la fila {FILA_AUTOFILL_ORIGEN} | "
                f"D64 {'establecido ✅' if resumen.get('ficha_d64_establecido') else 'sin cambios'}"
            )
        if pasos.get("paso4"):
            st.write(
                f"**Paso 4:** G115={resumen.get('axa_g115', '(sin cambio)')} | "
                f"G120={resumen.get('axa_g120', '(sin cambio)')} | "
                f"G124={resumen.get('axa_g124', '(sin cambio)')}"
            )
        if pasos.get("paso5"):
            st.write(
                f"**Paso 5:** {resumen.get('axb_filas_insertadas', 0)} fila(s) insertada(s) en 'Ax.B' "
                f"debajo de la fila {FILA_AUTOFILL_AXB}"
            )
        if pasos.get("paso6"):
            if resumen.get("listax_membresia") == "CON MEMBRESÍA":
                st.write("**Paso 6:** CON MEMBRESÍA → G17 con relleno blanco")
            elif resumen.get("listax_membresia") == "SIN MEMBRESIA":
                st.write(
                    f"**Paso 6:** SIN MEMBRESIA → G19 ('{resumen.get('listax_g19_movido_a_g17')}') "
                    f"movido a G17, y C19:F19 borradas"
                )

    if st.button("🔄 Nuevo procesamiento", type="primary"):
        st.session_state["resultado_bytes"] = None
        _ir_a_paso(0)

# =========================================================================
# CHATBOT DE APRENDIZAJE — siempre visible, fuera del wizard
# =========================================================================
st.divider()
st.subheader("💬 Habla con tu Chatbot para enseñarle o corregirlo")

chat_input = st.text_input("Escribe una instrucción (Ejemplo: 'Aprende: Ignorar telefono' o 'Aprende: Ignorar DNI')")

if chat_input:
    if st.button("Enviar al Bot"):
        texto_bajo = chat_input.lower()
        if "aprende:" in texto_bajo or "ignorar" in texto_bajo:
            guardar_en_memoria(chat_input)
            st.success(f"🤖 Bot: ¡Entendido! He memorizado la regla: '{chat_input}'.")
            st.rerun()
        else:
            st.warning("🤖 Bot: No reconocí esa instrucción. Usa un formato como: 'Aprende: Ignorar telefono'.")

# Fondo de imagen por proyecto (si el proyecto ya elegido tiene un archivo
# "fondo.jpeg/jpg/png" guardado). Va al FINAL del script (no arriba, junto al
# <style> principal): "proyecto_actual" recién queda actualizado para esta
# corrida después de que el Paso 0 (si está activo) llama a
# selector_proyecto() -- llamar esto antes leería el valor de la corrida
# anterior. La posición en el HTML no afecta la cascada de CSS (lo que
# importa es que este <style> aparezca DESPUÉS del principal en el
# documento, y así es, sin importar en qué parte del script se emita).
aplicar_fondo_proyecto(st.session_state.get("proyecto_actual"))
