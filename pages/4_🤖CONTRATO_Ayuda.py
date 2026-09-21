# -*- coding: utf-8 -*-
"""Página "🤖CONTRATO – Ayuda": guía de uso, de SOLO LECTURA.

No escribe ningún archivo, no toca Excel/Word, no ejecuta ningún
procesamiento -- por diseño, para que un error acá jamás pueda afectar a
las otras páginas (cada página de Streamlit corre aislada de las demás).

Lo que SÍ se actualiza solo: la lista de módulos y de proyectos (se leen
en vivo de ui_utils.MODULOS y templates_manager.obtener_proyectos()/
obtener_plantillas_proyecto -- las mismas fuentes que usa Inicio.py, no
una copia aparte).

Lo que NO se actualiza solo: el texto explicativo de "cómo usar cada
cosa" de acá abajo. Si el día de mañana cambia cómo funciona un Paso,
hay que pedir que se actualice este texto a mano -- ningún código puede
darse cuenta solo de que la explicación quedó desactualizada.
"""

import base64
import os

import streamlit as st

from auth_utils import requerir_login, mostrar_boton_logout
from templates_manager import obtener_proyectos, obtener_plantillas_proyecto
from ui_utils import (
    aplicar_tema_base,
    aplicar_fondo_proyecto,
    mostrar_logo,
    RUTA_LOGO,
    RUTA_GUIA_PDF,
    MODULOS,
    ACCIONES_RAPIDAS,
    PREGUNTAS_FRECUENTES,
    ETIQUETAS_ARCHIVOS_PLANTILLA,
)

st.set_page_config(page_title="🤖CONTRATO – Ayuda", page_icon=RUTA_LOGO, layout="wide")

authenticator = requerir_login()
mostrar_boton_logout(authenticator)

mostrar_logo()

# =========================================================================
# ESTILO — mismo sistema de tema que Excel/Word/PDF (aplicar_tema_base),
# no el de Inicio.py: esta es una página de trabajo/consulta, no el
# dashboard. Solo se agrega un estilo mínimo propio para el encabezado.
# =========================================================================
aplicar_tema_base()
st.markdown(
    """
    <style>
    [data-testid="stExpander"] {
        background: var(--cb-card) !important;
        border: 1px solid var(--cb-border) !important;
        border-radius: 14px !important;
        overflow: hidden;
    }
    .stButton > button {
        border-radius: 10px !important; font-weight: 600 !important;
        border: 1px solid var(--cb-border) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("❓ Ayuda")
st.caption("Guía de uso de CONTRATOBOT. Esta página es solo de consulta — no procesa nada.")

# =========================================================================
# GUÍA EN PDF — opcional: se agrega a mano en la carpeta principal de la
# app como "GUIA_PASO_A_PASO.pdf" (ver ui_utils.RUTA_GUIA_PDF). Si no
# existe, esta sección no se muestra -- no rompe nada, no hace falta
# tocar código para agregarla/sacarla, solo poner o borrar el archivo.
# =========================================================================
if os.path.exists(RUTA_GUIA_PDF):
    st.subheader("Guía paso a paso (PDF)")
    with st.container(border=True):
        with open(RUTA_GUIA_PDF, "rb") as f:
            _pdf_bytes = f.read()
        st.download_button(
            label="⬇️ Descargar guía en PDF",
            data=_pdf_bytes,
            file_name="GUIA_PASO_A_PASO.pdf",
            mime="application/pdf",
            type="primary",
        )
        _pdf_b64 = base64.b64encode(_pdf_bytes).decode("ascii")
        st.html(
            f'<iframe src="data:application/pdf;base64,{_pdf_b64}" width="100%" height="800" '
            f'style="border:1px solid var(--cb-border); border-radius:12px;"></iframe>'
        )
        st.caption(
            "Si el visor de acá arriba no se ve bien en tu navegador, usá el botón "
            "\"Descargar guía en PDF\"."
        )

# =========================================================================
# MÓDULOS — datos reales de ui_utils.MODULOS (la misma lista que usan las
# tarjetas de Inicio.py).
# =========================================================================
st.subheader("Módulos")
st.caption("Título, ícono y descripción de acá abajo vienen de la misma lista que usan las tarjetas de Inicio.")

# Instrucciones paso a paso por módulo (texto fijo, ver nota al principio
# del archivo sobre qué se actualiza solo y qué no). Clave = "pagina" de
# ui_utils.MODULOS, para que cada módulo real de la lista encuentre sus
# pasos sin depender del orden.
PASOS_POR_MODULO = {
    "pages/1_🤖CONTRATO_Excel.py": {
        "pasos": (
            "1. **📁 Proyecto y plantilla** — elegí el proyecto (o \"Otro / no listado\"), subí "
            "el Excel de ORIGEN del cliente y confirmá que el Excel MODELO esté cargado.\n"
            "2. **🧾 Datos de este cliente/cotización** — datos puntuales de esa cotización.\n"
            "3. **⚙️ Configuración de Pasos** — tildá qué Pasos correr (vienen todos tildados "
            "por defecto): Paso 2 Cotizador, Paso 3 FICHA, Paso 4 medidas del plano (Adelante/"
            "Derecha/Izquierda/Fondo + Sede), Paso 5 Ax.B, Paso 6 Membresía.\n"
            "4. **📋 Resumen antes de procesar** — revisá que esté todo bien antes de Procesar.\n"
            "5. **✅ Resultado** — descargá el Excel (sale nombrado \"COTIZADOR - <lote> "
            "<nombre del cliente>.xlsx\") y revisá el registro de actividad. Si el lote y el "
            "nombre se pudieron leer, esta misma copia se archiva sola — ver \"Archivo "
            "histórico\" más abajo."
        ),
        "extra": (
            "💬 Al final de esta página hay un cuadro de texto para \"enseñarle\" reglas al bot "
            "(ej. escribir \"Aprende: Ignorar telefono\") — sirve para que ciertos campos dejen "
            "de generar alertas de auditoría."
        ),
    },
    "pages/2_🤖CONTRATO_Word.py": {
        "pasos": (
            "1. **📁 Proyecto y plantillas** — elegí el proyecto y subí el Excel ya generado "
            "por el módulo Excel; elegí cuáles de los 4 documentos base generar.\n"
            "2. **🧾 Datos de este cliente/cotización** — igual que en Excel. Si vas a escribir "
            "un Número de Membresía, acá aparece un aviso con el último que se usó (y un "
            "historial de los últimos 20), para saber en qué número vas.\n"
            "3. **⚙️ Configuración del contrato** — 🏠 Tipo de propiedad (único propietario, "
            "casados, copropietarios...), 📕 Membresía (CON/SIN, cambia el ANEXO del "
            "documento), 🪪 Número de Membresía (reemplaza \"GTAA\" si aplica), y 📎 si pegar "
            "o no las imágenes de Excel dentro del CONTRATO.\n"
            "4. **📋 Resumen antes de procesar**.\n"
            "5. **✅ Resultado** — un botón DESCARGAR por cada documento generado (nombrados "
            "\"CONTRATO - <lote> <nombre>.docx\", \"REGLAMENTO - ...\", etc.), que también se "
            "archivan solos — ver \"Archivo histórico\" más abajo."
        ),
        "extra": (
            "El botón \"🚀 Procesar CONTRATO\" recalcula el Excel con Excel real primero "
            "(puede tardar); por eso la home mide y muestra el tiempo promedio de esta "
            "operación. El CONTRATO nunca sale con más de 15 páginas: si un nombre/texto muy "
            "largo hace que se pase a una página 16, esa página se recorta sola antes de "
            "mostrar el resultado."
        ),
    },
    "pages/3_🤖CONTRATO_PDF.py": {
        "pasos": (
            "1. Subí hasta **4 archivos .docx** a la vez.\n"
            "2. Tocá **Convertir a PDF**.\n"
            "3. Descargá cada PDF — sale con **el mismo nombre** con el que subiste el .docx "
            "(solo cambia la extensión)."
        ),
        "extra": (
            "Si subís más de 4 archivos, avisa y no deja convertir hasta que saques alguno. "
            "Si el .docx que subiste tiene el formato de nombre \"TIPO - <lote> <nombre>\" (el "
            "que descargan los módulos Excel/Word), el PDF se archiva solo en la misma carpeta "
            "del cliente — ver \"Archivo histórico\" más abajo."
        ),
    },
}

for modulo in MODULOS:
    detalle = PASOS_POR_MODULO.get(modulo["pagina"])
    with st.container(border=True):
        st.markdown(f'**{modulo["icono"]} {modulo["titulo"]}**')
        st.write(modulo["descripcion"])
        if detalle:
            st.markdown(detalle["pasos"])
            st.caption(detalle["extra"])
        st.page_link(modulo["pagina"], label=f'Abrir {modulo["titulo"]} →')

# =========================================================================
# LA PÁGINA DE INICIO
# =========================================================================
st.subheader("La página de Inicio")
with st.container(border=True):
    st.markdown(
        "- **🔍 Buscador** (arriba a la izquierda): filtra los módulos y las acciones rápidas "
        "mientras escribís (ej. escribí \"pdf\").\n"
        "- **🔔 Notificaciones**: avisa si hay errores registrados al procesar Excel o Word.\n"
        "- **⚙️ Configuración**: tu usuario, y el interruptor de **🧪 Modo de pruebas** (ver "
        "abajo).\n"
        "- **✏️ Crear nuevo contrato**: te lleva directo al módulo Word.\n"
        "- **Estadísticas**: dos apartados (Word y Excel) con contratos/Excel generados, "
        "de hoy, tiempo promedio y errores — son datos reales, se actualizan solos cada vez "
        "que procesás algo."
    )
    st.caption("Acciones rápidas de la home:")
    cols_acciones = st.columns(len(ACCIONES_RAPIDAS))
    for columna, accion in zip(cols_acciones, ACCIONES_RAPIDAS):
        with columna:
            st.page_link(accion["pagina"], label=f'{accion["icono"]}  {accion["label"]}')
    st.page_link("Inicio.py", label="Ir a Inicio →")

with st.container(border=True):
    st.markdown("**🗂️ ¿Qué es el archivo histórico?**")
    st.write(
        "Cada vez que descargás un Excel o un Word (CONTRATO, REGLAMENTO, MEMBRESIA o "
        "DERECHOHABIENTES) y la app pudo leer el lote y el nombre del cliente, se guarda "
        "automáticamente una copia en `C:\\JESUS LEGAL\\CONTRATOS\\<proyecto>\\<lote> "
        "<nombre del cliente>\\`, con el mismo nombre que descargaste. No hace falta hacer "
        "nada aparte — pasa solo, en segundo plano, después de procesar. Si por algún motivo "
        "no se pudo archivar (falta el lote/nombre en el Excel, o esa carpeta no existe en "
        "esta computadora), el Paso final de Excel/Word avisa el motivo en vez de fallar en "
        "silencio."
    )
    st.write(
        "Los números de Membresía que vas escribiendo también quedan en un historial propio "
        "(no en esa carpeta) — por eso el Paso 2 del módulo Word te muestra el último usado."
    )
    st.caption(
        "Con el Modo de pruebas activo, nada de esto se guarda (ni copias, ni números de "
        "membresía) — ver más abajo."
    )

with st.container(border=True):
    st.markdown("**🧪 ¿Qué es el Modo de pruebas?**")
    st.write(
        "Se activa desde ⚙️ Configuración, en la página de Inicio. Mientras está prendido, "
        "los Excel y Word que proceses **no se suman** a las estadísticas de \"generados\" ni "
        "de \"errores\", **no se archivan** en `C:\\JESUS LEGAL\\CONTRATOS\\` y **no se "
        "registra** el número de Membresía en el historial — útil para probar la app sin "
        "ensuciar los datos reales. "
        "Es una configuración compartida (afecta a cualquiera que use la app, no solo a vos), "
        "y queda prendida hasta que alguien la apague — por eso todas las páginas muestran un "
        "aviso 🧪 bien visible mientras está activa."
    )

# =========================================================================
# PROYECTOS CONFIGURADOS — datos reales de templates_manager.py: qué
# plantillas tiene guardada cada proyecto ahora mismo.
# =========================================================================
st.subheader("Proyectos configurados")
st.caption(
    "Las plantillas se agregan a mano en la carpeta PLANTILLAS/<PROYECTO>/ — esta página no "
    "permite subirlas, solo muestra si ya están."
)
for proyecto in obtener_proyectos():
    plantillas = obtener_plantillas_proyecto(proyecto)
    resumen = " · ".join(
        f"{etiqueta} {'✅' if plantillas[archivo]['existe'] else '❌'}"
        for etiqueta, archivo in ETIQUETAS_ARCHIVOS_PLANTILLA
    )
    with st.container(border=True):
        st.markdown(f"**{proyecto}**")
        st.caption(resumen)

# =========================================================================
# PREGUNTAS FRECUENTES
# =========================================================================
st.subheader("Preguntas frecuentes")
for item in PREGUNTAS_FRECUENTES:
    with st.expander(item["pregunta"]):
        st.write(item["respuesta"])

aplicar_fondo_proyecto(st.session_state.get("proyecto_actual"))
