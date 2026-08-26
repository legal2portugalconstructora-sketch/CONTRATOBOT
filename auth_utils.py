# -*- coding: utf-8 -*-
"""auth_utils.py — Login compartido para todas las páginas de ContratoBot.

Cada página de la app (Inicio.py y cada archivo en pages/) debe llamar a
requerir_login() justo después de st.set_page_config(...) y antes de
dibujar cualquier otro contenido. Si el usuario no inició sesión, la
función muestra el formulario de login y detiene la ejecución del script
con st.stop(), de forma que nada de la página quede visible.
"""

import os

import streamlit as st
import streamlit_authenticator as stauth
import yaml

RUTA_CONFIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.yaml")


def _cargar_config():
    if not os.path.exists(RUTA_CONFIG):
        st.error(
            "No se encontró config.yaml. Copiá config.example.yaml a "
            "config.yaml y completá los usuarios reales (ver README.md)."
        )
        st.stop()
    with open(RUTA_CONFIG, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def obtener_authenticator() -> stauth.Authenticate:
    """Crea (o reutiliza) el objeto Authenticate para esta sesión."""
    if "authenticator" not in st.session_state:
        config = _cargar_config()
        st.session_state["authenticator"] = stauth.Authenticate(
            config["credentials"],
            config["cookie"]["name"],
            config["cookie"]["key"],
            config["cookie"]["expiry_days"],
        )
    return st.session_state["authenticator"]


def requerir_login() -> stauth.Authenticate:
    """Bloquea la página hasta que el usuario inicie sesión.

    Devuelve el authenticator ya autenticado, para que la página pueda
    usarlo (por ejemplo, para el botón de logout en el sidebar).
    """
    authenticator = obtener_authenticator()
    authenticator.login(location="main")

    estado = st.session_state.get("authentication_status")

    if estado is False:
        st.error("Usuario o contraseña incorrectos.")
        st.stop()
    elif estado is None:
        st.info("Ingresá tu usuario y contraseña para continuar.")
        st.stop()

    return authenticator


def mostrar_boton_logout(authenticator: stauth.Authenticate) -> None:
    """Agrega el botón 'Cerrar sesión' y el saludo al usuario en el sidebar."""
    with st.sidebar:
        nombre = st.session_state.get("name", "")
        if nombre:
            st.caption(f"Sesión iniciada: **{nombre}**")
        authenticator.logout("Cerrar sesión", location="sidebar")
