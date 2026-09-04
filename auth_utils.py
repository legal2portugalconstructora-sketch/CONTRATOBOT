# -*- coding: utf-8 -*-
"""auth_utils.py — Compatibilidad con el login que tenía la app.

Antes, cada página llamaba a requerir_login() (que exigía usuario y
contraseña, ver config.yaml) y a mostrar_boton_logout(). Se sacó el login
por completo: la app ya no pide credenciales para entrar. Estas dos
funciones se dejan como no-operación (en vez de editar las 5 páginas que
las llaman) para no tocar nada más -- no bloquean nada ni muestran ningún
formulario.
"""


def requerir_login():
    return None


def mostrar_boton_logout(authenticator) -> None:
    return
