# -*- coding: utf-8 -*-
"""Prueba de la búsqueda de Inicio.py (módulos, acciones rápidas y FAQ):
tiene que tolerar errores de tipeo chicos y falta de tildes (algo muy común
al escribir rápido en español), sin por eso empezar a traer resultados sin
ninguna relación con lo buscado. Ver Inicio.py, función _coincide (usa
rapidfuzz.fuzz.partial_ratio sobre texto sin tildes).

Corré con: py tests\\test_busqueda_difusa.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from streamlit.testing.v1 import AppTest

RUTA_INICIO = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Inicio.py")


def _buscar(termino):
    at = AppTest.from_file(RUTA_INICIO, default_timeout=30)
    at.run()
    assert not at.exception, f"Inicio.py falló al cargar: {at.exception}"
    at.text_input(key="home_busqueda").set_value(termino).run()
    assert not at.exception, f"Falló al buscar '{termino}': {at.exception}"
    return (
        "\n".join(el.value for el in at.markdown)
        + "\n" + "\n".join(el.value for el in at.caption)
        + "\n" + "\n".join(el.value for el in at.info)
    )


if __name__ == "__main__":
    texto = _buscar("membresía")
    assert "membres" in texto.lower(), "Con tilde correcta no encontró nada"
    print("Con tilde correcta -- OK")

    texto = _buscar("membresia")
    assert "membres" in texto.lower(), "Sin tilde no encontró nada"
    print("Sin tilde ('membresia') encuentra resultados -- OK")

    texto = _buscar("membresa")
    assert "membres" in texto.lower(), "No toleró un error de tipeo chico"
    print("Con error de tipeo ('membresa') encuentra resultados -- OK")

    texto = _buscar("xyzqwerty123")
    assert "sin resultados" in texto.lower(), "Un término sin ninguna relación no debería dar resultados"
    print("Término sin relación -- 'Sin resultados', como se espera -- OK")

    texto = _buscar("")
    assert "sin resultados" not in texto.lower(), "La búsqueda vacía no debería mostrar 'Sin resultados'"
    print("Búsqueda vacía no rompe nada -- OK")

    print("TODO OK")
