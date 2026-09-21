# -*- coding: utf-8 -*-
"""Prueba de word_com_bridge._cerrar_sin_fallar: una función de limpieza
COM (Close()/Quit()/Delete() de un doc/app/chart, usada en los 'finally'
de este módulo) que falla NUNCA debe propagar esa excepción hacia afuera.

Esto importa porque en Python, si un 'finally' lanza una excepción, esa
excepción REEMPLAZA silenciosamente lo que haya pasado en el 'try' -- hasta
un 'return' con el resultado ya listo. Se vio en la práctica: un error
transitorio de COM al cerrar Excel real (AttributeError
"Excel.Application.Quit", justo el tipo de error que NO es un
pywintypes.com_error) tapaba un recálculo que en realidad había terminado
bien, y la app le mostraba al usuario un error sobre un archivo que sí se
había generado correctamente.

No requiere Excel/Word instalados (no usa COM de verdad, solo funciones de
prueba que fallan a propósito).

Corré con: py tests\\test_cerrar_sin_fallar.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from word_com_bridge import _cerrar_sin_fallar

if __name__ == "__main__":
    # --- Falla SIEMPRE (ej. el objeto COM quedó en un estado roto) -- no
    # debe propagar la excepción, aunque reintente unas pocas veces. ---
    llamadas = {"n": 0}

    def siempre_falla():
        llamadas["n"] += 1
        raise AttributeError("Excel.Application.Quit")

    _cerrar_sin_fallar(siempre_falla)  # no debe lanzar nada
    assert llamadas["n"] > 1, "debería haber reintentado más de una vez"
    print("Falla persistente en la limpieza: no se propaga -- OK")

    # --- Falla las primeras veces y después funciona (error transitorio
    # real) -- debe terminar sin excepción. ---
    llamadas2 = {"n": 0}

    def falla_dos_veces():
        llamadas2["n"] += 1
        if llamadas2["n"] < 3:
            raise AttributeError("Word.Application.Quit")
        return "ok"

    _cerrar_sin_fallar(falla_dos_veces)
    assert llamadas2["n"] == 3
    print("Falla transitoria en la limpieza (se recupera sola) -- OK")

    print("TODO OK")
