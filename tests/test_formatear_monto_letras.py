# -*- coding: utf-8 -*-
"""Prueba de excel_pipeline.formatear_monto_letras -- en particular, la
corrección de montos que quedan a 1 centavo de un número redondo por
acarreo de redondeo de Excel (ej. 64999.99 -> "...SESENTA Y CINCO MIL...
CON 00/100...", no "...SESENTA Y CUATRO MIL NOVECIENTOS NOVENTA Y
NUEVE... CON 99/100...").

Corré con: py tests\\test_formatear_monto_letras.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_pipeline import formatear_monto_letras

CASOS = [
    (64999.99, "SESENTA Y CINCO MIL CON 00/100"),    # el caso real que motivó el arreglo
    (65000.01, "SESENTA Y CINCO MIL CON 00/100"),    # simétrico, 1 centavo arriba
    (65000.00, "SESENTA Y CINCO MIL CON 00/100"),    # ya redondo, no debería cambiar
    (50000.60, "CINCUENTA MIL CON 60/100"),          # NO debe tocarse (60 centavos reales)
    (64999.98, "SESENTA Y CUATRO MIL NOVECIENTOS NOVENTA Y NUEVE CON 98/100"),  # fuera del umbral (2 centavos)
]

if __name__ == "__main__":
    for valor, esperado_contiene in CASOS:
        resultado = formatear_monto_letras(valor)
        assert esperado_contiene in resultado, (
            f"FALLO: formatear_monto_letras({valor}) = {resultado!r}, "
            f"esperaba que contuviera {esperado_contiene!r}"
        )
        print(f"{valor} -> {resultado}  OK")
    print("TODO OK")
