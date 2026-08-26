# -*- coding: utf-8 -*-
"""Prueba de excel_pipeline._formatear_celdas_segundo_titular_copropietario:
si DATOS!E33 (ORIGEN) tiene un dato real, en FICHA/Ax.A/Ax.B/Ax.E del
MODELO, cualquier celda que diga exactamente "SEGUNDO TITULAR/COPROPIETARIO"
hace que la celda de ARRIBA reciba fuente Negro, y las 2 celdas de ABAJO
reciban fuente Automático -- incluyendo el caso de celdas combinadas y el
caso límite de que la etiqueta esté en la fila 1 (no hay "arriba").

Corré con: py tests\\test_segundo_titular.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook
from excel_pipeline import _procesar_paso1_datos_wb, FILA_INICIO_DATOS, FILA_FIN_DATOS


def _es_auto(celda):
    return celda.font.color is not None and celda.font.color.type == "auto"


def _es_negro(celda):
    return (
        celda.font.color is not None
        and celda.font.color.type == "rgb"
        and celda.font.color.rgb in ("FF000000", "00000000", "000000")
    )


if __name__ == "__main__":
    wb_ori = Workbook()
    ws_datos_ori = wb_ori.active
    ws_datos_ori.title = "DATOS"
    for fila in range(FILA_INICIO_DATOS, FILA_FIN_DATOS + 1):
        ws_datos_ori[f"C{fila}"] = f"ETIQUETA{fila}"
        ws_datos_ori[f"E{fila}"] = "-"
    ws_datos_ori["E33"] = "JUAN SEGUNDO"  # dato real -> dispara el formateo

    wb_mod = Workbook()
    ws_datos_mod = wb_mod.active
    ws_datos_mod.title = "DATOS"
    for fila in range(FILA_INICIO_DATOS, FILA_FIN_DATOS + 1):
        ws_datos_mod[f"C{fila}"] = f"ETIQUETA{fila}"

    ws_axa = wb_mod.create_sheet("Ax.A")
    ws_axa["B9"] = "arriba"
    ws_axa["B10"] = "SEGUNDO TITULAR/COPROPIETARIO"
    ws_axa["B11"] = "nombre aca"
    ws_axa["B12"] = "dni aca"

    ws_axb = wb_mod.create_sheet("Ax.B")
    ws_axb["C19"] = "arriba combinada"
    ws_axb.merge_cells("C19:D19")
    ws_axb["C20"] = "SEGUNDO TITULAR/COPROPIETARIO"
    ws_axb.merge_cells("C21:D21")  # celda combinada justo debajo
    ws_axb["C21"] = "nombre combinado"
    ws_axb["C22"] = "dni aca"

    ws_axe = wb_mod.create_sheet("Ax.E")
    ws_axe["A5"] = "SEGUNDO TITULAR/COPROPIETARIO"
    ws_axe["A6"] = "nombre aca"
    ws_axe["A7"] = "dni aca"
    # Caso límite: etiqueta en la fila 1 (no hay "arriba") -- no debe explotar
    ws_axe["C1"] = "SEGUNDO TITULAR/COPROPIETARIO"
    ws_axe["C2"] = "nombre fila1"
    ws_axe["C3"] = "dni fila1"

    ws_ficha = wb_mod.create_sheet("FICHA")
    ws_ficha["A1"] = "arriba ficha"
    ws_ficha["A2"] = "SEGUNDO TITULAR/COPROPIETARIO"
    ws_ficha["A3"] = "nombre ficha"
    ws_ficha["A4"] = "dni ficha"

    alertas, resumen = [], {"filas_procesadas": 0, "filas_con_guion_ignoradas": 0}
    _procesar_paso1_datos_wb(wb_ori, wb_mod, [], alertas, resumen)

    assert resumen.get("segundo_titular_celdas_formateadas") == 5, (
        f"FALLO: se esperaban 5 coincidencias (Ax.A, Ax.B, Ax.E x2, FICHA), "
        f"vinieron {resumen.get('segundo_titular_celdas_formateadas')}"
    )
    print("Coincidencias encontradas: 5, OK")

    assert _es_negro(ws_axa["B9"]), "FALLO: Ax.A!B9 (arriba) debería ser negro"
    assert _es_auto(ws_axa["B11"]) and _es_auto(ws_axa["B12"]), "FALLO: Ax.A!B11/B12 (abajo) deberían ser automático"
    print("Ax.A: arriba negro, abajo automático -- OK")

    assert _es_negro(ws_axb["C19"]), "FALLO: Ax.B!C19 (arriba, combinada) debería ser negro"
    assert _es_auto(ws_axb["C21"]) and _es_auto(ws_axb["C22"]), "FALLO: Ax.B!C21/C22 (abajo) deberían ser automático"
    print("Ax.B (con celdas combinadas): arriba negro, abajo automático -- OK")

    assert _es_auto(ws_axe["A6"]) and _es_auto(ws_axe["A7"]), "FALLO: Ax.E!A6/A7 deberían ser automático"
    assert _es_auto(ws_axe["C2"]) and _es_auto(ws_axe["C3"]), "FALLO: caso fila 1 no debería romper el 'abajo'"
    print("Ax.E: caso normal y caso límite (etiqueta en fila 1) -- OK")

    assert _es_negro(ws_ficha["A1"]), "FALLO: FICHA!A1 (arriba) debería ser negro -- FICHA debe procesarse"
    assert _es_auto(ws_ficha["A3"]) and _es_auto(ws_ficha["A4"]), "FALLO: FICHA!A3/A4 (abajo) deberían ser automático"
    print("FICHA (incluida en el escaneo): arriba negro, abajo automático -- OK")

    print("TODO OK")
