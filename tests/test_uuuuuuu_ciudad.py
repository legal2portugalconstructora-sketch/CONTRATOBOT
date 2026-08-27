# -*- coding: utf-8 -*-
"""Prueba de word_pipeline.leer_datos_para_contrato: el placeholder
UUUUUUU debe salir como "en la ciudad de <Sede> el <fecha en letras>"
cuando se puede determinar la Sede (Arequipa/Ilo) a partir de la fórmula
escrita en la celda "g124" (ver excel_pipeline.procesar_paso4_axa) -- y
debe funcionar igual en COSTA DEL VALLE (celdas por defecto) y en
proyectos con celdas remapeadas (ej. COSTA AZUL). Si no se puede
determinar la Sede (ej. el Paso 4 no se corrió), debe quedar solo la
fecha, sin romper nada.

Corré con: py tests\\test_uuuuuuu_ciudad.py
"""
import io
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook
from word_pipeline import leer_datos_para_contrato
from templates_manager import CELDAS_AXA_POR_DEFECTO, obtener_celdas_axa


def _armar_excel(celdas_axa, escribir_g124=True, sede="Arequipa", fecha=date(2026, 8, 15), formula_o132=None):
    wb = Workbook()
    ws_datos = wb.active
    ws_datos.title = "DATOS"
    ws_datos["E11"] = "JUAN PEREZ"
    ws_datos["E13"] = "12345678"
    ws_datos["E22"] = "AV. TEST 123"
    ws_datos["H26"] = "AREQUIPA"
    ws_datos["H27"] = "AREQUIPA"
    ws_datos["H28"] = "AREQUIPA"

    ws_cotiz = wb.create_sheet("Cotizador")
    if fecha is not None:
        ws_cotiz["I16"] = fecha

    ws_axa = wb.create_sheet("Ax.A")
    ws_axa[celdas_axa["o132"]] = formula_o132 or "=Cotizador!I16"
    if escribir_g124:
        ws_axa[celdas_axa["g124"]] = celdas_axa["sedes_validas"][sede]

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


if __name__ == "__main__":
    celdas_cdv = CELDAS_AXA_POR_DEFECTO

    datos = leer_datos_para_contrato(_armar_excel(celdas_cdv, sede="Arequipa"), celdas_axa=celdas_cdv)
    assert datos["UUUUUUU"] == "en la ciudad de Arequipa el 15 de agosto del 2026", datos["UUUUUUU"]
    print("COSTA DEL VALLE, Arequipa:", datos["UUUUUUU"], "OK")

    datos = leer_datos_para_contrato(_armar_excel(celdas_cdv, sede="Ilo"), celdas_axa=celdas_cdv)
    assert datos["UUUUUUU"] == "en la ciudad de Ilo el 15 de agosto del 2026", datos["UUUUUUU"]
    print("COSTA DEL VALLE, Ilo:", datos["UUUUUUU"], "OK")

    celdas_ca = obtener_celdas_axa("COSTA AZUL")
    datos = leer_datos_para_contrato(_armar_excel(celdas_ca, sede="Arequipa"), celdas_axa=celdas_ca)
    assert datos["UUUUUUU"] == "en la ciudad de Arequipa el 15 de agosto del 2026", datos["UUUUUUU"]
    print("COSTA AZUL (celdas remapeadas), Arequipa:", datos["UUUUUUU"], "OK")

    datos = leer_datos_para_contrato(_armar_excel(celdas_cdv, escribir_g124=False), celdas_axa=celdas_cdv)
    assert datos["UUUUUUU"] == "15 de agosto del 2026", datos["UUUUUUU"]
    print("Sin sede detectada (fallback, solo fecha):", datos["UUUUUUU"], "OK")

    # Sede detectada pero SIN fecha (Cotizador!I16 vacío) -- no debe salir
    # "en la ciudad de Arequipa el None" (bug real encontrado al agregar
    # FINCA DEL VALLE).
    datos = leer_datos_para_contrato(_armar_excel(celdas_cdv, sede="Arequipa", fecha=None), celdas_axa=celdas_cdv)
    assert datos["UUUUUUU"] is None, f"FALLO: se esperaba None, vino {datos['UUUUUUU']!r}"
    print("Sede detectada pero sin fecha (no dice '...el None'):", datos["UUUUUUU"], "OK")

    # Referencia a libro externo (ej. "=[1]Cotizador!I16") -- visto de
    # verdad en el MODELO de FINCA DEL VALLE. Debe resolverse igual que
    # una referencia local.
    datos = leer_datos_para_contrato(
        _armar_excel(celdas_cdv, sede="Ilo", formula_o132="=[1]Cotizador!I16"), celdas_axa=celdas_cdv,
    )
    assert datos["UUUUUUU"] == "en la ciudad de Ilo el 15 de agosto del 2026", datos["UUUUUUU"]
    print("Referencia a libro externo '=[1]Cotizador!I16' se resuelve igual que una local:", datos["UUUUUUU"], "OK")

    print("TODO OK")
