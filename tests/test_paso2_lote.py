# -*- coding: utf-8 -*-
"""Prueba de excel_pipeline.procesar_paso2_cotizador: la celda del LOTE
(Cotizador!E15) se copia del ORIGEN al MODELO y se pasa a mayúsculas, y
la celda E14 (que NO es el lote, un error de configuración anterior la
usaba por accidente) queda intacta, tal cual estaba en el MODELO.

También prueba word_pipeline.extraer_lote_y_nombre (en historial_contratos.py)
leyendo esa misma celda E15 para el historial de contratos.

Corré con: py tests\\test_paso2_lote.py
"""
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook
from excel_pipeline import procesar_paso2_cotizador, _lee
from historial_contratos import extraer_lote_y_nombre


def _armar_origen_minimo():
    wb = Workbook()
    ws = wb.active
    ws.title = "Cotizador"
    ws["E14"] = "NO_TOCAR"
    ws["E15"] = "l-2"  # minúscula a propósito, para confirmar que se normaliza
    ws["E16"] = "=I20"
    ws["I20"] = "x"
    ws["I21"] = 100
    ws["I25"] = 1
    ws["I33"] = "ok"
    ws["E35"] = "asesor"
    ws["D33"] = "x"
    return wb


if __name__ == "__main__":
    wb_ori = _armar_origen_minimo()

    wb_mod = Workbook()
    ws_mod = wb_mod.active
    ws_mod.title = "Cotizador"
    ws_mod["I33"] = "ok"
    ws_mod["E14"] = "VALOR_ORIGINAL_MODELO"

    alertas, resumen = [], {}
    procesar_paso2_cotizador(wb_ori, wb_mod, alertas, resumen)

    valor_e14 = _lee(wb_mod["Cotizador"], "E14")
    assert valor_e14 == "VALOR_ORIGINAL_MODELO", f"FALLO: E14 se tocó (quedó {valor_e14!r})"
    print("E14 queda intacto:", repr(valor_e14), "OK")

    valor_e15 = _lee(wb_mod["Cotizador"], "E15")
    assert valor_e15 == "L-2", f"FALLO: E15 debería ser 'L-2', vino {valor_e15!r}"
    print("E15 (lote) copiado y normalizado:", repr(valor_e15), "OK")

    # extraer_lote_y_nombre debe leer el lote de la MISMA celda (E15)
    wb_final = Workbook()
    ws_datos = wb_final.active
    ws_datos.title = "DATOS"
    ws_datos["E11"] = "CLIENTE PRUEBA"
    ws_cotiz = wb_final.create_sheet("Cotizador")
    ws_cotiz["E15"] = "L-7"
    ws_cotiz["E14"] = "BASURA_NO_USAR"
    buf = io.BytesIO()
    wb_final.save(buf)

    lote, nombre = extraer_lote_y_nombre(buf.getvalue())
    assert lote == "L-7" and nombre == "CLIENTE PRUEBA", f"FALLO: extraer_lote_y_nombre devolvió ({lote!r}, {nombre!r})"
    print("extraer_lote_y_nombre lee de E15:", (lote, nombre), "OK")

    print("TODO OK")
