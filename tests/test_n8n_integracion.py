# -*- coding: utf-8 -*-
"""Prueba de n8n_integracion.py: clave del webhook (X-Api-Key), rango de
numeración de ID_CONTRATO por PC (para que dos PCs no se pisen al
escribirle al mismo n8n en la nube), y los dos avisos nuevos en
pantalla (clave incorrecta / ID duplicado).

No depende de un n8n real ni de Excel/Word real -- todo lo que llama a
la red se reemplaza por una respuesta falsa (monkeypatch de
requests.post), y los archivos de estado (correlativos_contrato.json,
pendientes_n8n.json, config.yaml) se apuntan a una carpeta temporal
nueva, nunca a los reales.

Corré con: py tests\\test_n8n_integracion.py
"""
import io
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import requests

import n8n_integracion as n8n
from stats_tracker import establecer_modo_pruebas


def _escribir_config(texto):
    with open(n8n.RUTA_CONFIG, "w", encoding="utf-8") as f:
        f.write(texto)


def _excel_minimo_valido():
    """Un .xlsx mínimo con las hojas DATOS/Cotizador vacías -- suficiente
    para que armar_payload() no falle (sus celdas simplemente quedan
    vacías/0, ver n8n_integracion.armar_payload)."""
    from openpyxl import Workbook

    wb = Workbook()
    wb.active.title = "DATOS"
    wb.create_sheet("Cotizador")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class _RespuestaFalsa:
    def __init__(self, status_code, texto=""):
        self.status_code = status_code
        self.text = texto


if __name__ == "__main__":
    carpeta_tmp = tempfile.mkdtemp()
    n8n.RUTA_CORRELATIVOS_CONTRATO = os.path.join(carpeta_tmp, "correlativos_contrato.json")
    n8n.RUTA_PENDIENTES_N8N = os.path.join(carpeta_tmp, "pendientes_n8n.json")
    n8n.RUTA_CONFIG = os.path.join(carpeta_tmp, "config.yaml")
    _post_original = requests.post

    try:
        establecer_modo_pruebas(True)

        # --- 1) Sin claves nuevas: rango default 1-9999, sin api key ---
        _escribir_config("n8n_webhook_url: http://x\n")
        assert n8n._leer_rango_id_contrato() == (1, 9999, None)
        assert n8n._leer_api_key_webhook() is None
        print("Sin claves nuevas -- rango default 1-9999 y sin api key -- OK")

        # --- 2) Rango 5001-9999 con correlativo en 0 -> primero es 5001 ---
        _escribir_config("id_contrato_rango_inicio: 5001\nid_contrato_rango_fin: 9999\n")
        inicio, fin, error = n8n._leer_rango_id_contrato()
        assert (inicio, fin, error) == (5001, 9999, None)
        numero, error_num = n8n._siguiente_correlativo("ZZZ_A", 2026, True, inicio, fin)
        assert (numero, error_num) == (5001, None)
        print("Rango 5001-9999 con correlativo en 0 -> primero asignado es 5001 -- OK")

        # --- 3) Correlativo ya en 12 con rango 1-4999 -> siguiente es 13 ---
        numero = None
        for _ in range(12):
            numero, error_num = n8n._siguiente_correlativo("ZZZ_B", 2026, True, 1, 4999)
            assert error_num is None
        assert numero == 12
        numero, error_num = n8n._siguiente_correlativo("ZZZ_B", 2026, True, 1, 4999)
        assert numero == 13 and error_num is None
        print("Correlativo ya en 12 con rango 1-4999 -> siguiente es 13 -- OK")

        # --- 4) Correlativo en el tope del rango -> error, no incrementa ---
        numero, error_num = n8n._siguiente_correlativo("ZZZ_C", 2026, True, 1, 1)
        assert numero == 1 and error_num is None
        numero2, error_num2 = n8n._siguiente_correlativo("ZZZ_C", 2026, True, 1, 1)
        assert numero2 is None
        assert error_num2 is not None and error_num2.startswith("No se puede asignar ID_CONTRATO")
        guardado = n8n._cargar_correlativos()["pruebas"]["ZZZ_C-2026"]
        assert guardado == 1, "FALLO: el correlativo se incrementó aunque se pasó del rango"
        print("Correlativo en el tope del rango -> error claro y no incrementa nada -- OK")

        # --- 5) rango_inicio presente pero no es entero -> error, sin default ---
        _escribir_config("id_contrato_rango_inicio: 'no es un numero'\n")
        inicio, fin, error = n8n._leer_rango_id_contrato()
        assert inicio is None and fin is None
        assert error is not None and "id_contrato_rango_inicio" in error
        print("id_contrato_rango_inicio no entero -> error, no cae al default -- OK")

        # --- 6) fuera de 1-9999 / inicio > fin -> error, sin default ---
        _escribir_config("id_contrato_rango_fin: 10000\n")
        inicio, fin, error = n8n._leer_rango_id_contrato()
        assert inicio is None and fin is None and error is not None
        print("id_contrato_rango_fin fuera de 1-9999 -> error -- OK")

        _escribir_config("id_contrato_rango_inicio: 100\nid_contrato_rango_fin: 50\n")
        inicio, fin, error = n8n._leer_rango_id_contrato()
        assert inicio is None and fin is None and error is not None
        print("id_contrato_rango_inicio > id_contrato_rango_fin -> error -- OK")

        # --- 7) X-Api-Key solo cuando hay clave ---
        llamadas = []

        def _post_registra_headers(url, json=None, timeout=None, headers=None):
            llamadas.append(headers)
            return _RespuestaFalsa(200)

        requests.post = _post_registra_headers
        n8n._enviar("http://x", {"a": 1}, api_key=None)
        n8n._enviar("http://x", {"a": 1}, api_key="clave-secreta")
        assert llamadas[0] is None
        assert llamadas[1] == {"X-Api-Key": "clave-secreta"}
        print("_enviar manda X-Api-Key solo cuando hay una clave configurada -- OK")

        # --- 8a) registrar_contrato_en_n8n: 403 -> advertencia de clave incorrecta ---
        _escribir_config(
            "n8n_webhook_url: http://x\n"
            "n8n_webhook_api_key: clave-mala\n"
            "codigos_proyecto:\n  ZZZ_TEST_PROYECTO: ZZZ8A\n"
        )
        requests.post = lambda *a, **k: _RespuestaFalsa(403)
        excel_bytes = _excel_minimo_valido()
        estado, advertencias, id_contrato = n8n.registrar_contrato_en_n8n(
            excel_bytes, "ZZZ_TEST_PROYECTO", "L-1", "CLIENTE DE PRUEBA", "0001", "UNICO PROPIETARIO",
        )
        assert id_contrato is not None and id_contrato.startswith("PRUEBA-ZZZ8A-")
        assert any("n8n_webhook_api_key" in a and "incorrecta" in a for a in advertencias), advertencias
        print("registrar_contrato_en_n8n con 403 -> advertencia de clave incorrecta -- OK")

        # --- 8b) registrar_contrato_en_n8n: cuerpo con "ID duplicado" -> advertencia de conflicto ---
        _escribir_config(
            "n8n_webhook_url: http://x\n"
            "codigos_proyecto:\n  ZZZ_TEST_PROYECTO: ZZZ8B\n"
        )
        requests.post = lambda *a, **k: _RespuestaFalsa(409, "Error: ID duplicado en la hoja")
        estado, advertencias, id_contrato = n8n.registrar_contrato_en_n8n(
            excel_bytes, "ZZZ_TEST_PROYECTO", "L-1", "CLIENTE DE PRUEBA", "0001", "UNICO PROPIETARIO",
        )
        assert any("ID_CONTRATO duplicado" in a and "CONFLICTOS" in a for a in advertencias), advertencias
        assert estado == "Pendiente de registrar (n8n apagado)"
        pendientes = n8n._cargar_pendientes()
        assert any(p["ID_CONTRATO"] == id_contrato for p in pendientes), pendientes
        print("registrar_contrato_en_n8n con 'ID duplicado' en el cuerpo -> advertencia de conflicto -- OK")

        print("TODO OK")
    finally:
        requests.post = _post_original
        establecer_modo_pruebas(False)
        shutil.rmtree(carpeta_tmp, ignore_errors=True)
