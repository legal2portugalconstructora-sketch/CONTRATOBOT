# -*- coding: utf-8 -*-
"""Prueba de historial_contratos.py: guardado de copias archivadas,
respeto del Modo de pruebas, búsqueda de carpeta por prefijo (para el
módulo PDF), y el historial de números de membresía.

OJO: esta prueba escribe y borra una carpeta de prueba DENTRO de la
carpeta real de archivo histórico (historial_contratos.RAIZ_HISTORIAL,
hoy C:\\JESUS LEGAL\\CONTRATOS\\) -- usa un nombre de proyecto claramente
falso (ZZZ_TEST_...) y se borra a sí misma al final (incluso si algo
falla a mitad de camino), para no dejar basura ahí. Si RAIZ_HISTORIAL no
existe en la máquina donde se corre esto, la prueba avisa y no falla
(el resto de la app tampoco falla en ese caso -- ver guardar_copia).

Corré con: py tests\\test_historial_contratos.py
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import historial_contratos as hc
from stats_tracker import establecer_modo_pruebas

PROYECTO_TEST = "ZZZ_TEST_PROYECTO_BORRAR"


def _limpiar():
    """Borra solo la carpeta de proyecto de prueba. El historial de
    membresías se restaura aparte (ver el try/finally más abajo), no se
    toca acá -- nunca se borra el archivo real completo."""
    carpeta_proyecto_test = os.path.join(hc.RAIZ_HISTORIAL, PROYECTO_TEST)
    if os.path.isdir(carpeta_proyecto_test):
        shutil.rmtree(carpeta_proyecto_test)


if __name__ == "__main__":
    carpeta_padre = os.path.dirname(hc.RAIZ_HISTORIAL)
    if not os.path.isdir(carpeta_padre):
        print(f"AVISO: '{carpeta_padre}' no existe en esta máquina -- se salta la prueba de archivado.")
        sys.exit(0)

    establecer_modo_pruebas(False)
    try:
        # --- prefijo_archivo / saneo de caracteres inválidos ---
        assert hc.prefijo_archivo("L-2", "JUAN PEREZ TEST") == "L-2 JUAN PEREZ TEST"
        sucio = hc.prefijo_archivo("L-2/3", 'Cliente "Raro": Test')
        assert "/" not in sucio and '"' not in sucio and ":" not in sucio
        print("prefijo_archivo sanea caracteres inválidos -- OK")

        # --- guardar_copia (modo pruebas apagado) ---
        ruta, motivo = hc.guardar_copia(PROYECTO_TEST, "L-2", "JUAN PEREZ TEST", "COTIZADOR.xlsx", b"contenido")
        assert ruta is not None and motivo is None and os.path.exists(ruta)
        assert os.path.basename(ruta) == "COTIZADOR - L-2 JUAN PEREZ TEST.xlsx"
        print("guardar_copia guarda con el nombre correcto -- OK")

        # --- nombre_archivo_descarga (usado por los botones de descarga) ---
        assert hc.nombre_archivo_descarga("CONTRATO", "L-2", "JUAN PEREZ TEST", "docx") == "CONTRATO - L-2 JUAN PEREZ TEST.docx"
        assert hc.nombre_archivo_descarga("CONTRATO", "L-2", "JUAN PEREZ TEST", ".docx") == "CONTRATO - L-2 JUAN PEREZ TEST.docx"
        assert hc.nombre_archivo_descarga("CONTRATO", None, None, "docx") == "CONTRATO.docx"
        print("nombre_archivo_descarga arma el nombre en el orden nuevo -- OK")

        # --- guardar_copia respeta el Modo de pruebas ---
        establecer_modo_pruebas(True)
        ruta2, motivo2 = hc.guardar_copia(PROYECTO_TEST, "L-2", "JUAN PEREZ TEST", "CONTRATO.docx", b"x")
        assert ruta2 is None and motivo2 is not None and "pruebas" in motivo2
        establecer_modo_pruebas(False)
        print("guardar_copia no guarda nada con el Modo de pruebas activo -- OK")

        # --- motivos de fallo (sin lote / sin nombre) ---
        ruta3, motivo3 = hc.guardar_copia(PROYECTO_TEST, None, "JUAN PEREZ", "X.docx", b"x")
        assert ruta3 is None and "lote" in motivo3
        ruta4, motivo4 = hc.guardar_copia(PROYECTO_TEST, "L-2", None, "X.docx", b"x")
        assert ruta4 is None and "nombre del cliente" in motivo4
        print("guardar_copia explica el motivo cuando falta lote/nombre -- OK")

        # --- encontrar_carpeta_por_prefijo (usado por el módulo PDF) ---
        carpeta_encontrada = hc.encontrar_carpeta_por_prefijo("L-2 JUAN PEREZ TEST")
        assert carpeta_encontrada is not None and PROYECTO_TEST in carpeta_encontrada
        print("encontrar_carpeta_por_prefijo encuentra la carpeta ya creada -- OK")

        # --- historial de números de membresía (con limpieza cuidadosa) ---
        eventos_previos = hc._cargar_membresias()
        try:
            hc.registrar_numero_membresia(PROYECTO_TEST, "JUAN PEREZ TEST", "GTAA-99999")
            ultimo = hc.obtener_ultimo_numero_membresia()
            assert ultimo["numero_membresia"] == "GTAA-99999"
            print("registrar_numero_membresia / obtener_ultimo_numero_membresia -- OK")

            establecer_modo_pruebas(True)
            hc.registrar_numero_membresia(PROYECTO_TEST, "JUAN PEREZ TEST", "GTAA-00000-NO-DEBERIA-QUEDAR")
            ultimo2 = hc.obtener_ultimo_numero_membresia()
            assert ultimo2["numero_membresia"] == "GTAA-99999", "FALLO: se registró un número con el Modo de pruebas activo"
            establecer_modo_pruebas(False)
            print("registrar_numero_membresia respeta el Modo de pruebas -- OK")
        finally:
            # Restaurar el historial real de membresías tal cual estaba antes de la prueba.
            hc._guardar_membresias(eventos_previos)

        print("TODO OK")
    finally:
        establecer_modo_pruebas(False)
        _limpiar()
