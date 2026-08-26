# -*- coding: utf-8 -*-
"""Prueba de word_com_bridge._imagen_esta_en_blanco: detecta si un PNG es
un solo color liso de punta a punta (señal de que una captura de pantalla
de Excel salió en blanco -- ver la nota junto a Appearance=XL_SCREEN en
word_com_bridge.pegar_rangos_excel_como_imagen_en_word) sin dar falsos
positivos con una imagen que tiene contenido real.

No requiere Excel/Word instalados (no usa COM) -- solo Pillow.

Corré con: py tests\\test_imagen_en_blanco.py
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PIL import Image
from word_com_bridge import _imagen_esta_en_blanco

if __name__ == "__main__":
    rutas_temporales = []
    try:
        ruta_blanca = tempfile.mktemp(suffix=".png")
        Image.new("RGB", (200, 100), color=(255, 255, 255)).save(ruta_blanca)
        rutas_temporales.append(ruta_blanca)
        assert _imagen_esta_en_blanco(ruta_blanca), "FALLO: una imagen blanca sólida debería detectarse como en blanco"
        print("Imagen blanca sólida -> en blanco: True -- OK")

        ruta_gris = tempfile.mktemp(suffix=".png")
        Image.new("RGB", (200, 100), color=(120, 120, 120)).save(ruta_gris)
        rutas_temporales.append(ruta_gris)
        assert _imagen_esta_en_blanco(ruta_gris), "FALLO: cualquier color sólido debería detectarse como en blanco"
        print("Imagen gris sólida -> en blanco: True -- OK")

        ruta_contenido = tempfile.mktemp(suffix=".png")
        img = Image.new("RGB", (200, 100), color=(255, 255, 255))
        for x in range(0, 200, 10):
            for y in range(100):
                img.putpixel((x, y), (0, 0, 0))
        img.save(ruta_contenido)
        rutas_temporales.append(ruta_contenido)
        assert not _imagen_esta_en_blanco(ruta_contenido), "FALLO: una imagen con contenido real no debería marcarse en blanco"
        print("Imagen con contenido (rayas) -> en blanco: False -- OK")

        print("TODO OK")
    finally:
        for ruta in rutas_temporales:
            try:
                os.remove(ruta)
            except OSError:
                pass
