# -*- coding: utf-8 -*-
"""Prueba de convertir_docx_a_pdf_con_com (word_com_bridge.py): la
conversión a PDF tiene que preservar la resolución real de las imágenes
pegadas en el .docx, no recomprimirlas.

Bug real encontrado el 2026-09-22: usar Document.SaveAs(FileFormat=17)
para exportar a PDF aplica el perfil de calidad "para pantalla" por
defecto de Word, que recomprime MUY agresivamente las imágenes (probado:
una imagen de 4398x7586px quedaba en solo 500x863px, prácticamente
ilegible al hacer zoom en el PDF -- justo lo que el usuario reportó como
"sigue saliendo con muy baja resolución" DESPUÉS de ya haber subido la
resolución de captura en Excel). El fix fue cambiar a
Document.ExportAsFixedFormat(OptimizeFor=wdExportOptimizeForPrint), que
para la misma imagen de prueba preserva 1043x1799px (~4x más píxeles).

Requiere Windows + Word instalado (usa automatización COM real) y la
librería pymupdf (solo para esta prueba, no la usa la app).

Corré con: py tests\\test_pdf_calidad_imagen.py
"""
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from PIL import Image
import fitz  # pymupdf

from word_com_bridge import convertir_docx_a_pdf_con_com

# Umbral bien por debajo de lo que da hoy ExportAsFixedFormat (~1800px para
# una imagen fuente de 2400px), pero MUY por encima de lo que daba el bug
# de SaveAs (~500px sin importar el tamaño de la fuente) -- así la prueba
# detecta si alguna vez se vuelve a usar SaveAs por error, sin depender de
# un número exacto de píxeles que pueda variar levemente entre versiones
# de Word.
UMBRAL_MINIMO_PX = 1200


def _armar_docx_con_imagen_grande():
    """Imagen sintética de 2400x1800px con contenido variado (no un solo
    color liso), para que la compresión de la conversión a PDF sea
    representativa de una imagen real (no de un color plano, que
    comprimiría distinto)."""
    imagen = Image.new("RGB", (2400, 1800))
    pixeles = imagen.load()
    for x in range(0, 2400, 10):
        for y in range(0, 1800, 10):
            color = ((x * 7) % 256, (y * 13) % 256, ((x + y) * 3) % 256)
            for dx in range(10):
                for dy in range(10):
                    if x + dx < 2400 and y + dy < 1800:
                        pixeles[x + dx, y + dy] = color

    buf_imagen = io.BytesIO()
    imagen.save(buf_imagen, format="PNG")
    buf_imagen.seek(0)

    doc = Document()
    doc.add_picture(buf_imagen, width=None)  # tamaño nativo, como el pegado real
    buf_docx = io.BytesIO()
    doc.save(buf_docx)
    return buf_docx.getvalue()


if __name__ == "__main__":
    docx_bytes = _armar_docx_con_imagen_grande()
    print("Documento de prueba armado (imagen de 2400x1800px pegada).")

    pdf_bytes = convertir_docx_a_pdf_con_com(docx_bytes)
    print(f"PDF generado: {len(pdf_bytes) / 1024:.0f} KB")

    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        pagina = pdf[0]
        imagenes = pagina.get_images(full=True)
        assert len(imagenes) == 1, f"Se esperaba 1 imagen en el PDF, hay {len(imagenes)}"

        xref = imagenes[0][0]
        pix = fitz.Pixmap(pdf, xref)
        print(f"Imagen dentro del PDF: {pix.width}x{pix.height} px")
        assert pix.width >= UMBRAL_MINIMO_PX, (
            f"FALLO: la imagen del PDF salió en {pix.width}px de ancho, por debajo del umbral "
            f"({UMBRAL_MINIMO_PX}px) -- parece que la conversión está recomprimiendo demasiado "
            f"de nuevo (¿volvió a usarse doc.SaveAs en vez de doc.ExportAsFixedFormat?)."
        )
    finally:
        pdf.close()

    print("Resolución preservada correctamente en el PDF -- OK")
    print("TODO OK")
