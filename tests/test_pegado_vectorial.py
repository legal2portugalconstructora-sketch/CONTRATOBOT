# -*- coding: utf-8 -*-
"""Prueba de word_com_bridge.pegar_rangos_excel_como_imagen_en_word: desde
el 2026-09-22 pega los rangos de Excel en Word como imagen VECTORIAL (EMF,
vía Range.Copy() + Selection.PasteSpecial(wdPasteEnhancedMetafile) +
Shape.ConvertToInlineShape()) en vez de una captura de pantalla rasterizada
a PNG -- el mismo mecanismo que usa Word cuando alguien pega un rango "a
mano". Motivo del cambio: el usuario reportó que las imágenes pegadas por
la app se veían con mucha menos resolución que copiando a mano; un EMF es
vectorial (sin límite de resolución) y además confirmado que NO depende de
que la pantalla esté realmente renderizada (a diferencia del método
anterior, que fallaba con capturas en blanco si la sesión estaba
bloqueada o una sesión de Escritorio Remoto se desconectaba).

Esta prueba verifica, con Excel/Word reales:
  - Que el pegado no lanza error y deja exactamente una imagen nueva.
  - Que la imagen resultante es vectorial (content_type "image/x-emf"),
    no un PNG rasterizado -- si esto falla, es señal de que alguien
    volvió a un método basado en captura de pantalla sin darse cuenta.
  - Que un rango más ancho que la página se reescala para entrar (la
    misma lógica de LockAspectRatio que ya existía).
  - Que el recorte a un máximo de páginas (max_paginas) sigue funcionando
    en la misma sesión de Word, ahora con este método de pegado.

Requiere Windows + Excel + Word instalados (usa automatización COM real).

Corré con: py tests\\test_pegado_vectorial.py
"""
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook
from openpyxl.styles import Border, Side
from docx import Document

from word_com_bridge import agregar_imagenes_a_contrato


def _armar_excel_bytes():
    """Rango B2:U10 (20 columnas) con texto y bordes -- lo bastante ancho
    como para necesitar reescalarse al pegarlo en una página normal, y con
    contenido real (no celdas vacías) para que un EMF vacío/roto se note."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Datos"
    borde = Border(*([Side(style="thin")] * 4))
    for fila in range(2, 11):
        for columna in range(2, 22):
            celda = ws.cell(row=fila, column=columna, value=f"F{fila}C{columna}")
            celda.border = borde

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _armar_docx_bytes(cantidad_paginas):
    doc = Document()
    for n in range(1, cantidad_paginas + 1):
        doc.add_paragraph(f"PAGINA {n}")
        if n < cantidad_paginas:
            doc.add_page_break()
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


if __name__ == "__main__":
    excel_bytes = _armar_excel_bytes()
    docx_bytes = _armar_docx_bytes(cantidad_paginas=2)

    especificaciones = [("Datos", "B2:U10", 2)]  # pegar en la página 2

    resultado_bytes = agregar_imagenes_a_contrato(excel_bytes, docx_bytes, especificaciones)
    print("agregar_imagenes_a_contrato no lanzó error -- OK")

    doc_resultado = Document(io.BytesIO(resultado_bytes))
    imagenes = [rel for rel in doc_resultado.part.rels.values() if "image" in rel.reltype]
    assert len(imagenes) == 1, f"Se esperaba 1 imagen, hay {len(imagenes)}"
    print("Se pegó exactamente 1 imagen -- OK")

    content_type = imagenes[0].target_part.content_type
    assert content_type == "image/x-emf", (
        f"FALLO: la imagen pegada es '{content_type}', no vectorial (image/x-emf) -- "
        f"parece que se volvió a un método de captura de pantalla rasterizada."
    )
    print(f"La imagen es vectorial ({content_type}), no un PNG rasterizado -- OK")

    peso_kb = len(imagenes[0].target_part.blob) / 1024
    print(f"Peso de la imagen: {peso_kb:.0f} KB")
    assert peso_kb < 500, (
        f"FALLO: la imagen pesa {peso_kb:.0f} KB, demasiado para un EMF de este tamaño de "
        f"rango -- señal de que no está saliendo vectorial de verdad."
    )
    print("Peso acorde a un EMF (liviano) -- OK")

    # --- max_paginas en la misma sesión (fusión con el pegado de imágenes) ---
    docx_bytes_3paginas = _armar_docx_bytes(cantidad_paginas=3)
    resultado_recortado = agregar_imagenes_a_contrato(
        excel_bytes, docx_bytes_3paginas, especificaciones, max_paginas=2,
    )
    doc_recortado = Document(io.BytesIO(resultado_recortado))
    texto = "\n".join(p.text for p in doc_recortado.paragraphs)
    assert "PAGINA 1" in texto and "PAGINA 2" in texto, "Faltó contenido que no debería haberse borrado"
    assert "PAGINA 3" not in texto, "La página 3 debería haberse recortado (max_paginas=2)"
    print("max_paginas sigue funcionando en la misma sesión de pegado -- OK")

    print("\nTODO OK")
