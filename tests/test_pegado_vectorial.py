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
  - Que un rango CHICO (más chico que la página) se agranda para ocupar
    el máximo posible -- bug real encontrado el 2026-09-25: al pasar al
    pegado vectorial, una imagen nativamente más chica que la página se
    dejaba en su tamaño chico en vez de agrandarla, lo cual además hacía
    que el CONTRATO completo quedara más corto (14 páginas en vez de 15)
    y que la última imagen fallara con "no existe la página 15".
  - Que un rango ANCHO (más ancho que la página) se achica para entrar
    (la misma lógica de LockAspectRatio que ya existía).
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
    """Dos rangos en la misma hoja: uno CHICO (B2:D4, mucho más chico que
    una página) y uno ANCHO (F2:Y10, 20 columnas -- más ancho que una
    página), ambos con bordes y contenido real (no celdas vacías) para que
    un EMF vacío/roto se note."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Datos"
    borde = Border(*([Side(style="thin")] * 4))
    for fila in range(2, 5):
        for columna in range(2, 5):  # B2:D4
            celda = ws.cell(row=fila, column=columna, value=f"F{fila}C{columna}")
            celda.border = borde
    for fila in range(2, 11):
        for columna in range(6, 26):  # F2:Y10
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

    especificaciones = [
        ("Datos", "B2:D4", 1),    # rango CHICO -- pegar en la página 1
        ("Datos", "F2:Y10", 2),   # rango ANCHO -- pegar en la página 2
    ]

    resultado_bytes = agregar_imagenes_a_contrato(excel_bytes, docx_bytes, especificaciones)
    print("agregar_imagenes_a_contrato no lanzó error -- OK")

    doc_resultado = Document(io.BytesIO(resultado_bytes))
    imagenes_rels = [rel for rel in doc_resultado.part.rels.values() if "image" in rel.reltype]
    assert len(imagenes_rels) == 2, f"Se esperaban 2 imágenes, hay {len(imagenes_rels)}"
    print("Se pegaron exactamente 2 imágenes -- OK")

    for rel in imagenes_rels:
        content_type = rel.target_part.content_type
        assert content_type == "image/x-emf", (
            f"FALLO: una imagen pegada es '{content_type}', no vectorial (image/x-emf) -- "
            f"parece que se volvió a un método de captura de pantalla rasterizada."
        )
    print("Las 2 imágenes son vectoriales (image/x-emf), no PNG rasterizados -- OK")

    # Ancho útil de página (según la sección del propio documento resultante,
    # no un número fijo -- así la prueba no depende de un tamaño de página
    # particular). Ambas imágenes (la chica agrandada y la ancha achicada)
    # tienen que terminar ocupando ese ancho máximo, con un margen chico de
    # tolerancia por redondeo de puntos a EMU.
    seccion = doc_resultado.sections[0]
    ancho_disponible_emu = seccion.page_width - seccion.left_margin - seccion.right_margin
    tolerancia_emu = 20000  # ~0.02 pulgadas
    assert len(doc_resultado.inline_shapes) == 2, (
        f"Se esperaban 2 inline shapes, hay {len(doc_resultado.inline_shapes)}"
    )
    for forma, etiqueta in zip(doc_resultado.inline_shapes, ("chica", "ancha")):
        print(f"Imagen {etiqueta}: ancho = {forma.width} EMU (disponible = {ancho_disponible_emu} EMU)")
        assert abs(forma.width - ancho_disponible_emu) <= tolerancia_emu, (
            f"FALLO: la imagen {etiqueta} no ocupa el ancho máximo de la página "
            f"(ancho={forma.width}, disponible={ancho_disponible_emu}) -- si es la imagen "
            f"'chica', es el bug real del 2026-09-25: una imagen más chica que la página se "
            f"dejaba en su tamaño nativo en vez de agrandarla al máximo."
        )
    print("Ambas imágenes (chica agrandada y ancha achicada) ocupan el ancho máximo de la página -- OK")

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
