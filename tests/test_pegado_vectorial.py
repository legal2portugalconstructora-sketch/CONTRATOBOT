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

Desde el 2026-09-27 existe además un método de RESPALDO (captura de
pantalla rasterizada a PNG, igual al método viejo) para máquinas donde
Selection.PasteSpecial() de Word nunca inserta nada aunque el
portapapeles tenga el formato EMF disponible (confirmado en la práctica:
Chart.Paste() dentro de Excel sí funciona siempre en esa máquina, pero
cruzar a Word por el portapapeles nunca funciona ahí) -- ver la nota junto
a XL_SCREEN en word_com_bridge.py. Esta prueba por eso acepta CUALQUIERA
de los dos formatos (vectorial o PNG de respaldo): lo que importa es que
el pegado nunca falle del todo, no cuál mecanismo puntual usó esta
máquina.

Esta prueba verifica, con Excel/Word reales:
  - Que el pegado no lanza error y deja exactamente una imagen nueva.
  - Que la imagen resultante es vectorial (content_type "image/x-emf") O,
    si esta máquina cayó al método de respaldo, un PNG -- cualquier otro
    tipo sería una señal real de que algo se rompió.
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

    content_types = []
    for rel in imagenes_rels:
        content_type = rel.target_part.content_type
        assert content_type in ("image/x-emf", "image/png"), (
            f"FALLO: una imagen pegada es '{content_type}', ni vectorial (image/x-emf) ni el "
            f"PNG de respaldo -- esto sí es un tipo de archivo inesperado."
        )
        content_types.append(content_type)
    if all(ct == "image/x-emf" for ct in content_types):
        print("Las 2 imágenes son vectoriales (image/x-emf) -- OK (método normal)")
    else:
        print(
            f"Tipos de imagen: {content_types} -- esta máquina usó el método de respaldo (PNG) "
            f"para al menos una imagen. No es un fallo: solo pasa cuando Word no acepta "
            f"PasteSpecial() en esta máquina en particular."
        )

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
    for forma, etiqueta, content_type in zip(doc_resultado.inline_shapes, ("chica", "ancha"), content_types):
        print(f"Imagen {etiqueta}: ancho = {forma.width} EMU (disponible = {ancho_disponible_emu} EMU)")
        if content_type == "image/x-emf":
            # El método vectorial SIEMPRE reescala al máximo (agranda o
            # achica) -- ver el bug real del 2026-09-25 más arriba.
            assert abs(forma.width - ancho_disponible_emu) <= tolerancia_emu, (
                f"FALLO: la imagen {etiqueta} (vectorial) no ocupa el ancho máximo de la página "
                f"(ancho={forma.width}, disponible={ancho_disponible_emu})."
            )
        else:
            # El método de respaldo (PNG) solo achica si hace falta, igual
            # que el método viejo -- no fuerza a agrandar una imagen chica,
            # así que acá solo importa que NUNCA se pase del ancho de la
            # página, no que la ocupe entera.
            assert forma.width <= ancho_disponible_emu + tolerancia_emu, (
                f"FALLO: la imagen {etiqueta} (respaldo/PNG) se pasa del ancho disponible "
                f"(ancho={forma.width}, disponible={ancho_disponible_emu})."
            )
    print("Los anchos de las imágenes son correctos para el método que usó cada una -- OK")

    # --- max_paginas en la misma sesión (fusión con el pegado de imágenes) ---
    # OJO: las imágenes de esta prueba están armadas a propósito para
    # ocupar el máximo posible de la página (para poder probar el
    # reescalado) -- eso hace que el documento real termine con MÁS
    # páginas de las que tenía el .docx original (el CONTRATO real no
    # tiene este problema: cada imagen va a una página ya diseñada para
    # esa imagen puntual, con espacio de sobra). Por eso acá se verifica
    # la paginación REAL con Word (no se asume en qué página de texto
    # quedó cada cosa), igual que en tests/test_recorte_paginas.py.
    docx_bytes_3paginas = _armar_docx_bytes(cantidad_paginas=3)
    max_paginas_prueba = 3
    resultado_recortado = agregar_imagenes_a_contrato(
        excel_bytes, docx_bytes_3paginas, especificaciones, max_paginas=max_paginas_prueba,
    )

    import tempfile
    import pythoncom
    import win32com.client as win32
    from word_com_bridge import _com_retry, _cerrar_sin_fallar, _cerrar_automatizaciones_huerfanas

    tmp_recortado = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp_recortado.write(resultado_recortado)
    tmp_recortado.close()
    pythoncom.CoInitialize()
    _cerrar_automatizaciones_huerfanas()
    word_verif = win32.Dispatch("Word.Application")
    # Reintentar (no una asignación directa): recién cerrada la sesión de
    # Word anterior de esta misma prueba, una nueva puede tardar un
    # instante en aceptar llamadas COM -- ver la nota junto a esto mismo
    # en word_com_bridge.pegar_rangos_excel_como_imagen_en_word.
    _com_retry(lambda: setattr(word_verif, "Visible", False), reintentar_cualquier_error=True)
    _com_retry(lambda: setattr(word_verif, "DisplayAlerts", 0), reintentar_cualquier_error=True)
    doc_verif = word_verif.Documents.Open(os.path.abspath(tmp_recortado.name))
    doc_verif.ActiveWindow.View.Type = 3
    paginas_reales = doc_verif.ComputeStatistics(2)
    _cerrar_sin_fallar(lambda: doc_verif.Close(SaveChanges=False))
    _cerrar_sin_fallar(lambda: word_verif.Quit())
    pythoncom.CoUninitialize()
    os.remove(tmp_recortado.name)

    print(f"Páginas reales tras recortar a max_paginas={max_paginas_prueba}: {paginas_reales}")
    assert paginas_reales == max_paginas_prueba, (
        f"FALLO: se pidió recortar a {max_paginas_prueba} páginas, quedaron {paginas_reales}."
    )
    print("max_paginas sigue funcionando en la misma sesión de pegado -- OK")

    print("\nTODO OK")
