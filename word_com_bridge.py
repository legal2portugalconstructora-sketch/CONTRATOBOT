# -*- coding: utf-8 -*-
"""
word_com_bridge.py — Pegar rangos de Excel como imagen en Word, vía
automatización COM real (pywin32).

IMPORTANTE: esto SOLO funciona en Windows, con Microsoft Excel y Word
instalados y con licencia en la MISMA máquina donde corre esta app —
porque usa automatización COM para controlar las aplicaciones reales
(no hay forma de hacer esto en un servidor Linux/Streamlit Cloud). Si
no estás en Windows, la función avisa con un error claro en vez de
fallar de forma confusa.

Los imports de pywin32 (pythoncom, win32com.client, pywintypes) quedan
DENTRO de las funciones que los usan, para que este módulo se pueda
importar sin error en cualquier sistema operativo (incluso sin pywin32
instalado) — solo falla, con un mensaje claro, si de verdad se intenta
ejecutar una de las funciones que automatiza Excel/Word.
"""

import os
import tempfile

RANGO_LISTAX_IMAGEN = "B2:I78"
HOJA_LISTAX_IMAGEN = "List.Ax"
PAGINA_CONTRATO_IMAGEN = 10

# Los rangos/páginas de "Ax.A" (antes fijos acá) ahora son por proyecto --
# ver templates_manager.CELDAS_AXA_POR_DEFECTO["rango_imagen_parte1"] etc.
# y CELDAS_AXA_POR_PROYECTO, porque el layout de esa hoja varía de un
# proyecto a otro (ej. COSTA AZUL salta su fila 89 en vez de la 98).

RANGO_AXB_PARTE1_IMAGEN = "B2:G77"
PAGINA_AXB_PARTE1_IMAGEN = 13
RANGO_AXB_PARTE2_IMAGEN = "B78:G132"
PAGINA_AXB_PARTE2_IMAGEN = 14

HOJA_AXE_IMAGEN = "Ax.E"
RANGO_AXE_IMAGEN = "B2:K73"
PAGINA_AXE_IMAGEN = 15

# Zoom de la ventana de Excel al hacer CopyPicture(Appearance=xlScreen): esa
# captura depende de cómo se ve el rango EN PANTALLA en ese momento, así que
# subir el zoom antes de copiar aumenta los píxeles reales capturados para
# el mismo rango (a más zoom, más densidad de píxeles), sin cambiar el
# tamaño físico del rango ni requerir ningún otro cambio. 300% da ~288 DPI
# (de sobra para verse nítido, incluso impreso), sin generar archivos PNG
# innecesariamente pesados como pasa con 400%+.
ZOOM_CAPTURA_IMAGEN = 300

# Constantes numéricas fijas de Excel (documentadas por Microsoft), en vez
# de win32com.client.constants.*, para no depender del orden de dispatch.
# A nivel de módulo (no solo dentro de una función) porque las necesitan
# tanto pegar_rangos_excel_como_imagen_en_word como _capturar_rango_como_png.
XL_SCREEN = 1                    # XlPictureAppearance.xlScreen
XL_BITMAP = 2                    # XlCopyPictureFormat.xlBitmap

# XlCorruptLoad.xlRepairFile -- le pide a Excel que intente reparar el
# archivo si hace falta al abrirlo, en vez de fallar directo. Hace falta
# para algunos Excel reales (ej. con hojas/vínculos externos más
# complejos) que openpyxl no reescribe 100% igual al guardarlos -- Excel
# los abre bien a mano (con su propio diálogo de reparación, que en modo
# automatizado nunca aparece), pero Workbooks.Open() sin esto los rechaza
# directo con el error 1004 en vez de repararlos solo. Confirmado que la
# reparación no pierde ni duplica ninguna hoja (se probó con un archivo
# real que fallaba sin esto). No tiene efecto en archivos que ya abren
# bien -- Excel solo repara si de verdad hace falta.
XL_REPAIR_FILE = 1               # XlCorruptLoad.xlRepairFile
XL_OPEN_XML_WORKBOOK = 51        # XlFileFormat.xlOpenXMLWorkbook (.xlsx)

# Constantes numéricas fijas de Word (documentadas por Microsoft), a nivel
# de módulo porque las necesitan tanto pegar_rangos_excel_como_imagen_en_word
# como recortar_contrato_a_maximo_paginas.
WD_GOTO_PAGE = 1                 # WdGoToItem.wdGoToPage
WD_GOTO_ABSOLUTE = 1             # WdGoToDirection.wdGoToAbsolute
WD_COLLAPSE_START = 1            # WdCollapseDirection.wdCollapseStart
WD_STATISTIC_PAGES = 2           # WdStatistic.wdStatisticPages
WD_PRINT_VIEW = 3                # WdViewType.wdPrintView

# El CONTRATO está diseñado para ocupar exactamente 15 páginas (la última
# imagen pegada, la de "Ax.E", va en la página 15 -- ver PAGINA_AXE_IMAGEN
# más abajo). Un placeholder con un nombre/texto más largo de lo normal a
# veces hace que el documento "se pase" a una página 16 extra al final
# (con muy poco o nada de contenido real) -- ver
# recortar_contrato_a_maximo_paginas, que la borra siempre.
MAX_PAGINAS_CONTRATO = 15

# Cuántas veces reintentar la captura de UN rango si sale "en blanco" (ver
# _imagen_esta_en_blanco) antes de rendirse con un error claro. La causa
# más probable de una captura en blanco es que la pantalla de esta
# computadora no estuviera realmente dibujada en ese instante (bloqueada,
# en reposo, o una sesión de Escritorio Remoto desconectada) -- ver la
# nota junto a Appearance=XL_SCREEN más abajo.
_MAX_REINTENTOS_CAPTURA_BLANCO = 3


def _en_windows():
    return os.name == "nt"


_RPC_E_CALL_REJECTED = -2147418111  # 0x8001010A: "la llamada fue rechazada por el destinatario"


def _com_retry(funcion, intentos=8, espera=0.5, descripcion=None, reintentar_cualquier_error=False):
    """Ejecuta 'funcion' (un lambda sin argumentos que hace la llamada COM)
    reintentando si Excel/Word devuelve RPC_E_CALL_REJECTED (-2147418111).

    Este error NO significa que algo esté mal configurado: significa que la
    aplicación (Excel o Word) estaba ocupada en ese instante exacto (todavía
    procesando la llamada anterior, recalculando, terminando de abrir el
    archivo, etc.) y rechazó la nueva llamada por eso. Casi siempre, una
    fracción de segundo después la aplicación ya la acepta sin problema; por
    eso la solución estándar de pywin32 para este error puntual es reintentar
    con una pequeña pausa, en vez de tratarlo como un fallo real.

    Si 'descripcion' se indica, cualquier otro error de COM (que no sea ese
    rechazo transitorio) se re-lanza con ese texto como contexto — para
    saber, ante un error genérico de Excel/Word (ej. el 1004 de Excel,
    '-2147352567 (...) -2146827284'), CUÁL llamada puntual falló, en vez de
    un mensaje sin pistas de dónde vino.

    Si 'reintentar_cualquier_error' es True, TAMBIÉN se reintenta ante
    cualquier otra excepción (no solo el rechazo transitorio). Se usa para
    llamadas puntuales conocidas por fallar de forma intermitente y sensible
    al tiempo — ej. Chart.Paste con rangos grandes, que a veces falla con el
    error 1004 de Excel si el portapapeles todavía no terminó de poblarse
    (CopyPicture es asíncrono para rangos grandes) cuando se intenta pegar
    justo después."""
    import time
    import pywintypes

    ultimo_error = None
    for _ in range(intentos):
        try:
            return funcion()
        except pywintypes.com_error as exc:
            ultimo_error = exc
            if exc.args and exc.args[0] == _RPC_E_CALL_REJECTED:
                time.sleep(espera)
                continue
            if reintentar_cualquier_error:
                time.sleep(espera)
                continue
            if descripcion:
                raise RuntimeError(f"Falló '{descripcion}': {exc}") from exc
            raise
        except Exception as exc:
            if not reintentar_cualquier_error:
                raise
            ultimo_error = exc
            time.sleep(espera)
    if descripcion:
        raise RuntimeError(f"Falló '{descripcion}' tras {intentos} reintentos: {ultimo_error}") from ultimo_error
    raise ultimo_error


def _cerrar_sin_fallar(funcion):
    """Ejecuta 'funcion' (un lambda de limpieza -- Close()/Quit() de un
    documento/aplicación COM) SOLO dentro de un 'finally', e ignora
    cualquier error que quede tras reintentarla unas pocas veces.

    Hace falta porque en Python, si el bloque 'finally' de un try/finally
    lanza una excepción, esa excepción REEMPLAZA silenciosamente lo que
    haya pasado en el 'try' -- incluso un 'return' con el resultado ya
    listo. Sin esto, un error transitorio de COM al cerrar Excel/Word
    (ej. "Excel.Application.Quit" con AttributeError, visto en la
    práctica cuando la app tarda un instante en quedar lista para recibir
    la llamada) tapaba un procesamiento que en realidad había terminado
    BIEN, y la app le mostraba al usuario un error de "no se pudo
    recalcular/procesar" sobre un archivo que sí se generó correctamente.
    Cerrar la aplicación es de todos modos best-effort: si falla incluso
    reintentando, el proceso de Excel/Word queda huérfano en segundo
    plano (recuperable a mano), pero eso es preferible a perder un
    resultado válido o un error real más útil."""
    try:
        _com_retry(funcion, intentos=3, espera=0.3, reintentar_cualquier_error=True)
    except Exception:
        pass


def _imagen_esta_en_blanco(ruta_png):
    """True si 'ruta_png' es una imagen de UN SOLO color liso de punta a
    punta -- señal de que la "foto" de la pantalla salió en blanco (ver
    la nota junto a Appearance=XL_SCREEN en
    pegar_rangos_excel_como_imagen_en_word) en vez de mostrar el
    contenido real del rango. Un rango real, aunque tenga pocos datos,
    siempre tiene algo de variación (bordes, cuadrícula, texto), así que
    exigir un solo color exacto no debería dar falsos positivos con
    capturas válidas."""
    from PIL import Image

    with Image.open(ruta_png) as imagen:
        extremos = imagen.convert("RGB").getextrema()
    # getextrema() de una imagen RGB devuelve ((min_r,max_r),(min_g,max_g),
    # (min_b,max_b)); si en las 3 bandas el mínimo es igual al máximo, la
    # imagen entera es un solo color (sin ninguna variación de píxel a píxel).
    return all(minimo == maximo for minimo, maximo in extremos)


def _capturar_rango_como_png(excel, hoja, hoja_excel, rango, rango_obj):
    """Copia 'rango_obj' como imagen (CopyPicture -> pegarlo en un Chart
    temporal -> Chart.Export a PNG -- ver la nota larga más abajo sobre
    por qué se hace así en vez de exportar directo), reintentando hasta
    _MAX_REINTENTOS_CAPTURA_BLANCO veces SI la imagen exportada sale en
    blanco (ver _imagen_esta_en_blanco). Devuelve la ruta del .png (el
    llamador es responsable de borrarlo). Lanza RuntimeError con un
    mensaje claro si sigue en blanco después de todos los reintentos."""
    for intento in range(1, _MAX_REINTENTOS_CAPTURA_BLANCO + 1):
        _com_retry(
            lambda: rango_obj.CopyPicture(Appearance=XL_SCREEN, Format=XL_BITMAP),
            descripcion=f"Excel CopyPicture {hoja_excel}!{rango} (intento {intento})",
        )

        chart_obj = hoja.ChartObjects().Add(
            rango_obj.Left, rango_obj.Top + rango_obj.Height + 20,
            rango_obj.Width, rango_obj.Height,
        )
        chart = chart_obj.Chart
        ruta_png = None
        try:
            _com_retry(
                lambda: chart_obj.Activate(),
                descripcion=f"ChartObject.Activate {hoja_excel}!{rango} (intento {intento})",
            )
            _com_retry(
                lambda: chart.Paste(),
                descripcion=f"Chart.Paste {hoja_excel}!{rango} (intento {intento})",
                reintentar_cualquier_error=True,
            )
            excel.CutCopyMode = False

            tmp_png = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
            tmp_png.close()
            ruta_png = os.path.abspath(tmp_png.name)
            _com_retry(
                lambda: chart.Export(Filename=ruta_png, FilterName="PNG"),
                descripcion=f"Chart.Export {hoja_excel}!{rango} (intento {intento})",
            )

            if not os.path.exists(ruta_png) or os.path.getsize(ruta_png) == 0:
                raise RuntimeError(
                    f"No se pudo generar la imagen del rango '{hoja_excel}!{rango}': el "
                    f"archivo PNG exportado desde Excel no existe o está vacío."
                )

            if not _imagen_esta_en_blanco(ruta_png):
                exito, ruta_png = ruta_png, None  # que el finally no la borre: se devuelve
                return exito
        finally:
            _cerrar_sin_fallar(lambda: chart_obj.Delete())
            if ruta_png is not None:
                try:
                    os.remove(ruta_png)
                except OSError:
                    pass

    raise RuntimeError(
        f"La imagen de '{hoja_excel}!{rango}' salió en blanco {_MAX_REINTENTOS_CAPTURA_BLANCO} "
        f"veces seguidas. Es casi seguro que la pantalla de esta computadora no estuvo "
        f"realmente visible mientras se procesaba (bloqueada, en reposo, o una sesión de "
        f"Escritorio Remoto desconectada) -- evitá eso y volvé a procesar el contrato."
    )


def pegar_rangos_excel_como_imagen_en_word(ruta_excel, especificaciones, ruta_word):
    """Abre Excel y Word de verdad UNA sola vez, y pega en el Word varios
    rangos de Excel como imagen, uno por cada elemento de
    'especificaciones' (lista de tuplas (hoja_excel, rango, numero_pagina)),
    procesados EN ORDEN. Guarda el documento (sobrescribiendo 'ruta_word')
    solo al final, después de insertar TODAS las imágenes con éxito — si
    cualquier rango falla, no se guarda nada, para no dejar el documento
    con solo algunas imágenes pegadas a medias. Requiere Windows + Excel +
    Word instalados; usa pywin32.

    Detalles importantes para que el pegado sea confiable de verdad
    (no solo "en teoría"):
      - Mantiene Excel y Word VISIBLES mientras procesa. Con las apps
        invisibles, Word no calcula bien la paginación real y el copiado
        al portapapeles desde Excel es menos confiable — esto es lo más
        probable que causó que "no apareciera nada".
      - Antes de navegar a CADA página, recalcula cuántas páginas tiene
        REALMENTE el documento en ese momento (con ComputeStatistics):
        como las imágenes anteriores de esta misma sesión pueden haber
        hecho crecer el documento, el total de páginas puede cambiar
        entre un rango y el siguiente. Si no llega a la página pedida,
        Word NO da error — salta silenciosamente a la última página que
        exista. Por eso lo validamos nosotros y avisamos con un error
        claro en vez de pegar la imagen en un lugar equivocado sin que
        te enteres.
      - Usa números de constante fijos (no win32com.client.constants),
        que a veces no se resuelven bien según el orden en que se hace
        EnsureDispatch de cada aplicación.
      - Verifica, contando las imágenes del documento antes y después de
        CADA rango, que el pegado realmente insertó algo; si no, lanza un
        error en vez de guardar el documento sin cambios silenciosamente.
    """
    if not _en_windows():
        raise RuntimeError(
            "Pegar imágenes de Excel en Word requiere Windows con Microsoft Excel y "
            "Word instalados (usa automatización COM vía pywin32); esta máquina/servidor "
            "no es Windows, así que este paso no se puede ejecutar aquí."
        )

    try:
        import ctypes
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "No se encontró pywin32 instalado. Instálalo con: pip install pywin32"
        ) from exc

    # Streamlit ejecuta el script en un hilo secundario (no el hilo
    # principal del proceso), y COM exige que cada hilo que va a usar
    # objetos COM llame a CoInitialize() antes de crearlos. Sin esto,
    # win32com.client.Dispatch (o las llamadas posteriores a Excel/Word)
    # fallan con errores como "CoInitialize has not been called" o
    # pywintypes.com_error, aunque pywin32 esté bien instalado.
    pythoncom.CoInitialize()

    # El pegado de imágenes depende de que la pantalla de ESTA computadora
    # esté realmente dibujada en pantalla en el momento de cada captura
    # (ver la nota junto a Appearance=XL_SCREEN, más abajo) -- si Windows
    # apaga la pantalla o suspende el equipo por inactividad a mitad del
    # proceso, las capturas salen en blanco. SetThreadExecutionState le
    # pide a Windows que NO haga eso mientras dure este hilo. ES_CONTINUOUS
    # hace que el pedido quede activo hasta el próximo llamado (el de
    # abajo, en el finally, que lo saca) en vez de expirar solo.
    # OJO: esto NO evita que alguien bloquee la sesión a mano (Win+L) ni
    # que se desconecte una sesión de Escritorio Remoto -- Windows no deja
    # prevenir eso desde un proceso en segundo plano, por seguridad.
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )

    ruta_excel = os.path.abspath(ruta_excel)
    ruta_word = os.path.abspath(ruta_word)

    excel = None
    word = None
    wb = None
    doc = None
    try:
        # Se usa Dispatch (enlace tardío) en vez de gencache.EnsureDispatch
        # a propósito: EnsureDispatch necesita generar y cachear un wrapper
        # "makepy" para Excel/Word en disco (gen_py), y si ese caché está
        # corrupto, desactualizado (ej. tras una actualización de Office) o
        # sin permisos de escritura, falla con "This COM object can not
        # automate the makepy process". Dispatch no depende de ese caché en
        # absoluto; como aquí ya usamos constantes numéricas fijas (ver
        # arriba) en vez de win32com.client.constants.*, no perdemos nada
        # usando enlace tardío.
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        # UpdateLinks=0: el MODELO tiene un vínculo externo roto (referencia a
        # un Excel de la PC original donde se armó la plantilla, que ya no
        # existe en esta máquina). DisplayAlerts=False NO siempre alcanza para
        # suprimir el diálogo "¿Actualizar vínculos?" -- si ese diálogo queda
        # esperando una respuesta que nunca llega (nadie puede hacer clic en
        # automatización headless), Open() se queda colgado indefinidamente.
        # UpdateLinks=0 le dice a Excel que ni siquiera pregunte.
        wb = _com_retry(lambda: excel.Workbooks.Open(ruta_excel, UpdateLinks=0, CorruptLoad=XL_REPAIR_FILE))

        word = win32.Dispatch("Word.Application")
        word.Visible = True
        word.DisplayAlerts = 0  # wdAlertsNone: evita diálogos de Word que bloquearían las llamadas COM
        doc = _com_retry(lambda: word.Documents.Open(ruta_word))
        # Forzar el cálculo de la paginación REAL antes de navegar a una página.
        _com_retry(lambda: setattr(doc.ActiveWindow.View, "Type", WD_PRINT_VIEW))

        for hoja_excel, rango, numero_pagina in especificaciones:
            try:
                hoja = wb.Sheets(hoja_excel)
            except Exception as exc:
                raise ValueError(f"La hoja '{hoja_excel}' no existe en el Excel.") from exc

            try:
                rango_obj = hoja.Range(rango)
            except Exception as exc:
                raise ValueError(f"El rango '{rango}' no es válido en la hoja '{hoja_excel}'.") from exc

            # Range.CopyPicture necesita que su hoja esté ACTIVA en Excel en
            # ese momento; si no lo está (ej. porque el rango anterior era de
            # otra hoja, o la hoja activa al abrir el archivo era otra),
            # falla con el error 1004 de Excel ("Application-defined or
            # object-defined error", que pywin32 reporta como
            # "-2147352567 (...) -2146827284"). Application.Goto con
            # Scroll=True activa la hoja Y desplaza la vista para que el
            # rango quede visible, que es la forma estándar de evitar eso.
            excel.Goto(rango_obj, True)

            # Subir el zoom ANTES de CopyPicture (ver nota junto a
            # ZOOM_CAPTURA_IMAGEN, arriba): la apariencia de pantalla se
            # captura con más densidad de píxeles a mayor zoom, para el
            # mismo rango y el mismo tamaño físico final. Hace falta un
            # SEGUNDO Goto después de cambiar el zoom -- sin esto, la vista
            # no termina de "ponerse al día" con el nuevo zoom antes de
            # CopyPicture, y la captura sale con menos densidad de píxeles
            # de la esperada (confirmado empíricamente).
            _com_retry(
                lambda: setattr(excel.ActiveWindow, "Zoom", ZOOM_CAPTURA_IMAGEN),
                descripcion=f"Excel Zoom (para {hoja_excel}!{rango})",
            )
            _com_retry(
                lambda: excel.Goto(rango_obj, True),
                descripcion=f"Excel Goto tras cambiar zoom ({hoja_excel}!{rango})",
            )

            # Se usa Range.CopyPicture (no Range.Copy()) a propósito: el
            # contenido copiado se va a pegar con Chart.Paste() dentro de un
            # gráfico temporal, y Chart.Paste() SOLO pega una imagen si el
            # portapapeles contiene una imagen. Si se copiara el rango con
            # Copy() normal (datos de celda), Chart.Paste() no lo pega como
            # imagen: lo interpreta como una serie de datos nueva para
            # graficar, y el resultado es un gráfico con barras/líneas
            # generado a partir de los valores de las celdas — la "imagen
            # rara" que salía antes de este cambio, no una foto del rango.
            #
            # El recorte del borde inferior/derecho que normalmente tiene
            # CopyPicture (problema conocido de ese método al pegarlo o
            # exportarlo directamente) NO ocurre acá, porque la imagen no se
            # exporta/pega directamente: se pega dentro de un gráfico (Chart)
            # temporal y vacío creado en la misma hoja (con el mismo ancho y
            # alto que el rango), y es Chart.Export (el motor de renderizado
            # de gráficos, sin ese bug) el que genera el .png final en disco.
            # Ese gráfico se borra después de exportar, para no dejar basura
            # en el archivo de Excel original.
            #
            # Se usa Format=xlBitmap (raster) en vez de xlPicture (vectorial,
            # WMF/EMF): para rangos grandes/complejos, xlPicture puede hacer
            # que Chart.Paste falle con el error 1004 de Excel de forma
            # consistente (no es un problema de timing: reintentar no lo
            # arregla). Como de todos modos el resultado final es un .png
            # (raster) vía Chart.Export, copiar directamente en bitmap no
            # pierde nada y es más robusto para rangos grandes.
            #
            # Toda esta lógica (más el reintento si la captura sale en
            # blanco -- ver _imagen_esta_en_blanco) vive en
            # _capturar_rango_como_png para poder reintentarla completa
            # desde cero (CopyPicture incluido) sin repetir el código acá.
            ruta_png = _capturar_rango_como_png(excel, hoja, hoja_excel, rango, rango_obj)
            try:
                # Paginación real recalculada para ESTE rango (ver nota arriba
                # sobre por qué no alcanza con calcularla una sola vez al inicio).
                total_paginas = _com_retry(
                    lambda: doc.ComputeStatistics(WD_STATISTIC_PAGES),
                    descripcion=f"Word ComputeStatistics (antes de {hoja_excel}!{rango})",
                )
                if numero_pagina > total_paginas:
                    raise ValueError(
                        f"El documento solo tiene {total_paginas} página(s) en este momento; "
                        f"no existe la página {numero_pagina} para pegar '{hoja_excel}!{rango}' ahí. "
                        f"(Si el CONTRATO quedó más corto/largo tras cambios anteriores, puede que "
                        f"ya no llegue a esa página; dime a qué página debería ir en ese caso.)"
                    )

                figuras_antes = doc.InlineShapes.Count + doc.Shapes.Count

                seleccion = word.Selection
                # Ir al inicio de la página indicada, usando la paginación real
                # de Word (igual que Ctrl+G -> Página -> 10 -> Ir a).
                _com_retry(
                    lambda: seleccion.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=numero_pagina),
                    descripcion=f"Word GoTo página {numero_pagina} (para {hoja_excel}!{rango})",
                )
                _com_retry(
                    lambda: seleccion.Collapse(Direction=WD_COLLAPSE_START),
                    descripcion=f"Word Collapse (para {hoja_excel}!{rango})",
                )
                # Insertar el .png exportado como imagen en la posición actual del cursor.
                imagen_insertada = _com_retry(
                    lambda: seleccion.InlineShapes.AddPicture(FileName=ruta_png),
                    descripcion=f"Word AddPicture {hoja_excel}!{rango} → página {numero_pagina}",
                )

                # AddPicture inserta la imagen a su tamaño NATIVO (según el DPI
                # del .png exportado), sin importar el ancho de la página. Para
                # rangos anchos (ej. columnas B:N de "Ax.A", ~11 pulgadas) esto
                # la deja más ancha que el área útil de la página (márgen a
                # márgen) -- se sale del márgen derecho y queda "cortada"
                # visualmente en vez de reducida para que quepa. Se reescala
                # manteniendo la proporción (LockAspectRatio) SOLO si hace
                # falta, para no agrandar imágenes que ya entran bien.
                MSO_TRUE = -1  # msoTrue
                ancho_disponible = _com_retry(
                    lambda: doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin,
                    descripcion=f"Word PageSetup (para {hoja_excel}!{rango})",
                )
                if imagen_insertada.Width > ancho_disponible:
                    _com_retry(
                        lambda: setattr(imagen_insertada, "LockAspectRatio", MSO_TRUE),
                        descripcion=f"Word LockAspectRatio (para {hoja_excel}!{rango})",
                    )
                    _com_retry(
                        lambda: setattr(imagen_insertada, "Width", ancho_disponible),
                        descripcion=f"Word ajustar ancho de imagen (para {hoja_excel}!{rango})",
                    )

                figuras_despues = doc.InlineShapes.Count + doc.Shapes.Count
                if figuras_despues <= figuras_antes:
                    raise RuntimeError(
                        f"El pegado de '{hoja_excel}!{rango}' no insertó ninguna imagen nueva en el "
                        f"documento. Puede que otra aplicación haya usado el portapapeles justo en ese "
                        f"momento; intenta de nuevo sin tocar el mouse/teclado mientras se procesa."
                    )
            finally:
                # El chart temporal ya se borró dentro de
                # _capturar_rango_como_png; acá solo queda limpiar el .png
                # una vez que Word ya lo usó (AddPicture lo lee del disco).
                if ruta_png is not None:
                    try:
                        os.remove(ruta_png)
                    except OSError:
                        pass

        # Devolver el zoom a 100% -- prolijidad nomás, ya que el workbook se
        # cierra sin guardar (SaveChanges=False) así que esto no queda
        # grabado en el archivo de todos modos.
        _com_retry(lambda: setattr(excel.ActiveWindow, "Zoom", 100), descripcion="Excel Zoom (restaurar a 100%)")

        _com_retry(lambda: doc.Save(), descripcion="Word doc.Save()")
    finally:
        if doc is not None:
            _cerrar_sin_fallar(lambda: doc.Close(SaveChanges=False))
        if word is not None:
            _cerrar_sin_fallar(lambda: word.Quit())
        if wb is not None:
            _cerrar_sin_fallar(lambda: wb.Close(SaveChanges=False))
        if excel is not None:
            _cerrar_sin_fallar(lambda: excel.Quit())
        # Deshacer el pedido de "no apagar la pantalla/no suspender" de
        # más arriba -- sin el flag ES_CONTINUOUS solo (sin
        # ES_SYSTEM_REQUIRED/ES_DISPLAY_REQUIRED), Windows vuelve a su
        # comportamiento normal de ahorro de energía para el resto del
        # sistema, no solo para este hilo.
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        pythoncom.CoUninitialize()


def pegar_rango_excel_como_imagen_en_word(ruta_excel, hoja_excel, rango, ruta_word, numero_pagina):
    """Envoltorio para pegar UN solo rango (se mantiene por compatibilidad);
    ver pegar_rangos_excel_como_imagen_en_word para el caso general de
    varios rangos en una sola sesión de Excel/Word."""
    pegar_rangos_excel_como_imagen_en_word(ruta_excel, [(hoja_excel, rango, numero_pagina)], ruta_word)


def agregar_imagenes_a_contrato(excel_bytes, docx_bytes, especificaciones):
    """Envoltorio 'en memoria': recibe los bytes del Excel y del Word ya
    procesado, los escribe a archivos temporales (pywin32 necesita
    archivos reales en disco, no streams en memoria), llama a
    pegar_rangos_excel_como_imagen_en_word con TODAS las 'especificaciones'
    (lista de (hoja, rango, pagina)) en una sola sesión de Excel/Word, y
    devuelve los bytes del Word final. Limpia los archivos temporales al
    terminar."""
    tmp_excel = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp_word = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    try:
        tmp_excel.write(excel_bytes)
        tmp_excel.close()
        tmp_word.write(docx_bytes)
        tmp_word.close()

        pegar_rangos_excel_como_imagen_en_word(
            os.path.abspath(tmp_excel.name), especificaciones, os.path.abspath(tmp_word.name)
        )

        with open(tmp_word.name, "rb") as f:
            return f.read()
    finally:
        for ruta in (tmp_excel.name, tmp_word.name):
            try:
                os.remove(ruta)
            except OSError:
                pass


def recortar_contrato_a_maximo_paginas(docx_bytes, max_paginas=MAX_PAGINAS_CONTRATO):
    """Abre 'docx_bytes' en Word de verdad y, si el documento tiene MÁS de
    'max_paginas' páginas, borra todo el contenido desde el inicio de la
    página max_paginas+1 hasta el final del documento, y lo guarda.
    Devuelve los bytes resultantes (recortados, o los mismos de entrada
    sin cambios si el documento ya tenía max_paginas o menos -- igual hay
    que abrir Word para saber cuántas páginas tiene de verdad).

    Hace falta abrir Word de verdad porque python-docx no calcula
    paginación (depende del layout real al renderizar el documento);
    ComputeStatistics(wdStatisticPages) es la única forma confiable de
    saber cuántas páginas tiene REALMENTE. Requiere Windows + Word
    instalados; usa pywin32."""
    if not _en_windows():
        raise RuntimeError(
            "Recortar el CONTRATO a un máximo de páginas requiere Windows con "
            "Microsoft Word instalado (usa automatización COM vía pywin32); esta "
            "máquina/servidor no es Windows, así que este paso no se puede ejecutar aquí."
        )

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "No se encontró pywin32 instalado. Instálalo con: pip install pywin32"
        ) from exc

    pythoncom.CoInitialize()

    tmp_word = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp_word.write(docx_bytes)
    tmp_word.close()
    ruta_word = os.path.abspath(tmp_word.name)

    word = None
    doc = None
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = True
        word.DisplayAlerts = 0  # wdAlertsNone: evita diálogos de Word que bloquearían las llamadas COM
        doc = _com_retry(lambda: word.Documents.Open(ruta_word))
        # Forzar el cálculo de la paginación REAL (ver la nota junto a esto
        # mismo en pegar_rangos_excel_como_imagen_en_word: con la app
        # invisible, Word no calcula bien la paginación real).
        _com_retry(lambda: setattr(doc.ActiveWindow.View, "Type", WD_PRINT_VIEW))
        total_paginas = _com_retry(
            lambda: doc.ComputeStatistics(WD_STATISTIC_PAGES),
            descripcion="Word ComputeStatistics (para recortar páginas sobrantes)",
        )

        if total_paginas > max_paginas:
            seleccion = word.Selection
            _com_retry(
                lambda: seleccion.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=max_paginas + 1),
                descripcion=f"Word GoTo página {max_paginas + 1} (recorte de páginas sobrantes)",
            )
            _com_retry(
                lambda: seleccion.Collapse(Direction=WD_COLLAPSE_START),
                descripcion="Word Collapse (recorte de páginas sobrantes)",
            )
            rango_sobrante = doc.Range(Start=seleccion.Start, End=doc.Content.End)
            _com_retry(lambda: rango_sobrante.Delete(), descripcion="Word borrar páginas sobrantes")
            _com_retry(lambda: doc.Save(), descripcion="Word doc.Save() (tras recortar páginas sobrantes)")

        with open(ruta_word, "rb") as f:
            return f.read()
    finally:
        if doc is not None:
            _cerrar_sin_fallar(lambda: doc.Close(SaveChanges=False))
        if word is not None:
            _cerrar_sin_fallar(lambda: word.Quit())
        pythoncom.CoUninitialize()
        try:
            os.remove(ruta_word)
        except OSError:
            pass


def agregar_imagen_listax_a_contrato(excel_bytes, docx_bytes,
                                      hoja=HOJA_LISTAX_IMAGEN, rango=RANGO_LISTAX_IMAGEN,
                                      pagina=PAGINA_CONTRATO_IMAGEN):
    """Envoltorio específico para pegar List.Ax como imagen (se mantiene
    por compatibilidad); ver agregar_imagenes_a_contrato para el caso
    general de varios rangos."""
    return agregar_imagenes_a_contrato(excel_bytes, docx_bytes, [(hoja, rango, pagina)])


def recalcular_excel_con_com(excel_bytes):
    """Abre 'excel_bytes' en Excel de verdad, fuerza un recálculo COMPLETO
    de todas las fórmulas del libro (CalculateFullRebuild — recalcula todo
    desde cero, sin importar si el archivo tenía el modo de cálculo en
    Manual o Automático) y lo guarda. Devuelve los bytes ya recalculados.

    Esto es necesario porque openpyxl NO tiene motor de fórmulas: cuando
    lee un archivo (aunque sea con data_only=True) solo puede devolver el
    valor que quedó cacheado la última vez que el archivo se guardó desde
    Excel real. Si el Excel se generó/editó únicamente con openpyxl (como
    hace el Paso 1-6), fórmulas complejas (BUSCARV, SI, CONCATENAR, etc.)
    quedan con un valor cacheado viejo o vacío — openpyxl jamás las
    recalcula por su cuenta. Por eso, antes de leer los placeholders del
    CONTRATO/REGLAMENTO, se recalcula el archivo con Excel real.

    Requiere Windows + Excel instalados; usa pywin32."""
    if not _en_windows():
        raise RuntimeError(
            "Recalcular las fórmulas del Excel requiere Windows con Microsoft Excel "
            "instalado (usa automatización COM vía pywin32); esta máquina/servidor "
            "no es Windows, así que este paso no se puede ejecutar aquí."
        )

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "No se encontró pywin32 instalado. Instálalo con: pip install pywin32"
        ) from exc

    pythoncom.CoInitialize()

    tmp_excel = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    excel = None
    wb = None
    try:
        tmp_excel.write(excel_bytes)
        tmp_excel.close()
        ruta_excel = os.path.abspath(tmp_excel.name)

        # Sin Visible=True a propósito: a diferencia de pegar_rangos_excel_como_imagen_en_word,
        # aquí no hay portapapeles ni paginación de Word de por medio — es
        # solo abrir, recalcular y guardar, así que puede quedar invisible
        # (más rápido y sin abrir una ventana de Excel en pantalla).
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        # UpdateLinks=0: ver la nota junto a esta misma llamada en
        # pegar_rangos_excel_como_imagen_en_word -- evita que Open() se quede
        # esperando indefinidamente el diálogo "¿Actualizar vínculos?" por el
        # vínculo externo roto que trae el MODELO.
        wb = _com_retry(
            lambda: excel.Workbooks.Open(ruta_excel, UpdateLinks=0, CorruptLoad=XL_REPAIR_FILE),
            descripcion="Excel Workbooks.Open (recálculo)",
        )
        _com_retry(lambda: excel.CalculateFullRebuild(), descripcion="Excel CalculateFullRebuild")
        # SaveAs (al mismo 'ruta_excel', formato .xlsx explícito) en vez de
        # Save() a secas: si el archivo se abrió en modo reparado (ver
        # CorruptLoad=XL_REPAIR_FILE más arriba), Excel a veces rechaza un
        # Save() normal con el mismo error 1004 -- ya no "confía" en
        # sobrescribir en el lugar un archivo que tuvo que reparar. SaveAs
        # con el formato explícito funciona en ambos casos (reparado o no).
        _com_retry(
            lambda: wb.SaveAs(ruta_excel, FileFormat=XL_OPEN_XML_WORKBOOK),
            descripcion="Excel Workbook.SaveAs (recálculo)",
        )

        with open(ruta_excel, "rb") as f:
            return f.read()
    finally:
        if wb is not None:
            _cerrar_sin_fallar(lambda: wb.Close(SaveChanges=False))
        if excel is not None:
            _cerrar_sin_fallar(lambda: excel.Quit())
        pythoncom.CoUninitialize()
        try:
            os.remove(tmp_excel.name)
        except OSError:
            pass


WD_FORMAT_PDF = 17  # wdFormatPDF


def convertir_docx_a_pdf_con_com(docx_bytes):
    """Abre 'docx_bytes' en Word de verdad y lo exporta a PDF tal como se ve
    en Word (conserva imágenes, formato, paginación real, etc. -- a
    diferencia de un conversor genérico que solo lee el XML). Devuelve los
    bytes del PDF.

    Requiere Windows + Word instalado; usa pywin32."""
    if not _en_windows():
        raise RuntimeError(
            "Convertir un Word a PDF requiere Windows con Microsoft Word instalado "
            "(usa automatización COM vía pywin32); esta máquina/servidor no es "
            "Windows, así que este paso no se puede ejecutar aquí."
        )

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "No se encontró pywin32 instalado. Instálalo con: pip install pywin32"
        ) from exc

    pythoncom.CoInitialize()

    tmp_docx = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp_pdf.close()
    word = None
    doc = None
    try:
        tmp_docx.write(docx_bytes)
        tmp_docx.close()
        ruta_docx = os.path.abspath(tmp_docx.name)
        ruta_pdf = os.path.abspath(tmp_pdf.name)

        # Sin Visible=True a propósito (igual que recalcular_excel_con_com):
        # solo se abre, se exporta a PDF y se cierra, sin que nadie necesite
        # ver la ventana de Word en pantalla.
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone
        doc = _com_retry(lambda: word.Documents.Open(ruta_docx), descripcion="Word Documents.Open (a PDF)")
        _com_retry(
            lambda: doc.SaveAs(ruta_pdf, FileFormat=WD_FORMAT_PDF),
            descripcion="Word doc.SaveAs (a PDF)",
        )

        with open(ruta_pdf, "rb") as f:
            return f.read()
    finally:
        if doc is not None:
            _cerrar_sin_fallar(lambda: doc.Close(SaveChanges=False))
        if word is not None:
            _cerrar_sin_fallar(lambda: word.Quit())
        pythoncom.CoUninitialize()
        for ruta in (tmp_docx.name, tmp_pdf.name):
            try:
                os.remove(ruta)
            except OSError:
                pass
