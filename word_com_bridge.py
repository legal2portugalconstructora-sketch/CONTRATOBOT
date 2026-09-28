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
WD_PASTE_ENHANCED_METAFILE = 9   # WdPasteDataType.wdPasteEnhancedMetafile

# El CONTRATO está diseñado para ocupar exactamente 15 páginas (la última
# imagen pegada, la de "Ax.E", va en la página 15 -- ver PAGINA_AXE_IMAGEN
# más abajo). Un placeholder con un nombre/texto más largo de lo normal a
# veces hace que el documento "se pase" a una página 16 extra al final
# (con muy poco o nada de contenido real) -- ver
# recortar_contrato_a_maximo_paginas, que la borra siempre.
MAX_PAGINAS_CONTRATO = 15

def _en_windows():
    return os.name == "nt"


def _cerrar_automatizaciones_huerfanas():
    """Cierra EXCEL.EXE/WINWORD.EXE 'huérfanos' de un intento anterior de
    automatización que quedó a medio camino -- ANTES de abrir uno nuevo.

    Hace falta porque Excel/Word se registran como servidores COM
    'multi-use': si ya hay un proceso corriendo (aunque sea un huérfano
    trabado, ej. mostrando un diálogo bloqueado que nadie puede cerrar
    porque la automatización es headless), Dispatch("Excel.Application")
    se ENGANCHA a ESE MISMO proceso en vez de abrir uno nuevo -- heredando
    lo que sea que lo tenga trabado. Se vio en la práctica: un huérfano
    trabado hace que hasta 'excel.Visible = False' falle con un
    AttributeError ("Property 'Excel.Application.Visible' can not be
    set"), y reintentar (_com_retry) no ayuda porque sigue siendo el mismo
    objeto roto -- hay que sacar el huérfano de en medio ANTES, no
    reintentar después.

    Identifica huérfanos por la bandera '/automation' en la línea de
    comando -- Windows la agrega SOLO cuando el proceso se abrió vía COM,
    nunca cuando alguien lo abre a mano (confirmado en la práctica varias
    veces: un Excel/Word abierto por el usuario nunca la tiene) -- así que
    nunca toca una ventana real que el usuario tenga abierta. Best-effort:
    si algo de esto falla (ej. PowerShell no disponible), simplemente no
    hace nada y se sigue con el Dispatch normal."""
    import json
    import subprocess

    for nombre_proceso in ("EXCEL.EXE", "WINWORD.EXE"):
        try:
            resultado = subprocess.run(
                [
                    "powershell", "-NoProfile", "-NonInteractive", "-Command",
                    f"Get-CimInstance Win32_Process -Filter \"Name='{nombre_proceso}'\" "
                    "| Select-Object ProcessId,CommandLine | ConvertTo-Json -Compress",
                ],
                capture_output=True, text=True, timeout=15,
            )
            salida = resultado.stdout.strip()
            if not salida:
                continue
            procesos = json.loads(salida)
            if isinstance(procesos, dict):
                procesos = [procesos]
            for proc in procesos:
                command_line = (proc.get("CommandLine") or "").lower()
                pid = proc.get("ProcessId")
                if pid is None or "/automation" not in command_line:
                    continue
                try:
                    subprocess.run(
                        ["taskkill", "/PID", str(pid), "/F"],
                        capture_output=True, timeout=10,
                    )
                except Exception:
                    pass
        except Exception:
            pass


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


def _recortar_documento_word_abierto(word, doc, max_paginas):
    """Con 'doc' ya abierto en una sesión de Word real (y su App 'word'),
    borra el contenido desde el inicio de la página max_paginas+1 hasta el
    final SI el documento tiene más de 'max_paginas' páginas en este
    momento. NO guarda el documento -- eso queda a cargo del llamador,
    para poder guardar junto con otros cambios hechos en la misma sesión
    (ej. imágenes recién pegadas) en un solo doc.Save(), en vez de abrir
    Word de nuevo aparte solo para esto. Devuelve True si recortó algo.

    El llamador es responsable de haber forzado ya la vista de impresión
    (doc.ActiveWindow.View.Type = WD_PRINT_VIEW) antes de llamar a esto,
    para que ComputeStatistics calcule la paginación real."""
    total_paginas = _com_retry(
        lambda: doc.ComputeStatistics(WD_STATISTIC_PAGES),
        descripcion="Word ComputeStatistics (para recortar páginas sobrantes)",
    )
    if total_paginas <= max_paginas:
        return False

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
    return True


def pegar_rangos_excel_como_imagen_en_word(ruta_excel, especificaciones, ruta_word, max_paginas=None):
    """Abre Excel y Word de verdad UNA sola vez, y pega en el Word varios
    rangos de Excel como imagen, uno por cada elemento de
    'especificaciones' (lista de tuplas (hoja_excel, rango, numero_pagina)),
    procesados EN ORDEN. Guarda el documento (sobrescribiendo 'ruta_word')
    solo al final, después de insertar TODAS las imágenes con éxito — si
    cualquier rango falla, no se guarda nada, para no dejar el documento
    con solo algunas imágenes pegadas a medias. Requiere Windows + Excel +
    Word instalados; usa pywin32.

    Si 'max_paginas' no es None, ANTES de guardar recorta el documento a
    ese máximo de páginas en esta MISMA sesión de Word (ver
    _recortar_documento_word_abierto) -- así, cuando además hay que pegar
    imágenes, se evita abrir Word una segunda vez solo para el recorte.

    CÓMO SE PEGA CADA RANGO (reemplaza el método anterior, basado en
    captura de pantalla + PNG): se copia el rango con Range.Copy() normal
    (lo mismo que Ctrl+C) y se pega en Word con Selection.PasteSpecial(
    DataType=wdPasteEnhancedMetafile) -- el mismo mecanismo que usa Word
    cuando alguien pega "a mano" un rango de Excel como imagen. Ventajas
    confirmadas con pruebas reales contra el método anterior:
      - Nitidez SIN límite de resolución: un EMF es vectorial (guarda
        instrucciones de dibujo, no una cuadrícula de píxeles), así que se
        redibuja nítido a cualquier tamaño/zoom/impresión. El método
        anterior dependía del zoom de pantalla de Excel al capturar (tope
        duro de 400%).
      - Archivos mucho más livianos (decenas de KB contra varios MB por
        imagen, probado con los rangos grandes de "Ax.A").
      - NO depende de que la pantalla esté realmente dibujada (probado con
        Excel y Word ambos invisibles) -- el método anterior (CopyPicture
        con Appearance=xlScreen) fallaba con capturas en blanco si la
        sesión estaba bloqueada o una sesión de Escritorio Remoto se
        desconectaba a mitad del proceso; esa clase de bug ya no puede
        pasar con este método.
    Selection.PasteSpecial deja una figura FLOTANTE (no incrustada en el
    texto); se convierte a InlineShape con Shape.ConvertToInlineShape()
    justo después, para que quede anclada en la posición del cursor (igual
    que hacía AddPicture antes) y para reusar la misma lógica de ajuste de
    ancho de acá abajo.

    Detalles importantes para que el pegado sea confiable de verdad
    (no solo "en teoría"):
      - Mantiene Excel y Word VISIBLES mientras procesa: Word no calcula
        bien la paginación real (ComputeStatistics, ver más abajo) con las
        apps invisibles, aunque el pegado en sí ya no lo necesite.
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
      - Verifica, contando las figuras del documento antes y después de
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

    # Evita que Windows apague la pantalla/suspenda el equipo por
    # inactividad a mitad de un procesamiento de varios segundos. Ya no
    # hace falta para que el pegado en sí funcione (ver nota grande
    # arriba), pero de todos modos conviene que la máquina no se suspenda
    # sola mientras Excel/Word están siendo automatizados. ES_CONTINUOUS
    # hace que el pedido quede activo hasta el próximo llamado (el de
    # abajo, en el finally, que lo saca) en vez de expirar solo.
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )

    ruta_excel = os.path.abspath(ruta_excel)
    ruta_word = os.path.abspath(ruta_word)

    MSO_TRUE = -1  # msoTrue

    excel = None
    word = None
    wb = None
    doc = None
    try:
        # Ver la nota junto a _cerrar_automatizaciones_huerfanas: si queda
        # un Excel/Word huérfano trabado de un intento anterior, Dispatch
        # se engancha a ESE MISMO proceso en vez de abrir uno nuevo.
        _cerrar_automatizaciones_huerfanas()

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
        # Recién abierto, Excel a veces tarda una fracción de segundo en
        # quedar listo para aceptar llamadas COM -- asignarle una propiedad
        # en ese instante puede fallar con un AttributeError rarísimo
        # ("Property 'Excel.Application.Visible' can not be set", visto en
        # la práctica), no el pywintypes.com_error habitual. reintentar
        # cualquier error acá (no solo RPC_E_CALL_REJECTED) resuelve ese
        # caso transitorio en vez de abortar todo el procesamiento.
        _com_retry(lambda: setattr(excel, "Visible", True), reintentar_cualquier_error=True, descripcion="Excel Visible=True")
        _com_retry(lambda: setattr(excel, "DisplayAlerts", False), reintentar_cualquier_error=True, descripcion="Excel DisplayAlerts=False")
        # UpdateLinks=0: el MODELO tiene un vínculo externo roto (referencia a
        # un Excel de la PC original donde se armó la plantilla, que ya no
        # existe en esta máquina). DisplayAlerts=False NO siempre alcanza para
        # suprimir el diálogo "¿Actualizar vínculos?" -- si ese diálogo queda
        # esperando una respuesta que nunca llega (nadie puede hacer clic en
        # automatización headless), Open() se queda colgado indefinidamente.
        # UpdateLinks=0 le dice a Excel que ni siquiera pregunte.
        wb = _com_retry(lambda: excel.Workbooks.Open(ruta_excel, UpdateLinks=0, CorruptLoad=XL_REPAIR_FILE))

        word = win32.Dispatch("Word.Application")
        # Mismo motivo que con Excel más arriba: recién abierto, puede
        # tardar un instante en aceptar llamadas COM.
        _com_retry(lambda: setattr(word, "Visible", True), reintentar_cualquier_error=True, descripcion="Word Visible=True")
        _com_retry(lambda: setattr(word, "DisplayAlerts", 0), reintentar_cualquier_error=True, descripcion="Word DisplayAlerts=0")  # wdAlertsNone: evita diálogos de Word que bloquearían las llamadas COM
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

            # Range.Copy() necesita que su hoja esté ACTIVA en Excel en ese
            # momento; si no lo está (ej. porque el rango anterior era de
            # otra hoja, o la hoja activa al abrir el archivo era otra),
            # falla con el error 1004 de Excel. Application.Goto con
            # Scroll=True activa la hoja Y desplaza la vista para que el
            # rango quede visible, que es la forma estándar de evitar eso.
            _com_retry(lambda: excel.Goto(rango_obj, True), descripcion=f"Excel Goto ({hoja_excel}!{rango})")
            _com_retry(lambda: rango_obj.Copy(), descripcion=f"Excel Copy ({hoja_excel}!{rango})")

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

            figuras_antes = _com_retry(lambda: doc.InlineShapes.Count + doc.Shapes.Count)

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
            # reintentar_cualquier_error=True: el portapapeles puede tardar
            # un instante en terminar de poblarse para rangos grandes --
            # pegar justo después a veces falla la primera vez.
            _com_retry(
                lambda: seleccion.PasteSpecial(DataType=WD_PASTE_ENHANCED_METAFILE),
                descripcion=f"Word PasteSpecial EMF ({hoja_excel}!{rango} → página {numero_pagina})",
                reintentar_cualquier_error=True,
            )

            cantidad_shapes = _com_retry(lambda: doc.Shapes.Count)
            if cantidad_shapes == 0:
                raise RuntimeError(
                    f"El pegado de '{hoja_excel}!{rango}' no insertó ninguna imagen nueva en el "
                    f"documento. Puede que otra aplicación haya usado el portapapeles justo en ese "
                    f"momento; intenta de nuevo sin tocar el mouse/teclado mientras se procesa."
                )
            forma_flotante = _com_retry(lambda: doc.Shapes(cantidad_shapes))
            # ConvertToInlineShape ancla la figura recién pegada en la
            # posición del cursor (igual que hacía AddPicture antes), en vez
            # de dejarla flotando con posición absoluta en la página.
            imagen_insertada = _com_retry(
                lambda: forma_flotante.ConvertToInlineShape(),
                descripcion=f"Word ConvertToInlineShape ({hoja_excel}!{rango})",
            )

            # La imagen se inserta a su tamaño NATIVO (el tamaño físico real
            # del rango en Excel) -- al ser vectorial, ese tamaño nativo
            # suele ser bastante más CHICO que el área útil de la página
            # (a diferencia del método anterior, que exportaba un PNG
            # inflado por el zoom de captura y por eso casi siempre había
            # que achicarlo). Por eso ya no alcanza con achicar SOLO
            # cuando sobra: la reescalamos siempre al tamaño MÁXIMO que
            # entre en la página (ancho y alto disponibles, lo que sea
            # más restrictivo), manteniendo la proporción -- tanto para
            # agrandar como para achicar. Como es vectorial, agrandarla no
            # le pierde nitidez.
            ancho_disponible = _com_retry(
                lambda: doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin,
                descripcion=f"Word PageSetup ancho (para {hoja_excel}!{rango})",
            )
            alto_disponible = _com_retry(
                lambda: doc.PageSetup.PageHeight - doc.PageSetup.TopMargin - doc.PageSetup.BottomMargin,
                descripcion=f"Word PageSetup alto (para {hoja_excel}!{rango})",
            )
            _com_retry(
                lambda: setattr(imagen_insertada, "LockAspectRatio", MSO_TRUE),
                descripcion=f"Word LockAspectRatio (para {hoja_excel}!{rango})",
            )
            _com_retry(
                lambda: setattr(imagen_insertada, "Width", ancho_disponible),
                descripcion=f"Word ajustar ancho de imagen (para {hoja_excel}!{rango})",
            )
            # Si al llenar el ancho la imagen se pasa del alto disponible
            # (rangos más altos que anchos, ej. "Ax.A"), se vuelve a ajustar
            # por el alto en cambio -- LockAspectRatio hace que el ancho se
            # achique en proporción solo.
            altura_actual = _com_retry(lambda: imagen_insertada.Height)
            if altura_actual > alto_disponible:
                _com_retry(
                    lambda: setattr(imagen_insertada, "Height", alto_disponible),
                    descripcion=f"Word ajustar alto de imagen (para {hoja_excel}!{rango})",
                )

            figuras_despues = _com_retry(lambda: doc.InlineShapes.Count + doc.Shapes.Count)
            if figuras_despues <= figuras_antes:
                raise RuntimeError(
                    f"El pegado de '{hoja_excel}!{rango}' no dejó ninguna imagen nueva en el "
                    f"documento después de convertirla. Volvé a intentar procesando el contrato."
                )

        if max_paginas is not None:
            _recortar_documento_word_abierto(word, doc, max_paginas)

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


def agregar_imagenes_a_contrato(excel_bytes, docx_bytes, especificaciones, max_paginas=None):
    """Envoltorio 'en memoria': recibe los bytes del Excel y del Word ya
    procesado, los escribe a archivos temporales (pywin32 necesita
    archivos reales en disco, no streams en memoria), llama a
    pegar_rangos_excel_como_imagen_en_word con TODAS las 'especificaciones'
    (lista de (hoja, rango, pagina)) en una sola sesión de Excel/Word, y
    devuelve los bytes del Word final. Limpia los archivos temporales al
    terminar.

    Si 'max_paginas' no es None, se lo pasa a
    pegar_rangos_excel_como_imagen_en_word para recortar el documento a ese
    máximo de páginas en la MISMA sesión de Word, en vez de necesitar abrir
    Word de nuevo aparte para eso (ver recortar_contrato_a_maximo_paginas,
    para cuando NO hay imágenes que pegar)."""
    tmp_excel = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp_word = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    try:
        tmp_excel.write(excel_bytes)
        tmp_excel.close()
        tmp_word.write(docx_bytes)
        tmp_word.close()

        pegar_rangos_excel_como_imagen_en_word(
            os.path.abspath(tmp_excel.name), especificaciones, os.path.abspath(tmp_word.name),
            max_paginas=max_paginas,
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
        _cerrar_automatizaciones_huerfanas()
        word = win32.Dispatch("Word.Application")
        # Recién abierto, Word puede tardar un instante en aceptar llamadas
        # COM (ver la nota junto a esto mismo en
        # pegar_rangos_excel_como_imagen_en_word).
        _com_retry(lambda: setattr(word, "Visible", True), reintentar_cualquier_error=True, descripcion="Word Visible=True")
        _com_retry(lambda: setattr(word, "DisplayAlerts", 0), reintentar_cualquier_error=True, descripcion="Word DisplayAlerts=0")  # wdAlertsNone: evita diálogos de Word que bloquearían las llamadas COM
        doc = _com_retry(lambda: word.Documents.Open(ruta_word))
        # Forzar el cálculo de la paginación REAL (ver la nota junto a esto
        # mismo en pegar_rangos_excel_como_imagen_en_word: con la app
        # invisible, Word no calcula bien la paginación real).
        _com_retry(lambda: setattr(doc.ActiveWindow.View, "Type", WD_PRINT_VIEW))

        if _recortar_documento_word_abierto(word, doc, max_paginas):
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

        _cerrar_automatizaciones_huerfanas()

        # Sin Visible=True a propósito: a diferencia de pegar_rangos_excel_como_imagen_en_word,
        # aquí no hay portapapeles ni paginación de Word de por medio — es
        # solo abrir, recalcular y guardar, así que puede quedar invisible
        # (más rápido y sin abrir una ventana de Excel en pantalla).
        excel = win32.Dispatch("Excel.Application")
        # Recién abierto, Excel puede tardar un instante en aceptar llamadas
        # COM -- asignarle una propiedad justo en ese momento a veces falla
        # con un AttributeError rarísimo ("Property 'Excel.Application.
        # Visible' can not be set", visto en la práctica: el usuario reportó
        # justo este error acá), no el pywintypes.com_error habitual.
        # reintentar cualquier error (no solo RPC_E_CALL_REJECTED) resuelve
        # ese caso transitorio en vez de abortar todo el recálculo.
        _com_retry(lambda: setattr(excel, "Visible", False), reintentar_cualquier_error=True, descripcion="Excel Visible=False")
        _com_retry(lambda: setattr(excel, "DisplayAlerts", False), reintentar_cualquier_error=True, descripcion="Excel DisplayAlerts=False")
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

        _cerrar_automatizaciones_huerfanas()

        # Sin Visible=True a propósito (igual que recalcular_excel_con_com):
        # solo se abre, se exporta a PDF y se cierra, sin que nadie necesite
        # ver la ventana de Word en pantalla.
        word = win32.Dispatch("Word.Application")
        # Recién abierto, Word puede tardar un instante en aceptar llamadas
        # COM (ver la nota junto a esto mismo en
        # pegar_rangos_excel_como_imagen_en_word).
        _com_retry(lambda: setattr(word, "Visible", False), reintentar_cualquier_error=True, descripcion="Word Visible=False")
        _com_retry(lambda: setattr(word, "DisplayAlerts", 0), reintentar_cualquier_error=True, descripcion="Word DisplayAlerts=0")  # wdAlertsNone
        doc = _com_retry(lambda: word.Documents.Open(ruta_docx), descripcion="Word Documents.Open (a PDF)")
        # ExportAsFixedFormat (no el simple doc.SaveAs) a propósito: SaveAs
        # usa el perfil de calidad "para pantalla" por defecto de Word, que
        # recomprime MUY agresivamente las imágenes pegadas (probado: una
        # imagen de 4398x7586px quedaba en solo 500x863px, prácticamente
        # ilegible al hacer zoom). ExportAsFixedFormat con
        # OptimizeFor=wdExportOptimizeForPrint (0, no wdExportOptimizeFor
        # OnScreen=1) preserva mucha más resolución real (misma imagen de
        # prueba: 1043x1799px, ~4x más píxeles), a cambio de un PDF apenas
        # más pesado.
        WD_EXPORT_FORMAT_PDF = 17            # WdExportFormat.wdExportFormatPDF
        WD_EXPORT_OPTIMIZE_FOR_PRINT = 0     # WdExportOptimizeFor.wdExportOptimizeForPrint
        WD_EXPORT_ALL_DOCUMENT = 0           # WdExportRange.wdExportAllDocument
        WD_EXPORT_DOCUMENT_CONTENT = 0       # WdExportItem.wdExportDocumentContent
        _com_retry(
            lambda: doc.ExportAsFixedFormat(
                OutputFileName=ruta_pdf,
                ExportFormat=WD_EXPORT_FORMAT_PDF,
                OptimizeFor=WD_EXPORT_OPTIMIZE_FOR_PRINT,
                Range=WD_EXPORT_ALL_DOCUMENT,
                Item=WD_EXPORT_DOCUMENT_CONTENT,
                BitmapMissingFonts=True,
            ),
            descripcion="Word doc.ExportAsFixedFormat (a PDF)",
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
