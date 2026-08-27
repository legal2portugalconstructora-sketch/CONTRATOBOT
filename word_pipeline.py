# -*- coding: utf-8 -*-
"""
word_pipeline.py — Generación/edición de documentos Word (CONTRATO), sin
automatización COM: solo openpyxl + python-docx, verificable en cualquier
sistema operativo.
"""

import io
import re
from datetime import datetime, date

from openpyxl import load_workbook
from docx import Document
from docx.enum.dml import MSO_THEME_COLOR

from excel_pipeline import HOJA_AXA, _lee
from templates_manager import CELDAS_AXA_POR_DEFECTO

# =========================================================================
# PESTAÑA "🤖CONTRATO – Word" — Generación/edición de documentos Word
# =========================================================================
# Nombres de los 4 documentos Word que se suben en esta pestaña:
DOCUMENTOS_WORD = ["CONTRATO", "REGLAMENTO", "MEMBRESIA", "DERECHOHABIENTES"]

# Las 4 opciones de tipo de propiedad (selección única, mutuamente excluyentes)
OPCIONES_PROPIEDAD = [
    "UNICO PROPIETARIO",
    "CASADOS QUE VIVEN EN EL MISMO SITIO",
    "COPROPIETARIOS NO CASADOS QUE VIVEN EN EL MISMO SITIO",
    "COPROPIETARIOS NO CASADOS QUE VIVEN EN SITIOS DIFERENTES",
]

HOJA_DATOS_WORD = "DATOS"

# Los 3 párrafos que se eliminan del CONTRATO cuando se elige "UNICO
# PROPIETARIO" (son las variantes de cláusula para comprador casado o
# copropietarios, que no aplican si hay un único propietario).
PARRAFOS_A_ELIMINAR_UNICO_PROPIETARIO = [
    "RRRRRRR, identificado/a con DNI N° NNNNNNN, casado con KKKKKK identificada/o con DNI N° "
    "JJJJJJJ, ambos con domicilio para estos efectos en LLLLLLLLLL, distrito de WWWWW, "
    "provincia de EEEEEEE y Departamento de TTTTTTT, En adelante se le denominará EL "
    "COMPRADOR; conforme a las cláusulas siguientes.",
    "RRRRRRR, identificado/a con DNI N° NNNNNNN, como copropietario KKKKKK identificada/o con "
    "DNI N° JJJJJJJ, ambos con domicilio para estos efectos en LLLLLLLLLL, distrito de "
    "WWWWW, provincia de EEEEEEE y Departamento de TTTTTTT, En adelante se le denominará "
    "EL COMPRADOR; conforme a las cláusulas siguientes.",
    "RRRRRRR, identificado/a con DNI N° NNNNNNN, con domicilio para estos efectos en "
    "LLLLLLLLLL, distrito de WWWWW, provincia de EEEEEEE y Departamento de TTTTTTT como "
    "copropietario KKKKKK identificada/o con DNI N° JJJJJJJ, con domicilio para estos "
    "efectos en OOOOOO, distrito de PPPPP, provincia de ÑÑÑÑÑÑ y Departamento de "
    "XXXXXXX, En adelante se le denominará EL COMPRADOR; conforme a las cláusulas "
    "siguientes.",
]


def _normalizar_espacios(texto):
    return re.sub(r"\s+", " ", texto or "").strip()


# Excel siempre guarda las fórmulas con los nombres de función en inglés
# (ver nota junto a "sedes_validas" en templates_manager.CELDAS_AXA_POR_DEFECTO),
# así que solo hace falta reconocer UPPER/LOWER/PROPER, nunca sus
# equivalentes en español (MAYUSC/MINUSC/NOMPROPIO) aunque la interfaz de
# Excel esté en español.
_TRANSFORMACIONES_TEXTO_EXCEL = {
    "UPPER": str.upper,
    "LOWER": str.lower,
    "PROPER": str.title,
}
# Función opcional envolviendo la referencia, ej. "UPPER(O127)" -> grupos
# ("UPPER", "O127"); si no hay función, el cuerpo completo es la referencia.
_PATRON_FUNCION_TEXTO = re.compile(r"^(UPPER|LOWER|PROPER)\s*\((.+)\)$", re.IGNORECASE)
# Referencia a UNA celda, con hoja opcional (misma hoja si no se indica),
# ej. "O127", "Cotizador!I16" o "'Ax.A'!O132" (comillas si el nombre de
# la hoja tiene puntos/espacios, como hace Excel). El "(?:\[\d+\])?" al
# principio tolera (e ignora) el prefijo que Excel antepone cuando la
# fórmula quedó apuntando a un LIBRO EXTERNO (ej. "=[1]Cotizador!I16" --
# visto de verdad en el MODELO de FINCA DEL VALLE): como de todos modos
# 'wb' acá es siempre el mismo libro que se está procesando, y el nombre
# de hoja coincide con una hoja real de ESE libro, tiene más sentido leer
# el valor actual de esa hoja local que fallar o devolver el texto crudo
# de la fórmula.
_PATRON_REF_HOJA_CELDA = re.compile(
    r"^(?:\[\d+\])?(?:(?:'(?P<hoja_q>[^']+)'|(?P<hoja>[A-Za-z_][A-Za-z0-9_.]*))!)?"
    r"\$?(?P<columna>[A-Za-z]{1,3})\$?(?P<fila>\d{1,7})$"
)


def _resolver_texto_formula(wb, hoja_actual, valor_bruto):
    """Si 'valor_bruto' es la fórmula (texto crudo, sin recalcular) de una
    celda que simplemente REFERENCIA otra celda —de la misma hoja (ej.
    "=O127") o de otra hoja (ej. "=Cotizador!I16")—, opcionalmente
    envuelta en una transformación de texto (=UPPER(...), =PROPER(...),
    =LOWER(...)), devuelve el valor ACTUAL de esa celda referenciada (con
    la transformación aplicada en Python si corresponde) — en vez del
    resultado que quedó cacheado en el .xlsx, que openpyxl nunca recalcula
    y por eso puede haber quedado desactualizado, o en vez de devolver el
    texto literal de la fórmula. Si 'valor_bruto' no es una referencia
    simple reconocible, se devuelve tal cual (texto fijo u otro valor)."""
    if not isinstance(valor_bruto, str) or not valor_bruto.startswith("="):
        return valor_bruto

    cuerpo = valor_bruto.strip()[1:].strip()
    funcion = None
    match_funcion = _PATRON_FUNCION_TEXTO.match(cuerpo)
    if match_funcion:
        funcion = match_funcion.group(1).upper()
        cuerpo = match_funcion.group(2).strip()

    match_ref = _PATRON_REF_HOJA_CELDA.match(cuerpo)
    if not match_ref:
        return valor_bruto

    nombre_hoja = match_ref.group("hoja_q") or match_ref.group("hoja")
    if nombre_hoja:
        if nombre_hoja not in wb.sheetnames:
            raise ValueError(
                f"La fórmula '{valor_bruto}' referencia la hoja '{nombre_hoja}', que no "
                f"existe en el Excel subido."
            )
        hoja_ref = wb[nombre_hoja]
    else:
        hoja_ref = hoja_actual

    valor_celda = _lee(hoja_ref, f"{match_ref.group('columna')}{match_ref.group('fila')}")
    if valor_celda is None or funcion is None:
        return valor_celda
    return _TRANSFORMACIONES_TEXTO_EXCEL[funcion](str(valor_celda))


_MESES_ESPANOL = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]


def _fecha_en_letras(valor):
    """Convierte una fecha ('valor' es datetime/date, o texto en un
    formato de fecha reconocido) al formato 'D de MES del AAAA' en
    español, ej. '3 de julio del 2026'. Si 'valor' no se pudo interpretar
    como fecha, se devuelve tal cual (sin tocarlo)."""
    fecha = None
    if isinstance(valor, (datetime, date)):
        fecha = valor
    elif isinstance(valor, str):
        texto = valor.strip()
        for formato in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y"):
            try:
                fecha = datetime.strptime(texto, formato)
                break
            except ValueError:
                continue
    if fecha is None:
        return valor
    return f"{fecha.day} de {_MESES_ESPANOL[fecha.month - 1]} del {fecha.year}"


def _detectar_sede_desde_g124(sheet_axa_formulas, celdas_axa):
    """Determina qué Sede (Arequipa/Ilo) se eligió al procesar el Excel
    (Paso 4, excel_pipeline.procesar_paso4_axa), comparando la fórmula
    LITERAL que quedó escrita en celdas_axa["g124"] (ej. "=UPPER(O127)")
    contra las fórmulas de celdas_axa["sedes_validas"]. No hace falta que
    la Sede se vuelva a indicar en el módulo Word -- ya quedó "grabada" en
    esa fórmula. Devuelve el nombre de la sede tal cual aparece en
    'sedes_validas' (ej. "Arequipa"), o None si esa celda no tiene
    ninguna de las fórmulas esperadas (ej. el Paso 4 no se corrió, o no
    se indicó ninguna sede en su momento)."""
    celda_g124 = celdas_axa.get("g124")
    if not celda_g124:
        return None
    formula_g124 = _lee(sheet_axa_formulas, celda_g124)
    if not isinstance(formula_g124, str):
        return None
    formula_g124 = formula_g124.strip()
    for nombre_sede, formula_esperada in celdas_axa.get("sedes_validas", {}).items():
        if formula_g124 == formula_esperada:
            return nombre_sede
    return None


def leer_datos_para_contrato(archivo_excel, celdas_axa=None):
    """Lee del Excel subido en la pestaña Word los valores que
    reemplazan a los placeholders del CONTRATO. Devuelve un diccionario
    {placeholder: valor}.

    - De la hoja 'DATOS': RRRRRRR (E11), NNNNNNN (E13), LLLLLLLLLL (E22),
      WWWWW (H26), EEEEEEE (H27), TTTTTTT (H28).
    - De la hoja 'Ax.A': UUUUUUU (celdas_axa["o132"], "O132" por defecto).
      Si esa celda es una fórmula que referencia otra celda (misma hoja u
      otra, ej. "=Cotizador!I16"), opcionalmente envuelta en
      =UPPER(...)/=PROPER(...)/=LOWER(...), se resuelve leyendo el valor
      ACTUAL de esa celda referenciada (ver _resolver_texto_formula) en
      vez de copiar el texto literal de la fórmula o un resultado
      cacheado desactualizado. Si ese valor resulta ser una fecha, se
      escribe en letras ('D de MES del AAAA', ver _fecha_en_letras) en
      vez de como fecha/hora cruda. Si además se puede determinar la Sede
      elegida (ver _detectar_sede_desde_g124), se antepone "en la ciudad
      de <Sede> el " (ej. "en la ciudad de Arequipa el 15 de agosto del
      2026"); si no se puede determinar, queda solo la fecha, como antes.

    celdas_axa: mapeo de coordenadas de "Ax.A" a usar (ver
        templates_manager.obtener_celdas_axa); si no se pasa, se usan las
        de templates_manager.CELDAS_AXA_POR_DEFECTO (layout de COSTA DEL
        VALLE / COSTA AZUL).
    """
    celdas_axa = celdas_axa or CELDAS_AXA_POR_DEFECTO
    celda_o132 = celdas_axa["o132"]

    wb = load_workbook(archivo_excel, data_only=True)
    if HOJA_DATOS_WORD not in wb.sheetnames:
        raise ValueError(f"La hoja '{HOJA_DATOS_WORD}' no existe en el Excel subido.")
    sheet_datos = wb[HOJA_DATOS_WORD]

    if HOJA_AXA not in wb.sheetnames:
        raise ValueError(f"La hoja '{HOJA_AXA}' no existe en el Excel subido (necesaria para UUUUUUU).")
    sheet_axa = wb[HOJA_AXA]

    archivo_excel.seek(0)
    wb_formulas = load_workbook(archivo_excel, data_only=False)
    sheet_axa_formulas = wb_formulas[HOJA_AXA]
    valor_o132 = _resolver_texto_formula(wb, sheet_axa, _lee(sheet_axa_formulas, celda_o132))
    valor_o132 = _fecha_en_letras(valor_o132)

    # Sin "and valor_o132" quedaría literal "en la ciudad de Arequipa el
    # None" si la fecha (Cotizador!I16) vino vacía -- se prefiere no
    # anteponer nada en ese caso, igual que ya se comportaba antes de
    # agregar la ciudad.
    sede_detectada = _detectar_sede_desde_g124(sheet_axa_formulas, celdas_axa)
    if sede_detectada and valor_o132:
        valor_o132 = f"en la ciudad de {sede_detectada} el {valor_o132}"

    return {
        "RRRRRRR": _lee(sheet_datos, "E11"),
        "NNNNNNN": _lee(sheet_datos, "E13"),
        "LLLLLLLLLL": _lee(sheet_datos, "E22"),
        "WWWWW": _lee(sheet_datos, "H26"),
        "EEEEEEE": _lee(sheet_datos, "H27"),
        "TTTTTTT": _lee(sheet_datos, "H28"),
        "UUUUUUU": valor_o132,
    }


def _eliminar_parrafo(paragraph):
    elemento = paragraph._element
    elemento.getparent().remove(elemento)


# Umbral mínimo de letras repetidas consecutivas para considerar que es un
# "placeholder" (ej. "RRRRRRR") y no una coincidencia normal del idioma
# (en español casi nunca hay 4+ letras idénticas seguidas). Usar un patrón
# flexible ("R{4,}" en vez de "RRRRRRR" literal) hace que el reemplazo y el
# borrado de párrafos NO dependan de contar a mano cuántas letras tiene
# cada placeholder en el documento real — evita errores como el de
# "LLLLLLLLLLL" vs "LLLLLLLLLL" que ya tuvimos.
UMBRAL_LETRAS_PLACEHOLDER = 4


def _patron_flexible_desde_texto(texto):
    """Convierte un texto con placeholders (letras repetidas, ej.
    'RRRRRRR') en un patrón regex donde cada corrida de 4+ letras
    idénticas se vuelve un comodín (coincide con esa misma letra
    repetida CUALQUIER cantidad de veces >=1), y cada espacio se vuelve
    '\\s+' (tolera espacios extra/distintos). El resto del texto se
    escapa literalmente. Así, para decidir qué párrafo borrar, ya no
    importa si el placeholder real tiene 6, 7 u 11 letras."""
    texto = _normalizar_espacios(texto)
    partes = []
    i, n = 0, len(texto)
    while i < n:
        c = texto[i]
        j = i
        while j < n and texto[j] == c:
            j += 1
        largo = j - i
        if c.isalpha() and largo >= UMBRAL_LETRAS_PLACEHOLDER:
            partes.append(re.escape(c) + "+")
        elif c.isspace():
            partes.append(r"\s+")
        else:
            partes.append(re.escape(texto[i:j]))
        i = j
    return "".join(partes)


def _patron_placeholder(placeholder):
    """Patrón regex para UN placeholder puntual (ej. 'RRRRRRR' o
    'LLLLLLLLLL'): coincide con la letra que lo compone repetida 4 o
    más veces, sin importar la cantidad exacta."""
    letra = placeholder[0]
    return re.compile(re.escape(letra) + "{" + str(UMBRAL_LETRAS_PLACEHOLDER) + ",}")


def eliminar_parrafos_por_texto(documento, textos_a_eliminar):
    """Elimina, del cuerpo principal del documento, cualquier párrafo
    cuyo texto (normalizado y con los placeholders convertidos a
    comodín) coincida con alguno de los textos dados. Devuelve cuántos
    párrafos se eliminaron."""
    patrones = [re.compile(_patron_flexible_desde_texto(t)) for t in textos_a_eliminar]
    eliminados = 0
    for paragraph in list(documento.paragraphs):
        texto_parrafo = _normalizar_espacios(paragraph.text)
        if not texto_parrafo:
            continue
        if any(patron.search(texto_parrafo) for patron in patrones):
            _eliminar_parrafo(paragraph)
            eliminados += 1
    return eliminados


def _reemplazar_en_parrafo(paragraph, patron_regex, reemplazar):
    if not patron_regex.search(paragraph.text):
        return
    reemplazar = "" if reemplazar is None else str(reemplazar)
    # 1) Intento simple: el placeholder completo vive dentro de UN solo run
    #    (lo más común) -> se reemplaza ahí mismo, conservando su formato.
    for run in paragraph.runs:
        if patron_regex.search(run.text):
            run.text = patron_regex.sub(reemplazar, run.text)
    # 2) Si el placeholder quedó partido entre varios runs (Word a veces
    #    separa el texto en runs distintos), se reconstruye el párrafo
    #    completo: se une todo el texto, se reemplaza, y se vuelca en el
    #    primer run (los demás quedan vacíos). Se pierde el formato por
    #    run intermedio en ese caso puntual, pero el contenido queda
    #    correcto.
    if patron_regex.search(paragraph.text) and paragraph.runs:
        texto_completo = "".join(r.text for r in paragraph.runs)
        texto_nuevo = patron_regex.sub(reemplazar, texto_completo)
        paragraph.runs[0].text = texto_nuevo
        for run in paragraph.runs[1:]:
            run.text = ""


def reemplazar_en_documento(documento, reemplazos):
    """Aplica _reemplazar_en_parrafo en todos los párrafos del cuerpo
    principal, de las tablas (celda por celda) y de encabezados/pies de
    página del documento. 'reemplazos' es {placeholder: valor}; el
    placeholder se usa solo para derivar la letra y armar el patrón
    flexible (ver _patron_placeholder)."""
    patrones = {placeholder: _patron_placeholder(placeholder) for placeholder in reemplazos}

    def _procesar_parrafos(parrafos):
        for paragraph in parrafos:
            for placeholder, valor in reemplazos.items():
                _reemplazar_en_parrafo(paragraph, patrones[placeholder], valor)

    _procesar_parrafos(documento.paragraphs)
    for tabla in documento.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                _procesar_parrafos(celda.paragraphs)
    for section in documento.sections:
        _procesar_parrafos(section.header.paragraphs)
        _procesar_parrafos(section.footer.paragraphs)


def _detectar_placeholders_sin_valor(valores):
    """Devuelve la lista de placeholders (ordenada) cuyo valor vino vacío
    del Excel -- celda vacía (None), solo espacios, o el texto literal
    'None' -- para poder avisarle al usuario en la interfaz que ese dato
    no se pudo completar en el documento, en vez de dejar que pase
    desapercibido como un espacio en blanco."""
    faltantes = []
    for placeholder, valor in valores.items():
        if valor is None:
            faltantes.append(placeholder)
        elif isinstance(valor, str) and valor.strip() in ("", "None"):
            faltantes.append(placeholder)
    return faltantes


def procesar_paso1_unico_propietario(archivo_excel, archivo_contrato, membresia, celdas_axa=None):
    """Paso 1: aplica al documento CONTRATO las instrucciones para la opción
    'UNICO PROPIETARIO':
      1. Borra los 3 párrafos de cláusulas de comprador casado/copropietarios.
      2. Reemplaza en todo el documento los placeholders RRRRRRR, NNNNNNN,
         LLLLLLLLLL, WWWWW, EEEEEEE y TTTTTTT por los valores de la hoja
         'DATOS' del Excel (E11, E13, E22, H26, H27, H28 respectivamente),
         y UUUUUUU por el valor de la hoja 'Ax.A' (O132).
      3. Reemplaza toda aparición de "JDIP" por "ANEXO F" o "ANEXO E" según
         'membresia' (ver _reemplazar_jdip_segun_membresia).
    Devuelve (bytes_del_docx_resultante, info) donde info trae detalles
    para mostrar en la interfaz.
    """
    valores = leer_datos_para_contrato(archivo_excel, celdas_axa=celdas_axa)

    archivo_contrato.seek(0)
    documento = Document(archivo_contrato)

    parrafos_eliminados = eliminar_parrafos_por_texto(documento, PARRAFOS_A_ELIMINAR_UNICO_PROPIETARIO)
    reemplazar_en_documento(documento, valores)
    texto_jdip = _reemplazar_jdip_segun_membresia(documento, membresia)

    buffer = io.BytesIO()
    documento.save(buffer)
    buffer.seek(0)

    info = {
        "parrafos_eliminados": parrafos_eliminados,
        "valores_usados": valores,
        "jdip_reemplazado_por": texto_jdip,
        "placeholders_sin_valor": _detectar_placeholders_sin_valor(valores),
    }
    return buffer.getvalue(), info


# =========================================================================
# PASO 2 — CONTRATO para "CASADOS QUE VIVEN EN EL MISMO SITIO"
# =========================================================================
# Los 3 párrafos que se eliminan del CONTRATO cuando se elige "CASADOS QUE
# VIVEN EN EL MISMO SITIO" (variantes de cláusula para único propietario o
# copropietarios, que no aplican si los compradores son cónyuges que viven
# en el mismo domicilio; esto deja en el documento la variante para
# cónyuges, que no se borra).
PARRAFOS_A_ELIMINAR_CASADOS_MISMO_SITIO = [
    "RRRRRRR, identificado/a con DNI N° NNNNNNN con domicilio para estos efectos en "
    "LLLLLLLLLL, distrito de WWWWW, provincia de EEEEEEE y Departamento de TTTTTTT, En "
    "adelante se le denominará EL COMPRADOR; conforme a las cláusulas siguientes.",
    "RRRRRRR, identificado/a con DNI N° NNNNNNN, como copropietario KKKKKK identificada/o con "
    "DNI N° JJJJJJJ, ambos con domicilio para estos efectos en LLLLLLLLLLL, distrito de "
    "WWWWW, provincia de EEEEEEE y Departamento de TTTTTTT, En adelante se le denominará "
    "EL COMPRADOR; conforme a las cláusulas siguientes.",
    "RRRRRRR, identificado/a con DNI N° NNNNNNN, con domicilio para estos efectos en "
    "LLLLLLLLLLL, distrito de WWWWW, provincia de EEEEEEE y Departamento de TTTTTTT como "
    "copropietario KKKKKK identificada/o con DNI N° JJJJJJJ, con domicilio para estos "
    "efectos en OOOOOO, distrito de PPPPP, provincia de ÑÑÑÑÑÑ y Departamento de "
    "XXXXXXX, En adelante se le denominará EL COMPRADOR; conforme a las cláusulas "
    "siguientes.",
]

# Textos a los que, en el Paso 2, hay que ponerles color de fuente
# "Negro, Texto 1" (viven en la variante de cónyuge que queda en el
# documento; KKKKKK/JJJJJJJ son datos del cónyuge que se llenan a mano, no
# vienen del Excel). "TITULAR / EL COMPRADOR" es texto fijo; los demás usan
# comodines (ver _patron_texto_con_comodines) para no depender de contar a
# mano cuántos guiones bajos o letras repetidas tiene cada uno en el
# documento real.
TEXTOS_COLOR_NEGRO_TEXTO1_CASADOS_MISMO_SITIO = [
    "______________________________________________",
    "TITULAR / EL COMPRADOR",
    "KKKKKK",
    "DNI N° JJJJJJJ",
]


def _patron_texto_con_comodines(texto):
    """Igual que _patron_flexible_desde_texto (letras repetidas 4+ veces
    -> comodín, espacios flexibles), pero además trata corridas de 3+
    guiones bajos ('___') como comodín (para líneas de firma en blanco,
    cuya cantidad exacta de guiones no queremos tener que contar a mano).
    Se usa para BUSCAR texto al que aplicar formato, no para reemplazarlo
    ni para decidir qué párrafo borrar, así que vive aparte de
    _patron_flexible_desde_texto."""
    texto = _normalizar_espacios(texto)
    partes = []
    i, n = 0, len(texto)
    while i < n:
        c = texto[i]
        j = i
        while j < n and texto[j] == c:
            j += 1
        largo = j - i
        if c.isalpha() and largo >= UMBRAL_LETRAS_PLACEHOLDER:
            partes.append(re.escape(c) + "+")
        elif c == "_" and largo >= 3:
            partes.append("_+")
        elif c.isspace():
            partes.append(r"\s+")
        else:
            partes.append(re.escape(texto[i:j]))
        i = j
    return "".join(partes)


def _establecer_color_negro_texto1_en_run(run):
    """Fuerza el color de fuente 'Negro, Texto 1' en 'run' (el mismo
    swatch de tema que Word escribe si lo eligen a mano en el selector de
    color de fuente: <w:color w:val="000000" w:themeColor="text1"/>).
    Se prefirió esto sobre 'Automático' (color.rgb = None, o w:val="auto")
    porque, si el color visible del texto viene heredado de un ESTILO de
    Word (de carácter o de párrafo) y no de formato manual, ninguna de
    esas dos formas de "Automático" se veía reflejada de forma
    consistente; escribir un color de TEMA explícito (con su propio
    w:val de respaldo) sí gana siempre, sin importar de dónde venía el
    color original."""
    run.font.color.theme_color = MSO_THEME_COLOR.TEXT_1


def _color_negro_texto1_en_parrafo(paragraph, patrones_regex):
    """Busca, dentro del texto COMPLETO del párrafo (uniendo todos sus
    runs), cualquier coincidencia de los patrones dados, y les pone color
    de fuente "Negro, Texto 1" (ver _establecer_color_negro_texto1_en_run)
    a todos los runs que se solapen con esa coincidencia. Si la
    coincidencia queda repartida entre varios runs, se recolorean todos
    los runs involucrados completos (misma limitación de granularidad que
    ya tiene _reemplazar_en_parrafo con texto partido entre runs)."""
    texto_completo = "".join(run.text for run in paragraph.runs)
    if not texto_completo:
        return
    limites = []
    posicion = 0
    for run in paragraph.runs:
        limites.append((posicion, posicion + len(run.text)))
        posicion += len(run.text)
    for patron in patrones_regex:
        for match in patron.finditer(texto_completo):
            inicio, fin = match.start(), match.end()
            for run, (run_inicio, run_fin) in zip(paragraph.runs, limites):
                if run_inicio < fin and run_fin > inicio:
                    _establecer_color_negro_texto1_en_run(run)


def establecer_color_negro_texto1_en_documento(documento, textos):
    """Aplica _color_negro_texto1_en_parrafo en todos los párrafos del
    cuerpo principal, de las tablas (celda por celda) y de
    encabezados/pies de página del documento, para cada texto en
    'textos' (ver _patron_texto_con_comodines)."""
    patrones = [re.compile(_patron_texto_con_comodines(t)) for t in textos]

    def _procesar_parrafos(parrafos):
        for paragraph in parrafos:
            _color_negro_texto1_en_parrafo(paragraph, patrones)

    _procesar_parrafos(documento.paragraphs)
    for tabla in documento.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                _procesar_parrafos(celda.paragraphs)
    for section in documento.sections:
        _procesar_parrafos(section.header.paragraphs)
        _procesar_parrafos(section.footer.paragraphs)


def leer_datos_conyuge_paso2(archivo_excel):
    """Lee de la hoja 'DATOS' los valores que reemplazan a KKKKKK (E33) y
    JJJJJJJ (E35): datos del cónyuge, que solo aplican al Paso 2 ('CASADOS
    QUE VIVEN EN EL MISMO SITIO'), para los 4 documentos (CONTRATO,
    REGLAMENTO, MEMBRESIA, DERECHOHABIENTES) — igual que los placeholders
    de leer_datos_para_contrato. Devuelve {placeholder: valor}."""
    wb = load_workbook(archivo_excel, data_only=True)
    if HOJA_DATOS_WORD not in wb.sheetnames:
        raise ValueError(f"La hoja '{HOJA_DATOS_WORD}' no existe en el Excel subido.")
    sheet_datos = wb[HOJA_DATOS_WORD]
    return {
        "KKKKKK": _lee(sheet_datos, "E33"),
        "JJJJJJJ": _lee(sheet_datos, "E35"),
    }


def procesar_paso2_casados_mismo_sitio(archivo_excel, archivo_contrato, membresia, celdas_axa=None):
    """Paso 2: aplica al documento CONTRATO las instrucciones para la
    opción 'CASADOS QUE VIVEN EN EL MISMO SITIO':
      1. Borra los 3 párrafos de cláusulas de único propietario/copropietarios
         (ver PARRAFOS_A_ELIMINAR_CASADOS_MISMO_SITIO), dejando la variante
         de cónyuge en el documento.
      2. Reemplaza en todo el documento los mismos placeholders que el
         Paso 1 (RRRRRRR, NNNNNNN, LLLLLLLLLL, WWWWW, EEEEEEE, TTTTTTT,
         UUUUUUU) por los valores del Excel.
      3. Pone color de fuente "Negro, Texto 1" en los textos de
         TEXTOS_COLOR_NEGRO_TEXTO1_CASADOS_MISMO_SITIO (línea de firma en
         blanco, etiqueta "TITULAR / EL COMPRADOR", KKKKKK y
         "DNI N° JJJJJJJ") -- ANTES de reemplazar KKKKKK/JJJJJJJ, para que
         la búsqueda de esos textos literales todavía los encuentre.
      4. Reemplaza KKKKKK y JJJJJJJ por los datos del cónyuge (ver
         leer_datos_conyuge_paso2).
      5. Reemplaza toda aparición de "JDIP" por "ANEXO F" o "ANEXO E" según
         'membresia' (ver _reemplazar_jdip_segun_membresia).
    Devuelve (bytes_del_docx_resultante, info) donde info trae detalles
    para mostrar en la interfaz.
    """
    valores = leer_datos_para_contrato(archivo_excel, celdas_axa=celdas_axa)
    archivo_excel.seek(0)
    valores_conyuge = leer_datos_conyuge_paso2(archivo_excel)

    archivo_contrato.seek(0)
    documento = Document(archivo_contrato)

    parrafos_eliminados = eliminar_parrafos_por_texto(documento, PARRAFOS_A_ELIMINAR_CASADOS_MISMO_SITIO)
    reemplazar_en_documento(documento, valores)
    establecer_color_negro_texto1_en_documento(documento, TEXTOS_COLOR_NEGRO_TEXTO1_CASADOS_MISMO_SITIO)
    reemplazar_en_documento(documento, valores_conyuge)
    texto_jdip = _reemplazar_jdip_segun_membresia(documento, membresia)

    buffer = io.BytesIO()
    documento.save(buffer)
    buffer.seek(0)

    valores_finales = {**valores, **valores_conyuge}
    info = {
        "parrafos_eliminados": parrafos_eliminados,
        "valores_usados": valores_finales,
        "jdip_reemplazado_por": texto_jdip,
        "placeholders_sin_valor": _detectar_placeholders_sin_valor(valores_finales),
    }
    return buffer.getvalue(), info


# =========================================================================
# PASO 3 — CONTRATO para "COPROPIETARIOS NO CASADOS QUE VIVEN EN EL MISMO
# SITIO". Igual que el Paso 2, solo cambian los 3 párrafos que se borran
# (deja en el documento la variante de copropietarios que viven en el
# mismo sitio, en vez de la variante de cónyuges).
# =========================================================================
PARRAFOS_A_ELIMINAR_COPROPIETARIOS_MISMO_SITIO = [
    "RRRRRRR, identificado/a con DNI N° NNNNNNN con domicilio para estos efectos en "
    "LLLLLLLLLL, distrito de WWWWW, provincia de EEEEEEE y Departamento de TTTTTTT, En "
    "adelante se le denominará EL COMPRADOR; conforme a las cláusulas siguientes.",
    "RRRRRRR, identificado/a con DNI N° NNNNNNN, casado con KKKKKK identificada/o con DNI N° "
    "JJJJJJJ, ambos con domicilio para estos efectos en LLLLLLLLLLL, distrito de WWWWW, "
    "provincia de EEEEEEE y Departamento de TTTTTTT, En adelante se le denominará EL "
    "COMPRADOR; conforme a las cláusulas siguientes.",
    "RRRRRRR, identificado/a con DNI N° NNNNNNN, con domicilio para estos efectos en "
    "LLLLLLLLLLL, distrito de WWWWW, provincia de EEEEEEE y Departamento de TTTTTTT como "
    "copropietario KKKKKK identificada/o con DNI N° JJJJJJJ, con domicilio para estos "
    "efectos en OOOOOO, distrito de PPPPP, provincia de ÑÑÑÑÑÑ y Departamento de "
    "XXXXXXX, En adelante se le denominará EL COMPRADOR; conforme a las cláusulas "
    "siguientes.",
]


def procesar_paso3_copropietarios_no_casados_mismo_sitio(archivo_excel, archivo_contrato, membresia, celdas_axa=None):
    """Paso 3: aplica al documento CONTRATO las instrucciones para la
    opción 'COPROPIETARIOS NO CASADOS QUE VIVEN EN EL MISMO SITIO'.
    Idéntico al Paso 2 (ver procesar_paso2_casados_mismo_sitio), salvo
    que borra los 3 párrafos de PARRAFOS_A_ELIMINAR_COPROPIETARIOS_MISMO_SITIO
    en vez de los del Paso 2. También reemplaza "JDIP" según 'membresia'
    (ver _reemplazar_jdip_segun_membresia). Devuelve (bytes_del_docx_resultante, info).
    """
    valores = leer_datos_para_contrato(archivo_excel, celdas_axa=celdas_axa)
    archivo_excel.seek(0)
    valores_conyuge = leer_datos_conyuge_paso2(archivo_excel)

    archivo_contrato.seek(0)
    documento = Document(archivo_contrato)

    parrafos_eliminados = eliminar_parrafos_por_texto(documento, PARRAFOS_A_ELIMINAR_COPROPIETARIOS_MISMO_SITIO)
    reemplazar_en_documento(documento, valores)
    establecer_color_negro_texto1_en_documento(documento, TEXTOS_COLOR_NEGRO_TEXTO1_CASADOS_MISMO_SITIO)
    reemplazar_en_documento(documento, valores_conyuge)
    texto_jdip = _reemplazar_jdip_segun_membresia(documento, membresia)

    buffer = io.BytesIO()
    documento.save(buffer)
    buffer.seek(0)

    valores_finales = {**valores, **valores_conyuge}
    info = {
        "parrafos_eliminados": parrafos_eliminados,
        "valores_usados": valores_finales,
        "jdip_reemplazado_por": texto_jdip,
        "placeholders_sin_valor": _detectar_placeholders_sin_valor(valores_finales),
    }
    return buffer.getvalue(), info


# =========================================================================
# PASO 4 — CONTRATO para "COPROPIETARIOS NO CASADOS QUE VIVEN EN SITIOS
# DIFERENTES". Igual que el Paso 3, solo cambian los 3 párrafos que se
# borran (deja en el documento la variante de copropietarios que viven en
# sitios diferentes), y se agregan 4 placeholders propios de esa variante
# (OOOOOO, PPPPP, ÑÑÑÑÑÑ, XXXXXXX: domicilio del segundo copropietario).
# =========================================================================
PARRAFOS_A_ELIMINAR_COPROPIETARIOS_SITIOS_DIFERENTES = [
    "RRRRRRR, identificado/a con DNI N° NNNNNNN con domicilio para estos efectos en "
    "LLLLLLLLLL, distrito de WWWWW, provincia de EEEEEEE y Departamento de TTTTTTT, En "
    "adelante se le denominará EL COMPRADOR; conforme a las cláusulas siguientes.",
    "RRRRRRR, identificado/a con DNI N° NNNNNNN, casado con KKKKKK identificada/o con DNI N° "
    "JJJJJJJ, ambos con domicilio para estos efectos en LLLLLLLLLLL, distrito de WWWWW, "
    "provincia de EEEEEEE y Departamento de TTTTTTT, En adelante se le denominará EL "
    "COMPRADOR; conforme a las cláusulas siguientes.",
    "RRRRRRR, identificado/a con DNI N° NNNNNNN, como copropietario KKKKKK identificada/o con "
    "DNI N° JJJJJJJ, ambos con domicilio para estos efectos en LLLLLLLLLL, distrito de "
    "WWWWW, provincia de EEEEEEE y Departamento de TTTTTTT, En adelante se le denominará "
    "EL COMPRADOR; conforme a las cláusulas siguientes.",
]


def leer_datos_segundo_copropietario_paso4(archivo_excel):
    """Lee de la hoja 'DATOS' los valores que reemplazan a OOOOOO (E44),
    PPPPP (H43), ÑÑÑÑÑÑ (H44) y XXXXXXX (H45): domicilio del segundo
    copropietario, que solo aplica al Paso 4 ('COPROPIETARIOS NO CASADOS
    QUE VIVEN EN SITIOS DIFERENTES'), para los 4 documentos (CONTRATO,
    REGLAMENTO, MEMBRESIA, DERECHOHABIENTES) -- igual que
    leer_datos_conyuge_paso2. Devuelve {placeholder: valor}."""
    wb = load_workbook(archivo_excel, data_only=True)
    if HOJA_DATOS_WORD not in wb.sheetnames:
        raise ValueError(f"La hoja '{HOJA_DATOS_WORD}' no existe en el Excel subido.")
    sheet_datos = wb[HOJA_DATOS_WORD]
    return {
        "OOOOOO": _lee(sheet_datos, "E44"),
        "PPPPP": _lee(sheet_datos, "H43"),
        "ÑÑÑÑÑÑ": _lee(sheet_datos, "H44"),
        "XXXXXXX": _lee(sheet_datos, "H45"),
    }


def leer_datos_extra_paso4(archivo_excel):
    """Junta los valores extra del Paso 4 para REGLAMENTO/MEMBRESIA/
    DERECHOHABIENTES: los del cónyuge/copropietario (KKKKKK, JJJJJJJ, ver
    leer_datos_conyuge_paso2) más los del segundo copropietario (OOOOOO,
    PPPPP, ÑÑÑÑÑÑ, XXXXXXX, ver leer_datos_segundo_copropietario_paso4)."""
    valores = leer_datos_conyuge_paso2(archivo_excel)
    archivo_excel.seek(0)
    valores.update(leer_datos_segundo_copropietario_paso4(archivo_excel))
    return valores


def procesar_paso4_copropietarios_no_casados_sitios_diferentes(archivo_excel, archivo_contrato, membresia, celdas_axa=None):
    """Paso 4: aplica al documento CONTRATO las instrucciones para la
    opción 'COPROPIETARIOS NO CASADOS QUE VIVEN EN SITIOS DIFERENTES'.
    Idéntico al Paso 3 (ver procesar_paso3_copropietarios_no_casados_mismo_sitio),
    salvo que:
      1. Borra los 3 párrafos de
         PARRAFOS_A_ELIMINAR_COPROPIETARIOS_SITIOS_DIFERENTES en vez de
         los del Paso 3 (deja la variante de copropietarios en sitios
         diferentes).
      2. Además reemplaza OOOOOO, PPPPP, ÑÑÑÑÑÑ y XXXXXXX por los datos
         del segundo copropietario (ver
         leer_datos_segundo_copropietario_paso4).
    También reemplaza "JDIP" según 'membresia' (ver
    _reemplazar_jdip_segun_membresia). Devuelve (bytes_del_docx_resultante, info).
    """
    valores = leer_datos_para_contrato(archivo_excel, celdas_axa=celdas_axa)
    archivo_excel.seek(0)
    valores_conyuge = leer_datos_conyuge_paso2(archivo_excel)
    archivo_excel.seek(0)
    valores_segundo_copropietario = leer_datos_segundo_copropietario_paso4(archivo_excel)

    archivo_contrato.seek(0)
    documento = Document(archivo_contrato)

    parrafos_eliminados = eliminar_parrafos_por_texto(documento, PARRAFOS_A_ELIMINAR_COPROPIETARIOS_SITIOS_DIFERENTES)
    reemplazar_en_documento(documento, valores)
    establecer_color_negro_texto1_en_documento(documento, TEXTOS_COLOR_NEGRO_TEXTO1_CASADOS_MISMO_SITIO)
    reemplazar_en_documento(documento, valores_conyuge)
    reemplazar_en_documento(documento, valores_segundo_copropietario)
    texto_jdip = _reemplazar_jdip_segun_membresia(documento, membresia)

    buffer = io.BytesIO()
    documento.save(buffer)
    buffer.seek(0)

    valores_finales = {**valores, **valores_conyuge, **valores_segundo_copropietario}
    info = {
        "parrafos_eliminados": parrafos_eliminados,
        "valores_usados": valores_finales,
        "jdip_reemplazado_por": texto_jdip,
        "placeholders_sin_valor": _detectar_placeholders_sin_valor(valores_finales),
    }
    return buffer.getvalue(), info


# =========================================================================
# REEMPLAZO DE "JDIP" SEGÚN MEMBRESÍA — el texto literal "JDIP" aparece en
# las plantillas reales de CONTRATO y REGLAMENTO (no en MEMBRESIA ni en
# DERECHOHABIENTES); se reemplaza por "ANEXO F" o "ANEXO E" según la
# opción de membresía elegida, para los 4 Pasos (tipos de propiedad) del
# CONTRATO y para REGLAMENTO por igual.
# =========================================================================
MEMBRESIA_REGLAMENTO_VALORES = {
    "CON MEMBRESÍA": "ANEXO F",
    "SIN MEMBRESÍA": "ANEXO E",
}


def _patron_texto_literal(texto):
    """Patrón regex que coincide EXACTAMENTE con 'texto' (sin comodines,
    a diferencia de _patron_placeholder), para reemplazar una palabra
    fija como 'JDIP' en vez de un placeholder de letras repetidas."""
    return re.compile(re.escape(texto))


def _reemplazar_texto_en_documento(documento, patron_regex, reemplazo):
    """Igual que reemplazar_en_documento, pero para UN solo patrón/valor
    (no un diccionario de placeholders): recorre los párrafos del cuerpo
    principal, las tablas (celda por celda) y los encabezados/pies de
    página, reemplazando cualquier coincidencia de 'patron_regex' por
    'reemplazo'."""
    def _procesar_parrafos(parrafos):
        for paragraph in parrafos:
            _reemplazar_en_parrafo(paragraph, patron_regex, reemplazo)

    _procesar_parrafos(documento.paragraphs)
    for tabla in documento.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                _procesar_parrafos(celda.paragraphs)
    for section in documento.sections:
        _procesar_parrafos(section.header.paragraphs)
        _procesar_parrafos(section.footer.paragraphs)


def _reemplazar_jdip_segun_membresia(documento, membresia):
    """Reemplaza toda aparición de "JDIP" por "ANEXO F" (si 'membresia' es
    "CON MEMBRESÍA") o por "ANEXO E" (si es "SIN MEMBRESÍA"). Devuelve el
    texto usado como reemplazo, para mostrarlo en la interfaz. Se usa
    tanto para el CONTRATO (los 4 Pasos) como para REGLAMENTO."""
    if membresia not in MEMBRESIA_REGLAMENTO_VALORES:
        raise ValueError(
            f"Membresía '{membresia}' no reconocida (debe ser 'CON MEMBRESÍA' o 'SIN MEMBRESÍA')."
        )
    texto_jdip = MEMBRESIA_REGLAMENTO_VALORES[membresia]
    _reemplazar_texto_en_documento(documento, _patron_texto_literal("JDIP"), texto_jdip)
    return texto_jdip


def procesar_reglamento(archivo_excel, archivo_reglamento, membresia, valores_extra=None, celdas_axa=None):
    """
    Aplica al documento REGLAMENTO (mismo Excel de datos que el CONTRATO,
    para cualquiera de los 4 Pasos/tipos de propiedad):
      1. Reemplaza en todo el documento los placeholders RRRRRRR, NNNNNNN,
         LLLLLLLLLL, WWWWW, EEEEEEE, TTTTTTT y UUUUUUU con los MISMOS
         valores que se usan para el CONTRATO (ver leer_datos_para_contrato).
      2. Solo si se pasan 'valores_extra' (Pasos 2, 3 y 4 -- el Paso 1,
         'UNICO PROPIETARIO', nunca pasa valores_extra): pone color de
         fuente "Negro, Texto 1" en los mismos textos que en el CONTRATO
         (ver TEXTOS_COLOR_NEGRO_TEXTO1_CASADOS_MISMO_SITIO) -- ANTES de
         reemplazar KKKKKK/JJJJJJJ, para que la búsqueda de esos textos
         literales todavía los encuentre -- y luego reemplaza
         'valores_extra' (ej. KKKKKK/JJJJJJJ del Paso 2, ver
         leer_datos_conyuge_paso2).
      4. Reemplaza toda aparición de "JDIP" por "ANEXO F" (si 'membresia'
         es "CON MEMBRESÍA") o por "ANEXO E" (si es "SIN MEMBRESÍA").

    Devuelve (bytes_del_docx_resultante, info) donde info trae detalles
    para mostrar en la interfaz.
    """
    valores = leer_datos_para_contrato(archivo_excel, celdas_axa=celdas_axa)

    archivo_reglamento.seek(0)
    documento = Document(archivo_reglamento)

    reemplazar_en_documento(documento, valores)
    if valores_extra:
        establecer_color_negro_texto1_en_documento(documento, TEXTOS_COLOR_NEGRO_TEXTO1_CASADOS_MISMO_SITIO)
        reemplazar_en_documento(documento, valores_extra)
        valores = {**valores, **valores_extra}

    texto_jdip = _reemplazar_jdip_segun_membresia(documento, membresia)

    buffer = io.BytesIO()
    documento.save(buffer)
    buffer.seek(0)

    info = {
        "valores_usados": valores,
        "jdip_reemplazado_por": texto_jdip,
        "placeholders_sin_valor": _detectar_placeholders_sin_valor(valores),
    }
    return buffer.getvalue(), info


# =========================================================================
# APARTADOS "MEMBRESIA" y "DERECHOHABIENTES" — mismos valores de
# placeholders que el CONTRATO (RRRRRRR, NNNNNNN, LLLLLLLLLL, WWWWW,
# EEEEEEE, TTTTTTT, UUUUUUU), sin el reemplazo de "JDIP" (eso es solo
# para REGLAMENTO). Aplica para los 4 Pasos (tipos de propiedad) por
# igual.
# =========================================================================


def _procesar_documento_con_valores_contrato(archivo_excel, archivo_documento, valores_extra=None, celdas_axa=None):
    """Reemplaza en 'archivo_documento' los placeholders RRRRRRR, NNNNNNN,
    LLLLLLLLLL, WWWWW, EEEEEEE, TTTTTTT y UUUUUUU con los mismos valores
    que se usan para el CONTRATO (ver leer_datos_para_contrato), más los
    de 'valores_extra' si se pasan (ej. KKKKKK/JJJJJJJ del Paso 2, ver
    leer_datos_conyuge_paso2). Devuelve (bytes_del_docx_resultante, info)."""
    valores = leer_datos_para_contrato(archivo_excel, celdas_axa=celdas_axa)
    if valores_extra:
        valores = {**valores, **valores_extra}

    archivo_documento.seek(0)
    documento = Document(archivo_documento)

    reemplazar_en_documento(documento, valores)

    buffer = io.BytesIO()
    documento.save(buffer)
    buffer.seek(0)

    info = {
        "valores_usados": valores,
        "placeholders_sin_valor": _detectar_placeholders_sin_valor(valores),
    }
    return buffer.getvalue(), info


def procesar_membresia(archivo_excel, archivo_membresia, numero_membresia, valores_extra=None, celdas_axa=None):
    """Aplica al documento MEMBRESIA:
      1. Los mismos reemplazos de placeholders que el CONTRATO, más los
         de 'valores_extra' si se pasan (ej. KKKKKK/JJJJJJJ del Paso 2,
         ver leer_datos_conyuge_paso2).
      2. Reemplaza toda aparición de "GTAA" por 'numero_membresia' (dato
         ingresado a mano en la interfaz, no viene del Excel).
    Devuelve (bytes_del_docx_resultante, info)."""
    valores = leer_datos_para_contrato(archivo_excel, celdas_axa=celdas_axa)
    if valores_extra:
        valores = {**valores, **valores_extra}

    archivo_membresia.seek(0)
    documento = Document(archivo_membresia)

    reemplazar_en_documento(documento, valores)
    _reemplazar_texto_en_documento(documento, _patron_texto_literal("GTAA"), numero_membresia)

    buffer = io.BytesIO()
    documento.save(buffer)
    buffer.seek(0)

    info = {
        "valores_usados": valores,
        "numero_membresia": numero_membresia,
        "placeholders_sin_valor": _detectar_placeholders_sin_valor(valores),
    }
    return buffer.getvalue(), info


def procesar_derechohabientes(archivo_excel, archivo_derechohabientes, valores_extra=None, celdas_axa=None):
    """Aplica al documento DERECHOHABIENTES los mismos reemplazos de
    placeholders que el CONTRATO, más los de 'valores_extra' si se pasan
    (ej. KKKKKK/JJJJJJJ del Paso 2, ver leer_datos_conyuge_paso2) (ver
    _procesar_documento_con_valores_contrato)."""
    return _procesar_documento_con_valores_contrato(archivo_excel, archivo_derechohabientes, valores_extra, celdas_axa)
