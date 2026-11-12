# -*- coding: utf-8 -*-
"""
excel_pipeline.py — Pasos 1 a 6 sobre los archivos Excel (ORIGEN/MODELO).

Contiene la memoria de aprendizaje, las utilidades de datos, y todas las
funciones de procesamiento de las hojas "DATOS", "Cotizador", "FICHA",
"Ax.A", "Ax.B" y "List.Ax". No depende de Word ni de pywin32.
"""

import os
import io
import json
import re
import threading
import zipfile
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Color
from openpyxl.utils import coordinate_to_tuple

from templates_manager import CELDAS_AXA_POR_DEFECTO

# =========================================================================
# CONFIGURACIÓN
# =========================================================================
ruta_actual = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_MEMORIA = os.path.join(ruta_actual, "memoria_aprendizaje.json")

# Protege a memoria_aprendizaje.json de que dos personas usando la app al
# mismo tiempo (un hilo por persona, mismo proceso de streamlit run)
# pisen/pierdan la escritura de la otra.
_LOCK_MEMORIA = threading.Lock()

# Si en tu caso SÍ quieres que la columna C (etiquetas) de "DATOS" se
# copie también del origen al modelo, pon esto en True. Por defecto en
# False, porque el modelo ya trae sus propias etiquetas y lo que se
# pidió copiar es el VALOR de la columna E.
COPIAR_ETIQUETAS_C = False

FILA_INICIO_DATOS = 11
FILA_FIN_DATOS = 50
CELDAS_TELEFONO = (27, 49)
CELDAS_DNI = (13, 35)

PATRON_FECHA_TEXTO = re.compile(r"^\d{1,4}[/-]\d{1,2}[/-]\d{1,4}$")

# Cualquiera de estos valores, solo o rodeado de espacios, se considera
# "sin dato" (equivalente a una celda vacía) y por lo tanto no dispara
# ninguna alerta de validación.
MARCADORES_VACIOS = {"-", "–", "—", "‐", "‑", "s/n", "n/a", "na"}

# Hoja usada en el Paso 2
HOJA_COTIZADOR = "Cotizador"

# Hoja usada en el Paso 3
HOJA_FICHA = "FICHA"
FILA_AUTOFILL_ORIGEN = 64
COLUMNAS_AUTOFILL = ["C", "D", "E", "F"]

# Hoja usada en el Paso 4
HOJA_AXA = "Ax.A"

# Hoja "Px": se copia ENTERA del ORIGEN al MODELO siempre, sin depender
# de qué Pasos estén activados (ver procesar_copia_hoja_px).
HOJA_PX = "Px"


# =========================================================================
# MEMORIA DE APRENDIZAJE (reglas tipo "ignorar teléfono", "ignorar DNI")
# =========================================================================
def cargar_memoria():
    if os.path.exists(ARCHIVO_MEMORIA):
        try:
            with open(ARCHIVO_MEMORIA, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "reglas_adicionales" in data:
                    return data
        except (json.JSONDecodeError, OSError):
            pass
    return {"reglas_adicionales": []}


def guardar_en_memoria(nueva_regla):
    with _LOCK_MEMORIA:
        memoria = cargar_memoria()
        if nueva_regla not in memoria["reglas_adicionales"]:
            memoria["reglas_adicionales"].append(nueva_regla)
            with open(ARCHIVO_MEMORIA, "w", encoding="utf-8") as f:
                json.dump(memoria, f, indent=4, ensure_ascii=False)


def borrar_memoria():
    with _LOCK_MEMORIA:
        with open(ARCHIVO_MEMORIA, "w", encoding="utf-8") as f:
            json.dump({"reglas_adicionales": []}, f, indent=4, ensure_ascii=False)


# =========================================================================
# UTILIDADES DE DATOS
# =========================================================================
def normalizar_texto(valor):
    if valor is None:
        return ""
    return str(valor).strip().upper()


def es_marcador_vacio(valor):
    """True si la celda está vacía o solo contiene un placeholder tipo
    guion ('-', '–', '—') o similar ('s/n', 'n/a'). Estas celdas se
    tratan como 'sin dato' y no deben generar alertas de validación."""
    if valor is None:
        return True
    texto = str(valor).strip().lower()
    return texto == "" or texto in MARCADORES_VACIOS


def celda_vacia(valor):
    """True si la celda no tiene contenido real (None o solo espacios)."""
    if valor is None:
        return True
    return str(valor).strip() == ""


def es_fecha(valor):
    """Detección robusta: primero por tipo real (openpyxl ya convierte
    automáticamente las celdas con formato de fecha a datetime/date),
    y como respaldo, un texto con patrón AAAA-MM-DD o DD/MM/AAAA."""
    if isinstance(valor, (datetime, date)):
        return True
    if isinstance(valor, str) and PATRON_FECHA_TEXTO.match(valor.strip()):
        return True
    return False


def formatear_fecha(valor):
    """Convierte un valor de fecha (datetime/date o texto) al formato DD/MM/AA."""
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%d/%m/%y")
    if isinstance(valor, str):
        texto = valor.strip()
        for formato in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y"):
            try:
                return datetime.strptime(texto, formato).strftime("%d/%m/%y")
            except ValueError:
                continue
    return valor


def solo_digitos(valor):
    return "".join(filter(str.isdigit, str(valor))) if valor is not None else ""


# =========================================================================
# CONVERSIÓN DE MONTOS A LETRAS (formato SOLES) — usado en el Paso 4
# =========================================================================
_UNIDADES_LETRAS = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE",
                     "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE",
                     "DIECIOCHO", "DIECINUEVE", "VEINTE"]
_VEINTES_ESPECIALES = {
    21: "VEINTIUNO", 22: "VEINTIDOS", 23: "VEINTITRES", 24: "VEINTICUATRO",
    25: "VEINTICINCO", 26: "VEINTISEIS", 27: "VEINTISIETE", 28: "VEINTIOCHO", 29: "VEINTINUEVE",
}
_DECENAS_LETRAS = ["", "", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"]
_CENTENAS_LETRAS = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS",
                     "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"]


def _apocopar_uno(texto):
    """Aplica la apócope de 'UNO' -> 'UN' (y 'VEINTIUNO' -> 'VEINTIUN') al
    último término de 'texto', tal como exige la gramática española antes
    de un sustantivo como 'MIL' o 'MILLONES' (ej. 'VEINTIUN MIL', no
    'VEINTIUNO MIL')."""
    palabras = texto.split(" ")
    if palabras[-1].endswith("UNO"):
        palabras[-1] = palabras[-1][:-1]
    return " ".join(palabras)


def _entero_a_letras(n):
    """Convierte un entero no negativo a su representación en letras,
    en español y en MAYÚSCULAS (sin tildes, para mantener consistencia
    con el resto del documento)."""
    if n == 0:
        return "CERO"
    if n < 0:
        return "MENOS " + _entero_a_letras(-n)
    if n < 21:
        return _UNIDADES_LETRAS[n]
    if n < 30:
        if n in _VEINTES_ESPECIALES:
            return _VEINTES_ESPECIALES[n]
        return _UNIDADES_LETRAS[n]
    if n < 100:
        decena, unidad = divmod(n, 10)
        if unidad == 0:
            return _DECENAS_LETRAS[decena]
        return f"{_DECENAS_LETRAS[decena]} Y {_UNIDADES_LETRAS[unidad]}"
    if n < 1000:
        if n == 100:
            return "CIEN"
        centena, resto = divmod(n, 100)
        texto = _CENTENAS_LETRAS[centena]
        if resto:
            texto += " " + _entero_a_letras(resto)
        return texto
    if n < 1_000_000:
        miles, resto = divmod(n, 1000)
        if miles == 1:
            texto = "MIL"
        else:
            texto = f"{_apocopar_uno(_entero_a_letras(miles))} MIL"
        if resto:
            texto += " " + _entero_a_letras(resto)
        return texto
    millones, resto = divmod(n, 1_000_000)
    if millones == 1:
        texto = "UN MILLON"
    else:
        texto = f"{_apocopar_uno(_entero_a_letras(millones))} MILLONES"
    if resto:
        texto += " " + _entero_a_letras(resto)
    return texto


# Corrige un problema real de cálculo de Excel: cuando el monto "debería"
# ser un número entero redondo (ej. 65000.00), algunas fórmulas de la
# hoja lo dejan a 1 centavo de distancia (ej. 64999.99) por acarreo de
# redondeos intermedios. Si el monto queda a este umbral o menos del
# entero más cercano, se usa ese entero en su lugar -- a propósito NO se
# redondea así CUALQUIER monto (ej. 50000.60 sigue mostrando "CON 60/100"
# tal cual), solo el caso puntual de quedar pegado a un número redondo.
UMBRAL_CORRECCION_MONTO_REDONDO = Decimal("0.01")


def formatear_monto_letras(valor, moneda="SOLES"):
    """Convierte un monto numérico (o texto numérico) al formato:
    '(CINCUENTA Y OCHO MIL SEISCIENTOS CON 00/100 SOLES)'
    Devuelve None si 'valor' no se pudo interpretar como número."""
    if valor is None:
        return None
    if isinstance(valor, str):
        texto_limpio = re.sub(r"[^0-9.,-]", "", valor.strip())
        texto_limpio = texto_limpio.replace(",", "")
        try:
            valor = float(texto_limpio)
        except ValueError:
            return None
    try:
        valor = float(valor)
    except (TypeError, ValueError):
        return None

    # Se usa Decimal (a partir del texto de 'valor', no del float directo)
    # para redondear al segundo decimal, en vez de round((valor - entero) * 100)
    # con aritmética de punto flotante: un valor como 15320.505 en realidad se
    # guarda en binario como algo como 15320.504999999997, así que round()
    # sobre el float redondeaba mal el segundo decimal (daba 50 en vez de 51).
    # Decimal(str(valor)) toma el decimal "tal cual se ve" y lo redondea de
    # forma exacta (ROUND_HALF_UP), incluyendo el acarreo a la parte entera
    # cuando corresponde (ej. 999999.995 -> 1000000.00).
    try:
        valor_decimal = Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return None

    # Ver nota junto a UMBRAL_CORRECCION_MONTO_REDONDO, arriba: si el monto
    # quedó a 1 centavo o menos del entero más cercano, se lo trata como
    # ese entero (corrige el acarreo de redondeo de Excel), no un
    # redondeo general de cualquier monto a centavos.
    entero_mas_cercano = valor_decimal.to_integral_value(rounding=ROUND_HALF_UP)
    if abs(valor_decimal - entero_mas_cercano) <= UMBRAL_CORRECCION_MONTO_REDONDO:
        valor_decimal = Decimal(entero_mas_cercano)

    entero = int(valor_decimal)
    centavos = int((valor_decimal - entero) * 100)

    texto_entero = _entero_a_letras(entero)
    return f"({texto_entero} CON {centavos:02d}/100 {moneda})"


# =========================================================================
# LECTURA/ESCRITURA SEGURA EN CELDAS COMBINADAS
# =========================================================================
# En openpyxl, cuando varias celdas están combinadas ("merged"), solo la
# celda de arriba-a-la-izquierda (la "ancla") guarda el valor real; las
# demás celdas del rango son objetos MergedCell cuyo atributo .value es
# de solo lectura (siempre None) y lanza AttributeError si se intenta
# escribir en ellas directamente. Estas dos funciones detectan si una
# coordenada cae dentro de un rango combinado y, si es así, leen/escriben
# siempre en la celda ancla, evitando el error
# "MergedCell object attribute 'value' is read-only".
def _celda_ancla(sheet, coord):
    fila, columna = coordinate_to_tuple(coord)
    for rango in sheet.merged_cells.ranges:
        if rango.min_row <= fila <= rango.max_row and rango.min_col <= columna <= rango.max_col:
            return sheet.cell(row=rango.min_row, column=rango.min_col)
    return sheet[coord]


def _lee(sheet, coord):
    """Lee el valor de una celda, resolviendo celdas combinadas."""
    return _celda_ancla(sheet, coord).value


def _escribe(sheet, coord, valor):
    """Escribe el valor en una celda, resolviendo celdas combinadas."""
    _celda_ancla(sheet, coord).value = valor


# =========================================================================
# PASO 1 — HOJA "DATOS"  (opera sobre workbooks ya abiertos)
# =========================================================================
# Si E33 de "DATOS" (segundo titular/copropietario) trae un dato real, en
# las hojas de HOJAS_ESCANEO_SEGUNDO_TITULAR (FICHA, Ax.A, Ax.B, Ax.E) se
# busca este texto EXACTO. A la celda de ARRIBA (misma columna, una fila
# antes) se le pone color de fuente Negro; a las 2 celdas de ABAJO (misma
# columna, una y dos filas después) se les pone color de fuente
# Automático -- no a la celda con el texto en sí.
TEXTO_SEGUNDO_TITULAR_COPROPIETARIO = "SEGUNDO TITULAR/COPROPIETARIO"
# HOJAS_ESCANEO_SEGUNDO_TITULAR se define más abajo, después de HOJA_AXB (Paso 5).

# Tope de columnas a recorrer por hoja al buscar
# TEXTO_SEGUNDO_TITULAR_COPROPIETARIO. "Ax.A" reporta su rango usado
# inflado hasta la columna XFD (16384) por formato aplicado a filas
# enteras, aunque el dato real nunca pasa de la columna AA (27). 60
# columnas da margen de sobra sobre cualquier hoja real y evita recorrer
# millones de celdas de más.
LIMITE_COLUMNAS_ESCANEO_CONYUGE = 60


def _celda_ancla_por_posicion(sheet, fila, columna):
    """Igual que _celda_ancla, pero a partir de (fila, columna) numéricos
    en vez de una coordenada tipo 'A1' -- resuelve a la celda ancla si esa
    posición cae dentro de un rango de celdas combinadas (necesario porque
    escribir en el atributo .font de una MergedCell que no es la ancla
    lanza AttributeError)."""
    for rango in sheet.merged_cells.ranges:
        if rango.min_row <= fila <= rango.max_row and rango.min_col <= columna <= rango.max_col:
            return sheet.cell(row=rango.min_row, column=rango.min_col)
    return sheet.cell(row=fila, column=columna)


# Referencia simple a UNA celda de otra hoja del mismo libro, ej.
# "=Ax.A!$I$145" o "=+Ax.A!I145" -- algunos proyectos (ej. COSTA DEL
# LAGO) no repiten el texto de TEXTO_SEGUNDO_TITULAR_COPROPIETARIO en
# cada hoja: en vez de eso, la celda tiene esta fórmula apuntando a la
# hoja que sí tiene el texto literal (típicamente "Ax.A"). Ver
# _valor_mostrado_en_celda, abajo.
_PATRON_REFERENCIA_SIMPLE_A_OTRA_CELDA = re.compile(
    r"^=\+?'?([A-Za-z0-9_. ]+)'?!\$?([A-Za-z]{1,3})\$?(\d+)$"
)


def _valor_mostrado_en_celda(wb_modelo, celda):
    """Devuelve el texto que Excel mostraría en 'celda' -- igual a
    'celda.value' si es un valor literal, pero si es una fórmula que
    apunta a UNA sola celda de otra hoja del mismo libro (ver patrón
    arriba), resuelve esa referencia y devuelve el valor literal de la
    celda referenciada en su lugar (solo un nivel: si la celda
    referenciada es a su vez otra fórmula, se devuelve tal cual, sin
    seguir encadenando). Necesario porque wb_modelo se carga con
    data_only=False (hace falta para poder escribirlo después), así que
    una fórmula nunca trae el valor ya calculado por Excel."""
    valor = celda.value
    if not (isinstance(valor, str) and valor.startswith("=")):
        return valor
    coincidencia = _PATRON_REFERENCIA_SIMPLE_A_OTRA_CELDA.match(valor.strip())
    if not coincidencia:
        return valor
    nombre_hoja_ref, columna_ref, fila_ref = coincidencia.groups()
    if nombre_hoja_ref not in wb_modelo.sheetnames:
        return valor
    return _lee(wb_modelo[nombre_hoja_ref], f"{columna_ref}{fila_ref}")


def _formatear_celdas_segundo_titular_copropietario(wb_modelo):
    """Busca en las hojas de HOJAS_ESCANEO_SEGUNDO_TITULAR (FICHA, Ax.A,
    Ax.B, Ax.E; hasta LIMITE_COLUMNAS_ESCANEO_CONYUGE columnas por hoja,
    ver comentario arriba) las celdas cuyo valor (ya resuelto, ver
    _valor_mostrado_en_celda) sea exactamente
    TEXTO_SEGUNDO_TITULAR_COPROPIETARIO. A la celda de ARRIBA (misma
    columna, una fila antes) le pone color de fuente Negro; a la celda
    de la ETIQUETA misma y a las DOS celdas de ABAJO (misma columna, una
    y dos filas después) les pone color de fuente Automático -- todas
    resolviendo a la celda ancla si son parte de un rango combinado.
    Devuelve cuántas coincidencias del texto se encontraron (sin importar
    cuántas celdas alrededor de cada una se hayan formateado)."""
    from copy import copy
    coincidencias_formateadas = 0
    for nombre_hoja in HOJAS_ESCANEO_SEGUNDO_TITULAR:
        if nombre_hoja not in wb_modelo.sheetnames:
            continue
        sheet = wb_modelo[nombre_hoja]
        max_col = min(sheet.max_column, LIMITE_COLUMNAS_ESCANEO_CONYUGE)
        for fila in sheet.iter_rows(max_col=max_col):
            for celda in fila:
                valor = _valor_mostrado_en_celda(wb_modelo, celda)
                if isinstance(valor, str) and valor.strip() == TEXTO_SEGUNDO_TITULAR_COPROPIETARIO:
                    if celda.row > 1:
                        celda_arriba = _celda_ancla_por_posicion(sheet, celda.row - 1, celda.column)
                        nueva_font_arriba = copy(celda_arriba.font)
                        nueva_font_arriba.color = Color(rgb="FF000000")
                        celda_arriba.font = nueva_font_arriba

                    celda_etiqueta = _celda_ancla_por_posicion(sheet, celda.row, celda.column)
                    nueva_font_etiqueta = copy(celda_etiqueta.font)
                    nueva_font_etiqueta.color = Color(auto=True)
                    celda_etiqueta.font = nueva_font_etiqueta

                    for desplazamiento in (1, 2):
                        celda_debajo = _celda_ancla_por_posicion(
                            sheet, celda.row + desplazamiento, celda.column
                        )
                        nueva_font_debajo = copy(celda_debajo.font)
                        nueva_font_debajo.color = Color(auto=True)
                        celda_debajo.font = nueva_font_debajo

                    coincidencias_formateadas += 1
    return coincidencias_formateadas


def _procesar_paso1_datos_wb(wb_origen, wb_modelo, reglas_memoria, alertas, resumen):
    color_blanco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    # HOJAS PROTEGIDAS (nunca se abren ni se tocan):
    # "FICHA DE DATOS", "PEP", "UIF", "Px", "C.List", "AV", "DJ", "CROQUIS", "Courier"

    if "DATOS" not in wb_origen.sheetnames:
        alertas.append("[DATOS] La hoja 'DATOS' no existe en el archivo de ORIGEN.")
        return
    if "DATOS" not in wb_modelo.sheetnames:
        alertas.append("[DATOS] La hoja 'DATOS' no existe en el archivo MODELO.")
        return

    sheet_ori = wb_origen["DATOS"]
    sheet_mod = wb_modelo["DATOS"]

    for fila in range(FILA_INICIO_DATOS, FILA_FIN_DATOS + 1):
        if COPIAR_ETIQUETAS_C:
            _escribe(sheet_mod, f"C{fila}", _lee(sheet_ori, f"C{fila}"))

        etiqueta_origen = normalizar_texto(_lee(sheet_ori, f"C{fila}"))
        etiqueta_modelo = normalizar_texto(_lee(sheet_mod, f"C{fila}"))
        valor_e = _lee(sheet_ori, f"E{fila}")

        # --- ¿La fila está desalineada? (nombre/dato en la celda equivocada) ---
        if etiqueta_origen and etiqueta_modelo and etiqueta_origen != etiqueta_modelo:
            alertas.append(
                f"[DATOS] Fila {fila}: la etiqueta de origen ('{_lee(sheet_ori, f'C{fila}')}') "
                f"no coincide con la del modelo ('{_lee(sheet_mod, f'C{fila}')}'). "
                f"Verifica que el dato no esté desplazado de fila."
            )

        # --- Guion / placeholder de "sin dato": se copia tal cual y no valida nada ---
        if es_marcador_vacio(valor_e):
            valor_final = normalizar_texto(valor_e) if valor_e is not None else ""
            _escribe(sheet_mod, f"E{fila}", valor_final)
            sheet_mod[f"E{fila}"].fill = color_blanco
            resumen["filas_procesadas"] += 1
            if valor_e is not None:
                resumen["filas_con_guion_ignoradas"] += 1
            continue  # sin más validaciones para esta celda

        # --- Formateo normal: fecha o mayúsculas ---
        if es_fecha(valor_e):
            valor_final = formatear_fecha(valor_e)
        else:
            valor_final = normalizar_texto(valor_e)

        _escribe(sheet_mod, f"E{fila}", valor_final)
        sheet_mod[f"E{fila}"].fill = color_blanco
        resumen["filas_procesadas"] += 1

        # --- Validación: nombre/razón social con números ---
        if ("NOMBRE" in etiqueta_origen or "APELLIDO" in etiqueta_origen or "RAZÓN SOCIAL" in etiqueta_origen) \
                and any(char.isdigit() for char in str(valor_final)):
            alertas.append(f"[DATOS] Fila {fila}: el nombre o razón social contiene números: '{valor_e}'.")

        # --- Validación: correo electrónico ---
        if "CORREO" in etiqueta_origen and "@" not in str(valor_final):
            alertas.append(f"[DATOS] Fila {fila}: el correo electrónico '{valor_e}' no parece válido (falta '@').")

        # --- Validación: teléfono (9 dígitos) ---
        if fila in CELDAS_TELEFONO:
            digitos = solo_digitos(valor_final)
            if len(digitos) != 9 and not any("IGNORAR TELEFONO" in r.upper() for r in reglas_memoria):
                alertas.append(f"[DATOS] Celda E{fila}: el teléfono '{valor_e}' no tiene exactamente 9 dígitos.")

        # --- Validación: DNI (8 dígitos) ---
        if fila in CELDAS_DNI:
            digitos = solo_digitos(valor_final)
            if len(digitos) != 8 and not any("IGNORAR DNI" in r.upper() for r in reglas_memoria):
                alertas.append(f"[DATOS] Celda E{fila}: el DNI '{valor_e}' no tiene exactamente 8 dígitos.")

    # --- Formato condicional de "SEGUNDO TITULAR/COPROPIETARIO" según E33 ---
    # Si E33 (ORIGEN) tiene un dato real (no es un marcador de "sin dato"
    # como "-"), hay segundo titular/copropietario: se les pone color de
    # fuente Automático a las 2 celdas debajo de esa etiqueta, en las
    # hojas del MODELO que la tengan (ver
    # _formatear_celdas_segundo_titular_copropietario).
    valor_e33 = _lee(sheet_ori, "E33")
    if not es_marcador_vacio(valor_e33):
        resumen["segundo_titular_celdas_formateadas"] = _formatear_celdas_segundo_titular_copropietario(wb_modelo)


def procesar_paso1_datos(archivo_origen, archivo_modelo, reglas_memoria):
    """Versión standalone del Paso 1 (se mantiene por compatibilidad,
    por si se quiere seguir corriendo solo este paso)."""
    alertas = []
    resumen = {"filas_procesadas": 0, "filas_con_guion_ignoradas": 0}

    wb_origen = load_workbook(archivo_origen, data_only=True)
    archivo_modelo.seek(0)
    wb_modelo = load_workbook(archivo_modelo, data_only=False)

    _procesar_paso1_datos_wb(wb_origen, wb_modelo, reglas_memoria, alertas, resumen)

    return _guardar_bytes(wb_modelo), alertas, resumen


# =========================================================================
# PASO 2 — HOJA "Cotizador"  (opera sobre workbooks ya abiertos)
# =========================================================================
def _copiar_celdas_combinadas(sheet_ori, sheet_mod):
    """Replica en el modelo los rangos de celdas combinadas que existan
    en la hoja de origen. Se hace ANTES de copiar cualquier valor,
    tal como se pidió."""
    for rango in list(sheet_ori.merged_cells.ranges):
        rango_txt = str(rango)
        ya_combinada = any(rango_txt == str(r) for r in sheet_mod.merged_cells.ranges)
        if not ya_combinada:
            try:
                sheet_mod.merge_cells(rango_txt)
            except Exception:
                # Si el rango se solapa con una combinación ya existente en
                # el modelo, se ignora en vez de detener todo el proceso.
                pass


def procesar_paso2_cotizador(wb_origen, wb_modelo, alertas, resumen, fecha_manual=None):
    """
    fecha_manual: si se indica (string DD/MM/AA o un objeto date/datetime),
    se usa en vez de la fórmula "=HOY()" para la celda I16.

    Devuelve True si el Paso 2 se completó SIN alertas (condición para
    copiar al final los rangos E35:I35 y E37:I37).
    """
    alertas_previas = len(alertas)

    if HOJA_COTIZADOR not in wb_origen.sheetnames:
        alertas.append(f"[COTIZADOR] La hoja '{HOJA_COTIZADOR}' no existe en el archivo de ORIGEN.")
        return False
    if HOJA_COTIZADOR not in wb_modelo.sheetnames:
        alertas.append(f"[COTIZADOR] La hoja '{HOJA_COTIZADOR}' no existe en el archivo MODELO.")
        return False

    sheet_ori = wb_origen[HOJA_COTIZADOR]
    sheet_mod = wb_modelo[HOJA_COTIZADOR]

    # --- 1) Celdas combinadas primero, antes de copiar cualquier dato ---
    _copiar_celdas_combinadas(sheet_ori, sheet_mod)

    # --- 2) Copiar E15 en mayúsculas -- esta es la celda del LOTE, que
    #          también usa historial_contratos.extraer_lote_y_nombre para
    #          nombrar la carpeta/archivos del historial de contratos
    #          (no hace falta copiarla aparte para eso, ya se copia acá). ---
    valor_e15 = _lee(sheet_ori, "E15")
    _escribe(sheet_mod, "E15", normalizar_texto(valor_e15))
    resumen["cotizador_e15_copiado"] = True

    # --- 3) Verificar (sin alertar) que E16 == I20 en el ORIGEN y en el MODELO.
    #         Esta comparación NO copia E16 a ningún lado y NO genera alertas;
    #         solo se usa internamente para decidir si se escribe I16 (paso 4). ---
    e16_ori, i20_ori = _lee(sheet_ori, "E16"), _lee(sheet_ori, "I20")
    e16_mod, i20_mod = _lee(sheet_mod, "E16"), _lee(sheet_mod, "I20")

    ok_origen = normalizar_texto(e16_ori) == normalizar_texto(i20_ori)
    ok_modelo = normalizar_texto(e16_mod) == normalizar_texto(i20_mod)

    verificacion_e16_i20_ok = ok_origen and ok_modelo

    # --- 4) I16: si se indicó una fecha MANUAL, se escribe siempre (el
    #         usuario la eligió a propósito en la interfaz, así que no debe
    #         depender de la verificación E16==I20). El "=HOY()" automático
    #         sigue dependiendo de esa verificación, como antes. ---
    if fecha_manual:
        if isinstance(fecha_manual, (datetime, date)):
            _escribe(sheet_mod, "I16", fecha_manual.strftime("%d/%m/%y"))
        else:
            _escribe(sheet_mod, "I16", str(fecha_manual))
        resumen["cotizador_i16_escrito"] = _lee(sheet_mod, "I16")
    elif verificacion_e16_i20_ok:
        _escribe(sheet_mod, "I16", "=HOY()")
        resumen["cotizador_i16_escrito"] = _lee(sheet_mod, "I16")

    # --- 5) Copiar I21 ---
    _escribe(sheet_mod, "I21", _lee(sheet_ori, "I21"))

    # --- 6) Copiar C26:C29, omitiendo celdas vacías en el origen ---
    for fila in range(26, 30):
        valor_c = _lee(sheet_ori, f"C{fila}")
        if not celda_vacia(valor_c):
            _escribe(sheet_mod, f"C{fila}", valor_c)

    # --- 7) Copiar D25:D29 como fecha DD/MM/AA, alertar si D vacío pero C con dato,
    #         y alinear a la derecha ---
    alineacion_derecha = Alignment(horizontal="right")
    for fila in range(25, 30):
        valor_d = _lee(sheet_ori, f"D{fila}")
        valor_c_misma_fila = _lee(sheet_ori, f"C{fila}") if fila >= 26 else None

        if celda_vacia(valor_d):
            if fila >= 26 and not celda_vacia(valor_c_misma_fila):
                alertas.append(
                    f"[COTIZADOR] Fila {fila}: la celda C{fila} tiene información "
                    f"('{valor_c_misma_fila}') pero D{fila} está vacía."
                )
            _escribe(sheet_mod, f"D{fila}", valor_d)
        else:
            if es_fecha(valor_d):
                _escribe(sheet_mod, f"D{fila}", formatear_fecha(valor_d))
            else:
                _escribe(sheet_mod, f"D{fila}", valor_d)

        sheet_mod[f"D{fila}"].alignment = alineacion_derecha

    # --- 8) Copiar E25:E29, alertar si E vacío pero C o D con dato ---
    for fila in range(25, 30):
        valor_e = _lee(sheet_ori, f"E{fila}")
        valor_c_misma_fila = _lee(sheet_ori, f"C{fila}") if fila >= 26 else None
        valor_d_misma_fila = _lee(sheet_ori, f"D{fila}")

        if celda_vacia(valor_e):
            hay_c = not celda_vacia(valor_c_misma_fila)
            hay_d = not celda_vacia(valor_d_misma_fila)
            if hay_c or hay_d:
                alertas.append(
                    f"[COTIZADOR] Fila {fila}: falta un dato en E{fila} "
                    f"(C{fila}='{valor_c_misma_fila}', D{fila}='{valor_d_misma_fila}')."
                )
        _escribe(sheet_mod, f"E{fila}", valor_e)

    # --- 8.5) Copiar I25 siempre (sin condiciones) ---
    _escribe(sheet_mod, "I25", _lee(sheet_ori, "I25"))

    # --- 8.6) Copiar E35 (nombre del asesor) siempre, sin condiciones --
    #          a diferencia del resto de la fila 35 (F35:I35), que solo se
    #          copia más abajo si el Paso 2 no tuvo alertas. ---
    _escribe(sheet_mod, "E35", _lee(sheet_ori, "E35"))
    resumen["cotizador_e35_copiado"] = True

    # --- 9) Copiar D33 del ORIGEN al MODELO tal cual (con formato de fecha si es fecha) ---
    valor_d33 = _lee(sheet_ori, "D33")
    if es_fecha(valor_d33):
        _escribe(sheet_mod, "D33", formatear_fecha(valor_d33))
    else:
        _escribe(sheet_mod, "D33", valor_d33)

    # --- 10) I33 del modelo (tras las modificaciones) debe igualar I33 del origen ---
    valor_i33_ori = _lee(sheet_ori, "I33")
    valor_i33_mod = _lee(sheet_mod, "I33")
    i33_mod_es_formula = isinstance(valor_i33_mod, str) and valor_i33_mod.strip().startswith("=")

    if i33_mod_es_formula:
        alertas.append(
            "[COTIZADOR] I33 del modelo contiene una fórmula; abre el archivo en Excel para que "
            "recalcule y así poder confirmar que coincide con I33 del origen "
            f"('{valor_i33_ori}')."
        )
    elif normalizar_texto(valor_i33_mod) != normalizar_texto(valor_i33_ori):
        alertas.append(
            f"[COTIZADOR] I33 del modelo ('{valor_i33_mod}') no coincide con I33 del origen "
            f"('{valor_i33_ori}')."
        )

    # --- 11) Si todo salió bien (sin alertas nuevas en este paso), copiar E35:I35 y E37:I37 ---
    hubo_alertas_en_paso2 = len(alertas) > alertas_previas
    if not hubo_alertas_en_paso2:
        for fila in (35, 37):
            for col in ("E", "F", "G", "H", "I"):
                celda = f"{col}{fila}"
                _escribe(sheet_mod, celda, _lee(sheet_ori, celda))
        resumen["cotizador_rangos_finales_copiados"] = True
    else:
        resumen["cotizador_rangos_finales_copiados"] = False
        alertas.append("[COTIZADOR] Poner asesor a mano")

    return not hubo_alertas_en_paso2


# =========================================================================
# PASO 3 — HOJA "FICHA"  (emulación de AutoFill con openpyxl)
# =========================================================================
# IMPORTANTE (leer antes de tocar este bloque):
# openpyxl NO tiene el método Range.AutoFill de Excel (ese método solo
# existe dentro de la aplicación real de Excel vía COM/win32). Lo que
# hacemos aquí es una EMULACIÓN manual que reproduce lo más fielmente
# posible lo que Excel hace al arrastrar el controlador de relleno hacia
# abajo:
#   - Fórmulas: se copian ajustando las referencias de FILA relativas
#     (una referencia con "$" antes del número de fila, ej. C$64, se
#     queda fija; una referencia normal, ej. C64, se desplaza).
#   - Formato (fuente, bordes, relleno, alineación, formato numérico,
#     protección): se copia tal cual de la fila origen a cada fila nueva.
#   - Celdas combinadas: se detecta qué bloques de columnas están
#     combinados en la fila origen (dentro de C:F) y se repite el mismo
#     patrón de combinación en cada fila nueva.
#   - Validación de datos y formato condicional: se intenta extender los
#     rangos existentes que cubren la fila origen para que también
#     cubran las filas nuevas.
# Limitaciones conocidas de esta emulación (a diferencia del AutoFill
# real de Excel):
#   - No hay detección de "series" (1,2,3 -> 4,5,6 o Enero -> Febrero);
#     los valores constantes (no fórmula) se copian tal cual.
#   - El ajuste de fórmulas es un regex sobre referencias de celda; no
#     entiende fórmulas 3D, referencias a rangos con nombre definido, ni
#     texto entre comillas que contenga algo parecido a una celda (esas
#     partes se dejan intactas a propósito).
#   - openpyxl no reescribe automáticamente fórmulas EN OTRAS partes del
#     libro que apunten a filas debajo de la 64; solo ajusta lo que
#     nosotros copiamos explícitamente en C64:F64.
# =========================================================================

_PATRON_REF_CELDA = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})")


def _contar_no_vacias(sheet, coords):
    return sum(1 for c in coords if not celda_vacia(_lee(sheet, c)))


def _ajustar_formula_relativa(formula, offset_filas):
    """Desplaza las referencias de fila NO absolutas (sin '$' antes del
    número de fila) en 'offset_filas'. Las referencias con '$' antes de
    la fila (ej. C$64) se mantienen fijas, igual que en Excel. Se
    respeta el texto entre comillas dobles para no tocar cadenas."""
    partes = formula.split('"')
    for i in range(0, len(partes), 2):  # solo los tramos FUERA de comillas

        def _reemplazar(m):
            col_abs, col, fila_abs, fila = m.groups()
            if fila_abs == "$":
                return m.group(0)
            nueva_fila = int(fila) + offset_filas
            return f"{col_abs}{col}{fila_abs}{nueva_fila}"

        partes[i] = _PATRON_REF_CELDA.sub(_reemplazar, partes[i])
    return '"'.join(partes)


def _copiar_estilo_celda(celda_origen, celda_destino):
    from copy import copy
    celda_destino.font = copy(celda_origen.font)
    celda_destino.border = copy(celda_origen.border)
    celda_destino.fill = copy(celda_origen.fill)
    celda_destino.alignment = copy(celda_origen.alignment)
    celda_destino.number_format = celda_origen.number_format
    celda_destino.protection = copy(celda_origen.protection)


def _bloques_combinados_en_fila(sheet, fila, columnas):
    """Devuelve los bloques de celdas combinadas que existen en 'fila'
    y caen dentro de la lista de letras de columna indicada, como pares
    (col_letra_inicio, col_letra_fin). Ej.: si C64:D64 está combinada,
    devuelve [('C', 'D')]."""
    from openpyxl.utils import column_index_from_string, get_column_letter

    col_min = min(column_index_from_string(c) for c in columnas)
    col_max = max(column_index_from_string(c) for c in columnas)
    bloques = []
    for rng in sheet.merged_cells.ranges:
        if rng.min_row == fila and rng.max_row == fila and rng.min_col >= col_min and rng.max_col <= col_max:
            bloques.append((get_column_letter(rng.min_col), get_column_letter(rng.max_col)))
    return bloques


def _extender_validaciones(sheet, fila_origen, filas_nuevas, columnas):
    """Intenta extender cada validación de datos que cubra alguna celda
    de 'columnas' en 'fila_origen' para que también cubra esas mismas
    columnas en cada una de las 'filas_nuevas'."""
    for dv in list(sheet.data_validations.dataValidation):
        celdas_origen_cubiertas = [c for c in columnas if f"{c}{fila_origen}" in dv.sqref]
        if not celdas_origen_cubiertas:
            continue
        for fila_nueva in filas_nuevas:
            for c in celdas_origen_cubiertas:
                dv.add(f"{c}{fila_nueva}")


def _extender_formato_condicional(sheet, fila_origen, filas_nuevas, columnas):
    """Intenta extender cada regla de formato condicional que cubra
    alguna celda de 'columnas' en 'fila_origen' para que también cubra
    esas columnas en cada una de las 'filas_nuevas'. Es un 'best effort':
    si el rango interno de openpyxl no permite extenderse limpiamente,
    se ignora esa regla en particular en vez de detener el proceso."""
    try:
        for cf in list(sheet.conditional_formatting):
            celdas_origen_cubiertas = [c for c in columnas if f"{c}{fila_origen}" in cf.sqref]
            if not celdas_origen_cubiertas:
                continue
            nuevos_rangos = " ".join(
                f"{c}{fila_nueva}" for fila_nueva in filas_nuevas for c in celdas_origen_cubiertas
            )
            try:
                cf.sqref = f"{cf.sqref} {nuevos_rangos}"
            except Exception:
                pass
    except Exception:
        pass


def _insertar_filas_preservando_combinadas(sheet, fila_insercion, cantidad, fila_referencia_altura=None):
    """Inserta 'cantidad' filas en 'fila_insercion' sin romper/dividir las
    celdas que ya estaban combinadas en la hoja. openpyxl no siempre
    reubica correctamente los rangos combinados al usar insert_rows()
    (algunas combinaciones existentes pueden quedar mal formadas o
    'partidas'), así que aquí se descombinan todas ANTES de insertar,
    se hace la inserción, y luego se vuelven a combinar en su posición
    correcta: las que estaban en filas >= fila_insercion se desplazan
    'cantidad' filas hacia abajo; las que estaban antes de ese punto
    (como la propia fila 64) se restauran tal cual, sin moverse.

    Si se indica 'fila_referencia_altura', además se preserva el ALTO de
    fila (row height): openpyxl no reubica el alto de fila al usar
    insert_rows() (solo mueve el contenido de las celdas), así que sin
    esto las filas nuevas quedarían con el alto por defecto en vez de
    heredar el de la fila de referencia (ej. la 64), como hace Excel al
    insertar una fila completa arrastrando el controlador de relleno o
    con "Insertar" desde el menú contextual."""
    combinadas_previas = [str(r) for r in sheet.merged_cells.ranges]
    for rango_txt in combinadas_previas:
        sheet.unmerge_cells(rango_txt)

    filas_existentes = None
    if fila_referencia_altura is not None:
        filas_existentes = {
            num_fila: {
                "height": dim.height,
                "hidden": dim.hidden,
                "outlineLevel": dim.outlineLevel,
            }
            for num_fila, dim in sheet.row_dimensions.items()
        }

    sheet.insert_rows(fila_insercion, amount=cantidad)

    if filas_existentes is not None:
        for num_fila, propiedades in filas_existentes.items():
            nueva_fila = num_fila + cantidad if num_fila >= fila_insercion else num_fila
            dim_nueva = sheet.row_dimensions[nueva_fila]
            dim_nueva.height = propiedades["height"]
            dim_nueva.hidden = propiedades["hidden"]
            dim_nueva.outlineLevel = propiedades["outlineLevel"]

        altura_referencia = filas_existentes.get(fila_referencia_altura)
        if altura_referencia is not None:
            for fila_nueva in range(fila_insercion, fila_insercion + cantidad):
                sheet.row_dimensions[fila_nueva].height = altura_referencia["height"]

    from openpyxl.utils.cell import range_boundaries
    from openpyxl.utils import get_column_letter

    for rango_txt in combinadas_previas:
        col_min, fila_min, col_max, fila_max = range_boundaries(rango_txt)
        if fila_min >= fila_insercion:
            fila_min += cantidad
            fila_max += cantidad
        nuevo_rango = (
            f"{get_column_letter(col_min)}{fila_min}:{get_column_letter(col_max)}{fila_max}"
        )
        try:
            sheet.merge_cells(nuevo_rango)
        except Exception:
            pass


def procesar_paso3_ficha(wb_origen, wb_modelo, alertas, resumen):
    """
    Paso 3, hoja 'FICHA' del archivo MODELO:
      1. Cuenta cuántas de las celdas C27, C28, C29 de 'Cotizador' (en el
         ORIGEN) tienen datos.
      2. Si D25:D26 está combinada en 'Cotizador' (ORIGEN), copia ese
         valor combinado a FICHA!D64 (MODELO).
      3. Si el conteo del punto 1 es mayor a 0, inserta esa cantidad de
         filas debajo de la fila 64 de 'FICHA' y emula un AutoFill de
         C64:F64 sobre esas filas nuevas (ver notas arriba del bloque).
    """
    if HOJA_COTIZADOR not in wb_origen.sheetnames:
        alertas.append(f"[FICHA] La hoja '{HOJA_COTIZADOR}' no existe en el archivo de ORIGEN.")
        return
    if HOJA_FICHA not in wb_modelo.sheetnames:
        alertas.append(f"[FICHA] La hoja '{HOJA_FICHA}' no existe en el archivo MODELO.")
        return

    sheet_cotiz_ori = wb_origen[HOJA_COTIZADOR]
    sheet_ficha_mod = wb_modelo[HOJA_FICHA]

    # --- 1) Contar datos en C27:C29 de Cotizador (ORIGEN) ---
    cantidad_filas = _contar_no_vacias(sheet_cotiz_ori, ["C27", "C28", "C29"])
    resumen["ficha_filas_a_insertar"] = cantidad_filas

    # --- 2) Si D25:D26 está combinada en Cotizador (ORIGEN), copiar su valor a FICHA!D64 ---
    combinada_d25_d26 = any(str(r) == "D25:D26" for r in sheet_cotiz_ori.merged_cells.ranges)
    if combinada_d25_d26:
        valor_d25_d26 = _lee(sheet_cotiz_ori, "D25")
        _escribe(sheet_ficha_mod, "D64", valor_d25_d26)
        resumen["ficha_d64_establecido"] = True
    else:
        resumen["ficha_d64_establecido"] = False

    if cantidad_filas == 0:
        resumen["ficha_filas_insertadas"] = 0
        return

    # --- 3) Insertar filas debajo de la 64, sin romper combinaciones ya existentes ---
    primera_fila_nueva = FILA_AUTOFILL_ORIGEN + 1
    _insertar_filas_preservando_combinadas(
        sheet_ficha_mod, primera_fila_nueva, cantidad_filas, fila_referencia_altura=FILA_AUTOFILL_ORIGEN
    )
    filas_nuevas = list(range(primera_fila_nueva, primera_fila_nueva + cantidad_filas))
    resumen["ficha_filas_insertadas"] = cantidad_filas

    alertas.append(
        f"[FICHA] Se insertaron {cantidad_filas} fila(s) debajo de la fila {FILA_AUTOFILL_ORIGEN} "
        f"y se emuló el AutoFill de C{FILA_AUTOFILL_ORIGEN}:F{FILA_AUTOFILL_ORIGEN}. "
        f"openpyxl no ajusta automáticamente fórmulas de OTRAS celdas del libro que "
        f"referencien filas debajo de la {FILA_AUTOFILL_ORIGEN}; revisa manualmente si "
        f"existen ese tipo de referencias en 'FICHA' u otras hojas."
    )

    # --- Columnas de TODA la fila (no solo C:F), para que la fila nueva
    #     quede formateada de punta a punta como si se hubiera insertado
    #     a mano ("Insertar fila completa" en Excel copia el formato -- y
    #     el patrón de celdas combinadas -- de la fila de referencia en
    #     todas sus columnas, no solo en las que tienen datos de AutoFill). ---
    from openpyxl.utils import get_column_letter

    max_columna = max(sheet_ficha_mod.max_column, 1)
    columnas_fila_completa = [get_column_letter(i) for i in range(1, max_columna + 1)]

    # --- Bloques combinados de TODA la fila origen, antes de emular el fill ---
    bloques_combinados = _bloques_combinados_en_fila(sheet_ficha_mod, FILA_AUTOFILL_ORIGEN, columnas_fila_completa)

    # --- Copiar valor/fórmula + estilo de C64:F64 (datos del AutoFill) y
    #     el estilo del resto de la fila 64 (formato de fila completa) a
    #     cada fila nueva ---
    for fila_nueva in filas_nuevas:
        offset = fila_nueva - FILA_AUTOFILL_ORIGEN
        for col in COLUMNAS_AUTOFILL:
            celda_origen = sheet_ficha_mod[f"{col}{FILA_AUTOFILL_ORIGEN}"]
            celda_destino = sheet_ficha_mod[f"{col}{fila_nueva}"]

            valor_origen = celda_origen.value
            if isinstance(valor_origen, str) and valor_origen.startswith("="):
                celda_destino.value = _ajustar_formula_relativa(valor_origen, offset)
            else:
                celda_destino.value = valor_origen

            _copiar_estilo_celda(celda_origen, celda_destino)

        for col in columnas_fila_completa:
            if col in COLUMNAS_AUTOFILL:
                continue  # ya se copiaron arriba, con su valor/fórmula incluido
            _copiar_estilo_celda(
                sheet_ficha_mod[f"{col}{FILA_AUTOFILL_ORIGEN}"], sheet_ficha_mod[f"{col}{fila_nueva}"]
            )

        # Replicar el mismo patrón de celdas combinadas de la fila origen
        for inicio, fin in bloques_combinados:
            try:
                sheet_ficha_mod.merge_cells(f"{inicio}{fila_nueva}:{fin}{fila_nueva}")
            except Exception:
                pass

    # --- Extender validación de datos y formato condicional a las filas
    #     nuevas, en TODA la fila (no solo C:F) ---
    _extender_validaciones(sheet_ficha_mod, FILA_AUTOFILL_ORIGEN, filas_nuevas, columnas_fila_completa)
    _extender_formato_condicional(sheet_ficha_mod, FILA_AUTOFILL_ORIGEN, filas_nuevas, columnas_fila_completa)


# =========================================================================
# PASO 4 — HOJA "Ax.A"
# =========================================================================
# Las fórmulas de sede (Arequipa/Ilo) referencian celdas propias de cada
# proyecto -- ver templates_manager.CELDAS_AXA_POR_DEFECTO["sedes_validas"]
# (con la nota sobre UPPER vs MAYUSC) y CELDAS_AXA_POR_PROYECTO.


def procesar_paso4_axa(wb_origen, wb_modelo, alertas, resumen, valores_e51_e54=None, sede=None, celdas_axa=None):
    """
    valores_e51_e54: lista/tupla de hasta 4 valores (para las celdas de
        celdas_axa["e51_e54"]) escritos a mano desde la interfaz. Los
        valores None o cadenas vacías se omiten (no se sobrescribe esa celda).
    sede: "Arequipa" o "Ilo" (o None para no tocar celdas_axa["g124"]).
    celdas_axa: mapeo de coordenadas de la hoja "Ax.A" a usar (ver
        templates_manager.obtener_celdas_axa); si no se pasa, se usan las
        de templates_manager.CELDAS_AXA_POR_DEFECTO (layout de COSTA DEL
        VALLE / COSTA AZUL), para no romper a quien llame esta función sin
        indicar proyecto.
    """
    celdas_axa = celdas_axa or CELDAS_AXA_POR_DEFECTO

    if HOJA_AXA not in wb_modelo.sheetnames:
        alertas.append(f"[AX.A] La hoja '{HOJA_AXA}' no existe en el archivo MODELO.")
        return

    sheet_axa_mod = wb_modelo[HOJA_AXA]

    # --- 0) Asegurar que todas las filas tengan un alto explícito. Esta
    #         hoja también tiene la propiedad de Excel "zeroHeight" activada
    #         (igual que "Ax.B", ver procesar_paso5_axb): cualquier fila SIN
    #         alto explícito se guarda -y se ve, al reabrir en Excel- con
    #         alto 0 (como oculta), aunque nunca se haya tocado. Las filas
    #         que YA tenían un alto explícito quedan tal cual (no se pisa
    #         ningún valor real del archivo original). ---
    _asegurar_alto_explicito_en_todas_las_filas(sheet_axa_mod)

    # --- 1) Valores manuales E51:E54 (vienen del apartado del chatbot) ---
    if valores_e51_e54:
        for celda, valor in zip(celdas_axa["e51_e54"], valores_e51_e54):
            if valor is not None and str(valor).strip() != "":
                _escribe(sheet_axa_mod, celda, valor)
        resumen["axa_e51_e54_escrito"] = True
    else:
        resumen["axa_e51_e54_escrito"] = False

    # --- 2) Cotizador!I23 (ORIGEN) -> letras -> Ax.A!G115 (MODELO) ---
    if HOJA_COTIZADOR not in wb_origen.sheetnames:
        alertas.append(
            f"[AX.A] La hoja '{HOJA_COTIZADOR}' no existe en el archivo de ORIGEN; no se pudo calcular G115."
        )
    else:
        sheet_cotiz_ori = wb_origen[HOJA_COTIZADOR]
        valor_i23 = _lee(sheet_cotiz_ori, "I23")
        texto_i23 = formatear_monto_letras(valor_i23)
        if texto_i23 is None:
            alertas.append(
                f"[AX.A] No se pudo convertir a letras el valor de Cotizador!I23 ('{valor_i23}')."
            )
        else:
            _escribe(sheet_axa_mod, celdas_axa["g115"], texto_i23)
            resumen["axa_g115"] = texto_i23

    # --- 3) Ax.A!J119 (ORIGEN, valor numérico ya calculado, NO la fórmula) -> letras -> Ax.A!G120 (MODELO) ---
    if HOJA_AXA not in wb_origen.sheetnames:
        alertas.append(
            f"[AX.A] La hoja '{HOJA_AXA}' no existe en el archivo de ORIGEN; no se pudo calcular G120."
        )
    else:
        sheet_axa_ori = wb_origen[HOJA_AXA]
        # wb_origen se carga con data_only=True, así que _lee() aquí devuelve
        # el valor numérico ya calculado por Excel, nunca el texto "=...".
        valor_j119 = _lee(sheet_axa_ori, celdas_axa["j119_origen"])
        texto_j119 = formatear_monto_letras(valor_j119)
        if texto_j119 is None:
            alertas.append(
                f"[AX.A] No se pudo convertir a letras el valor de Ax.A!{celdas_axa['j119_origen']} ('{valor_j119}')."
            )
        else:
            _escribe(sheet_axa_mod, celdas_axa["g120"], texto_j119)
            resumen["axa_g120"] = texto_j119

    # --- 4) Sede (Arequipa / Ilo) -> fórmula en Ax.A!G124 (o su equivalente) ---
    sedes_validas = celdas_axa["sedes_validas"]
    if sede in sedes_validas:
        _escribe(sheet_axa_mod, celdas_axa["g124"], sedes_validas[sede])
        resumen["axa_g124"] = sedes_validas[sede]
    elif sede:
        alertas.append(f"[AX.A] Sede '{sede}' no reconocida (debe ser 'Arequipa' o 'Ilo'); no se modificó G124.")


# =========================================================================
# PASO 5 — HOJA "Ax.B"  (emulación de AutoFill con openpyxl)
# =========================================================================
# Usa exactamente la misma lógica de emulación de AutoFill que el Paso 3
# (ver las notas detalladas junto a ese bloque, más arriba en este
# archivo, sobre lo que SÍ y lo que NO reproduce esta emulación frente
# al Range.AutoFill real de Excel).
HOJA_AXB = "Ax.B"
FILA_AUTOFILL_AXB = 34
COLUMNAS_AUTOFILL_AXB = ["B", "C", "D", "E"]

# Hojas donde se busca TEXTO_SEGUNDO_TITULAR_COPROPIETARIO (ver Paso 1,
# más arriba en este archivo, junto a
# _formatear_celdas_segundo_titular_copropietario).
HOJA_AXE = "Ax.E"
HOJAS_ESCANEO_SEGUNDO_TITULAR = [HOJA_FICHA, HOJA_AXA, HOJA_AXB, HOJA_AXE]


def _insertar_filas_preservando_formato_axb(sheet, fila_insercion, cantidad, fila_referencia_altura):
    """Igual que _insertar_filas_preservando_combinadas, pero ADEMÁS
    conserva el ALTO de cada fila (row height) al desplazarlas hacia
    abajo. openpyxl no reubica el alto de fila al usar insert_rows()
    (solo mueve el contenido de las celdas), así que sin este ajuste el
    'cuadro de abajo' queda con las filas del alto por defecto en vez
    de su alto original -> se ve descuadrado. Las filas nuevas (recién
    insertadas) heredan el alto de 'fila_referencia_altura' (la fila
    origen del AutoFill, ej. la 34), igual que hace Excel al insertar
    filas arrastrando el controlador de relleno."""
    # --- Celdas combinadas: se descombinan antes de insertar (igual que en el Paso 3) ---
    combinadas_previas = [str(r) for r in sheet.merged_cells.ranges]
    for rango_txt in combinadas_previas:
        sheet.unmerge_cells(rango_txt)

    # --- Alturas de fila existentes, antes de insertar ---
    filas_existentes = {
        num_fila: {
            "height": dim.height,
            "hidden": dim.hidden,
            "outlineLevel": dim.outlineLevel,
        }
        for num_fila, dim in sheet.row_dimensions.items()
    }

    sheet.insert_rows(fila_insercion, amount=cantidad)

    # --- Reubicar las alturas: las de fila_insercion en adelante se desplazan 'cantidad' filas ---
    for num_fila, propiedades in filas_existentes.items():
        nueva_fila = num_fila + cantidad if num_fila >= fila_insercion else num_fila
        dim_nueva = sheet.row_dimensions[nueva_fila]
        dim_nueva.height = propiedades["height"]
        dim_nueva.hidden = propiedades["hidden"]
        dim_nueva.outlineLevel = propiedades["outlineLevel"]

    # --- Las filas nuevas (vacías) heredan el alto de la fila de referencia ---
    altura_referencia = filas_existentes.get(fila_referencia_altura)
    if altura_referencia is not None:
        for fila_nueva in range(fila_insercion, fila_insercion + cantidad):
            dim_nueva = sheet.row_dimensions[fila_nueva]
            dim_nueva.height = altura_referencia["height"]

    # --- Restaurar celdas combinadas en su posición correcta (igual que en el Paso 3) ---
    from openpyxl.utils.cell import range_boundaries
    from openpyxl.utils import get_column_letter

    for rango_txt in combinadas_previas:
        col_min, fila_min, col_max, fila_max = range_boundaries(rango_txt)
        if fila_min >= fila_insercion:
            fila_min += cantidad
            fila_max += cantidad
        nuevo_rango = (
            f"{get_column_letter(col_min)}{fila_min}:{get_column_letter(col_max)}{fila_max}"
        )
        try:
            sheet.merge_cells(nuevo_rango)
        except Exception:
            pass


def _expandir_rangos_que_terminan_en(texto, fila_limite, fila_nueva_limite):
    """Dentro de 'texto' (una fórmula), busca RANGOS (dos referencias
    unidas por ':') cuyo extremo final sea EXACTAMENTE la fila
    'fila_limite' (sin '$') y lo extiende a 'fila_nueva_limite'. A
    diferencia de _desplazar_referencias_filas, esto NO toca referencias
    sueltas a una sola celda (ej. 'B34' fuera de un rango) — solo el
    extremo final de un rango tipo 'D25:D34' -> 'D25:D38'. Así se
    reproduce el comportamiento real de Excel cuando insertas filas
    justo en el límite de un rango usado en una fórmula (ej. una suma
    total), sin alterar referencias sueltas a esa misma fila que no
    forman parte de un rango."""
    patron_rango = re.compile(r"(\$?[A-Za-z]{1,3}\$?\d+):(\$?[A-Za-z]{1,3})(\$?)(\d+)")
    partes = texto.split('"')
    for i in range(0, len(partes), 2):

        def _reemplazar(m):
            inicio, col_fin, fila_abs, fila = m.groups()
            if fila_abs == "$" or int(fila) != fila_limite:
                return m.group(0)
            return f"{inicio}:{col_fin}{fila_nueva_limite}"

        partes[i] = patron_rango.sub(_reemplazar, partes[i])
    return '"'.join(partes)


def _desplazar_referencias_filas(texto, fila_insercion, cantidad):
    """Desplaza, dentro de 'texto' (una fórmula o una lista de rangos
    tipo sqref), toda referencia de celda cuya FILA sea >= fila_insercion
    y NO esté anclada con '$' antes del número de fila, sumándole
    'cantidad'. Es la misma regla que aplica Excel de forma nativa
    cuando insertas filas de verdad: cualquier referencia que apunte a
    una fila que se movió, se actualiza para seguir apuntando al mismo
    contenido en su nueva posición. Respeta el texto entre comillas para
    no tocar cadenas dentro de fórmulas."""
    partes = texto.split('"')
    for i in range(0, len(partes), 2):  # solo los tramos FUERA de comillas

        def _reemplazar(m):
            col_abs, col, fila_abs, fila = m.groups()
            if fila_abs == "$":
                return m.group(0)
            fila_int = int(fila)
            if fila_int >= fila_insercion:
                return f"{col_abs}{col}{fila_abs}{fila_int + cantidad}"
            return m.group(0)

        partes[i] = _PATRON_REF_CELDA.sub(_reemplazar, partes[i])
    return '"'.join(partes)


def _desplazar_formulas_de_la_hoja(sheet, fila_insercion, cantidad, fila_limite_a_expandir=None):
    """Recorre TODAS las celdas de la hoja y, para cada fórmula:
      1) ajusta las referencias que apunten a filas >= fila_insercion
         (que ya se movieron físicamente por el insert_rows), y
      2) si se indica 'fila_limite_a_expandir', además extiende
         cualquier RANGO que terminaba exactamente en esa fila (ej. la
         fila fuente del AutoFill) para que abarque las filas nuevas
         (ej. '=SUMA(D25:D34)' -> '=SUMA(D25:D38)')."""
    for fila in sheet.iter_rows():
        for celda in fila:
            valor = celda.value
            if isinstance(valor, str) and valor.startswith("="):
                nuevo_valor = _desplazar_referencias_filas(valor, fila_insercion, cantidad)
                if fila_limite_a_expandir is not None:
                    nuevo_valor = _expandir_rangos_que_terminan_en(
                        nuevo_valor, fila_limite_a_expandir, fila_limite_a_expandir + cantidad
                    )
                if nuevo_valor != valor:
                    celda.value = nuevo_valor


def _desplazar_validaciones_y_formato_condicional(sheet, fila_insercion, cantidad):
    """Igual que _desplazar_formulas_de_la_hoja, pero para los rangos
    (sqref) de las validaciones de datos y del formato condicional, que
    tampoco se reubican solos al insertar filas."""
    for dv in sheet.data_validations.dataValidation:
        sqref_actual = str(dv.sqref)
        nuevo_sqref = _desplazar_referencias_filas(sqref_actual, fila_insercion, cantidad)
        if nuevo_sqref != sqref_actual:
            try:
                dv.sqref = nuevo_sqref
            except Exception:
                pass
    try:
        for cf in sheet.conditional_formatting:
            sqref_actual = str(cf.sqref)
            nuevo_sqref = _desplazar_referencias_filas(sqref_actual, fila_insercion, cantidad)
            if nuevo_sqref != sqref_actual:
                try:
                    cf.sqref = nuevo_sqref
                except Exception:
                    pass
    except Exception:
        pass


def _asegurar_alto_explicito_en_todas_las_filas(sheet):
    """Fuerza que TODAS las filas usadas de 'sheet' (1 a sheet.max_row)
    tengan un alto explícito: si ya tenían uno, se deja tal cual; si no
    tenían ninguno, se les pone el alto por defecto de la propia hoja
    (sheet_format.defaultRowHeight). Necesario en hojas con la propiedad
    "zeroHeight" activada, donde cualquier fila SIN alto explícito se
    guarda -y se ve, al reabrir en Excel- con alto 0 (como oculta),
    incluso si esa fila nunca fue tocada por el procesamiento."""
    alto_por_defecto = sheet.sheet_format.defaultRowHeight or 15
    for num_fila in range(1, sheet.max_row + 1):
        dim = sheet.row_dimensions[num_fila]
        if dim.height is None:
            dim.height = alto_por_defecto


def procesar_paso5_axb(wb_origen, wb_modelo, alertas, resumen):
    """
    Paso 5, hoja 'Ax.B' del archivo MODELO:
      1. Cuenta cuántas de las celdas C27, C28, C29 de 'Cotizador' (en el
         ORIGEN) tienen datos (mismo conteo que se usa en el Paso 3).
      2. Si ese conteo es mayor a 0, inserta esa cantidad de filas DEBAJO
         de la fila 34 (la fila 34 en sí NUNCA se mueve; sigue siendo la
         fuente fija del AutoFill, igual que en la primera versión de
         este paso) y rellena esas filas nuevas emulando el AutoFill de
         B34:E34: fórmulas con referencias relativas ajustadas,
         formato completo, el mismo patrón de celdas combinadas,
         validación de datos y formato condicional extendidos.
      3. Además, ajusta TODAS las demás fórmulas de la hoja (y los
         rangos de validación de datos / formato condicional) que
         hacían referencia a filas por debajo de la 34, para que sigan
         apuntando al contenido correcto tras el desplazamiento — y
         extiende cualquier rango que terminaba exactamente en la fila
         34 (como una suma total) para que abarque también las filas
         nuevas.
    """
    if HOJA_COTIZADOR not in wb_origen.sheetnames:
        alertas.append(f"[AX.B] La hoja '{HOJA_COTIZADOR}' no existe en el archivo de ORIGEN.")
        return
    if HOJA_AXB not in wb_modelo.sheetnames:
        alertas.append(f"[AX.B] La hoja '{HOJA_AXB}' no existe en el archivo MODELO.")
        return

    sheet_cotiz_ori = wb_origen[HOJA_COTIZADOR]
    sheet_axb_mod = wb_modelo[HOJA_AXB]

    # --- 0) Asegurar que todas las filas tengan un alto explícito, se
    #         inserten filas o no. Esta hoja tiene la propiedad de Excel
    #         "zeroHeight" activada (alto=0 para cualquier fila SIN alto
    #         explícito): openpyxl, al volver a GUARDAR el archivo, escribe
    #         esas filas sin alto tal cual -- y Excel las muestra con alto
    #         0 (como ocultas) al reabrirlo, aunque nunca se hayan tocado.
    #         Esto pasa incluso sin insertar ninguna fila (es un problema
    #         de ida y vuelta de guardado, no del AutoFill de este Paso). ---
    _asegurar_alto_explicito_en_todas_las_filas(sheet_axb_mod)

    # --- 0.1) C34 siempre alineada a la derecha, se inserten filas o no ---
    alineacion_derecha_axb = Alignment(horizontal="right")
    sheet_axb_mod[f"C{FILA_AUTOFILL_AXB}"].alignment = alineacion_derecha_axb

    # --- 1) Contar datos en C27:C29 de Cotizador (ORIGEN) ---
    cantidad_filas = _contar_no_vacias(sheet_cotiz_ori, ["C27", "C28", "C29"])
    resumen["axb_filas_a_insertar"] = cantidad_filas

    if cantidad_filas == 0:
        resumen["axb_filas_insertadas"] = 0
        return

    # --- 2) Insertar filas DEBAJO de la 34; la fila 34 nunca se mueve ---
    primera_fila_nueva = FILA_AUTOFILL_AXB + 1
    _insertar_filas_preservando_formato_axb(
        sheet_axb_mod, primera_fila_nueva, cantidad_filas, fila_referencia_altura=FILA_AUTOFILL_AXB
    )
    filas_nuevas = list(range(primera_fila_nueva, primera_fila_nueva + cantidad_filas))
    resumen["axb_filas_insertadas"] = cantidad_filas

    # --- 3) Ajustar TODAS las demás fórmulas/validaciones de la hoja que quedaron desplazadas,
    #         y extender cualquier rango que terminaba justo en la fila 34 ---
    _desplazar_formulas_de_la_hoja(
        sheet_axb_mod, primera_fila_nueva, cantidad_filas, fila_limite_a_expandir=FILA_AUTOFILL_AXB
    )
    _desplazar_validaciones_y_formato_condicional(sheet_axb_mod, primera_fila_nueva, cantidad_filas)

    alertas.append(
        f"[AX.B] Se insertaron {cantidad_filas} fila(s) debajo de la fila {FILA_AUTOFILL_AXB} "
        f"y se emuló el AutoFill de B{FILA_AUTOFILL_AXB}:E{FILA_AUTOFILL_AXB}. También se ajustaron "
        f"automáticamente las demás fórmulas de 'Ax.B' que apuntaban a filas desplazadas, y los "
        f"rangos que terminaban justo en la fila {FILA_AUTOFILL_AXB} se extendieron para abarcar las "
        f"filas nuevas. OJO: esto solo cubre la hoja 'Ax.B'; si OTRA hoja tiene fórmulas que "
        f"referencian celdas de 'Ax.B' por debajo de la fila {FILA_AUTOFILL_AXB}, revísalas a mano."
    )

    # --- Bloques combinados de la fila 34 (fija), antes de emular el fill ---
    bloques_combinados = _bloques_combinados_en_fila(sheet_axb_mod, FILA_AUTOFILL_AXB, COLUMNAS_AUTOFILL_AXB)

    # --- Copiar valor/fórmula + estilo de B34:E34 a cada fila nueva (igual que antes) ---
    for fila_nueva in filas_nuevas:
        offset = fila_nueva - FILA_AUTOFILL_AXB
        for col in COLUMNAS_AUTOFILL_AXB:
            celda_origen = sheet_axb_mod[f"{col}{FILA_AUTOFILL_AXB}"]
            celda_destino = sheet_axb_mod[f"{col}{fila_nueva}"]

            valor_origen = celda_origen.value
            if isinstance(valor_origen, str) and valor_origen.startswith("="):
                celda_destino.value = _ajustar_formula_relativa(valor_origen, offset)
            else:
                celda_destino.value = valor_origen

            _copiar_estilo_celda(celda_origen, celda_destino)
            if col == "C":
                celda_destino.alignment = alineacion_derecha_axb

        # Replicar el mismo patrón de celdas combinadas de la fila 34
        for inicio, fin in bloques_combinados:
            try:
                sheet_axb_mod.merge_cells(f"{inicio}{fila_nueva}:{fin}{fila_nueva}")
            except Exception:
                pass

    # --- Extender validación de datos y formato condicional a las filas nuevas ---
    _extender_validaciones(sheet_axb_mod, FILA_AUTOFILL_AXB, filas_nuevas, COLUMNAS_AUTOFILL_AXB)
    _extender_formato_condicional(sheet_axb_mod, FILA_AUTOFILL_AXB, filas_nuevas, COLUMNAS_AUTOFILL_AXB)


# =========================================================================
# PASO 6 — HOJA "List.Ax"
# =========================================================================
HOJA_LISTAX = "List.Ax"
MEMBRESIAS_VALIDAS = {"CON MEMBRESÍA", "SIN MEMBRESIA"}


def procesar_paso6_listax(wb_modelo, alertas, resumen, membresia=None):
    """
    Paso 6, hoja 'List.Ax' del archivo MODELO:
      - "CON MEMBRESÍA": el relleno de G17 se pone en blanco.
      - "SIN MEMBRESIA": el contenido de G19 se MUEVE a G17 (se copia y
        luego se borra de G19), y además se borra el contenido de
        C19:F19.
    """
    if HOJA_LISTAX not in wb_modelo.sheetnames:
        alertas.append(f"[LIST.AX] La hoja '{HOJA_LISTAX}' no existe en el archivo MODELO.")
        return
    if not membresia:
        return
    if membresia not in MEMBRESIAS_VALIDAS:
        alertas.append(f"[LIST.AX] Opción '{membresia}' no reconocida (debe ser 'CON MEMBRESÍA' o 'SIN MEMBRESIA').")
        return

    sheet = wb_modelo[HOJA_LISTAX]
    resumen["listax_membresia"] = membresia

    if membresia == "CON MEMBRESÍA":
        color_blanco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        _celda_ancla(sheet, "G17").fill = color_blanco
        resumen["listax_g17_relleno_blanco"] = True
    else:  # "SIN MEMBRESIA"
        valor_g19 = _lee(sheet, "G19")
        _escribe(sheet, "G17", valor_g19)
        _escribe(sheet, "G19", None)
        for col in ("C", "D", "E", "F"):
            _escribe(sheet, f"{col}19", None)
        resumen["listax_g19_movido_a_g17"] = valor_g19


# =========================================================================
# HOJA "Px" — se copia ENTERA del ORIGEN al MODELO, reemplazando la que
# ya existe ahí, siempre (sin depender de qué Pasos estén activados).
# =========================================================================
def _clonar_hoja_completa(sheet_origen, sheet_destino):
    """Copia todo el contenido de 'sheet_origen' a 'sheet_destino' (ya
    vacía, recién creada): valores (las fórmulas se copian como su
    resultado ya calculado, porque 'sheet_origen' viene de un workbook
    cargado con data_only=True, igual que el resto del ORIGEN en este
    módulo), estilo de cada celda, celdas combinadas, y alto/ancho de
    filas y columnas."""
    for fila in sheet_origen.iter_rows():
        for celda in fila:
            celda_destino = sheet_destino.cell(row=celda.row, column=celda.column, value=celda.value)
            _copiar_estilo_celda(celda, celda_destino)

    for rango in sheet_origen.merged_cells.ranges:
        sheet_destino.merge_cells(str(rango))

    for num_fila, dim in sheet_origen.row_dimensions.items():
        dim_destino = sheet_destino.row_dimensions[num_fila]
        dim_destino.height = dim.height
        dim_destino.hidden = dim.hidden
        dim_destino.outlineLevel = dim.outlineLevel

    for letra_col, dim in sheet_origen.column_dimensions.items():
        dim_destino = sheet_destino.column_dimensions[letra_col]
        dim_destino.width = dim.width
        dim_destino.hidden = dim.hidden
        dim_destino.outlineLevel = dim.outlineLevel


def procesar_copia_hoja_px(wb_origen, wb_modelo, alertas, resumen):
    """Copia ENTERA la hoja 'Px' del ORIGEN al MODELO, reemplazando la que
    ya existe en el MODELO (mismas celdas, estilos, combinaciones y
    alto/ancho de filas y columnas). Se hace siempre, para los 4 Pasos y
    sin depender de qué otros Pasos estén activados."""
    if HOJA_PX not in wb_origen.sheetnames:
        alertas.append(f"[PX] La hoja '{HOJA_PX}' no existe en el archivo de ORIGEN.")
        return
    if HOJA_PX not in wb_modelo.sheetnames:
        alertas.append(f"[PX] La hoja '{HOJA_PX}' no existe en el archivo MODELO.")
        return

    sheet_origen = wb_origen[HOJA_PX]
    indice_original = wb_modelo.sheetnames.index(HOJA_PX)
    del wb_modelo[HOJA_PX]
    sheet_destino = wb_modelo.create_sheet(HOJA_PX, indice_original)

    _clonar_hoja_completa(sheet_origen, sheet_destino)
    resumen["px_copiada"] = True


# =========================================================================
# ORQUESTACIÓN: PASO 1 + PASO 2 EN UNA SOLA PASADA
# =========================================================================
def procesar_todo(archivo_origen, archivo_modelo, reglas_memoria, ejecutar_paso2=True,
                   fecha_manual=None, ejecutar_paso3=True, ejecutar_paso4=True,
                   valores_e51_e54=None, sede=None, ejecutar_paso5=True,
                   ejecutar_paso6=True, membresia=None, celdas_axa=None):
    alertas = []
    resumen = {
        "filas_procesadas": 0,
        "filas_con_guion_ignoradas": 0,
    }

    wb_origen = load_workbook(archivo_origen, data_only=True)
    archivo_modelo.seek(0)
    wb_modelo = load_workbook(archivo_modelo, data_only=False)

    _procesar_paso1_datos_wb(wb_origen, wb_modelo, reglas_memoria, alertas, resumen)
    procesar_copia_hoja_px(wb_origen, wb_modelo, alertas, resumen)

    if ejecutar_paso2:
        procesar_paso2_cotizador(wb_origen, wb_modelo, alertas, resumen, fecha_manual=fecha_manual)

    if ejecutar_paso3:
        procesar_paso3_ficha(wb_origen, wb_modelo, alertas, resumen)

    if ejecutar_paso4:
        procesar_paso4_axa(
            wb_origen, wb_modelo, alertas, resumen,
            valores_e51_e54=valores_e51_e54, sede=sede, celdas_axa=celdas_axa,
        )

    if ejecutar_paso5:
        procesar_paso5_axb(wb_origen, wb_modelo, alertas, resumen)

    if ejecutar_paso6:
        procesar_paso6_listax(wb_modelo, alertas, resumen, membresia=membresia)

    return _guardar_bytes(wb_modelo), alertas, resumen


def _guardar_bytes(wb_modelo):
    buffer = io.BytesIO()
    wb_modelo.save(buffer)
    buffer.seek(0)
    return _suprimir_marcas_de_error_excel(buffer.getvalue())


# Bloque <ignoredErrors> que se inserta en CADA hoja del .xlsx ya guardado,
# cubriendo el rango completo posible de una hoja (A1:XFD1048576), para
# que Excel no muestre el triangulito/marca verde de "revisión de errores"
# (número guardado como texto, fórmula inconsistente con las de alrededor,
# etc.) en ninguna celda. openpyxl define el esquema de <ignoredErrors>
# (ver openpyxl.worksheet.errors) pero NUNCA lo conecta a la escritura del
# archivo, así que no hay forma de pedírselo vía su API normal -- hay que
# insertarlo a mano en el XML de cada hoja dentro del .zip del .xlsx ya
# guardado, sin tocar ningún otro contenido del archivo. Esto NO oculta
# errores reales de fórmula (#REF!, #DIV/0!, etc.), que se siguen viendo
# igual -- solo apaga las sugerencias/advertencias de estilo de Excel.
_BLOQUE_IGNORED_ERRORS_XML = (
    '<ignoredErrors>'
    '<ignoredError sqref="A1:XFD1048576" numberStoredAsText="1" formula="1" '
    'formulaRange="1" twoDigitTextYear="1" emptyCellReference="1" '
    'listDataValidation="1" calculatedColumn="1" unlockedFormula="1"/>'
    '</ignoredErrors>'
)
# <ignoredErrors> debe ir, según el esquema de Excel, después de
# rowBreaks/colBreaks y antes de drawing/legacyDrawing/tableParts/extLst
# (lo que exista primero en cada hoja); si no hay ninguno de esos, se
# inserta justo antes de cerrar </worksheet>.
_PATRON_PUNTO_INSERCION_IGNORED_ERRORS = re.compile(r"<(drawing|legacyDrawing|tableParts|extLst)\b")


def _suprimir_marcas_de_error_excel(bytes_xlsx):
    """Postprocesa el .xlsx ya guardado agregándole, a CADA hoja, el
    bloque <ignoredErrors> de arriba. Devuelve los bytes del .xlsx
    modificado."""
    entrada = io.BytesIO(bytes_xlsx)
    salida = io.BytesIO()

    with zipfile.ZipFile(entrada, "r") as zip_entrada:
        with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as zip_salida:
            for item in zip_entrada.infolist():
                contenido = zip_entrada.read(item.filename)
                if re.match(r"xl/worksheets/sheet\d+\.xml$", item.filename):
                    texto = contenido.decode("utf-8")
                    match = _PATRON_PUNTO_INSERCION_IGNORED_ERRORS.search(texto)
                    if match:
                        texto = texto[:match.start()] + _BLOQUE_IGNORED_ERRORS_XML + texto[match.start():]
                    else:
                        texto = texto.replace("</worksheet>", _BLOQUE_IGNORED_ERRORS_XML + "</worksheet>")
                    contenido = texto.encode("utf-8")
                zip_salida.writestr(item, contenido)

    salida.seek(0)
    return salida.getvalue()
