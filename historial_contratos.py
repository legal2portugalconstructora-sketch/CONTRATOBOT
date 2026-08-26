# -*- coding: utf-8 -*-
"""historial_contratos.py — Archivo histórico de documentos generados
(una copia de cada Excel/Word/PDF, organizada por proyecto y cliente en
C:\\JESUS LEGAL\\CONTRATOS\\) y registro de números de membresía usados.

Todo acá es "mejor esfuerzo": si algo falla (la carpeta no existe en esta
máquina, no hay permisos, etc.), se ignora en silencio y nunca interrumpe
el procesamiento real -- lo único que se pierde es la copia archivada o
el registro del historial, nunca el documento que el usuario ya tiene en
sus manos (la descarga normal por el navegador no depende de nada de acá).

Todo esto respeta el "Modo de pruebas" (ver stats_tracker.py): mientras
está activo, ni se archivan copias ni se registran números de membresía,
para no ensuciar el historial real con pruebas."""

import io
import json
import os
from datetime import datetime

from stats_tracker import modo_pruebas_activo

RAIZ_HISTORIAL = r"C:\JESUS LEGAL\CONTRATOS"

RUTA_HISTORIAL_MEMBRESIAS = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "historial_membresias.json"
)

_CARACTERES_INVALIDOS_WINDOWS = '\\/:*?"<>|'


def _sanear_nombre(texto):
    """Reemplaza los caracteres que Windows no permite en nombres de
    archivo/carpeta (\\ / : * ? " < > |) por '-', y recorta espacios."""
    texto = str(texto if texto is not None else "").strip()
    for caracter in _CARACTERES_INVALIDOS_WINDOWS:
        texto = texto.replace(caracter, "-")
    return texto or "SIN-DATO"


def extraer_lote_y_nombre(excel_bytes):
    """Lee el lote (hoja Cotizador, celda E15) y el nombre del cliente
    (hoja DATOS, celda E11) de un Excel ya generado por el módulo Excel
    (ORIGEN ya copiado a MODELO) -- para armar el nombre de
    carpeta/archivo del historial. Devuelve (lote, nombre); cualquiera de
    los dos puede salir None si la celda está vacía o la hoja no existe."""
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    lote = wb["Cotizador"]["E15"].value if "Cotizador" in wb.sheetnames else None
    nombre = wb["DATOS"]["E11"].value if "DATOS" in wb.sheetnames else None
    return lote, nombre


def prefijo_archivo(lote, nombre_cliente):
    """Arma el prefijo tipo 'L-2 JUAN PEREZ' a partir del lote (tal cual
    viene en la celda, ej. 'L-2') y el nombre del cliente."""
    return _sanear_nombre(f"{lote} {nombre_cliente}".strip())


def carpeta_cliente(proyecto, lote, nombre_cliente, crear=True):
    """Devuelve la ruta de carpeta del historial para este cliente:
    RAIZ_HISTORIAL/<proyecto>/<lote nombre>/ -- la crea si no existe y
    'crear' es True. Devuelve (ruta, prefijo)."""
    prefijo = prefijo_archivo(lote, nombre_cliente)
    ruta = os.path.join(RAIZ_HISTORIAL, _sanear_nombre(proyecto), prefijo)
    if crear:
        os.makedirs(ruta, exist_ok=True)
    return ruta, prefijo


def guardar_copia(proyecto, lote, nombre_cliente, nombre_base_archivo, bytes_contenido):
    """Guarda una copia de 'bytes_contenido' en
    CONTRATOS/<proyecto>/<lote nombre>/<lote nombre> - <nombre_base_archivo>.
    Nunca lanza excepción hacia afuera -- si algo falla (Modo de pruebas
    activo, falta lote/nombre, RAIZ_HISTORIAL no accesible en esta
    máquina, etc.), no interrumpe el procesamiento real.

    Devuelve (ruta_o_None, motivo_o_None): si se guardó, ruta tiene la
    ruta y motivo es None; si NO se guardó, ruta es None y motivo trae un
    texto explicando por qué (para poder avisarle al usuario en la
    interfaz, en vez de que el archivado falle en silencio sin que nadie
    se entere)."""
    if modo_pruebas_activo():
        return None, "el Modo de pruebas está activo"
    if not lote or not str(lote).strip():
        return None, "no se encontró el lote (celda Cotizador!E15 está vacía en el Excel)"
    if not nombre_cliente or not str(nombre_cliente).strip():
        return None, "no se encontró el nombre del cliente (celda DATOS!E11 está vacía en el Excel)"
    try:
        carpeta, prefijo = carpeta_cliente(proyecto, lote, nombre_cliente)
        ruta_archivo = os.path.join(carpeta, f"{prefijo} - {nombre_base_archivo}")
        with open(ruta_archivo, "wb") as f:
            f.write(bytes_contenido)
        return ruta_archivo, None
    except Exception as exc:
        return None, f"error al guardar en '{RAIZ_HISTORIAL}': {exc}"


def encontrar_carpeta_por_prefijo(prefijo):
    """Busca, en TODOS los proyectos de RAIZ_HISTORIAL, una carpeta de
    cliente cuyo nombre sea exactamente 'prefijo' (ej. 'L-2 JUAN PEREZ').
    Se usa desde el módulo PDF, que no sabe a qué proyecto pertenece un
    .docx que le suben. Devuelve la ruta si la encuentra, o None."""
    if not os.path.isdir(RAIZ_HISTORIAL):
        return None
    try:
        for proyecto in os.listdir(RAIZ_HISTORIAL):
            candidato = os.path.join(RAIZ_HISTORIAL, proyecto, prefijo)
            if os.path.isdir(candidato):
                return candidato
    except OSError:
        return None
    return None


def guardar_copia_en_carpeta(carpeta, nombre_archivo, bytes_contenido):
    """Guarda 'bytes_contenido' directo en 'carpeta' (ya existente) con
    el nombre 'nombre_archivo'. No hace nada (y no falla) si el Modo de
    pruebas está activo."""
    if modo_pruebas_activo():
        return None
    try:
        ruta_archivo = os.path.join(carpeta, nombre_archivo)
        with open(ruta_archivo, "wb") as f:
            f.write(bytes_contenido)
        return ruta_archivo
    except Exception:
        return None


# =========================================================================
# HISTORIAL DE NÚMEROS DE MEMBRESÍA -- para saber "en qué número estamos"
# antes de escribir uno nuevo a mano en el Paso 2 del módulo Word.
# =========================================================================
_MAX_MEMBRESIAS_GUARDADAS = 200


def _cargar_membresias():
    if not os.path.exists(RUTA_HISTORIAL_MEMBRESIAS):
        return {"eventos": []}
    try:
        with open(RUTA_HISTORIAL_MEMBRESIAS, "r", encoding="utf-8") as f:
            datos = json.load(f)
    except (json.JSONDecodeError, OSError):
        return {"eventos": []}
    if not isinstance(datos, dict) or not isinstance(datos.get("eventos"), list):
        return {"eventos": []}
    return datos


def _guardar_membresias(datos):
    with open(RUTA_HISTORIAL_MEMBRESIAS, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)


def registrar_numero_membresia(proyecto, nombre_cliente, numero_membresia):
    """Agrega un evento al historial de números de membresía usados. No
    hace nada (y no falla) si el Modo de pruebas está activo, o si
    'numero_membresia' viene vacío."""
    if modo_pruebas_activo():
        return
    if not numero_membresia or not str(numero_membresia).strip():
        return
    try:
        datos = _cargar_membresias()
        datos["eventos"].append({
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "proyecto": proyecto,
            "nombre_cliente": nombre_cliente,
            "numero_membresia": str(numero_membresia).strip(),
        })
        datos["eventos"] = datos["eventos"][-_MAX_MEMBRESIAS_GUARDADAS:]
        _guardar_membresias(datos)
    except Exception:
        pass


def obtener_historial_membresias(limite=20):
    """Devuelve los últimos 'limite' eventos del historial de números de
    membresía, más recientes primero."""
    eventos = _cargar_membresias()["eventos"]
    return list(reversed(eventos))[:limite]


def obtener_ultimo_numero_membresia():
    """Devuelve el evento más reciente del historial de membresías, o
    None si todavía no hay ninguno."""
    historial = obtener_historial_membresias(limite=1)
    return historial[0] if historial else None
