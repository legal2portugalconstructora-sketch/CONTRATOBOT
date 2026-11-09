# -*- coding: utf-8 -*-
"""n8n_integracion.py — Avisa a un flujo de n8n (registro maestro externo,
fuera de esta app) cada vez que se genera un CONTRATO en Word.

Todo acá es "mejor esfuerzo", mismo criterio que historial_contratos.py:
si n8n está apagado, no responde a tiempo, o responde con un error, NUNCA
se corta la generación del contrato -- el aviso queda guardado en
pendientes_n8n.json y se reintenta solo (la próxima vez que se genere un
contrato, o al abrir la app)."""

import io
import json
import os
from datetime import date, datetime
from zoneinfo import ZoneInfo

import yaml

from excel_pipeline import es_fecha, es_marcador_vacio, solo_digitos
from historial_contratos import carpeta_cliente
from stats_tracker import modo_pruebas_activo

RUTA_CONFIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.yaml")
RUTA_PENDIENTES_N8N = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pendientes_n8n.json")
RUTA_CORRELATIVOS_CONTRATO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "correlativos_contrato.json")

_MAX_PENDIENTES_GUARDADOS = 200
_TIMEOUT_SEGUNDOS = 15
_ZONA_LIMA = ZoneInfo("America/Lima")


def _leer_url_webhook():
    """Lee 'n8n_webhook_url' de config.yaml. Devuelve None (no rompe
    nada) si el archivo no existe, no es YAML válido, o no tiene esa
    clave -- el llamador lo trata igual que "n8n no está configurado"."""
    try:
        with open(RUTA_CONFIG, "r", encoding="utf-8") as f:
            datos = yaml.safe_load(f)
        if isinstance(datos, dict):
            url = datos.get("n8n_webhook_url")
            if url and str(url).strip():
                return str(url).strip()
    except (OSError, yaml.YAMLError):
        pass
    return None


def obtener_codigos_proyecto():
    """Lee 'codigos_proyecto' de config.yaml (ej. {'COSTA DEL VALLE':
    'CDV'}) -> {PROYECTO EN MAYÚSCULAS: CODIGO}. Devuelve {} (no rompe
    nada) si falta el archivo o la clave. Pública: también la usa la
    pantalla de Configuración (ver Inicio.py) para el ajuste manual del
    correlativo."""
    try:
        with open(RUTA_CONFIG, "r", encoding="utf-8") as f:
            datos = yaml.safe_load(f)
        if isinstance(datos, dict) and isinstance(datos.get("codigos_proyecto"), dict):
            return {
                str(proyecto).strip().upper(): str(codigo).strip()
                for proyecto, codigo in datos["codigos_proyecto"].items()
            }
    except (OSError, yaml.YAMLError):
        pass
    return {}


# =========================================================================
# NORMALIZACIÓN DE VALORES -- para que el payload sea siempre JSON válido
# (nunca un datetime/date crudo, nunca un número donde se pide texto).
# =========================================================================
def _texto(valor):
    """Siempre devuelve un string (nunca None, nunca un tipo crudo de
    openpyxl) -- para los 20 campos de texto del payload."""
    if valor is None:
        return ""
    return str(valor).strip()


def _texto_documento(valor, tipo_doc, etiqueta=""):
    """Igual que _texto(), pero si 'tipo_doc' contiene "DNI": si quedó con
    7 dígitos (Excel le quitó el cero inicial, ej. '5208' guardado como
    número en vez de '05208'), lo completa a 8 con un cero a la
    izquierda. Si quedó con menos de 7 o más de 8, se envía tal cual y se
    devuelve una advertencia para mostrar en pantalla (puede ser un dato
    mal tipeado, no un problema de Excel que se pueda corregir solo).
    Devuelve (texto, advertencia_o_None)."""
    texto = _texto(valor)
    if "DNI" in _texto(tipo_doc).upper():
        digitos = solo_digitos(texto)
        cantidad = len(digitos)
        if cantidad == 7:
            return digitos.zfill(8), None
        if 0 < cantidad < 7 or cantidad > 8:
            return texto, f"Revisa el DNI{etiqueta}: tiene {cantidad} dígitos"
    return texto, None


def _numero(valor):
    """Siempre un float (nunca texto, nunca None) -- para PRECIO,
    INICIAL_x_MONTO e INICIAL_TOTAL."""
    if valor is None or valor == "":
        return 0.0
    try:
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


def _fecha_ddmmaaaa(valor):
    """"" si la celda está vacía; si no, DD/MM/AAAA. No reusa
    excel_pipeline.formatear_fecha() a propósito: esa da año de 2 dígitos
    (DD/MM/AA) y acá n8n pidió 4 dígitos."""
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        return ""
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, str) and es_fecha(valor):
        texto = valor.strip()
        for formato in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y"):
            try:
                return datetime.strptime(texto, formato).strftime("%d/%m/%Y")
            except ValueError:
                continue
    return _texto(valor)


_RELACION_POR_OPCION = {
    "UNICO PROPIETARIO": "",
    "CASADOS QUE VIVEN EN EL MISMO SITIO": "CONYUGE",
    "COPROPIETARIOS NO CASADOS QUE VIVEN EN EL MISMO SITIO": "COPROPIETARIO",
    "COPROPIETARIOS NO CASADOS QUE VIVEN EN SITIOS DIFERENTES": "COPROPIETARIO",
}


def armar_payload(excel_bytes, proyecto, lote, nombre_cliente, numero_membresia, opcion_propiedad, id_contrato):
    """Arma el dict de 32 claves que se envía a n8n, leyendo el Excel ya
    recalculado (el mismo que usa historial_contratos.extraer_lote_y_nombre).
    'id_contrato' ya viene resuelto por el llamador (ver
    _obtener_o_asignar_id_contrato) -- esta función solo lo coloca en el
    payload, no decide cómo se arma. Nunca lanza excepción -- si algo
    puntual falla al leer una celda, esa celda queda vacía/0 en vez de
    cortar todo el envío.

    Devuelve (payload, advertencias): 'advertencias' es una lista de
    textos para mostrar en pantalla (ej. un DNI con una cantidad de
    dígitos rara que no se pudo corregir solo, ver _texto_documento)."""
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    sheet_datos = wb["DATOS"] if "DATOS" in wb.sheetnames else None
    sheet_cot = wb["Cotizador"] if "Cotizador" in wb.sheetnames else None

    def _datos(celda):
        return sheet_datos[celda].value if sheet_datos is not None else None

    def _cot(celda):
        return sheet_cot[celda].value if sheet_cot is not None else None

    advertencias = []

    tipo_doc = _texto(_datos("E12"))
    nro_doc, _advertencia_nro_doc = _texto_documento(_datos("E13"), tipo_doc)
    if _advertencia_nro_doc:
        advertencias.append(_advertencia_nro_doc)

    hay_copropietario = not es_marcador_vacio(_datos("E33"))
    copropietario = _texto(_datos("E33")) if hay_copropietario else ""
    copropietario_tipo_doc = _texto(_datos("E34")) if hay_copropietario else ""
    if hay_copropietario:
        copropietario_nro_doc, _advertencia_coprop_doc = _texto_documento(
            _datos("E35"), copropietario_tipo_doc, etiqueta=" del copropietario"
        )
        if _advertencia_coprop_doc:
            advertencias.append(_advertencia_coprop_doc)
    else:
        copropietario_nro_doc = ""
    copropietario_celular = _texto(_datos("E49")) if hay_copropietario else ""

    montos_iniciales = [_numero(_cot(f"I{fila}")) for fila in (25, 26, 27, 28, 29)]

    ruta_carpeta, _ = carpeta_cliente(proyecto, lote, nombre_cliente, crear=False)

    payload = {
        "ID_CONTRATO": id_contrato,
        "FECHA_GENERACION": datetime.now(_ZONA_LIMA).strftime("%Y-%m-%d %H:%M"),
        "PROYECTO": _texto(proyecto),
        "LOTE": _texto(lote),
        "NRO_MEMBRESIA": _texto(numero_membresia),
        "TITULAR": _texto(nombre_cliente),
        "TIPO_DOC": tipo_doc,
        "NRO_DOC": nro_doc,
        "ESTADO_CIVIL": _texto(_datos("E16")),
        "DOMICILIO": _texto(_datos("E22")),
        "DISTRITO_PROV_DEP": _texto(_datos("E23")),
        "CELULAR": _texto(_datos("E27")),
        "CORREO": _texto(_datos("E29")),
        "COPROPIETARIO": copropietario,
        "RELACION": _RELACION_POR_OPCION.get(_texto(opcion_propiedad).upper(), ""),
        "COPROP_TIPO_DOC": copropietario_tipo_doc,
        "COPROP_NRO_DOC": copropietario_nro_doc,
        "COPROP_CELULAR": copropietario_celular,
        "PRECIO": _numero(_cot("I23")),
        "INICIAL_TOTAL": round(sum(montos_iniciales), 2),
        "INICIAL_1_MONTO": montos_iniciales[0],
        "INICIAL_1_FECHA": _fecha_ddmmaaaa(_cot("D25")),
        "INICIAL_2_MONTO": montos_iniciales[1],
        "INICIAL_2_FECHA": _fecha_ddmmaaaa(_cot("D26")),
        "INICIAL_3_MONTO": montos_iniciales[2],
        "INICIAL_3_FECHA": _fecha_ddmmaaaa(_cot("D27")),
        "INICIAL_4_MONTO": montos_iniciales[3],
        "INICIAL_4_FECHA": _fecha_ddmmaaaa(_cot("D28")),
        "INICIAL_5_MONTO": montos_iniciales[4],
        "INICIAL_5_FECHA": _fecha_ddmmaaaa(_cot("D29")),
        "RUTA_CARPETA": ruta_carpeta,
        "ESTADO": "PRUEBA" if modo_pruebas_activo() else "GENERADO",
    }
    return payload, advertencias


# =========================================================================
# ENVÍO + PENDIENTES
# =========================================================================
def _cargar_pendientes():
    if not os.path.exists(RUTA_PENDIENTES_N8N):
        return []
    try:
        with open(RUTA_PENDIENTES_N8N, "r", encoding="utf-8") as f:
            datos = json.load(f)
    except (json.JSONDecodeError, OSError):
        return []
    return datos if isinstance(datos, list) else []


def _guardar_pendientes(pendientes):
    with open(RUTA_PENDIENTES_N8N, "w", encoding="utf-8") as f:
        json.dump(pendientes[-_MAX_PENDIENTES_GUARDADOS:], f, ensure_ascii=False, indent=2)


def _enviar(url, payload):
    """POST a 'url' con 'payload' como JSON. Éxito SOLO si responde
    exactamente 200 (cualquier otro código -- 4xx, 5xx, incluso otro
    2xx -- cuenta como fallo). Nunca lanza excepción: timeouts, el host
    caído, DNS roto, etc. se devuelven como (False, motivo)."""
    import requests

    try:
        respuesta = requests.post(url, json=payload, timeout=_TIMEOUT_SEGUNDOS)
    except requests.RequestException as exc:
        return False, str(exc)
    if respuesta.status_code == 200:
        return True, None
    return False, f"n8n respondió con código {respuesta.status_code}"


def reintentar_pendientes_n8n():
    """Reintenta reenviar TODOS los pendientes guardados. Se usa tanto
    antes de enviar un contrato nuevo como al abrir la app (ver
    Inicio.py). Silencioso si no hay URL configurada o no hay nada
    pendiente."""
    url = _leer_url_webhook()
    if not url:
        return
    pendientes = _cargar_pendientes()
    if not pendientes:
        return
    quedan = []
    for item in pendientes:
        exito, _ = _enviar(url, item)
        if not exito:
            quedan.append(item)
    _guardar_pendientes(quedan)


# =========================================================================
# ID_CONTRATO -- numerador propio de la app (CODIGO-AAAA-NNNN), ver
# n8n_integracion.py en el plan: corre SIEMPRE que se genera un CONTRATO,
# aunque n8n no esté configurado -- no es solo un dato para el payload.
# =========================================================================
def _cargar_correlativos():
    if not os.path.exists(RUTA_CORRELATIVOS_CONTRATO):
        return {"reales": {}, "pruebas": {}}
    try:
        with open(RUTA_CORRELATIVOS_CONTRATO, "r", encoding="utf-8") as f:
            datos = json.load(f)
    except (json.JSONDecodeError, OSError):
        return {"reales": {}, "pruebas": {}}
    if not isinstance(datos, dict):
        return {"reales": {}, "pruebas": {}}
    datos.setdefault("reales", {})
    datos.setdefault("pruebas", {})
    return datos


def _guardar_correlativos(datos):
    with open(RUTA_CORRELATIVOS_CONTRATO, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)


def _siguiente_correlativo(codigo, anio, pruebas):
    """Incrementa en 1 el correlativo de 'codigo-anio' (grupo 'pruebas' o
    'reales' según corresponda), lo guarda, y devuelve el nuevo número."""
    datos = _cargar_correlativos()
    grupo = "pruebas" if pruebas else "reales"
    clave = f"{codigo}-{anio}"
    siguiente = datos[grupo].get(clave, 0) + 1
    datos[grupo][clave] = siguiente
    _guardar_correlativos(datos)
    return siguiente


def obtener_correlativo_actual(codigo, anio):
    """Devuelve el correlativo 'reales' guardado para 'codigo-anio' (0 si
    todavía no se asignó ninguno) -- el próximo ID automático sería este
    número + 1. Pública: la usa la pantalla de Configuración (ver
    Inicio.py) para mostrar el valor actual antes de ajustarlo a mano."""
    return _cargar_correlativos()["reales"].get(f"{codigo}-{anio}", 0)


def fijar_correlativo_real(codigo, anio, valor):
    """Ajusta a mano el correlativo 'reales' de 'codigo-anio' a 'valor'.
    Uso excepcional: cuando se asignó un ID_CONTRATO a mano fuera de la
    app (ej. en la matriz de contratos), para que el siguiente ID
    automático de ese proyecto+año no repita un número ya usado. Nunca
    toca el contador de 'pruebas'."""
    datos = _cargar_correlativos()
    datos["reales"][f"{codigo}-{anio}"] = int(valor)
    _guardar_correlativos(datos)


def _obtener_o_asignar_id_contrato(proyecto, lote, nombre_cliente):
    """Devuelve (id_contrato, error_o_None), formato CODIGO-AAAA-NNNN
    (ej. "CDV-2026-0001"), o "PRUEBA-CODIGO-AAAA-NNNN" con Modo de
    pruebas activo (contador separado, para no gastar números reales).

    - Fuera de Modo de pruebas: si la carpeta de este cliente ya tiene un
      ID_CONTRATO.txt (de una generación anterior), lo reutiliza tal
      cual. Si no, asigna el siguiente correlativo del año actual para
      el código de este proyecto y lo guarda en ID_CONTRATO.txt dentro
      de esa carpeta.
    - En Modo de pruebas: nunca lee ni escribe ID_CONTRATO.txt (para no
      tocar la carpeta de un cliente real) -- genera un ID de prueba
      nuevo cada vez.
    - Si el proyecto no tiene código en config.yaml ('codigos_proyecto'):
      (None, "Falta el código del proyecto {proyecto} en config.yaml"),
      sin gastar ningún correlativo.

    Cada operación de disco va en su propio try/except OSError: si la
    carpeta de archivo (ver historial_contratos.RAIZ_HISTORIAL) no es
    accesible en esta máquina, el ID se asigna igual (con el correlativo
    ya gastado), solo que no queda guardado para la próxima regeneración
    -- mismo criterio de "mejor esfuerzo" que el resto de este módulo."""
    pruebas = modo_pruebas_activo()
    codigos = obtener_codigos_proyecto()
    codigo = codigos.get(_texto(proyecto).strip().upper())
    if not codigo:
        return None, f"Falta el código del proyecto {_texto(proyecto)} en config.yaml"

    anio = datetime.now(_ZONA_LIMA).year
    ruta_id = None

    if not pruebas:
        try:
            carpeta, _ = carpeta_cliente(proyecto, lote, nombre_cliente, crear=True)
            ruta_id = os.path.join(carpeta, "ID_CONTRATO.txt")
            if os.path.isfile(ruta_id):
                with open(ruta_id, "r", encoding="utf-8") as f:
                    existente = f.read().strip()
                if existente:
                    return existente, None
        except OSError:
            pass

    correlativo = _siguiente_correlativo(codigo, anio, pruebas)
    id_contrato = f"{codigo}-{anio}-{correlativo:04d}"
    if pruebas:
        return f"PRUEBA-{id_contrato}", None

    if ruta_id:
        try:
            with open(ruta_id, "w", encoding="utf-8") as f:
                f.write(id_contrato)
        except OSError:
            pass
    return id_contrato, None


def registrar_contrato_en_n8n(excel_bytes, proyecto, lote, nombre_cliente, numero_membresia, opcion_propiedad):
    """Función principal: asigna/reutiliza el ID_CONTRATO (ver
    _obtener_o_asignar_id_contrato, corre siempre, con o sin n8n
    configurado), reintenta pendientes viejos, arma y envía el payload de
    ESTE contrato. Devuelve (texto_corto, advertencias, id_contrato) para
    mostrar en pantalla. Nunca lanza excepción hacia afuera."""
    try:
        id_contrato, error_id = _obtener_o_asignar_id_contrato(proyecto, lote, nombre_cliente)
        if error_id:
            return error_id, [], None

        url = _leer_url_webhook()
        if not url:
            return "n8n no está configurado (falta 'n8n_webhook_url' en config.yaml)", [], id_contrato

        reintentar_pendientes_n8n()

        payload, advertencias = armar_payload(
            excel_bytes, proyecto, lote, nombre_cliente, numero_membresia, opcion_propiedad, id_contrato
        )
        exito, _motivo = _enviar(url, payload)
        if exito:
            return "Registrado en el registro maestro", advertencias, id_contrato

        pendientes = _cargar_pendientes()
        pendientes.append(payload)
        _guardar_pendientes(pendientes)
        return "Pendiente de registrar (n8n apagado)", advertencias, id_contrato
    except Exception as exc:
        return f"No se pudo registrar en n8n: {exc}", [], None
