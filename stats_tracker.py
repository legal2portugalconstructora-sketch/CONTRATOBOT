# -*- coding: utf-8 -*-
"""stats_tracker.py — Registro liviano (un .json en disco, sin base de
datos) de cada CONTRATO generado desde pages/2_🤖CONTRATO_Word.py, para
poder mostrar estadísticas reales (no inventadas) en la página de inicio.

No guarda ningún dato del cliente/contrato -- solo si salió bien o mal, y
cuánto tardó, para poder calcular "contratos generados", "de hoy", "tiempo
promedio" y "errores"."""

import json
import os
import threading
from datetime import date, datetime

RUTA_STATS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "contrato_stats.json")
RUTA_MODO_PRUEBAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "modo_pruebas.json")

_MAX_EVENTOS_GUARDADOS = 500  # de sobra para "hoy"/promedio; evita que el .json crezca sin límite

# Protege a contrato_stats.json/modo_pruebas.json de que dos personas
# usando la app al mismo tiempo (un hilo por persona, mismo proceso de
# streamlit run) pisen/pierdan la escritura de la otra.
_LOCK_STATS = threading.Lock()


def modo_pruebas_activo() -> bool:
    """True si el "Modo de pruebas" está prendido (ver el checkbox en
    Configuración de Inicio.py). Es un archivo en disco, no session_state,
    a propósito: así se mantiene prendido aunque se reinicie el servidor
    (algo que pasa seguido en esta app, cada vez que se toca un .py que no
    es una página) y es igual para todos los que usan la app, no una
    configuración personal por navegador."""
    try:
        with open(RUTA_MODO_PRUEBAS, "r", encoding="utf-8") as f:
            return bool(json.load(f).get("activo", False))
    except (OSError, json.JSONDecodeError):
        return False


def establecer_modo_pruebas(activo: bool):
    with _LOCK_STATS:
        with open(RUTA_MODO_PRUEBAS, "w", encoding="utf-8") as f:
            json.dump({"activo": bool(activo)}, f)


def _cargar():
    if not os.path.exists(RUTA_STATS):
        return {"eventos": []}
    try:
        with open(RUTA_STATS, "r", encoding="utf-8") as f:
            datos = json.load(f)
    except (json.JSONDecodeError, OSError):
        return {"eventos": []}
    if not isinstance(datos, dict) or not isinstance(datos.get("eventos"), list):
        return {"eventos": []}
    return datos


def _guardar(datos):
    with open(RUTA_STATS, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)


def registrar_generacion(exito: bool, segundos: float, categoria: str = "word"):
    """Agrega un evento (llamar justo después de terminar de procesar,
    tanto si salió bien como si no). 'categoria' separa las estadísticas
    de distintas páginas -- hoy "word" (pages/2, CONTRATO) y "excel"
    (pages/1, procesar_todo) -- para que cada una tenga su propio
    apartado en la página de inicio. Nunca lanza excepción hacia afuera --
    si falla al escribir el .json, se ignora en silencio, para que esto
    jamás pueda romper el procesamiento real.

    Si el "Modo de pruebas" está activo, no guarda nada -- así las páginas
    que llaman a esta función no necesitan saber nada del modo de pruebas."""
    if modo_pruebas_activo():
        return
    try:
        with _LOCK_STATS:
            datos = _cargar()
            datos["eventos"].append({
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "categoria": categoria,
                "exito": bool(exito),
                "segundos": round(float(segundos), 2),
            })
            datos["eventos"] = datos["eventos"][-_MAX_EVENTOS_GUARDADOS:]
            _guardar(datos)
    except Exception:
        pass


def obtener_resumen(categoria: str = "word"):
    """Devuelve un dict con las 4 estadísticas de un apartado de la
    página de inicio, calculadas de los eventos reales guardados de esa
    'categoria' (todo en cero si todavía no se procesó nada de ese tipo
    desde que existe este archivo). Los eventos guardados antes de que
    existiera "categoria" se tratan como "word" (compatibilidad)."""
    eventos = [e for e in _cargar()["eventos"] if e.get("categoria", "word") == categoria]
    exitosos = [e for e in eventos if e.get("exito")]
    errores = [e for e in eventos if not e.get("exito")]
    hoy = date.today().isoformat()
    de_hoy = [e for e in exitosos if str(e.get("timestamp", "")).startswith(hoy)]
    tiempos = [e["segundos"] for e in exitosos if e.get("segundos")]
    promedio = sum(tiempos) / len(tiempos) if tiempos else 0.0
    return {
        "total_generados": len(exitosos),
        "generados_hoy": len(de_hoy),
        "tiempo_promedio_seg": round(promedio, 1),
        "errores": len(errores),
        "hay_datos": len(eventos) > 0,
    }
